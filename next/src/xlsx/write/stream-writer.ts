// Constant-memory streaming XLSX writer.
//
// Design
// - The writer owns a private core `Workbook`: its StyleRegistry interns every style, and each
//   `addWorksheet()` registers a real (cell-less) `Worksheet` that carries metadata — columns,
//   views/freeze, merges, autoFilter, tab color, state — so all sheet-level XML comes from the
//   same code as the in-memory path.
// - Rows are buffered in a *scratch* Worksheet (same Workbook, so values are classified, dates
//   converted and row/column styles inherited exactly like `ws.addRow`). Every `flushRows` rows
//   the scratch rows are encoded with the shared `encodeRow` (at their absolute row index), the
//   XML goes into the zip entry, and the scratch sheet is replaced by an empty one. Memory per
//   sheet is O(flushRows), independent of the total row count.
// - Sheets are written one at a time, in order: adding a new worksheet commits the open one.
// - Workbook-level parts (workbook.xml, styles.xml, sharedStrings.xml, rels, content types,
//   docProps) are written by `end()`, after all sheets (zip entry order is irrelevant to Excel).
// - strings: 'inline' is constant memory; 'shared' keeps the growing string table in RAM until
//   end() (smaller files, memory ∝ distinct strings). Hyperlink targets are kept per sheet until
//   commit (memory ∝ hyperlink count).
// - Output goes to `sink` (awaited sequentially when it returns a promise) or to `readable`.
//   addRow() is synchronous; call `await sheet.flush()` periodically (or use `addRows`) to apply
//   backpressure (with native CompressionStream deflate, compression is asynchronous: rows
//   added in a synchronous loop queue their XML until the next await, so flush() also bounds
//   that queue). In readable mode commit()/end() never wait for the consumer (so writing
//   everything before reading cannot deadlock); flush()/addRows() do.

import { SharedStrings } from '../../core/sharedstrings.js';
import type { CellValue, Style, WriteOptions } from '../../core/types.js';
import { Workbook, type WorksheetOptions } from '../../core/workbook.js';
import {
  type ColumnDef,
  type RowProps,
  Worksheet,
  type WorksheetView,
} from '../../core/worksheet.js';
import { nativeCompressor } from '../../zip/index.js';
import { activeTabOf, bookParts, CHUNK, sheetEntry } from './book.js';
import { type ResolvedWriteOptions, resolveWriteOptions } from './options.js';
import { PackageParts } from './package.js';
import type { BookInfo } from './parts.js';
import {
  afterSheetParts,
  encodeRow,
  newSheetCtx,
  type SheetCtx,
  type SheetPlan,
  sheetHead,
  sheetPlan,
  sheetTail,
} from './sheet.js';
import { ZipOut } from './zipout.js';

export type ChunkSink = (chunk: Uint8Array) => void | Promise<void>;

export interface XlsxWriterOptions extends WriteOptions {
  /** Receives zip output. Without it, read `writer.readable`. */
  sink?: ChunkSink;
  /** Rows buffered per sheet before they are encoded and flushed (default 1000). */
  flushRows?: number;
  /** Bytes queued in `readable` before flush() waits (default 1 MiB). */
  highWaterMark?: number;
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
}

export interface StreamSheetOptions extends WorksheetOptions {
  columns?: readonly ColumnDef[] | undefined;
  views?: WorksheetView[] | undefined;
  /** Freeze rows (and columns): `1` or `{ rows: 1, cols: 2 }`. */
  freeze?: number | { rows: number; cols?: number } | undefined;
  autoFilter?: string | undefined;
}

type RowValues = readonly CellValue[] | Readonly<Record<string, CellValue>>;

export class StreamSheet {
  /** Metadata holder (columns, views, merges, autoFilter, state). Never contains cells. */
  readonly worksheet: Worksheet;
  private readonly _w: XlsxWriter;
  /** Zero-based sheet position (selects xl/worksheets/sheetN.xml). */
  readonly index: number;
  private readonly _ctx: SheetCtx;
  private _scratch: Worksheet | undefined = undefined;
  /** Absolute 0-based row index of scratch row 0. */
  private _base = 0;
  private _started = false;
  private _committed = false;
  private _buf = '';

  /** @internal Use `writer.addWorksheet()`. */
  constructor(w: XlsxWriter, ws: Worksheet, index: number, ctx: SheetCtx) {
    this._w = w;
    this.worksheet = ws;
    this.index = index;
    this._ctx = ctx;
  }

  get name(): string {
    return this.worksheet.name;
  }

  get committed(): boolean {
    return this._committed;
  }

  /** 1-based number of the last row added (0 when none). */
  get rowCount(): number {
    return this._base + (this._scratch ? this._scratch.rowData.length : 0);
  }

  /**
   * Append a row (array → columns A.., object → keyed via `columns[].key`). Returns its 1-based
   * row number. Rows are encoded in batches; they cannot be modified once added.
   */
  addRow(values: RowValues, style?: Style | number, props?: RowProps): number {
    this._w._check(this);
    let sc = this._scratch;
    if (sc === undefined) sc = this._newScratch();
    const n = sc.addRow(values, style);
    if (props) sc.setRowProps(n, props);
    const abs = this._base + n;
    if (n >= this._w._flushRows) this._encode();
    return abs;
  }

  /** Append rows from a (sync or async) iterable, applying backpressure between batches. */
  async addRows(
    rows: Iterable<RowValues> | AsyncIterable<RowValues>,
    style?: Style,
  ): Promise<void> {
    const limit = this._w._flushRows;
    let k = 0;
    for await (const r of rows) {
      this.addRow(r, style);
      if (++k >= limit) {
        k = 0;
        await this._w._drain();
      }
    }
  }

  /** Merge a range ("A1:C2" or 1-based top, left, bottom, right); written at commit. */
  mergeCells(range: string): void;
  mergeCells(top: number, left: number, bottom: number, right: number): void;
  mergeCells(a: string | number, l?: number, b?: number, r?: number): void {
    if (typeof a === 'string') this.worksheet.mergeCells(a);
    else this.worksheet.mergeCells(a, l!, b!, r!);
  }

  /** Encode buffered rows and wait until the output consumer has caught up. */
  async flush(): Promise<void> {
    if (!this._committed) {
      this._w._check(this);
      this._encode();
      this._emit(false);
    }
    await this._w._drain();
  }

  /** Finish the sheet (writes the sheet tail and rels). Further rows are rejected. */
  async commit(): Promise<void> {
    this._commitSync();
    await this._w._settle();
  }

  /** @internal */
  _commitSync(): void {
    if (this._committed) return;
    this._encode();
    this._buf += sheetTail(this._ctx);
    this._emit(true);
    this._committed = true;
    // feature parts (comments, tables, ...) and the sheet rels go right after the sheet
    for (const p of afterSheetParts(this._ctx, this._w._pkg)) this._w._zip.add(p.path, p.data);
    this._ctx.links.length = 0;
    this._w._sheetDone(this, this._ctx.formulas);
  }

  private _newScratch(): Worksheet {
    const ws = this.worksheet;
    const sc = new Worksheet(ws.workbook, ws.id, ws.name);
    const cols = ws.columnData;
    if (cols.length > 0) {
      const styles = ws.workbook.styles;
      const defs: ColumnDef[] = [];
      for (let i = 0; i < cols.length; i++) {
        const c = cols[i];
        defs.push(
          c ? { key: c.key, style: c.styleIndex ? styles.get(c.styleIndex) : undefined } : {},
        );
      }
      sc.columns = defs;
    }
    // rows placed on the metadata sheet before the first flush (e.g. the header row written by
    // `columns` with headers) are emitted first; buffered rows follow them
    if (!this._started) this._base = ws.rowData.length;
    this._scratch = sc;
    return sc;
  }

  private _start(): void {
    this._started = true;
    const ws = this.worksheet;
    this._ctx.tabSelected = this._w._claimActive(this);
    this._buf = sheetHead(this._ctx);
    const pre = ws.rowData;
    for (let r = 0; r < pre.length; r++) {
      const row = pre[r];
      if (row !== undefined) this._buf += encodeRow(row, r, this._ctx);
    }
    if (this._scratch === undefined) this._base = pre.length;
  }

  private _encode(): void {
    if (!this._started) this._start();
    const sc = this._scratch;
    if (sc === undefined) return;
    const rows = sc.rowData;
    const base = this._base;
    const ctx = this._ctx;
    let buf = this._buf;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (row === undefined) continue;
      buf += encodeRow(row, base + r, ctx);
      if (buf.length >= CHUNK) {
        this._buf = buf;
        this._emit(false);
        buf = '';
      }
    }
    this._buf = buf;
    this._base = base + rows.length;
    this._scratch = undefined;
  }

  private _emit(last: boolean): void {
    if (!last && this._buf.length === 0) return;
    this._w._zip.write(this._ctx.path, this._buf, last);
    this._buf = '';
  }
}

export class XlsxWriter {
  /** Private workbook: styles + sheet metadata. */
  readonly workbook: Workbook;
  /** Zip output when no `sink` was given. */
  readonly readable: ReadableStream<Uint8Array> | undefined;
  /** @internal */
  readonly _zip: ZipOut;
  /** @internal */
  readonly _flushRows: number;
  /** @internal Feature parts / content types / part counters. */
  readonly _pkg: PackageParts = new PackageParts();
  private readonly _o: ResolvedWriteOptions;
  private readonly _plan: SheetPlan;
  private readonly _sst: SharedStrings | undefined;
  private readonly _sheets: StreamSheet[] = [];
  private readonly _sink: ChunkSink | undefined;
  private _open: StreamSheet | undefined = undefined;
  private _formulas = false;
  private _ended = false;
  private _error: unknown = undefined;
  // sink mode: sequential delivery queue
  private _queue: Uint8Array[] = [];
  private _busy: Promise<void> | undefined = undefined;
  // readable mode
  private _ctrl: ReadableStreamDefaultController<Uint8Array> | undefined = undefined;
  private _waiter: (() => void) | undefined = undefined;
  private _closed = false;
  private _active = -1;

  constructor(opts: XlsxWriterOptions = {}) {
    this._o = resolveWriteOptions(opts);
    this._plan = sheetPlan(this._o.writers);
    this.workbook = new Workbook({ date1904: opts.date1904 === true });
    const wb = this.workbook;
    if (opts.creator !== undefined) wb.creator = opts.creator;
    if (opts.lastModifiedBy !== undefined) wb.lastModifiedBy = opts.lastModifiedBy;
    if (opts.created) wb.created = opts.created;
    if (opts.modified) wb.modified = opts.modified;
    this._flushRows = Math.max(1, opts.flushRows ?? 1000);
    this._sst = this._o.strings === 'shared' ? new SharedStrings() : undefined;
    this._sink = opts.sink;
    if (opts.sink) {
      this.readable = undefined;
    } else {
      const hwm = opts.highWaterMark ?? 1 << 20;
      this.readable = new ReadableStream<Uint8Array>(
        {
          start: (c) => {
            this._ctrl = c;
          },
          pull: () => this._wake(),
          cancel: (reason) => {
            this._closed = true;
            this._error ??= reason ?? new Error('stream cancelled');
            this._wake();
          },
        },
        { highWaterMark: hwm, size: (c) => c.byteLength },
      );
    }
    // native CompressionStream when available: deflate runs off the main thread
    this._zip = new ZipOut(
      (chunk, final) => this._out(chunk, final),
      this._o.compression,
      nativeCompressor(),
    );
  }

  /** Add a worksheet. Commits the currently open sheet (sheets are written sequentially). */
  addWorksheet(name?: string, opts: StreamSheetOptions = {}): StreamSheet {
    this._throwIfFailed();
    if (this._ended) throw new Error('XlsxWriter already ended');
    if (this._open) this._open._commitSync();
    const ws = this.workbook.addWorksheet(name, opts);
    if (opts.columns) ws.columns = opts.columns;
    if (opts.views) ws.views = opts.views;
    if (opts.freeze !== undefined) {
      if (typeof opts.freeze === 'number') ws.freeze(opts.freeze);
      else ws.freeze(opts.freeze.rows, opts.freeze.cols ?? 0);
    }
    if (opts.autoFilter) ws.autoFilter = opts.autoFilter;
    const index = this._sheets.length;
    const ctx = newSheetCtx({
      plan: this._plan,
      sst: this._sst,
      pkg: this._pkg,
      ws,
      index,
      streaming: true,
    });
    const s = new StreamSheet(this, ws, index, ctx);
    this._sheets.push(s);
    this._open = s;
    return s;
  }

  get worksheets(): readonly StreamSheet[] {
    return this._sheets;
  }

  /** Commit the open sheet, write workbook-level parts and finish the zip. */
  async end(): Promise<void> {
    this._throwIfFailed();
    if (this._ended) return;
    if (this._open) this._open._commitSync();
    this._ended = true;
    const wb = this.workbook;
    const sheets = wb.worksheets;
    if (sheets.length === 0) throw new Error('XlsxWriter: no worksheets were added');
    const shared = this._sst !== undefined;
    const book: BookInfo = {
      sheets: sheets.map(sheetEntry),
      date1904: wb.date1904,
      sharedStrings: shared,
      formulas: this._formulas,
      activeTab: this._activeIndex(),
      creator: wb.creator,
      lastModifiedBy: wb.lastModifiedBy,
      created: wb.created,
      modified: wb.modified,
      definedNames: wb.definedNames,
    };
    const z = this._zip;
    for (const p of bookParts(this._o.writers, wb, book, this._pkg, this._sst))
      z.add(p.path, p.data);
    z.end();
    await this._settle();
  }

  // ---- internals ----

  private _activeIndex(): number {
    return this._active >= 0 ? this._active : activeTabOf(this.workbook.worksheets);
  }

  /**
   * @internal Called when a sheet starts: the first visible sheet becomes the active tab (later
   * sheets are not known yet, so `views[].tabSelected` cannot move the selection forward).
   */
  _claimActive(s: StreamSheet): boolean {
    if (this._active < 0 && s.worksheet.state === 'visible') this._active = s.index;
    return this._active === s.index;
  }

  /** @internal */
  _check(s: StreamSheet): void {
    this._throwIfFailed();
    if (s.committed) throw new Error(`Worksheet "${s.name}" is already committed`);
    if (this._open !== s) throw new Error(`Worksheet "${s.name}" is not the open sheet`);
  }

  /** @internal */
  _sheetDone(s: StreamSheet, formulas: boolean): void {
    if (formulas) this._formulas = true;
    if (this._open === s) this._open = undefined;
  }

  private _throwIfFailed(): void {
    if (this._error !== undefined) throw this._error;
  }

  private _out(chunk: Uint8Array, final: boolean): void {
    if (this._sink) {
      if (chunk.length > 0) this._deliver(chunk);
      return;
    }
    const c = this._ctrl!;
    if (!this._closed) {
      if (chunk.length > 0) c.enqueue(chunk);
      if (final) {
        this._closed = true;
        c.close();
      }
    }
  }

  private _deliver(chunk: Uint8Array): void {
    if (this._busy) {
      this._queue.push(chunk);
      return;
    }
    let r: void | Promise<void>;
    try {
      r = this._sink!(chunk);
    } catch (e) {
      this._error ??= e;
      throw e;
    }
    if (r && typeof (r as Promise<void>).then === 'function') {
      this._busy = (r as Promise<void>).then(
        () => this._next(),
        (e) => this._fail(e),
      );
    }
  }

  private _fail(e: unknown): void {
    this._error ??= e;
    this._queue = [];
    this._busy = undefined;
  }

  private _next(): Promise<void> | undefined {
    const q = this._queue;
    let i = 0;
    try {
      while (i < q.length) {
        const r = this._sink!(q[i++]!);
        if (r && typeof (r as Promise<void>).then === 'function') {
          q.splice(0, i);
          return (r as Promise<void>).then(
            () => this._next(),
            (e) => this._fail(e),
          );
        }
      }
    } catch (e) {
      this._fail(e);
      return undefined;
    }
    q.length = 0;
    this._busy = undefined;
    return undefined;
  }

  private _wake(): void {
    const w = this._waiter;
    this._waiter = undefined;
    w?.();
  }

  /** @internal Wait for the consumer (sink queue drained / readable below its high-water mark). */
  async _drain(): Promise<void> {
    const z = this._zip.ready();
    if (z) await z;
    if (this._sink) {
      while (this._busy) await this._busy;
    } else {
      const c = this._ctrl!;
      while (!this._closed && (c.desiredSize ?? 1) <= 0) {
        await new Promise<void>((res) => {
          this._waiter = res;
        });
      }
    }
    this._throwIfFailed();
  }

  /** @internal commit()/end(): sink mode waits for delivery; readable mode never blocks. */
  async _settle(): Promise<void> {
    await this._zip.idle();
    if (this._sink) await this._drain();
    else this._throwIfFailed();
  }
}

/**
 * Create a streaming XLSX writer running only the feature writers in `opts.features` (none by
 * default; the package root's createXlsxWriter passes all built-ins). Its workbook is new and private: content preserved by
 * `readXlsx` (preserveUnknown) is never carried over — use writeXlsx / xlsxStream for that.
 * ```ts
 * const w = createXlsxWriter({ sink: (chunk) => file.write(chunk) });
 * const s = w.addWorksheet('Data', { columns: [{ header: 'Id', key: 'id', width: 10 }] });
 * for (const r of rows) s.addRow(r);
 * await s.commit();
 * await w.end();
 * ```
 */
export function createXlsxWriter(opts?: XlsxWriterOptions): XlsxWriter {
  return new XlsxWriter(opts);
}
