// XLSX reader: DOM mode (`readXlsx`) and streaming row mode (`xlsxRows`). Feature-agnostic:
// `readXlsx` runs only the feature readers in `opts.features` (the `./lite` entry exports this
// module as is; ./index.ts wraps it with every built-in feature for the package root).

import type { CellValue, DefinedName, ReadOptions } from '../../core/types.js';
import { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, type XmlHandler } from '../../xml/index.js';
import {
  unzipBlob,
  unzipLazy,
  unzipStream,
  type ZipEntryOut,
  ZipLimitError,
} from '../../zip/index.js';
import { decodeChunks, decodePart, relTypeIs } from '../common/r-opc.js';
import type {
  FeatureReader,
  PackageReadCtx,
  SheetReadCtx,
  SheetReader,
  SubtreeHandler,
} from '../features/types.js';
import { CT_PATH, maybeRelevant, type ResolvedSheet, XlsxPackage } from './package.js';
import { PartBuffer } from './part-buffer.js';
import { Preserver } from './preserve.js';
import { type SheetContext, SheetHandler } from './sheet-part.js';
import { SharedStringsHandler, type StringItem } from './strings-part.js';
import { StylesHandler } from './styles-part.js';
import { WorkbookHandler } from './workbook-part.js';
import {
  EncryptedFileError,
  normalizeError,
  parseStrict,
  StrictXml,
  XlsxParseError,
} from './xml-util.js';

export { translateFormula } from './formula.js';
export type { SheetEntry } from './workbook-part.js';
export { EncryptedFileError, SheetLimitError, XlsxParseError } from './xml-util.js';

/** One streamed row. */
export interface XlsxRow {
  /** Sheet name. */
  readonly sheet: string;
  /** 0-based position of the sheet in the workbook's tab order. */
  readonly sheetIndex: number;
  /** 1-based row number. */
  readonly row: number;
  /**
   * Cell values, sparse and 0-BASED by column (values[0] = column A); empty cells are holes.
   * Dates (numeric cells with a date format, or t="d") are `Date`s; shared strings are resolved;
   * formulas are `Formula` objects with their cached result. Hyperlinks are NOT applied (they
   * are stored after the cell data in the sheet XML).
   */
  readonly values: CellValue[];
}

// ---------------------------------------------------------------- shared helpers

function toBytes(data: Uint8Array | ArrayBuffer): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

function filterDatabase(names: readonly DefinedName[], index: number): string | undefined {
  for (const n of names) {
    if (n.name === '_xlnm._FilterDatabase' && n.localSheetId === index) {
      const v = n.value;
      const bang = v.lastIndexOf('!');
      const ref = (bang >= 0 ? v.slice(bang + 1) : v).replace(/\$/g, '');
      return /^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(ref) ? ref : undefined;
    }
  }
  return undefined;
}

function addSheet(wb: Workbook, s: ResolvedSheet): Worksheet {
  const opts = s.state !== 'visible' ? { state: s.state } : {};
  try {
    return wb.addWorksheet(s.name, opts);
  } catch {
    // Name rejected by Excel rules (too long, duplicate, invalid chars) — keep a usable name.
    let base =
      s.name
        .replace(/[\\/?*[\]:]/g, '_')
        .replace(/^'+|'+$/g, '')
        .slice(0, 28) || 'Sheet';
    if (base.toLowerCase() === 'history') base = 'History_';
    let name = base;
    for (let i = 2; wb.getWorksheet(name); i++) name = `${base}(${i})`;
    return wb.addWorksheet(name, opts);
  }
}

const CORE_FIELDS = [
  'title',
  'subject',
  'keywords',
  'description',
  'category',
  'contentStatus',
  'language',
];

/** Company / Manager of docProps/app.xml (the rest of it is regenerated). */
function parseApp(xml: string, wb: Workbook): void {
  let field = '';
  let text = '';
  try {
    parseStrict(
      xml,
      {
        open(name) {
          field = localName(name);
          text = '';
        },
        text(t) {
          text += t;
        },
        close() {
          if (text && (field === 'Company' || field === 'Manager')) {
            wb[field === 'Company' ? 'company' : 'manager'] = text;
          }
          text = '';
        },
      },
      'docProps/app.xml',
    );
  } catch {
    // metadata only; ignore a broken app.xml
  }
}

function parseCore(xml: string, wb: Workbook): void {
  let field = '';
  let text = '';
  parseStrict(
    xml,
    {
      open(name) {
        field = localName(name);
        text = '';
      },
      text(t) {
        text += t;
      },
      close() {
        if (text === '') return;
        if (field === 'creator') wb.creator = text;
        else if (CORE_FIELDS.includes(field))
          (wb as unknown as Record<string, string>)[field] = text;
        else if (field === 'lastModifiedBy') wb.lastModifiedBy = text;
        else if (field === 'revision') {
          if (+text - +text === 0) wb.revision = +text;
        } else if (field === 'created' || field === 'modified') {
          const d = new Date(text);
          if (!Number.isNaN(d.getTime())) wb[field] = d;
        }
        text = '';
      },
    },
    'docProps/core.xml',
  );
}

// ---------------------------------------------------------------- DOM mode

/**
 * Feature sheet readers by element path. Several features may claim the same path (e.g.
 * "extLst/ext"): they are combined into one reader that fans events out to each.
 */
function sheetFeatures(readers: readonly FeatureReader[]): {
  map: Map<string, SheetReader>;
  all: SheetReader[];
} {
  const multi = new Map<string, SheetReader[]>();
  const all: SheetReader[] = [];
  for (const f of readers) {
    const r = f.sheet;
    if (!r) continue;
    all.push(r);
    for (const el of r.elements ?? []) {
      const list = multi.get(el);
      if (list) list.push(r);
      else multi.set(el, [r]);
    }
  }
  const map = new Map<string, SheetReader>();
  for (const [el, list] of multi) map.set(el, list.length === 1 ? list[0]! : fanOut(list));
  return { map, all };
}

function fanOut(list: SheetReader[]): SheetReader {
  return {
    open(ctx, el, a) {
      const subs: SubtreeHandler[] = [];
      for (const r of list) {
        const h = r.open?.(ctx, el, a);
        if (h) subs.push(h);
      }
      if (subs.length === 0) return undefined;
      return {
        open: (n, at, sc) => {
          for (const h of subs) h.open?.(n, at, sc);
        },
        close: (n) => {
          for (const h of subs) h.close?.(n);
        },
        text: (t) => {
          for (const h of subs) h.text?.(t);
        },
        end: () => {
          for (const h of subs) h.end?.();
        },
      };
    },
  };
}

/**
 * Read a whole XLSX file into a {@link Workbook} (synchronous), modelling only the features
 * whose readers are in `opts.features` (none by default). Elements and parts of the other
 * features are kept raw (`preserveUnknown`) and written back unchanged.
 * - `opts.sheets`: only load these sheet names (case-insensitive); others are not created.
 * - `opts.limits`: zip-bomb limits (entry count, sizes, compression ratio).
 * - `opts.maxRows` / `opts.maxCols`: throw SheetLimitError at the first cell beyond them.
 * - `opts.preserveUnknown` (default true): keep what the model does not represent (charts,
 *   shapes, pivot tables, threaded comments, VBA, custom XML, unknown worksheet elements, ...)
 *   so `writeXlsx` / `xlsxStream` write it back (see `listPreservedParts`). Kept parts stay
 *   compressed in the input buffer (inflated when written): do not mutate `data` afterwards.
 *   `false` drops all of it (and does not retain the input buffer).
 * `profile` is ignored.
 * Throws ZipLimitError / ZipPathError for hostile archives, XlsxParseError for broken XML and
 * EncryptedFileError for password-protected files (decrypt with `decryptXlsx` from `./crypto`),
 * SheetLimitError when a sheet exceeds `maxRows` / `maxCols`.
 */
export function readXlsx(data: Uint8Array | ArrayBuffer, opts: ReadOptions = {}): Workbook {
  try {
    return readXlsxImpl(data, opts);
  } catch (e) {
    throw normalizeError(e);
  }
}

function readXlsxImpl(data: Uint8Array | ArrayBuffer, opts: ReadOptions): Workbook {
  const bytes = toBytes(data);
  // OLE compound file signature D0 CF 11 E0: encrypted OOXML container.
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0)
    throw new EncryptedFileError();
  // Lazy: only the parts we need are inflated, one at a time; big parts (sheets, shared
  // strings) are inflated and decoded incrementally, so their full XML never exists in memory.
  const files = unzipLazy(bytes, opts.limits);
  if (files.size === 0)
    throw new XlsxParseError('Not an XLSX file: empty or unreadable zip archive');
  const pkg = new XlsxPackage();
  for (const name of files.keys()) pkg.addName(name);
  const text = (path: string): string | undefined => {
    const e = files.get(pkg.entry(path));
    return e === undefined ? undefined : decodePart(e.bytes());
  };

  const ct = text(CT_PATH);
  if (ct !== undefined) pkg.addMeta(CT_PATH, ct);
  for (const name of files.keys()) {
    if (name.endsWith('.rels') && name !== CT_PATH)
      pkg.addMeta(name, decodePart(files.get(name)!.bytes()));
  }
  pkg.final = true;

  const wbPath = pkg.workbookPath()!;
  const wbh = new WorkbookHandler();
  const wbXml = text(wbPath);
  if (wbXml !== undefined) {
    parseStrict(wbXml, wbh, wbPath);
    pkg.sheetEntries = wbh.sheets;
  }
  const wb = new Workbook({ date1904: wbh.date1904, localDates: opts.localDates });
  if (wbh.fullCalcOnLoad) wb.calcProperties = { fullCalcOnLoad: true };
  if (wbh.views.length) wb.views = wbh.views;

  const app = text('docProps/app.xml');
  if (app !== undefined) parseApp(app, wb);
  const core = text('docProps/core.xml');
  if (core !== undefined) {
    try {
      parseCore(core, wb);
    } catch {
      // metadata only; ignore broken core properties
    }
  }

  let xfMap: Uint32Array | undefined;
  let dxfMap: Uint32Array | undefined;
  /** Styles handler when the file has named cell styles to keep (preserveUnknown). */
  let named: StylesHandler | undefined;
  const stylesPath = pkg.stylesPath();
  const stylesXml = stylesPath ? text(stylesPath) : undefined;
  if (stylesXml !== undefined) {
    const sh = new StylesHandler();
    parseStrict(stylesXml, sh, stylesPath!);
    // named cell styles are kept (raw) with preserveUnknown: so is each xf's link to them
    if (opts.preserveUnknown !== false && (sh.styleXfs > 1 || sh.cellStyles > 1)) named = sh;
    xfMap = sh.apply(wb.styles, named !== undefined);
    dxfMap = sh.applyDxfs(wb.styles);
  }
  const pctx: PackageReadCtx = {
    workbook: wb,
    readText: text,
    readBytes: (path) => files.get(pkg.entry(path))?.bytes(),
    relsOf: (path) => pkg.relsOf(path),
  };
  const readers = opts.features ?? [];
  const { map: features, all: featureReaders } = sheetFeatures(readers);
  let pres: Preserver | undefined;
  if (opts.preserveUnknown !== false) {
    const owned = ['docProps/core.xml', 'docProps/app.xml', wbPath];
    if (stylesPath) owned.push(stylesPath);
    const sp = pkg.sstPath();
    if (sp) owned.push(sp);
    for (const s of pkg.sheets() ?? []) owned.push(s.path);
    for (const r of pkg.relsOf(wbPath)?.values() ?? []) {
      if (!r.external && relTypeIs(r.type, 'calcChain')) owned.push(r.path);
    }
    pres = new Preserver(files, pkg, owned, dxfMap);
  }

  let strings: StringItem[] = [];
  const sstPath = pkg.sstPath();
  const sstEntry = sstPath ? files.get(sstPath) : undefined;
  if (sstEntry !== undefined) {
    const h = new SharedStringsHandler();
    const p = new StrictXml(h, sstPath!);
    decodeChunks(sstEntry.chunks(), (s) => p.write(s));
    p.end();
    strings = h.items;
  }

  const filter = opts.sheets ? new Set(opts.sheets.map((s) => s.toLowerCase())) : undefined;
  /** File sheet index (workbook.xml order) → index in wb.worksheets. */
  const sheetMap = new Map<number, number>();
  for (const s of pkg.sheets() ?? []) {
    if (filter && !filter.has(s.name.toLowerCase())) continue;
    const entry = files.get(s.path);
    if (entry === undefined) continue;
    const ws = addSheet(wb, s);
    sheetMap.set(s.index, wb.worksheets.length - 1);
    const rels = pkg.relsOf(s.path);
    const fctx: SheetReadCtx = {
      ...pctx,
      ws,
      path: s.path,
      rels,
      xf: (v) => (xfMap ? (xfMap[+v] ?? 0) : 0),
      dxf: (v) => (dxfMap ? (dxfMap[+v] ?? 0) : 0),
      scratch: new Map(),
    };
    const cap = pres?.sheet(rels);
    const ctx: SheetContext = {
      strings,
      xfMap,
      styles: wb.styles,
      date1904: wb.date1904,
      rels,
      features,
      featureCtx: fctx,
      capture: cap,
      name: s.name,
      maxRows: opts.maxRows,
      maxCols: opts.maxCols,
    };
    const h = new SheetHandler(ctx, ws);
    const p = new StrictXml(h, s.path);
    decodeChunks(entry.chunks(), (chunk) => p.write(chunk));
    p.end();
    h.finish();
    for (const r of featureReaders) r.finish?.(fctx);
    if (cap) pres!.finishSheet(cap, ws, s.sheetId);
    if (ws.autoFilter === undefined) {
      const af = filterDatabase(wbh.definedNames, s.index);
      if (af) ws.autoFilter = af;
    }
  }
  wb.definedNames = modelNames(wbh.definedNames, sheetMap);
  for (const f of readers) f.book?.(pctx);
  pres?.book(wb, wbPath, wbXml, wbh.sheets, filter, stylesXml, dxfMap, named, sheetMap);
  return wb;
}

/**
 * Defined names for the model: sheet scopes remapped to `wb.worksheets` indices (names scoped to
 * sheets that were not loaded are dropped) and `_xlnm._FilterDatabase` left out (it lives on as
 * `worksheet.autoFilter` and is regenerated by the writer).
 */
function modelNames(names: readonly DefinedName[], sheetMap: Map<number, number>): DefinedName[] {
  const out: DefinedName[] = [];
  for (const n of names) {
    if (n.name === '_xlnm._FilterDatabase') continue;
    if (n.localSheetId === undefined) {
      out.push(n);
      continue;
    }
    const i = sheetMap.get(n.localSheetId);
    if (i !== undefined) out.push(i === n.localSheetId ? n : { ...n, localSheetId: i });
  }
  return out;
}

// ---------------------------------------------------------------- streaming mode

type Source =
  | Uint8Array
  | ArrayBuffer
  | Blob
  | AsyncIterable<Uint8Array>
  | ReadableStream<Uint8Array>;

/** An archive entry that can be read at any time, in any order (random-access sources). */
interface RandomEntry {
  readonly path: string;
  read(unbounded: boolean): AsyncIterable<Uint8Array> | Iterable<Uint8Array>;
}

/**
 * Entries of a random-access archive in the order xlsxRows needs: content types and every
 * .rels first, then the workbook, styles and shared-strings parts (resolved from the rels just
 * read — the consumer processes each entry before asking for the next), then the rest in
 * archive order. Sheets then never wait for metadata, whatever order the file stores them in.
 */
async function* metadataFirst(
  entries: readonly RandomEntry[],
  pkg: XlsxPackage,
): AsyncGenerator<ZipEntryOut, void, undefined> {
  for (const e of entries) pkg.addName(e.path);
  const byName = new Map(entries.map((e) => [e.path.toLowerCase(), e]));
  const done = new Set<RandomEntry>();
  const open = (e: RandomEntry): ZipEntryOut => {
    done.add(e);
    let unbounded = false;
    return {
      path: e.path,
      unbounded: () => {
        unbounded = true;
      },
      chunks: {
        async *[Symbol.asyncIterator]() {
          yield* e.read(unbounded);
        },
      },
    };
  };
  for (const e of entries) if (e.path === CT_PATH || e.path.endsWith('.rels')) yield open(e);
  for (const part of [() => pkg.workbookPath(), () => pkg.stylesPath(), () => pkg.sstPath()]) {
    const p = part();
    const e = p ? byName.get(p.toLowerCase()) : undefined;
    if (e !== undefined && !done.has(e)) yield open(e);
  }
  for (const e of entries) if (!done.has(e)) yield open(e);
}

/**
 * Inflating in-memory bytes is synchronous: hand the event loop a turn about every 50 ms so
 * timers and I/O keep running during a long read.
 */
async function* yielding(it: Iterable<Uint8Array>): AsyncGenerator<Uint8Array, void, undefined> {
  let t = Date.now();
  for (const c of it) {
    yield c;
    if (Date.now() - t > 50) {
      await new Promise((r) => setTimeout(r, 0));
      t = Date.now();
    }
  }
}

/**
 * Archive entries for xlsxRows. Bytes and Blobs are random-access: they are indexed by the
 * central directory and read metadata-first. Streams are read in archive order; so are bytes
 * whose central directory cannot be read (truncated / streamed-only archives).
 */
async function* entriesOf(
  source: Source,
  opts: ReadOptions,
  pkg: XlsxPackage,
): AsyncGenerator<ZipEntryOut, void, undefined> {
  const lim = opts.limits;
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    const m = await unzipBlob(source, lim, { deferSizeCheck: true });
    yield* metadataFirst(
      [...m.values()].map((e) => ({ path: e.path, read: (u) => e.chunks(u) })),
      pkg,
    );
    return;
  }
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
    const bytes = toBytes(source);
    let m: ReturnType<typeof unzipLazy> | undefined;
    try {
      m = unzipLazy(bytes, lim, { deferSizeCheck: true });
    } catch (e) {
      // no readable central directory: fall back to reading local headers in order
      if (!(e instanceof ZipLimitError) || !/^Malformed|Truncated/.test(e.message)) throw e;
    }
    if (m !== undefined) {
      yield* metadataFirst(
        [...m.values()].map((e) => ({
          path: e.path,
          read: (u) => yielding(e.chunks(256 * 1024, u)),
        })),
        pkg,
      );
      return;
    }
  }
  yield* unzipStream(fromSource(source as Exclude<Source, Blob>), lim);
}

async function* fromSource(src: Source): AsyncGenerator<Uint8Array, void, undefined> {
  if (src instanceof Uint8Array || src instanceof ArrayBuffer) {
    yield toBytes(src);
    return;
  }
  if (typeof (src as ReadableStream<Uint8Array>).getReader === 'function') {
    const reader = (src as ReadableStream<Uint8Array>).getReader();
    try {
      for (;;) {
        const r = await reader.read();
        if (r.done) return;
        yield r.value;
      }
    } finally {
      reader.releaseLock();
    }
  }
  yield* src as AsyncIterable<Uint8Array>;
}

async function readAllText(chunks: AsyncIterable<Uint8Array>): Promise<string[]> {
  const out: string[] = [];
  let dec: TextDecoder | undefined;
  for await (const c of chunks) {
    if (dec === undefined) {
      dec =
        c.length >= 2 && c[0] === 0xff && c[1] === 0xfe
          ? new TextDecoder('utf-16le')
          : c.length >= 2 && c[0] === 0xfe && c[1] === 0xff
            ? new TextDecoder('utf-16be')
            : new TextDecoder('utf-8');
    }
    const s = dec.decode(c, { stream: true });
    if (s.length > 0) out.push(s);
  }
  if (dec) {
    const s = dec.decode();
    if (s.length > 0) out.push(s);
  }
  return out;
}

/**
 * Streaming state. Parts are processed in zip order when possible; a part whose role or whose
 * prerequisites are not known yet is buffered (see PartBuffer) and processed as soon as
 * the missing metadata arrives (or at the end of the archive).
 */
class StreamReader {
  readonly pkg: XlsxPackage = new XlsxPackage();
  private wbDone = false;
  private stylesDone = false;
  private sstDone = false;
  private date1904 = false;
  private readonly wb = new Workbook();
  private xfMap: Uint32Array | undefined;
  private strings: StringItem[] = [];
  /** Buffered parts whose role or prerequisites were unknown when they arrived. */
  readonly buffered = new Map<string, PartBuffer>();
  private readonly filter: Set<string> | undefined;
  private readonly compress: boolean;
  private readonly onSheetEnd: ReadOptions['onSheetEnd'];
  private readonly localDates: boolean;
  private readonly maxRows: number | undefined;
  private readonly maxCols: number | undefined;

  constructor(opts: ReadOptions) {
    this.localDates = opts.localDates === true;
    this.maxRows = opts.maxRows;
    this.maxCols = opts.maxCols;
    this.onSheetEnd = opts.onSheetEnd;
    this.filter = opts.sheets ? new Set(opts.sheets.map((s) => s.toLowerCase())) : undefined;
    this.compress = opts.profile === 'low-memory';
  }

  /** Buffer an entry until it can be processed. */
  async buffer(path: string, chunks: AsyncIterable<Uint8Array>): Promise<void> {
    const b = new PartBuffer(this.compress);
    for await (const c of chunks) b.push(c);
    b.end();
    this.buffered.set(path, b);
  }

  /** True when every metadata part needed to decode sheet cells has been processed. */
  ready(): boolean {
    if (!this.wbDone) return false;
    const st = this.pkg.stylesPath();
    const ss = this.pkg.sstPath();
    if (st === undefined || ss === undefined) return false;
    return (st === null || this.stylesDone) && (ss === null || this.sstDone);
  }

  sheetFor(path: string): ResolvedSheet | undefined {
    const lp = path.toLowerCase();
    return this.pkg.sheets()?.find((s) => s.path.toLowerCase() === lp);
  }

  wanted(s: ResolvedSheet): boolean {
    return !this.filter || this.filter.has(s.name.toLowerCase());
  }

  /** Parse a metadata part given as string chunks. */
  meta(role: 'workbook' | 'styles' | 'sst', path: string, chunks: Iterable<string>): void {
    const m = this.metaParser(role, path);
    for (const c of chunks) m.p.write(c);
    m.done();
  }

  metaParser(role: 'workbook' | 'styles' | 'sst', path: string): { p: StrictXml; done(): void } {
    const h =
      role === 'workbook'
        ? new WorkbookHandler()
        : role === 'styles'
          ? new StylesHandler()
          : new SharedStringsHandler();
    const p = new StrictXml(h as XmlHandler, path);
    return {
      p,
      done: () => {
        p.end();
        this.metaDone(role, h);
      },
    };
  }

  private metaDone(
    role: 'workbook' | 'styles' | 'sst',
    h: WorkbookHandler | StylesHandler | SharedStringsHandler,
  ): void {
    if (role === 'workbook') {
      const w = h as WorkbookHandler;
      this.date1904 = w.date1904;
      this.pkg.sheetEntries = w.sheets;
      this.wbDone = true;
    } else if (role === 'styles') {
      this.xfMap = (h as StylesHandler).apply(this.wb.styles);
      this.stylesDone = true;
    } else {
      this.strings = (h as SharedStringsHandler).items;
      this.sstDone = true;
    }
  }

  /** Mark the end of the archive: apply fallbacks; parts that never appeared count as absent. */
  finalize(): void {
    this.pkg.final = true;
    // a workbook part that never appeared (and is not buffered) is absent: synthesize sheets
    const wb = this.pkg.workbookPath();
    if (!this.wbDone && wb !== undefined && !this.buffered.has(wb)) this.wbDone = true;
  }

  /** Process buffered metadata parts whose role is now known. */
  drainBufferedMeta(): void {
    for (const [path, chunks] of [...this.buffered]) {
      const role = this.pkg.roleOf(path);
      if (role === 'workbook' || role === 'styles' || role === 'sst') {
        this.buffered.delete(path);
        this.meta(role, path, chunks.text());
      } else if (role === 'ignore') this.buffered.delete(path);
    }
    if (this.pkg.final) {
      // parts that are referenced but never appeared in the archive count as absent
      const st = this.pkg.stylesPath();
      const ss = this.pkg.sstPath();
      if (st === null || (st !== undefined && !this.buffered.has(st))) this.stylesDone = true;
      if (ss === null || (ss !== undefined && !this.buffered.has(ss))) this.sstDone = true;
    }
  }

  /** Buffered sheets ready to emit, in tab order. */
  readySheets(): { sheet: ResolvedSheet; chunks: PartBuffer }[] {
    const out: { sheet: ResolvedSheet; chunks: PartBuffer }[] = [];
    if (!this.ready()) return out;
    for (const s of this.pkg.sheets() ?? []) {
      const chunks = this.buffered.get(s.path);
      if (chunks === undefined) continue;
      this.buffered.delete(s.path);
      if (this.wanted(s)) out.push({ sheet: s, chunks });
    }
    return out;
  }

  /** A sheet parser that pushes finished rows into `out`. */
  sheetParser(
    s: ResolvedSheet,
    out: XlsxRow[],
  ): { write(t: string): void; end(): void; done(): void } {
    const ctx: SheetContext = {
      strings: this.strings,
      xfMap: this.xfMap,
      styles: this.wb.styles,
      date1904: this.date1904,
      localDates: this.localDates,
      name: s.name,
      maxRows: this.maxRows,
      maxCols: this.maxCols,
    };
    const name = s.name;
    const index = s.index;
    const h = new SheetHandler(ctx, undefined, (r0, values) => {
      out.push({ sheet: name, sheetIndex: index, row: r0 + 1, values });
    });
    const p = new StrictXml(h, s.path);
    const cb = this.onSheetEnd;
    return {
      write: (t: string) => p.write(t),
      end: () => p.end(),
      done: () => cb?.({ sheet: name, sheetIndex: index, merges: h.merges }),
    };
  }
}

/**
 * Stream the rows of every worksheet (or `opts.sheets`) of an XLSX file.
 *
 * Sources: `Uint8Array` / `ArrayBuffer`, `Blob` / `File`, `ReadableStream<Uint8Array>` or any
 * `AsyncIterable<Uint8Array>`.
 *
 * Memory:
 * - Bytes and Blobs are random-access: the central directory is read first, then the content
 *   types, rels, workbook, styles and shared strings, then the sheets, each inflated as it is
 *   parsed. Sheet data uses constant memory whatever order the archive stores its parts in
 *   (Excel 16.x, ExcelJS and `createXlsxWriter` store worksheets before styles.xml). For a
 *   large local file, pass a lazily read Blob: `xlsxRows(await fs.openAsBlob(path))` in Node
 *   (>= 19.8), a `File` from an `<input>` in browsers.
 * - Streams are read in archive order. Sheets stream at constant memory when workbook.xml,
 *   styles.xml and sharedStrings.xml come first (`writeXlsx` / `xlsxStream` output); a sheet
 *   stored before them is buffered (raw UTF-8 bytes, off the JS heap) until the metadata
 *   arrives. With `profile: 'low-memory'` buffered XML is re-deflated (level 1), ~5-10x less
 *   memory for some extra CPU.
 * - The shared-string table is always held in memory.
 *
 * Limits: a sheet parsed as it inflates keeps nothing, so `maxEntryBytes` / `maxTotalBytes` do
 * not apply to it, only `maxCompressionRatio`; metadata and buffered parts get every limit.
 *
 * Order: random-access sources and streams emit sheets in archive order (buffered sheets in tab
 * order as soon as they can be decoded). Rows are emitted in file order (row numbers ascending
 * in well-formed files). Rows without any value are skipped.
 *
 * `maxRows` / `maxCols`: a sheet with a cell beyond them throws SheetLimitError while it is
 * parsed (rows yielded before stay yielded; the rest of the file is not read).
 */
export async function* xlsxRows(
  source: Source,
  opts: ReadOptions = {},
): AsyncGenerator<XlsxRow, void, undefined> {
  try {
    yield* xlsxRowsImpl(source, opts);
  } catch (e) {
    throw normalizeError(e);
  }
}

/** Appended to byte-limit errors of sheets a stream source had to buffer (#89). */
const STREAM_HINT =
  'This sheet is stored before styles.xml / sharedStrings.xml, so a stream source must buffer it. ' +
  'Pass a Blob (`await fs.openAsBlob(path)` in Node, a File in browsers) or the bytes to stream it at constant memory';

async function* xlsxRowsImpl(
  source: Source,
  opts: ReadOptions,
): AsyncGenerator<XlsxRow, void, undefined> {
  const st = new StreamReader(opts);
  const pkg = st.pkg;
  const out: XlsxRow[] = [];
  const seekable =
    source instanceof Uint8Array ||
    source instanceof ArrayBuffer ||
    (typeof Blob !== 'undefined' && source instanceof Blob);

  /** Buffer a part; a stream source hitting a byte limit on a sheet learns how to avoid it. */
  const buffer = async (path: string, chunks: AsyncIterable<Uint8Array>): Promise<void> => {
    try {
      await st.buffer(path, chunks);
    } catch (e) {
      if (
        !seekable &&
        e instanceof ZipLimitError &&
        /too large|total uncompressed/.test(e.message) &&
        /(^|\/)worksheets\//i.test(path)
      ) {
        throw new ZipLimitError(`${e.message}. ${STREAM_HINT}`);
      }
      throw e;
    }
  };

  function* flush(): Generator<{ sheet: ResolvedSheet; chunks: PartBuffer }> {
    st.drainBufferedMeta();
    yield* st.readySheets();
  }

  async function* emitBuffered(): AsyncGenerator<XlsxRow, void, undefined> {
    for (const { sheet, chunks } of flush()) {
      const p = st.sheetParser(sheet, out);
      for (const s of chunks.text()) {
        p.write(s);
        if (out.length > 0) {
          yield* out;
          out.length = 0;
        }
      }
      p.end();
      yield* out;
      out.length = 0;
      p.done();
    }
  }

  for await (const entry of entriesOf(source, opts, pkg)) {
    const path = entry.path;
    pkg.addName(path);
    if (path === CT_PATH || path.endsWith('.rels')) {
      pkg.addMeta(path, (await readAllText(entry.chunks)).join(''));
      yield* emitBuffered();
      continue;
    }
    const role = pkg.roleOf(path);
    if (role === 'workbook' || role === 'styles' || role === 'sst') {
      const m = st.metaParser(role, path);
      const dec = new TextDecoder();
      for await (const c of entry.chunks) {
        const s = dec.decode(c, { stream: true });
        if (s) m.p.write(s);
      }
      const tail = dec.decode();
      if (tail) m.p.write(tail);
      m.done();
      yield* emitBuffered();
      continue;
    }
    if (role === 'sheet') {
      const sheet = st.sheetFor(path);
      if (sheet && !st.wanted(sheet)) continue;
      if (sheet && st.ready()) {
        // parsed as it inflates, nothing kept: only the ratio limit applies (#74)
        entry.unbounded();
        const p = st.sheetParser(sheet, out);
        const dec = new TextDecoder();
        for await (const c of entry.chunks) {
          const s = dec.decode(c, { stream: true });
          if (s) p.write(s);
          if (out.length > 0) {
            yield* out;
            out.length = 0;
          }
        }
        const tail = dec.decode();
        if (tail) p.write(tail);
        p.end();
        yield* out;
        out.length = 0;
        p.done();
        continue;
      }
      await buffer(path, entry.chunks);
      continue;
    }
    if (role === 'unknown' && maybeRelevant(path)) {
      await buffer(path, entry.chunks);
      yield* emitBuffered();
    }
    // role 'ignore' (or irrelevant unknown): leave the chunks unconsumed; unzipStream skips them.
  }
  st.finalize();
  yield* emitBuffered();
}
