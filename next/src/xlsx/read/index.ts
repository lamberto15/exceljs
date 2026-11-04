// XLSX reader: DOM mode (`readXlsx`) and streaming row mode (`xlsxRows`).

import type { CellValue, DefinedName, ReadOptions } from '../../core/types.js';
import { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, type XmlHandler } from '../../xml/index.js';
import { unzipLazy, unzipStream } from '../../zip/index.js';
import { decodeChunks, decodePart } from '../common/r-opc.js';
import { CT_PATH, maybeRelevant, type ResolvedSheet, XlsxPackage } from './package.js';
import { PartBuffer } from './part-buffer.js';
import { type SheetContext, SheetHandler } from './sheet-part.js';
import { SharedStringsHandler, type StringItem } from './strings-part.js';
import { StylesHandler } from './styles-part.js';
import { WorkbookHandler } from './workbook-part.js';
import { parseStrict, StrictXml, XlsxParseError } from './xml-util.js';

export { translateFormula } from './formula.js';
export type { SheetEntry } from './workbook-part.js';
export { XlsxParseError } from './xml-util.js';

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
        else if (field === 'lastModifiedBy') wb.lastModifiedBy = text;
        else if (field === 'created' || field === 'modified') {
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
 * Read a whole XLSX file into a {@link Workbook} (synchronous).
 * - `opts.sheets`: only load these sheet names (case-insensitive); others are not created.
 * - `opts.limits`: zip-bomb limits (entry count, sizes, compression ratio).
 * `profile` / `preserveUnknown` are currently ignored (unknown parts are not preserved).
 * Throws ZipLimitError / ZipPathError for hostile archives and XlsxParseError for broken XML.
 */
export function readXlsx(data: Uint8Array | ArrayBuffer, opts: ReadOptions = {}): Workbook {
  // Lazy: only the parts we need are inflated, one at a time; big parts (sheets, shared
  // strings) are inflated and decoded incrementally, so their full XML never exists in memory.
  const files = unzipLazy(toBytes(data), opts.limits);
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
  const wb = new Workbook({ date1904: wbh.date1904 });

  const core = text('docProps/core.xml');
  if (core !== undefined) {
    try {
      parseCore(core, wb);
    } catch {
      // metadata only; ignore broken core properties
    }
  }

  let xfMap: Uint32Array | undefined;
  const stylesPath = pkg.stylesPath();
  const stylesXml = stylesPath ? text(stylesPath) : undefined;
  if (stylesXml !== undefined) {
    const sh = new StylesHandler();
    parseStrict(stylesXml, sh, stylesPath!);
    xfMap = sh.apply(wb.styles);
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
    const ctx: SheetContext = {
      strings,
      xfMap,
      styles: wb.styles,
      date1904: wb.date1904,
      rels: pkg.relsOf(s.path),
    };
    const h = new SheetHandler(ctx, ws);
    const p = new StrictXml(h, s.path);
    decodeChunks(entry.chunks(), (chunk) => p.write(chunk));
    p.end();
    h.finish();
    if (ws.autoFilter === undefined) {
      const af = filterDatabase(wbh.definedNames, s.index);
      if (af) ws.autoFilter = af;
    }
  }
  wb.definedNames = modelNames(wbh.definedNames, sheetMap);
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

type Source = Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>;

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

  constructor(opts: ReadOptions) {
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
  sheetParser(s: ResolvedSheet, out: XlsxRow[]): StrictXml {
    const ctx: SheetContext = {
      strings: this.strings,
      xfMap: this.xfMap,
      styles: this.wb.styles,
      date1904: this.date1904,
    };
    const name = s.name;
    const index = s.index;
    const h = new SheetHandler(ctx, undefined, (r0, values) => {
      out.push({ sheet: name, sheetIndex: index, row: r0 + 1, values });
    });
    return new StrictXml(h, s.path);
  }
}

/**
 * Stream the rows of every worksheet (or `opts.sheets`) of an XLSX file.
 *
 * Memory: constant for sheet data when the archive stores workbook.xml, styles.xml and
 * sharedStrings.xml (and their relationship parts) BEFORE the worksheets — the order Excel and
 * exceljs use. When a worksheet precedes any of them (Google Sheets, LibreOffice, some
 * generators), that worksheet's decompressed XML is buffered in memory (as raw UTF-8 bytes, off
 * the JS heap) until the metadata is available — exceljs spools such sheets to temp files
 * instead. With `profile: 'low-memory'` the buffered XML is re-deflated (level 1), cutting that
 * memory ~5-10x for some extra CPU. The shared-string table is always held in memory.
 *
 * Sheets are emitted in zip order when streamed directly; buffered sheets are emitted in tab
 * order as soon as they can be decoded. Rows are emitted in file order (row numbers ascending in
 * well-formed files). Rows without any value are skipped.
 */
export async function* xlsxRows(
  source: Source,
  opts: ReadOptions = {},
): AsyncGenerator<XlsxRow, void, undefined> {
  const st = new StreamReader(opts);
  const pkg = st.pkg;
  const out: XlsxRow[] = [];

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
    }
  }

  for await (const entry of unzipStream(fromSource(source), opts.limits)) {
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
        continue;
      }
      await st.buffer(path, entry.chunks);
      continue;
    }
    if (role === 'unknown' && maybeRelevant(path)) {
      await st.buffer(path, entry.chunks);
      yield* emitBuffered();
    }
    // role 'ignore' (or irrelevant unknown): leave the chunks unconsumed; unzipStream skips them.
  }
  st.finalize();
  yield* emitBuffered();
}
