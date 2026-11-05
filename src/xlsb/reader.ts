// .xlsb reader: the zip / OPC package of .xlsx with BIFF12 binary parts (MS-XLSB) instead of
// XML. Records are `id` (1-2 bytes) + `size` (1-4 bytes, 7 bits each, high bit = more) +
// payload. Produces the same Workbook model as readXlsx; read-only.

import {
  addSheet,
  alignment,
  applyNames,
  argb,
  BiffParseError,
  ERRORS,
  edge,
  fill,
  quoteSheet,
  type RawName,
  rkValue,
  StyleSink,
  type XfRecord,
} from '../biff/common.js';
import { decodeFormula, type FormulaContext } from '../biff/formula.js';
import { type BinaryReadOptions, CellWriter, key } from '../biff/model.js';
import { encodeRange } from '../core/address.js';
import type { Color, Font, RichTextRun } from '../core/types.js';
import { Workbook } from '../core/workbook.js';
import { CellKind } from '../core/worksheet.js';
import { decodePart, parseRels, type Relationship, relTypeIs } from '../xlsx/common/r-opc.js';
import { type LazyEntry, unzipLazy } from '../zip/index.js';

type StringItem = string | { richText: RichTextRun[] };

/** Field reader over one record payload. */
class Rec {
  readonly dv: DataView;
  p: number;
  constructor(
    readonly d: Uint8Array,
    start: number,
    readonly end: number,
  ) {
    this.dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
    this.p = start;
  }
  u8(): number {
    if (this.p >= this.end) throw new RangeError('record overrun');
    return this.d[this.p++]!;
  }
  u16(): number {
    const v = this.dv.getUint16(this.p, true);
    this.p += 2;
    return v;
  }
  u32(): number {
    const v = this.dv.getUint32(this.p, true);
    this.p += 4;
    return v;
  }
  f64(): number {
    const v = this.dv.getFloat64(this.p, true);
    this.p += 8;
    return v;
  }
  /** XLWideString; XLNullableWideString gives undefined for 0xFFFFFFFF. */
  str(): string | undefined {
    const n = this.u32();
    if (n === 0xffffffff) return undefined;
    if (this.p + 2 * n > this.end) throw new RangeError('string overrun');
    let s = '';
    for (let i = 0; i < n; i += 4096) {
      const m = Math.min(n, i + 4096) - i;
      const codes = new Array<number>(m);
      for (let k = 0; k < m; k++) codes[k] = this.dv.getUint16(this.p + 2 * (i + k), true);
      s += String.fromCharCode(...codes);
    }
    this.p += 2 * n;
    return s;
  }
  bytes(n: number): Uint8Array {
    if (n > this.end - this.p) throw new RangeError('record overrun');
    const b = this.d.subarray(this.p, this.p + n);
    this.p += n;
    return b;
  }
  /** CellParsedFormula / NameParsedFormula: [rgce, rgcb]. */
  formula(): [Uint8Array, Uint8Array] {
    const rgce = this.bytes(this.u32());
    const rgcb = this.bytes(this.u32());
    return [rgce, rgcb];
  }
}

/** Calls `cb(id, rec)` for every record of a BIFF12 part. */
function records(d: Uint8Array, cb: (id: number, r: Rec) => void): void {
  let p = 0;
  const n = d.length;
  while (p < n) {
    let id = d[p++]!;
    if (id & 0x80) id = (id & 0x7f) | ((d[p++]! & 0x7f) << 7);
    let len = 0;
    for (let s = 0; s < 28; s += 7) {
      const b = d[p++]!;
      len |= (b & 0x7f) << s;
      if (!(b & 0x80)) break;
    }
    if (p > n || len > n - p) throw new RangeError('record overrun');
    cb(id, new Rec(d, p, p + len));
    p += len;
  }
}

/** BrtColor (8 bytes) → model color; automatic colors are omitted. */
function color(r: Rec): Color | undefined {
  const flags = r.u8();
  const index = r.u8();
  const tint = r.dv.getInt16(r.p, true) / 32767;
  const [red, green, blue, alpha] = [r.d[r.p + 2]!, r.d[r.p + 3]!, r.d[r.p + 4]!, r.d[r.p + 5]!];
  r.p += 6;
  const type = flags >> 1;
  const c: { argb?: string; theme?: number; indexed?: number; tint?: number } = {};
  if (type === 1) c.indexed = index;
  else if (type === 2) {
    c.argb = argb(alpha, red, green, blue);
  } else if (type === 3) c.theme = index;
  else return undefined;
  if (tint) c.tint = tint;
  return c;
}

const UNDERLINE: Record<number, Font['underline']> = {
  1: 'single',
  2: 'double',
  33: 'singleAccounting',
  34: 'doubleAccounting',
};

function font(r: Rec): Font {
  const f: { -readonly [K in keyof Font]: Font[K] } = {};
  const size = r.u16() / 20;
  const flags = r.u16();
  const weight = r.u16();
  const sss = r.u16();
  const uls = r.u8();
  const family = r.u8();
  const charset = r.u8();
  r.u8();
  const c = color(r);
  const scheme = r.u8();
  const name = r.str();
  if (name) f.name = name;
  f.size = size;
  if (weight >= 700) f.bold = true;
  if (flags & 2) f.italic = true;
  if (flags & 8) f.strike = true;
  if (flags & 0x10) f.outline = true;
  if (flags & 0x20) f.shadow = true;
  if (flags & 0x40) f.condense = true;
  if (flags & 0x80) f.extend = true;
  if (UNDERLINE[uls]) f.underline = UNDERLINE[uls];
  if (sss === 1) f.vertAlign = 'superscript';
  else if (sss === 2) f.vertAlign = 'subscript';
  if (c) f.color = c;
  if (family) f.family = family;
  if (charset) f.charset = charset;
  if (scheme === 1) f.scheme = 'major';
  else if (scheme === 2) f.scheme = 'minor';
  return f;
}

/** Fill colors Excel uses as "not set" (system foreground / background). */
const isDefault = (c: Color | undefined, i: number): boolean =>
  c !== undefined && c.indexed === i && c.argb === undefined && c.theme === undefined;

interface Styles {
  sink: StyleSink;
  xfs: XfRecord[];
  fonts: Font[];
  map: number[];
}

function readStyles(d: Uint8Array | undefined, wb: Workbook): Styles {
  const formats = new Map<number, string>();
  const fonts: Font[] = [];
  const fills: XfRecord['fill'][] = [];
  const borders: XfRecord['border'][] = [];
  const xfs: XfRecord[] = [];
  let cellXfs = false;
  if (d) {
    records(d, (id, r) => {
      switch (id) {
        case 0x2c: // BrtFmt
          formats.set(r.u16(), r.str() ?? '');
          break;
        case 0x2b: // BrtFont
          fonts.push(font(r));
          break;
        case 0x2d: {
          // BrtFill (gradients are read as no fill)
          const fls = r.u32();
          const fg = color(r);
          const bg = color(r);
          fills.push(
            fill(fls, isDefault(fg, 64) ? undefined : fg, isDefault(bg, 65) ? undefined : bg),
          );
          break;
        }
        case 0x2e: {
          // BrtBorder: flags, then top / bottom / left / right / diagonal (style, reserved, color)
          const flags = r.u8();
          const e = [0, 1, 2, 3, 4].map(() => {
            const dg = r.u8();
            r.u8();
            return edge(dg, color(r));
          });
          const b: { -readonly [K in keyof XfRecord['border']]: XfRecord['border'][K] } = {};
          if (e[0]) b.top = e[0];
          if (e[1]) b.bottom = e[1];
          if (e[2]) b.left = e[2];
          if (e[3]) b.right = e[3];
          if (e[4] && flags & 3) {
            b.diagonal = {
              ...e[4],
              ...(flags & 2 ? { up: true } : {}),
              ...(flags & 1 ? { down: true } : {}),
            };
          }
          borders.push(b);
          break;
        }
        case 0x269: // BrtBeginCellXFs
          cellXfs = true;
          break;
        case 0x26a:
          cellXfs = false;
          break;
        case 0x2f: {
          // BrtXF
          if (!cellXfs) break;
          r.u16();
          const numFmtId = r.u16();
          const fontIx = r.u16();
          const fillIx = r.u16();
          const borderIx = r.u16();
          const trot = r.u8();
          const indent = r.u8();
          const g = r.u16();
          xfs.push({
            numFmtId,
            font: fonts[fontIx],
            fill: fills[fillIx],
            border: borders[borderIx] ?? {},
            alignment: alignment(
              g & 7,
              (g >> 3) & 7,
              (g & 0x40) !== 0,
              (g & 0x100) !== 0,
              indent,
              trot,
            ),
            protection: { locked: (g & 0x1000) !== 0, hidden: (g & 0x2000) !== 0 },
          });
          break;
        }
      }
    });
  }
  const sink = new StyleSink(wb.styles, fonts[0] ?? wb.styles.defaultFont, formats);
  return { sink, xfs, fonts, map: xfs.map((x) => sink.intern(x)) };
}

function readStrings(d: Uint8Array | undefined, fonts: readonly Font[]): StringItem[] {
  const out: StringItem[] = [];
  if (!d) return out;
  records(d, (id, r) => {
    if (id !== 0x13) return; // BrtSSTItem
    out.push(richStr(r, fonts));
  });
  return out;
}

/** RichStr: flags, text, then (when rich) the formatting runs. */
function richStr(r: Rec, fonts: readonly Font[]): StringItem {
  const flags = r.u8();
  const text = r.str() ?? '';
  if (!(flags & 1)) return text;
  const n = r.u32();
  if (n === 0 || n > (r.end - r.p) / 4) return text;
  const starts: number[] = [];
  const fs: (Font | undefined)[] = [];
  for (let i = 0; i < n; i++) {
    starts.push(r.u16());
    fs.push(fonts[r.u16()]);
  }
  const runs: RichTextRun[] = [];
  if (starts[0]! > 0) runs.push({ text: text.slice(0, starts[0]) });
  for (let i = 0; i < n; i++) {
    const t = text.slice(starts[i], i + 1 < n ? starts[i + 1] : text.length);
    if (t === '' && i + 1 < n) continue;
    const f = fs[i];
    runs.push(f ? { text: t, font: { ...f } } : { text: t });
  }
  return { richText: runs };
}

interface Book {
  date1904: boolean;
  sheets: { name: string; state: 'visible' | 'hidden' | 'veryHidden'; rId: string | undefined }[];
  /** XTI entries: [supbook, first sheet, last sheet]. */
  xti: [number, number, number][];
  /** Supporting links: true = this workbook. */
  supbooks: boolean[];
  names: {
    name: string;
    sheet: number | undefined;
    flags: number;
    rgce: Uint8Array;
    rgcb: Uint8Array;
  }[];
}

function readBook(d: Uint8Array): Book {
  const book: Book = { date1904: false, sheets: [], xti: [], supbooks: [], names: [] };
  records(d, (id, r) => {
    switch (id) {
      case 0x99: // BrtWbProp
        book.date1904 = (r.u32() & 1) !== 0;
        break;
      case 0x9c: {
        // BrtBundleSh
        const hs = r.u32();
        r.u32();
        const rId = r.str();
        book.sheets.push({
          name: r.str() ?? `Sheet${book.sheets.length + 1}`,
          state: hs === 1 ? 'hidden' : hs === 2 ? 'veryHidden' : 'visible',
          rId,
        });
        break;
      }
      case 0x165: // BrtSupSelf
      case 0x166: // BrtSupSame
        book.supbooks.push(true);
        break;
      case 0x167: // BrtSupBookSrc (external workbook)
      case 0x168: // BrtSupAddin
        book.supbooks.push(false);
        break;
      case 0x16a: {
        // BrtExternSheet
        const n = r.u32();
        for (let i = 0; i < n; i++) book.xti.push([r.u32(), r.u32() | 0, r.u32() | 0]);
        break;
      }
      case 0x27: {
        // BrtName
        const flags = r.u32();
        r.u8();
        const itab = r.u32();
        const name = r.str() ?? '';
        const [rgce, rgcb] = r.formula();
        book.names.push({ name, sheet: itab === 0xffffffff ? undefined : itab, flags, rgce, rgcb });
        break;
      }
    }
  });
  return book;
}

/**
 * Read a whole .xlsb file (Excel binary workbook) into a {@link Workbook}, synchronously: sheets
 * (order, names, hidden state), cell values, shared strings incl. rich text, styles, merges,
 * column widths, row heights, hidden rows / columns, defined names, hyperlinks, the 1904 date
 * system, and formulas (decompiled to text, with their cached results; formulas the decompiler
 * does not support keep only their cached value). Anything else in the file (charts, pivot
 * tables, comments, ...) is not read. Throws BiffParseError for malformed data, ZipLimitError /
 * ZipPathError for hostile archives.
 */
export function readXlsb(data: Uint8Array | ArrayBuffer, opts: BinaryReadOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) {
    throw new BiffParseError(
      "Not a zip package: password-protected .xlsb files are OLE containers (decrypt with decryptXlsx from 'turbosheet/crypto' first); legacy .xls files are read by readXls",
    );
  }
  try {
    return read(bytes, opts);
  } catch (e) {
    if (e instanceof RangeError) throw new BiffParseError(`Corrupt .xlsb data: ${e.message}`);
    throw e;
  }
}

function read(bytes: Uint8Array, opts: BinaryReadOptions): Workbook {
  const files = unzipLazy(bytes, opts.limits);
  const lower = new Map<string, LazyEntry>();
  for (const [k, v] of files) lower.set(k.toLowerCase(), v);
  const get = (path: string): Uint8Array | undefined => lower.get(path.toLowerCase())?.bytes();
  const rels = (part: string): Map<string, Relationship> => {
    const i = part.lastIndexOf('/');
    const x = get(`${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`);
    return x ? parseRels(decodePart(x), part) : new Map();
  };
  let wbPath: string | undefined;
  for (const r of rels('').values()) if (relTypeIs(r.type, 'officeDocument')) wbPath = r.path;
  wbPath ??= 'xl/workbook.bin';
  const wbBytes = get(wbPath);
  if (!wbBytes || !wbPath.toLowerCase().endsWith('.bin')) {
    throw new BiffParseError('Not an .xlsb file: no binary workbook part (use readXlsx for .xlsx)');
  }
  const book = readBook(wbBytes);
  const wbRels = rels(wbPath);
  const part = (type: string): string | undefined => {
    for (const r of wbRels.values()) if (!r.external && relTypeIs(r.type, type)) return r.path;
    return undefined;
  };
  const wb = new Workbook({ date1904: book.date1904, localDates: opts.localDates });
  const partBytes = (type: string): Uint8Array | undefined => {
    const path = part(type);
    return path === undefined ? undefined : get(path);
  };
  const st = readStyles(partBytes('styles'), wb);
  const strings = readStrings(partBytes('sharedStrings'), st.fonts);

  const names = book.names.map((n) => n.name);
  const ctx: FormulaContext = {
    b12: true,
    sheet(ixti) {
      const x = book.xti[ixti];
      if (!x || !book.supbooks[x[0]] || x[1] < 0) return undefined;
      const a = book.sheets[x[1]]?.name;
      const b = book.sheets[x[2]]?.name;
      if (a === undefined || b === undefined) return undefined;
      if (a === b) return quoteSheet(a);
      const qa = quoteSheet(a);
      const qb = quoteSheet(b);
      return qa === a && qb === b ? `${a}:${b}` : `'${`${a}:${b}`.replace(/'/g, "''")}'`;
    },
    name: (i) => names[i - 1],
  };

  const filter = opts.sheets ? new Set(opts.sheets.map((s) => s.toLowerCase())) : undefined;
  const sheetMap = new Map<number, number>();
  book.sheets.forEach((s, i) => {
    if (filter && !filter.has(s.name.toLowerCase())) return;
    const rel = s.rId !== undefined ? wbRels.get(s.rId) : undefined;
    const d = rel && !rel.external ? get(rel.path) : undefined;
    if (!d || !relTypeIs(rel!.type, 'worksheet')) return; // chart sheets, macro sheets
    const ws = addSheet(wb, s.name, s.state);
    sheetMap.set(i, wb.worksheets.length - 1);
    readSheet(d, new CellWriter(ws, wb.styles, ctx), st.map, strings, st.fonts, rels(rel!.path));
  });

  const raw: RawName[] = [];
  for (const n of book.names) {
    // hidden future-function names (`_xlfn.X`) and empty placeholders are not model names
    if (n.rgce.length === 0 || (n.flags & 2 && n.name.startsWith('_xlfn.'))) continue;
    const value = decodeFormula(ctx, n.rgce, n.rgcb, 0, 0);
    if (value === undefined) continue;
    const builtin = (n.flags & 0x20) !== 0 && !n.name.startsWith('_xlnm.');
    raw.push({
      name: builtin ? `_xlnm.${n.name}` : n.name,
      value,
      sheet: n.sheet,
      hidden: (n.flags & 1) !== 0,
    });
  }
  applyNames(wb, raw, sheetMap);
  return wb;
}

function readSheet(
  d: Uint8Array,
  w: CellWriter,
  xfMap: readonly number[],
  strings: readonly StringItem[],
  fonts: readonly Font[],
  rels: Map<string, Relationship>,
): void {
  const ws = w.ws;
  let row = -1;
  let col = -1;
  const xf = (v: number): number => xfMap[v & 0xffffff] ?? 0;
  const links: [number, number, number, number, string | undefined, string, string][] = [];
  records(d, (id, r) => {
    // cells: Cell = column u32 + style u32 (24 bits); BrtShort* omit the column (previous + 1)
    if (id <= 0x12 || id === 0x3e) {
      if (id === 0) {
        // BrtRowHdr
        row = r.u32();
        const s = r.u32();
        const height = r.u16();
        const g = r.u16();
        col = -1;
        const custom = (g & 0x2000) !== 0;
        const hidden = (g & 0x1000) !== 0;
        const outline = (g >> 8) & 7;
        const collapsed = (g & 0x800) !== 0;
        const styled = (g & 0x4000) !== 0 && xf(s) !== 0;
        if (custom || hidden || outline || collapsed || styled) {
          ws.setRowProps(row + 1, {
            height: custom ? height / 20 : undefined,
            hidden,
            outlineLevel: outline,
            collapsed,
            styleIndex: styled ? xf(s) : 0,
          });
        }
        return;
      }
      if (row < 0) return;
      const short = id >= 0x0c && id <= 0x12;
      col = short ? col + 1 : r.u32();
      const s = xf(r.u32());
      const kind = short ? id - 0x0b : id;
      switch (kind) {
        case 1:
          w.value(row, col, CellKind.Empty, 0, undefined, s);
          break;
        case 2:
          w.number(row, col, rkValue(r.u32()), s);
          break;
        case 3:
          w.value(row, col, CellKind.Error, 0, ERRORS[r.u8()] ?? '#N/A', s);
          break;
        case 4:
          w.value(row, col, CellKind.Boolean, r.u8() ? 1 : 0, undefined, s);
          break;
        case 5:
          w.number(row, col, r.f64(), s);
          break;
        case 6:
          w.value(row, col, CellKind.String, 0, r.str() ?? '', s);
          break;
        case 7: {
          const item = strings[r.u32()];
          if (item === undefined) w.value(row, col, CellKind.Empty, 0, undefined, s);
          else if (typeof item === 'string') w.value(row, col, CellKind.String, 0, item, s);
          else w.value(row, col, CellKind.RichText, 0, item, s);
          break;
        }
        case 0x3e: {
          const item = richStr(r, fonts);
          w.value(
            row,
            col,
            typeof item === 'string' ? CellKind.String : CellKind.RichText,
            0,
            item,
            s,
          );
          break;
        }
        default: {
          // BrtFmlaString / Num / Bool / Error: cached value, flags, CellParsedFormula
          let k: number = CellKind.Number;
          let num = 0;
          let obj: unknown;
          if (id === 8) {
            k = CellKind.String;
            obj = r.str() ?? '';
          } else if (id === 9) num = r.f64();
          else if (id === 10) {
            k = CellKind.Boolean;
            num = r.u8() ? 1 : 0;
          } else {
            k = CellKind.Error;
            obj = ERRORS[r.u8()] ?? '#N/A';
          }
          r.u16();
          const [rgce, rgcb] = r.formula();
          let exp: number | undefined;
          if (rgce.length === 5 && rgce[0] === 1) {
            // PtgExp: row in the tokens, column in the extra data
            const dv = new DataView(rgce.buffer, rgce.byteOffset, 5);
            const mc =
              rgcb.length >= 4
                ? new DataView(rgcb.buffer, rgcb.byteOffset, 4).getUint32(0, true)
                : col;
            exp = key(dv.getUint32(1, true), mc);
          }
          w.formula(row, col, k, num, obj, s, rgce, rgcb, exp);
        }
      }
      return;
    }
    switch (id) {
      case 0x1aa: // BrtArrFmla
      case 0x1ab: {
        // BrtShrFmla
        const top = r.u32();
        const bottom = r.u32();
        const left = r.u32();
        const right = r.u32();
        if (id === 0x1aa) r.u8();
        const [rgce, rgcb] = r.formula();
        w.defs.set(key(top, left), {
          rgce,
          rgcb,
          array: id === 0x1aa ? encodeRange({ top, left, bottom, right }) : undefined,
        });
        break;
      }
      case 0x3c: {
        // BrtColInfo
        const first = r.u32();
        const last = Math.min(r.u32(), first + 1023, 16383);
        const width = r.u32() / 256;
        const s = xf(r.u32());
        const g = r.u16();
        const hidden = (g & 1) !== 0;
        const custom = (g & 2) !== 0;
        const outline = (g >> 8) & 7;
        const collapsed = (g & 0x1000) !== 0;
        for (let c = first; c <= last; c++) {
          if (!(custom || hidden || outline || collapsed || s)) break;
          const column = ws.getColumn(c + 1);
          if (custom || hidden) column.width = width;
          if (hidden) column.hidden = true;
          if (outline) column.outlineLevel = outline;
          if (collapsed) column.collapsed = true;
          if (s) column.styleIndex = s;
        }
        break;
      }
      case 0xb0: {
        // BrtMergeCell
        const top = r.u32();
        const bottom = r.u32();
        const left = r.u32();
        const right = r.u32();
        try {
          ws.mergeCells(top + 1, left + 1, bottom + 1, right + 1);
        } catch {
          // single-cell / overlapping merges are dropped (as by readXlsx)
        }
        break;
      }
      case 0x1ee: {
        // BrtHLink
        const top = r.u32();
        const bottom = r.u32();
        const left = r.u32();
        const right = r.u32();
        const rId = r.str();
        const location = r.str() ?? '';
        const tooltip = r.str() ?? '';
        links.push([top, bottom, left, right, rId, location, tooltip]);
        break;
      }
    }
  });
  w.finish();
  for (const [top, bottom, left, right, rId, location, tooltip] of links) {
    let target = rId ? rels.get(rId)?.target : undefined;
    if (location) target = target === undefined ? `#${location}` : `${target}#${location}`;
    if (target === undefined) continue;
    for (let rr = top; rr <= Math.min(bottom, top + 1023); rr++) {
      for (let cc = left; cc <= Math.min(right, left + 1023); cc++) {
        if (top === bottom && left === right) w.link(rr, cc, target, tooltip || undefined);
        else if (ws.rowData[rr]?.find(cc) ?? -1 >= 0) w.link(rr, cc, target, tooltip || undefined);
      }
    }
  }
}
