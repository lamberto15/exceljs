// .xls reader: the BIFF8 "Workbook" stream (MS-XLS) of an OLE compound file (Excel 97-2003).
// Records are `type` u16 + `size` u16 + payload; payloads over 8224 bytes continue in CONTINUE
// records. The workbook globals (fonts, formats, XFs, shared strings, sheets, names) come first,
// then one substream per sheet at the offset its BOUNDSHEET record gives. Read-only; produces
// the same Workbook model as readXlsx.

import {
  addSheet,
  alignment,
  applyNames,
  argb,
  BiffParseError,
  BUILTIN_NAMES,
  ERRORS,
  edge,
  fill,
  latin1,
  quoteSheet,
  type RawName,
  rkValue,
  StyleSink,
  utf16,
  type XfRecord,
} from '../biff/common.js';
import { decodeFormula, type FormulaContext } from '../biff/formula.js';
import { type BinaryReadOptions, CellWriter, key } from '../biff/model.js';
import { encodeRange } from '../core/address.js';
import type { Color, Font, RichTextRun } from '../core/types.js';
import { Workbook } from '../core/workbook.js';
import { CellKind } from '../core/worksheet.js';
import { isCfb, readCfb } from '../crypto/cfb.js';

/** The file is password-protected (FILEPASS): .xls encryption (XOR / RC4) is not supported. */
export class EncryptedXlsError extends BiffParseError {
  constructor() {
    super('The .xls file is password-protected (encrypted .xls files are not supported)');
    this.name = 'EncryptedXlsError';
  }
}

type StringItem = string | { richText: RichTextRun[] };

/** Reads across a record and its CONTINUE records (strings restart their flags at a boundary). */
class Segs {
  private s = 0;
  p = 0;
  private cur: Uint8Array;
  constructor(private readonly segs: readonly Uint8Array[]) {
    this.cur = segs[0]!;
  }
  private next(): void {
    const n = this.segs[++this.s];
    if (n === undefined) throw new RangeError('record overrun');
    this.cur = n;
    this.p = 0;
  }
  u8(): number {
    if (this.p >= this.cur.length) this.next();
    return this.cur[this.p++]!;
  }
  u16(): number {
    return this.u8() | (this.u8() << 8);
  }
  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }
  skip(n: number): void {
    while (n > 0) {
      if (this.p >= this.cur.length) this.next();
      const take = Math.min(n, this.cur.length - this.p);
      this.p += take;
      n -= take;
    }
  }
  /** `n` characters; a CONTINUE boundary inside them starts with a new high-byte flag. */
  chars(n: number, wide: boolean): string {
    let out = '';
    while (n > 0) {
      if (this.p >= this.cur.length) {
        this.next();
        wide = (this.cur[this.p++]! & 1) !== 0;
      }
      const avail = wide ? (this.cur.length - this.p) >> 1 : this.cur.length - this.p;
      const take = Math.min(n, avail);
      if (take <= 0) throw new RangeError('string overrun');
      out += wide ? utf16(this.cur, this.p, take) : latin1(this.cur, this.p, take);
      this.p += wide ? 2 * take : take;
      n -= take;
    }
    return out;
  }
  /** XLUnicodeRichExtendedString (SST items, rich LABEL text). */
  richString(fonts: readonly (Font | undefined)[]): StringItem {
    const cch = this.u16();
    const flags = this.u8();
    const runs = flags & 8 ? this.u16() : 0;
    const ext = flags & 4 ? this.u32() : 0;
    const text = this.chars(cch, (flags & 1) !== 0);
    let item: StringItem = text;
    if (runs) {
      const starts: number[] = [];
      const fs: (Font | undefined)[] = [];
      for (let i = 0; i < runs; i++) {
        starts.push(this.u16());
        fs.push(fonts[this.u16()]);
      }
      const out: RichTextRun[] = [];
      if (starts[0]! > 0) out.push({ text: text.slice(0, starts[0]) });
      for (let i = 0; i < runs; i++) {
        const t = text.slice(starts[i], i + 1 < runs ? starts[i + 1] : text.length);
        if (t === '' && i + 1 < runs) continue;
        const f = fs[i];
        out.push(f ? { text: t, font: { ...f } } : { text: t });
      }
      item = { richText: out };
    }
    this.skip(ext);
    return item;
  }
}

/** A record: type, payload (CONTINUE payloads appended) and the separate segments. */
interface Rec {
  id: number;
  d: Uint8Array;
  dv: DataView;
  segs: Uint8Array[];
}

/** Calls `cb` for each record from `pos` until EOF (or the end of the stream). */
function records(s: Uint8Array, pos: number, cb: (r: Rec) => boolean | undefined): void {
  let p = pos;
  const n = s.length;
  while (p + 4 <= n) {
    const id = s[p]! | (s[p + 1]! << 8);
    let len = s[p + 2]! | (s[p + 3]! << 8);
    if (p + 4 + len > n) throw new RangeError('record overrun');
    const segs = [s.subarray(p + 4, p + 4 + len)];
    p += 4 + len;
    while (p + 4 <= n && s[p] === 0x3c && s[p + 1] === 0 && id !== 0x3c) {
      len = s[p + 2]! | (s[p + 3]! << 8);
      if (p + 4 + len > n) throw new RangeError('record overrun');
      segs.push(s.subarray(p + 4, p + 4 + len));
      p += 4 + len;
    }
    let d = segs[0]!;
    if (segs.length > 1) {
      d = new Uint8Array(segs.reduce((a, x) => a + x.length, 0));
      let o = 0;
      for (const x of segs) {
        d.set(x, o);
        o += x.length;
      }
    }
    if (cb({ id, d, dv: new DataView(d.buffer, d.byteOffset, d.byteLength), segs }) === false) {
      return;
    }
    if (id === 0x0a) return; // EOF
  }
}

/** XLUnicodeString (u16 length) / ShortXLUnicodeString (u8 length) at `p`: [text, end]. */
function xlString(d: Uint8Array, p: number, short: boolean): [string, number] {
  const n = short ? d[p]! : d[p]! | (d[p + 1]! << 8);
  const q = p + (short ? 1 : 2);
  const wide = d[q]! & 1;
  const end = q + 1 + (wide ? 2 * n : n);
  if (end > d.length) throw new RangeError('string overrun');
  return [wide ? utf16(d, q + 1, n) : latin1(d, q + 1, n), end];
}

const UNDERLINE: Record<number, Font['underline']> = {
  1: 'single',
  2: 'double',
  33: 'singleAccounting',
  34: 'doubleAccounting',
};

/** A palette color index → model color (0x7FFF = automatic). */
const icv = (i: number): Color | undefined => (i === 0x7fff ? undefined : { indexed: i });

function font(r: Rec): Font {
  const { d, dv } = r;
  const f: { -readonly [K in keyof Font]: Font[K] } = {};
  const flags = dv.getUint16(2, true);
  const name = xlString(d, 14, true)[0];
  if (name) f.name = name;
  f.size = dv.getUint16(0, true) / 20;
  if (dv.getUint16(6, true) >= 700) f.bold = true;
  if (flags & 2) f.italic = true;
  if (flags & 8) f.strike = true;
  if (flags & 0x10) f.outline = true;
  if (flags & 0x20) f.shadow = true;
  if (flags & 0x40) f.condense = true;
  if (flags & 0x80) f.extend = true;
  const sss = dv.getUint16(8, true);
  if (UNDERLINE[d[10]!]) f.underline = UNDERLINE[d[10]!];
  if (sss === 1) f.vertAlign = 'superscript';
  else if (sss === 2) f.vertAlign = 'subscript';
  const c = icv(dv.getUint16(4, true));
  if (c) f.color = c;
  if (d[11]) f.family = d[11];
  if (d[12]) f.charset = d[12];
  return f;
}

/** FullColorExt (XFExt colors): indexed, RGB or theme + tint. */
function fullColor(dv: DataView, p: number): Color | undefined {
  const type = dv.getUint16(p, true);
  const tint = dv.getInt16(p + 2, true) / 32767;
  const c: { argb?: string; theme?: number; indexed?: number; tint?: number } = {};
  if (type === 1) c.indexed = dv.getUint32(p + 4, true);
  else if (type === 2) {
    const u = (o: number): number => dv.getUint8(p + 4 + o);
    c.argb = argb(u(3), u(0), u(1), u(2)); // stored R G B A
  } else if (type === 3) c.theme = dv.getUint32(p + 4, true);
  else return undefined;
  if (tint) c.tint = tint;
  return c;
}

interface RawXf {
  font: number;
  fmt: number;
  locked: boolean;
  hidden: boolean;
  align: number;
  trot: number;
  misc: number;
  border1: number;
  border2: number;
  colors: number;
}

interface Globals {
  date1904: boolean;
  sheets: { name: string; state: 'visible' | 'hidden' | 'veryHidden'; pos: number; type: number }[];
  formats: Map<number, string>;
  fonts: (Font | undefined)[];
  xfs: RawXf[];
  /** XFExt colors by xf index and extension type. */
  ext: Map<number, Map<number, Color | number>>;
  strings: StringItem[];
  supbooks: { self: boolean; names: string[] }[];
  xti: [number, number, number][];
  names: {
    name: string;
    sheet: number | undefined;
    flags: number;
    rgce: Uint8Array;
    rgcb: Uint8Array;
  }[];
}

function readGlobals(s: Uint8Array): Globals {
  const g: Globals = {
    date1904: false,
    sheets: [],
    formats: new Map(),
    fonts: [],
    xfs: [],
    ext: new Map(),
    strings: [],
    supbooks: [],
    xti: [],
    names: [],
  };
  let first = true;
  records(s, 0, (r) => {
    const { d, dv } = r;
    if (first) {
      first = false;
      if (r.id !== 0x809 || dv.getUint16(0, true) !== 0x600) {
        throw new BiffParseError(
          'Not a BIFF8 workbook stream (Excel 95 and older .xls files are not supported)',
        );
      }
      return;
    }
    switch (r.id) {
      case 0x2f:
        throw new EncryptedXlsError();
      case 0x22: // DATEMODE
        g.date1904 = dv.getUint16(0, true) === 1;
        break;
      case 0x85: {
        // BOUNDSHEET
        const hs = d[4]! & 3;
        g.sheets.push({
          pos: dv.getUint32(0, true),
          state: hs === 1 ? 'hidden' : hs === 2 ? 'veryHidden' : 'visible',
          type: d[5]!,
          name: xlString(d, 6, true)[0],
        });
        break;
      }
      case 0x41e: // FORMAT
        g.formats.set(dv.getUint16(0, true), xlString(d, 2, false)[0]);
        break;
      case 0x31: // FONT (index 4 is never used)
        if (g.fonts.length === 4) g.fonts.push(undefined);
        g.fonts.push(font(r));
        break;
      case 0xe0: {
        // XF
        const f = dv.getUint16(4, true);
        g.xfs.push({
          font: dv.getUint16(0, true),
          fmt: dv.getUint16(2, true),
          locked: (f & 1) !== 0,
          hidden: (f & 2) !== 0,
          align: d[6]!,
          trot: d[7]!,
          misc: d[8]!,
          border1: dv.getUint32(10, true),
          border2: dv.getUint32(14, true),
          colors: dv.getUint16(18, true),
        });
        break;
      }
      case 0x87d: {
        // XFExt: frt header (12), reserved, ixfe, reserved, count, then ExtProps
        const exts = new Map<number, Color | number>();
        const n = dv.getUint16(18, true);
        let p = 20;
        for (let i = 0; i < n && p + 4 <= d.length; i++) {
          const type = dv.getUint16(p, true);
          const cb = dv.getUint16(p + 2, true);
          if (cb < 4) break;
          if (type === 14) exts.set(14, d[p + 4]!);
          else if (type >= 4 && type <= 13 && type !== 6 && p + 12 <= d.length) {
            const c = fullColor(dv, p + 4);
            if (c) exts.set(type, c);
          }
          p += cb;
        }
        g.ext.set(dv.getUint16(14, true), exts);
        break;
      }
      case 0xfc: {
        // SST (+ CONTINUE)
        const rd = new Segs(r.segs);
        rd.u32();
        const n = rd.u32();
        for (let i = 0; i < n; i++) g.strings.push(rd.richString(g.fonts));
        break;
      }
      case 0x1ae: {
        // SUPBOOK: 0x0401 = this workbook, 0x3A01 = add-in functions
        g.supbooks.push({ self: dv.getUint16(2, true) === 0x401, names: [] });
        break;
      }
      case 0x23: {
        // EXTERNNAME (of the last SUPBOOK): flags, reserved (4), cch u8, flags + chars
        const sb = g.supbooks[g.supbooks.length - 1];
        if (sb && d.length > 7) sb.names.push(xlString(d, 6, true)[0]);
        break;
      }
      case 0x17: {
        // EXTERNSHEET
        const n = dv.getUint16(0, true);
        for (let i = 0; i < n; i++) {
          const p = 2 + 6 * i;
          g.xti.push([dv.getUint16(p, true), dv.getInt16(p + 2, true), dv.getInt16(p + 4, true)]);
        }
        break;
      }
      case 0x18: {
        // NAME
        const flags = dv.getUint16(0, true);
        const cch = d[3]!;
        const cce = dv.getUint16(4, true);
        const itab = dv.getUint16(8, true);
        const wide = d[14]! & 1;
        const p = 15 + (wide ? 2 * cch : cch);
        let name = wide ? utf16(d, 15, cch) : latin1(d, 15, cch);
        if (flags & 0x20) {
          const b = BUILTIN_NAMES[name.charCodeAt(0)];
          name = b !== undefined ? `_xlnm.${b}` : name;
        }
        g.names.push({
          name,
          sheet: itab ? itab - 1 : undefined,
          flags,
          rgce: d.subarray(p, p + cce),
          rgcb: d.subarray(p + cce),
        });
        break;
      }
    }
  });
  return g;
}

/** Resolve the file's XFs into registry xfs, lazily (most .xls XFs are unused style XFs). */
function styleMap(g: Globals, wb: Workbook): (ixfe: number) => number {
  const fonts = g.fonts;
  const withExt = (
    base: Font | undefined,
    ext: Map<number, Color | number> | undefined,
  ): Font | undefined => {
    if (!base || !ext) return base;
    const c = ext.get(13);
    const scheme = ext.get(14);
    if (c === undefined && scheme === undefined) return base;
    const f: { -readonly [K in keyof Font]: Font[K] } = { ...base };
    if (typeof c === 'object') f.color = c;
    if (scheme === 1) f.scheme = 'major';
    else if (scheme === 2) f.scheme = 'minor';
    return f;
  };
  const sink = new StyleSink(
    wb.styles,
    withExt(fonts[g.xfs[0]?.font ?? 0] ?? fonts[0], g.ext.get(0)) ?? wb.styles.defaultFont,
    g.formats,
  );
  const cache = new Map<number, number>();
  return (ixfe) => {
    let v = cache.get(ixfe);
    if (v !== undefined) return v;
    const x = g.xfs[ixfe];
    if (!x) return 0;
    const ext = g.ext.get(ixfe);
    const col = (type: number, index: number): Color | undefined => {
      const c = ext?.get(type);
      return typeof c === 'object' ? c : { indexed: index };
    };
    const b1 = x.border1;
    const b2 = x.border2;
    const fg = x.colors & 0x7f;
    const bg = (x.colors >> 7) & 0x7f;
    const border: XfRecord['border'] = {};
    const top = edge((b1 >> 8) & 15, col(7, b2 & 0x7f));
    const bottom = edge((b1 >> 12) & 15, col(8, (b2 >> 7) & 0x7f));
    const left = edge(b1 & 15, col(9, (b1 >> 16) & 0x7f));
    const right = edge((b1 >> 4) & 15, col(10, (b1 >> 23) & 0x7f));
    const diag = edge((b2 >> 21) & 15, col(11, (b2 >> 14) & 0x7f));
    const b = border as { -readonly [K in keyof XfRecord['border']]: XfRecord['border'][K] };
    if (top) b.top = top;
    if (bottom) b.bottom = bottom;
    if (left) b.left = left;
    if (right) b.right = right;
    const dir = b1 >>> 30;
    if (diag && dir)
      b.diagonal = {
        ...diag,
        ...(dir & 2 ? { up: true } : {}),
        ...(dir & 1 ? { down: true } : {}),
      };
    v = sink.intern({
      numFmtId: x.fmt,
      font: withExt(fonts[x.font], ext),
      fill: fill(
        b2 >>> 26,
        ext?.has(4) ? col(4, fg) : fg === 64 ? undefined : { indexed: fg },
        ext?.has(5) ? col(5, bg) : bg === 65 ? undefined : { indexed: bg },
      ),
      border,
      alignment: alignment(
        x.align & 7,
        (x.align >> 4) & 7,
        (x.align & 8) !== 0,
        (x.misc & 0x10) !== 0,
        x.misc & 15,
        x.trot,
      ),
      protection: { locked: x.locked, hidden: x.hidden },
    });
    cache.set(ixfe, v);
    return v;
  };
}

/**
 * Read a legacy .xls file (Excel 97-2003, BIFF8) into a {@link Workbook}, synchronously: sheets
 * (order, names, hidden state), cell values, shared strings incl. rich text, styles (incl. the
 * full colors Excel 2007+ stores in XFExt records), merges, column widths, row heights, hidden
 * rows / columns, defined names, the 1904 date system, and formulas (decompiled to text, with
 * their cached results; formulas the decompiler does not support keep only their cached
 * value). Not read: charts, pictures, comments, hyperlinks, data validation, conditional
 * formatting, Excel 95 (BIFF5) and older files. Throws EncryptedXlsError for password-protected
 * files and BiffParseError for malformed data (CfbError for a broken compound file).
 */
export function readXls(data: Uint8Array | ArrayBuffer, opts: BinaryReadOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    throw new BiffParseError('Not an .xls file: this is a zip package (use readXlsx or readXlsb)');
  }
  if (!isCfb(bytes)) throw new BiffParseError('Not an .xls file: no OLE compound file signature');
  const cfb = readCfb(bytes, opts.limits);
  if (cfb.entries.some((e) => e.path === 'EncryptionInfo')) {
    throw new BiffParseError(
      "Not an .xls file: an encrypted .xlsx / .xlsb (decrypt with decryptXlsx from 'turbosheet/crypto')",
    );
  }
  const stream = cfb.read('Workbook') ?? cfb.read('WORKBOOK') ?? cfb.read('workbook');
  if (!stream) {
    throw new BiffParseError(
      cfb.read('Book')
        ? 'Excel 95 and older .xls files are not supported'
        : 'Not an .xls file: no Workbook stream',
    );
  }
  try {
    return read(stream, opts);
  } catch (e) {
    if (e instanceof RangeError) throw new BiffParseError(`Corrupt .xls data: ${e.message}`);
    throw e;
  }
}

function read(stream: Uint8Array, opts: BinaryReadOptions): Workbook {
  const g = readGlobals(stream);
  const wb = new Workbook({ date1904: g.date1904, localDates: opts.localDates });
  const xf = styleMap(g, wb);
  const nameText = (i: number): string | undefined => g.names[i - 1]?.name;
  const ctx: FormulaContext = {
    b12: false,
    sheet(ixti) {
      const x = g.xti[ixti];
      if (!x || !g.supbooks[x[0]]?.self || x[1] < 0) return undefined;
      const a = g.sheets[x[1]]?.name;
      const b = g.sheets[x[2]]?.name;
      if (a === undefined || b === undefined) return undefined;
      if (a === b) return quoteSheet(a);
      return quoteSheet(a) === a && quoteSheet(b) === b
        ? `${a}:${b}`
        : `'${`${a}:${b}`.replace(/'/g, "''")}'`;
    },
    name: nameText,
    nameX(ixti, i) {
      const sb = g.supbooks[g.xti[ixti]?.[0] ?? -1];
      if (!sb) return undefined;
      return sb.self ? nameText(i) : sb.names[i - 1];
    },
  };

  const filter = opts.sheets ? new Set(opts.sheets.map((s) => s.toLowerCase())) : undefined;
  const sheetMap = new Map<number, number>();
  g.sheets.forEach((s, i) => {
    if (s.type !== 0 || (filter && !filter.has(s.name.toLowerCase()))) return;
    if (s.pos >= stream.length) return;
    const ws = addSheet(wb, s.name, s.state);
    sheetMap.set(i, wb.worksheets.length - 1);
    readSheet(stream, s.pos, new CellWriter(ws, wb.styles, ctx), xf, g);
  });

  const raw: RawName[] = [];
  for (const n of g.names) {
    if (n.rgce.length === 0 || (n.flags & 2 && n.name.startsWith('_xlfn.'))) continue;
    const value = decodeFormula(ctx, n.rgce, n.rgcb, 0, 0);
    if (value === undefined) continue;
    raw.push({ name: n.name, value, sheet: n.sheet, hidden: (n.flags & 1) !== 0 });
  }
  applyNames(wb, raw, sheetMap);
  return wb;
}

function readSheet(
  s: Uint8Array,
  pos: number,
  w: CellWriter,
  xf: (ixfe: number) => number,
  g: Globals,
): void {
  const ws = w.ws;
  /** A FORMULA cell whose string result follows in a STRING record. */
  let str: [number, number, number, Uint8Array, Uint8Array, number | undefined] | undefined;
  const flush = (text: string): void => {
    if (!str) return;
    const [r, c, s0, rgce, rgcb, exp] = str;
    str = undefined;
    w.formula(r, c, CellKind.String, 0, text, s0, rgce, rgcb, exp);
  };
  const links: [number, number, number, number, string][] = [];
  const tips = new Map<number, string>();
  let first = true;
  records(s, pos, (rec) => {
    const { id, d, dv } = rec;
    if (first) {
      first = false;
      if (id !== 0x809) throw new BiffParseError('Sheet substream does not start with BOF');
      return;
    }
    if (str && id !== 0x207 && id !== 0x4bc && id !== 0x221) flush('');
    const r = d.length >= 2 ? dv.getUint16(0, true) : 0;
    const c = d.length >= 4 ? dv.getUint16(2, true) : 0;
    switch (id) {
      case 0x0a: // EOF
        return false;
      case 0xfd: {
        // LABELSST
        const item = g.strings[dv.getUint32(6, true)];
        const s0 = xf(dv.getUint16(4, true));
        if (item === undefined) w.value(r, c, CellKind.Empty, 0, undefined, s0);
        else
          w.value(
            r,
            c,
            typeof item === 'string' ? CellKind.String : CellKind.RichText,
            0,
            item,
            s0,
          );
        break;
      }
      case 0x204: // LABEL
      case 0xd6: // RSTRING (formatting runs ignored)
        w.value(r, c, CellKind.String, 0, xlString(d, 6, false)[0], xf(dv.getUint16(4, true)));
        break;
      case 0x203: // NUMBER
        w.number(r, c, dv.getFloat64(6, true), xf(dv.getUint16(4, true)));
        break;
      case 0x27e: // RK
        w.number(r, c, rkValue(dv.getUint32(6, true)), xf(dv.getUint16(4, true)));
        break;
      case 0xbd: {
        // MULRK
        const n = (d.length - 6) / 6;
        for (let i = 0; i < n; i++) {
          const p = 4 + 6 * i;
          w.number(r, c + i, rkValue(dv.getUint32(p + 2, true)), xf(dv.getUint16(p, true)));
        }
        break;
      }
      case 0x205: {
        // BOOLERR
        const s0 = xf(dv.getUint16(4, true));
        if (d[7]) w.value(r, c, CellKind.Error, 0, ERRORS[d[6]!] ?? '#N/A', s0);
        else w.value(r, c, CellKind.Boolean, d[6] ? 1 : 0, undefined, s0);
        break;
      }
      case 0x201: // BLANK
        w.value(r, c, CellKind.Empty, 0, undefined, xf(dv.getUint16(4, true)));
        break;
      case 0xbe: {
        // MULBLANK
        const n = (d.length - 6) / 2;
        for (let i = 0; i < n; i++) {
          w.value(r, c + i, CellKind.Empty, 0, undefined, xf(dv.getUint16(4 + 2 * i, true)));
        }
        break;
      }
      case 0x06: {
        // FORMULA: cached value (8 bytes), flags, chn, cce, tokens, extra data
        const s0 = xf(dv.getUint16(4, true));
        const cce = dv.getUint16(20, true);
        const rgce = d.subarray(22, 22 + cce);
        const rgcb = d.subarray(22 + cce);
        const exp =
          cce === 5 && rgce[0] === 1
            ? key(dv.getUint16(23, true), dv.getUint16(25, true))
            : undefined;
        if (dv.getUint16(12, true) === 0xffff) {
          const t = d[6];
          if (t === 0) {
            str = [r, c, s0, rgce, rgcb, exp];
            break;
          }
          if (t === 1)
            w.formula(r, c, CellKind.Boolean, d[8] ? 1 : 0, undefined, s0, rgce, rgcb, exp);
          else if (t === 2)
            w.formula(r, c, CellKind.Error, 0, ERRORS[d[8]!] ?? '#N/A', s0, rgce, rgcb, exp);
          else w.formula(r, c, CellKind.String, 0, '', s0, rgce, rgcb, exp);
        } else
          w.formula(r, c, CellKind.Number, dv.getFloat64(6, true), undefined, s0, rgce, rgcb, exp);
        break;
      }
      case 0x207: {
        // STRING: the text result of the preceding FORMULA
        const rd = new Segs(rec.segs);
        const n = rd.u16();
        flush(rd.chars(n, (rd.u8() & 1) !== 0));
        break;
      }
      case 0x4bc: // SHRFMLA
      case 0x221: {
        // ARRAY
        const top = dv.getUint16(0, true);
        const bottom = dv.getUint16(2, true);
        const left = d[4]!;
        const right = d[5]!;
        const p = id === 0x4bc ? 8 : 12;
        const cce = dv.getUint16(p, true);
        w.defs.set(key(top, left), {
          rgce: d.subarray(p + 2, p + 2 + cce),
          rgcb: d.subarray(p + 2 + cce),
          array: id === 0x221 ? encodeRange({ top, left, bottom, right }) : undefined,
        });
        break;
      }
      case 0xe5: {
        // MERGECELLS
        const n = dv.getUint16(0, true);
        for (let i = 0; i < n; i++) {
          const p = 2 + 8 * i;
          try {
            ws.mergeCells(
              dv.getUint16(p, true) + 1,
              dv.getUint16(p + 4, true) + 1,
              dv.getUint16(p + 2, true) + 1,
              dv.getUint16(p + 6, true) + 1,
            );
          } catch {
            // single-cell / overlapping merges are dropped (as by readXlsx)
          }
        }
        break;
      }
      case 0x7d: {
        // COLINFO
        const last = Math.min(dv.getUint16(2, true), r + 1023, 16383);
        const width = dv.getUint16(4, true) / 256;
        const s0 = xf(dv.getUint16(6, true));
        const f = dv.getUint16(8, true);
        const hidden = (f & 1) !== 0;
        const custom = (f & 2) !== 0;
        const outline = (f >> 8) & 7;
        const collapsed = (f & 0x1000) !== 0;
        if (!(custom || hidden || outline || collapsed || s0)) break;
        for (let i = r; i <= last; i++) {
          const column = ws.getColumn(i + 1);
          if (custom || hidden) column.width = width;
          if (hidden) column.hidden = true;
          if (outline) column.outlineLevel = outline;
          if (collapsed) column.collapsed = true;
          if (s0) column.styleIndex = s0;
        }
        break;
      }
      case 0x1b8: {
        // HLINK
        const target = hlink(d, dv);
        if (target !== undefined) {
          links.push([
            r,
            dv.getUint16(2, true),
            dv.getUint16(4, true),
            dv.getUint16(6, true),
            target,
          ]);
        }
        break;
      }
      case 0x800: // HLINKTOOLTIP: frt header, Ref8, null-terminated text
        if (d.length > 12) tips.set(key(dv.getUint16(2, true), dv.getUint16(6, true)), zstr(d, 10));
        break;
      case 0x208: {
        // ROW
        const f = dv.getUint16(12, true);
        const custom = (f & 0x40) !== 0;
        const hidden = (f & 0x20) !== 0;
        const outline = f & 7;
        const collapsed = (f & 0x10) !== 0;
        const s0 = f & 0x80 ? xf(dv.getUint16(14, true) & 0xfff) : 0;
        if (custom || hidden || outline || collapsed || s0) {
          ws.setRowProps(r + 1, {
            height: custom ? (dv.getUint16(6, true) & 0x7fff) / 20 : undefined,
            hidden,
            outlineLevel: outline,
            collapsed,
            styleIndex: s0,
          });
        }
        break;
      }
    }
  });
  flush('');
  w.finish();
  for (const [top, bottom, left, right, target] of links) {
    const tip = tips.get(key(top, left));
    for (let rr = top; rr <= Math.min(bottom, top + 1023); rr++) {
      for (let cc = left; cc <= Math.min(right, left + 1023); cc++) {
        if ((top === bottom && left === right) || (ws.rowData[rr]?.find(cc) ?? -1) >= 0) {
          w.link(rr, cc, target, tip);
        }
      }
    }
  }
}

/** Null-terminated UTF-16 text at `p`. */
function zstr(d: Uint8Array, p: number): string {
  let e = p;
  while (e + 1 < d.length && (d[e] || d[e + 1])) e += 2;
  return utf16(d, p, (e - p) >> 1);
}

const URL_MONIKER = 'e0c9ea79f9bace118c8200aa004ba90b';

/**
 * HLINK target (MS-OSHARED Hyperlink Object): a URL moniker or a moniker saved as a string,
 * plus `#location`; file monikers and other targets are not read.
 */
function hlink(d: Uint8Array, dv: DataView): string | undefined {
  const flags = dv.getUint32(28, true);
  let p = 32;
  /** HyperlinkString: u32 character count (with the terminating null) + UTF-16. */
  const hstr = (): string => {
    const n = dv.getUint32(p, true);
    if (p + 4 + 2 * n > d.length) throw new RangeError('hyperlink overrun');
    const v = utf16(d, p + 4, n).replace(/\0+$/, '');
    p += 4 + 2 * n;
    return v;
  };
  if (flags & 0x14) hstr(); // display name
  if (flags & 0x80) hstr(); // target frame
  let url: string | undefined;
  if (flags & 1) {
    if (flags & 0x100) url = hstr();
    else {
      let clsid = '';
      for (let i = 0; i < 16; i++) clsid += d[p + i]!.toString(16).padStart(2, '0');
      p += 16;
      if (clsid !== URL_MONIKER) return undefined;
      const n = dv.getUint32(p, true);
      url = zstr(d.subarray(0, p + 4 + n), p + 4);
      p += 4 + n;
    }
  }
  const loc = flags & 8 ? hstr() : '';
  if (url === undefined) return loc ? `#${loc}` : undefined;
  return loc ? `${url}#${loc}` : url;
}
