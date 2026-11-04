/**
 * Column auto-fit: estimates what Excel's "AutoFit Column Width" would choose, without a font
 * rasterizer. A pure model helper (not part of any writer), so it is tree-shaken unless imported.
 *
 * Model: widths are in "characters of the default font's maximum digit width" (Calibri 11 =
 * 7 px). Text width = sum of per-character advance widths of Calibri (a compact table in
 * 1/2048 em, ASCII only; CJK/emoji = 2 digits, everything else ~1 digit) divided by the digit
 * advance (1038), scaled by font size / default font size and 1.1 for bold. The cell shows its
 * DISPLAYED text (number formats are approximated, see `compileFmt`).
 */

import type { CellRange } from './address.js';
import { isDateFormat } from './numfmt.js';
import type { CellValue, Font } from './types.js';
import type { Worksheet } from './worksheet.js';
import { CellKind } from './worksheet.js';

export interface AutoFitOptions {
  /** Columns to fit (1-based numbers, letters or keys). Default: every used column. */
  columns?: readonly (number | string)[] | undefined;
  /** Minimum width for measured columns. Default 0; when > 0, empty requested columns get it too. */
  min?: number | undefined;
  /** Maximum width (default 100). */
  max?: number | undefined;
  /** Extra characters added to the text width (default 1.5, includes Excel's 5 px cell padding). */
  padding?: number | undefined;
  /** Measure the first row of the range (default true). */
  includeHeader?: boolean | undefined;
  /** Only rows from..to (1-based, inclusive). */
  rows?: { from?: number | undefined; to?: number | undefined } | undefined;
  /** Only scan the first N non-empty rows (of the range). */
  sample?: number | undefined;
  /** Write the widths to `ws.getColumn(n).width` (default true). */
  apply?: boolean | undefined;
}

export interface EstimateWidthOptions {
  min?: number | undefined;
  max?: number | undefined;
  padding?: number | undefined;
  /** Measure the first row (default true). */
  includeHeader?: boolean | undefined;
  /** Measure the first row in bold (a styled header). Default false. */
  headerBold?: boolean | undefined;
  /** Number format codes per column (0-based array), e.g. `['0.00', undefined, 'yyyy-mm-dd']`. */
  numFmts?: readonly (string | undefined)[] | undefined;
  /** Font of the body cells (default Calibri 11). */
  font?: Font | undefined;
  /** Only scan the first N rows. */
  sample?: number | undefined;
}

// ---- character widths (Calibri, 1/2048 em) -------------------------------------------------

const D = 1038; // digit advance
const COMMA = 511;
const DOT = 517;
const MINUS = 627;
const WIDE = 2 * D;
const OTHER = 1000;
const MAX_UNIT = 1822 * 1.01; // widest ASCII glyph ('W'): cheap "cannot beat the best" bound
const TRUE_U = 4424; // advance sums of "TRUE" and "FALSE"
const FALSE_U = 4928;

// ASCII 32..126
const T: readonly number[] = [
  463, 543, 821, 1020, 1038, 1464, 1397, 452, 621, 621, 1020, 1020, 511, 627, 517, 791, 1038, 1038,
  1038, 1038, 1038, 1038, 1038, 1038, 1038, 1038, 548, 548, 1020, 1020, 1020, 949, 1831, 1185, 1114,
  1092, 1260, 1000, 941, 1292, 1276, 516, 653, 1064, 861, 1751, 1322, 1356, 1058, 1378, 1112, 941,
  998, 1314, 1162, 1822, 1063, 998, 959, 622, 791, 622, 1020, 1020, 596, 981, 1076, 866, 1076, 1019,
  625, 964, 1076, 470, 490, 931, 470, 1636, 1076, 1080, 1076, 1076, 714, 801, 686, 1076, 925, 1464,
  887, 927, 809, 616, 943, 616, 1020,
];

function charUnits(c: number): number {
  if (c < 127) return c >= 32 ? T[c - 32]! : 0;
  if (
    (c >= 0x2e80 && c <= 0x9fff) ||
    (c >= 0xac00 && c <= 0xd7af) ||
    (c >= 0xf900 && c <= 0xfaff) ||
    (c >= 0xff00 && c <= 0xff60) ||
    (c >= 0xffe0 && c <= 0xffe6) ||
    (c >= 0x2600 && c <= 0x27bf) ||
    (c >= 0xd800 && c <= 0xdbff) // emoji etc. as a surrogate pair: the high half counts 2 digits
  )
    return WIDE;
  if (
    (c >= 0xdc00 && c <= 0xdfff) ||
    (c >= 0x300 && c <= 0x36f) ||
    (c >= 0x200b && c <= 0x200f) ||
    (c >= 0xfe00 && c <= 0xfe0f)
  )
    return 0;
  return OTHER;
}

/** Units of the widest line of `s` (lines split on \n). */
function textUnits(s: string): number {
  let line = 0;
  let max = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 10) {
      if (line > max) max = line;
      line = 0;
    } else line += charUnits(c);
  }
  return Math.max(line, max);
}

// ---- number formats ------------------------------------------------------------------------

interface Section {
  /** Units of literal text around the number (e.g. `$`, `%`, parens). */
  lit: number;
  /** false: the section has no number placeholder (fixed text). */
  num: boolean;
  minInt: number;
  dec: number;
  thousands: boolean;
  /** Divide by 1000 this many times (trailing commas). */
  scale: number;
  pct: number;
  /** Exponent digits for scientific notation, 0 = none. */
  exp: number;
}

interface Fmt {
  /** 0 General, 1 date (fixed width in `date`), 2 numeric sections. */
  t: 0 | 1 | 2;
  date: number;
  pos: Section;
  neg: Section | undefined;
  zero: Section | undefined;
}

const NO_SECTION: Section = {
  lit: 0,
  num: true,
  minInt: 1,
  dec: 0,
  thousands: false,
  scale: 0,
  pct: 0,
  exp: 0,
};
const GENERAL: Fmt = { t: 0, date: 0, pos: NO_SECTION, neg: undefined, zero: undefined };

function splitSections(code: string): string[] {
  const out: string[] = [];
  let start = 0;
  for (let i = 0; i < code.length; i++) {
    const c = code.charCodeAt(i);
    if (c === 34) {
      const e = code.indexOf('"', i + 1);
      if (e < 0) break;
      i = e;
    } else if (c === 92 || c === 95 || c === 42) i++;
    else if (c === 91) {
      const e = code.indexOf(']', i + 1);
      if (e < 0) break;
      i = e;
    } else if (c === 59) {
      out.push(code.slice(start, i));
      start = i + 1;
    }
  }
  out.push(code.slice(start));
  return out;
}

/** Date/time format → a representative text of the maximum width of its output. */
function dateSample(code: string): string {
  let s = '';
  const n = code.length;
  for (let i = 0; i < n; i++) {
    const ch = code[i]!;
    const l = ch.toLowerCase();
    if (ch === '"') {
      const e = code.indexOf('"', i + 1);
      if (e < 0) break;
      s += code.slice(i + 1, e);
      i = e;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      if (ch !== '*') s += ch === '_' ? ' ' : (code[i + 1] ?? '');
      i++;
    } else if (ch === '[') {
      const e = code.indexOf(']', i + 1);
      if (e < 0) break;
      if (/^(h+|m+|s+)$/i.test(code.slice(i + 1, e))) s += '000';
      i = e;
    } else if (l === 'y' || l === 'd' || l === 'm' || l === 'h' || l === 's') {
      let j = i;
      while (j + 1 < n && code[j + 1]!.toLowerCase() === l) j++;
      const len = j - i + 1;
      i = j;
      if (l === 'y') s += len > 2 ? '0000' : '00';
      else if (l === 'm') s += len >= 5 ? 'M' : len === 4 ? 'September' : len === 3 ? 'May' : '00';
      else if (l === 'd') s += len >= 4 ? 'Wednesday' : len === 3 ? 'Wed' : '00';
      else s += '00';
    } else if (l === 'a' && code.substr(i, 5).toLowerCase() === 'am/pm') {
      s += 'AM';
      i += 4;
    } else if (l === 'a' && code.substr(i, 3).toLowerCase() === 'a/p') {
      s += 'A';
      i += 2;
    } else s += ch;
  }
  return s;
}

function compileSection(sec: string): Section {
  let lit = 0;
  let minInt = 0;
  let dec = 0;
  let thousands = false;
  let scale = 0;
  let pct = 0;
  let exp = 0;
  let seenDigit = false;
  let afterDot = false;
  let num = false;
  const n = sec.length;
  for (let i = 0; i < n; i++) {
    const c = sec[i]!;
    if (c === '"') {
      const e = sec.indexOf('"', i + 1);
      if (e < 0) break;
      lit += textUnits(sec.slice(i + 1, e));
      i = e;
    } else if (c === '\\' || c === '_') {
      const x = sec[i + 1];
      if (x !== undefined) lit += charUnits(x.charCodeAt(0));
      i++;
    } else if (c === '*') i++;
    else if (c === '[') {
      const e = sec.indexOf(']', i + 1);
      if (e < 0) break;
      if (sec[i + 1] === '$') {
        // currency symbol: [$€-407]
        const dash = sec.indexOf('-', i + 2);
        lit += textUnits(sec.slice(i + 2, dash > 0 && dash < e ? dash : e));
      }
      i = e;
    } else if (c === '0' || c === '#' || c === '?') {
      num = true;
      seenDigit = true;
      if (afterDot) dec++;
      else if (c === '0') minInt++;
    } else if (c === '.') afterDot = true;
    else if (c === ',') {
      if (!seenDigit) lit += COMMA;
      else if (/[0#?]/.test(sec.slice(i + 1))) thousands = true;
      else scale++; // trailing commas scale by 1000
    } else if (c === '%') {
      pct++;
      lit += 1464;
    } else if ((c === 'E' || c === 'e') && (sec[i + 1] === '+' || sec[i + 1] === '-') && num) {
      i++;
      lit += 1000 + 1020;
      while (sec[i + 1] === '0' || sec[i + 1] === '#' || sec[i + 1] === '?') {
        exp++;
        i++;
      }
    } else if (c === '/') lit += 791;
    else if (c !== '@') lit += charUnits(c.charCodeAt(0));
  }
  return { lit, num, minInt, dec, thousands, scale, pct, exp };
}

function compileFmt(code: string): Fmt {
  if (code === '' || code === '@') return GENERAL;
  const secs = splitSections(code);
  const first = secs[0]!;
  if (first.trim().toLowerCase() === 'general') return GENERAL;
  if (isDateFormat(first)) {
    return {
      t: 1,
      date: textUnits(dateSample(first)),
      pos: NO_SECTION,
      neg: undefined,
      zero: undefined,
    };
  }
  const pos = compileSection(first);
  if (!pos.num && pos.lit === 0) return GENERAL;
  return {
    t: 2,
    date: 0,
    pos,
    neg: secs[1] !== undefined ? compileSection(secs[1]) : undefined,
    zero: secs[2] !== undefined ? compileSection(secs[2]) : undefined,
  };
}

/** Digit count of a non-negative integer-valued number. */
function digits(x: number): number {
  if (x < 10) return 1;
  let d = Math.floor(Math.log10(x)) + 1;
  if (x >= 10 ** d) d++;
  else if (x < 10 ** (d - 1)) d--;
  return d;
}

/** General format: shortest representation up to 11 characters, like Excel. */
function generalUnits(v: number): number {
  const a = v < 0 ? -v : v;
  const neg = v < 0 ? MINUS : 0;
  if (a === 0) return D;
  if (a >= 1e11 || a < 1e-9) {
    let s = a
      .toExponential(5)
      .replace(/\.?0+e/, 'e')
      .replace('e+', 'E+')
      .replace('e-', 'E-');
    if (/E[+-]\d$/.test(s)) s = `${s.slice(0, -1)}0${s.slice(-1)}`;
    return neg + textUnits(s);
  }
  if (a === Math.floor(a)) return neg + digits(a) * D;
  const id = a >= 1 ? digits(Math.floor(a)) : 1;
  let s = a.toFixed(Math.max(0, 10 - id));
  if (s.indexOf('.') >= 0) s = s.replace(/\.?0+$/, '');
  return neg + textUnits(s);
}

function sectionUnits(s: Section, a: number, showMinus: boolean): number {
  const m = showMinus ? MINUS : 0;
  if (!s.num) return s.lit + m;
  for (let k = 0; k < s.pct; k++) a *= 100;
  for (let k = 0; k < s.scale; k++) a /= 1000;
  if (s.exp > 0) {
    const e = a === 0 ? 0 : Math.floor(Math.log10(a));
    const ed = Math.max(s.exp, digits(Math.abs(e)));
    return (Math.max(s.minInt, 1) + s.dec + ed) * D + (s.dec > 0 ? DOT : 0) + s.lit + m;
  }
  const f = 10 ** s.dec;
  const r = Math.round(a * f) / f;
  const ip = Math.floor(r);
  let int = ip < 1 ? 0 : digits(ip);
  if (int < s.minInt) int = s.minInt;
  const sep = s.thousands && int > 3 ? Math.floor((int - 1) / 3) : 0;
  return (int + s.dec) * D + sep * COMMA + (s.dec > 0 ? DOT : 0) + s.lit + m;
}

function numberUnits(f: Fmt, v: number): number {
  if (f.t === 1) return f.date;
  if (f.t === 0) return generalUnits(v);
  if (v < 0) {
    if (f.neg !== undefined) return sectionUnits(f.neg, -v, false);
    return sectionUnits(f.pos, -v, true);
  }
  if (v === 0 && f.zero !== undefined) return sectionUnits(f.zero, 0, false);
  return sectionUnits(f.pos, v, false);
}

// ---- per-style info --------------------------------------------------------------------------

interface Info {
  /** Characters per unit: size ratio x bold factor / digit advance. */
  k: number;
  rot: boolean;
  /** Extra characters for the indent. */
  indent: number;
  fmt: Fmt;
}

const ROT_CHARS = 1.4; // a rotated cell needs about one line height: ~1.4 digits at 11 pt
const DEFAULT_FONT: Font = { size: 11 };

function makeInfo(
  font: Font | undefined,
  def: Font,
  numFmt: string | undefined,
  rotation: number | undefined,
  indent: number | undefined,
  fmts: Map<string, Fmt>,
): Info {
  const ratio = (font?.size ?? def.size ?? 11) / (def.size ?? 11);
  let fmt = GENERAL;
  if (numFmt !== undefined) {
    let c = fmts.get(numFmt);
    if (c === undefined) {
      c = compileFmt(numFmt);
      fmts.set(numFmt, c);
    }
    fmt = c;
  }
  return {
    k: (ratio * (font?.bold ? 1.1 : 1)) / D,
    rot: rotation === 90 || rotation === 180 || rotation === 255,
    indent: (indent ?? 0) * 1.3 * ratio,
    fmt,
  };
}

/** Rich text: the runs concatenated, measured with the widest run font. */
function richUnits(
  runs: readonly { text: string; font?: Font | undefined }[],
  base: Info,
  def: Font,
): number {
  let s = '';
  let k = base.k;
  for (const r of runs) {
    s += r.text;
    if (r.font) {
      const rk = makeInfo(r.font, def, undefined, 0, 0, new Map()).k;
      if (rk > k) k = rk;
    }
  }
  return textUnits(s) * k;
}

/** Measured width (characters) of one cell; -1 when it has nothing to measure. */
function measure(
  kind: number,
  num: number,
  obj: unknown,
  info: Info,
  def: Font,
  best: number,
): number {
  if (info.rot) return ROT_CHARS * info.k * D;
  let u: number;
  switch (kind) {
    case CellKind.String:
      if ((obj as string).length * MAX_UNIT * info.k <= best) return 0;
      u = textUnits(obj as string);
      break;
    case CellKind.Number:
    case CellKind.Date:
      u = numberUnits(info.fmt, num);
      break;
    case CellKind.Boolean:
      u = num ? TRUE_U : FALSE_U;
      break;
    case CellKind.Error:
      u = textUnits(obj as string);
      break;
    case CellKind.Hyperlink:
      u = textUnits((obj as { text: string }).text);
      break;
    case CellKind.RichText:
      return richUnits((obj as { richText: never[] }).richText, info, def) + info.indent;
    case CellKind.Formula: {
      const r = (obj as { result?: unknown }).result;
      if (typeof r === 'number') u = numberUnits(info.fmt, r);
      else if (typeof r === 'string') u = textUnits(r);
      else if (typeof r === 'boolean') u = r ? TRUE_U : FALSE_U;
      else if (typeof r === 'object' && r !== null) u = textUnits((r as { error: string }).error);
      else return -1;
      break;
    }
    default:
      return -1;
  }
  return u * info.k + info.indent;
}

function finish(w: number, min: number, max: number, padding: number): number {
  const v = Math.round((w + padding) * 100) / 100;
  return Math.min(max, Math.max(min, v));
}

const MAXC = 16384;

/**
 * Computes (and by default applies) the width of each column from its cell contents, like
 * Excel's AutoFit. Returns the widths indexed by 0-based column (`out[0]` is column A);
 * `undefined` = column not measured. Multi-column merges are skipped; rotated text gets a small
 * fixed width. See docs/features.md for the model and its limits.
 */
export function autoFitColumns(ws: Worksheet, opts: AutoFitOptions = {}): (number | undefined)[] {
  const min = opts.min ?? 0;
  const max = opts.max ?? 100;
  const padding = opts.padding ?? 1.5;
  const styles = ws.workbook.styles;
  const def = styles.defaultFont;
  const best = new Float64Array(MAXC).fill(-1);
  let want: Uint8Array | undefined;
  let requested: number[] | undefined;
  if (opts.columns) {
    want = new Uint8Array(MAXC);
    requested = [];
    for (const c of opts.columns) {
      const i = typeof c === 'number' ? c - 1 : ws.getColumn(c).number - 1;
      if (!(i >= 0 && i < MAXC)) throw new RangeError(`Column out of range: ${c}`);
      want[i] = 1;
      requested.push(i);
    }
  }
  // multi-column merges: only the top-left cell holds a value; skip it
  let skip: Set<number> | undefined;
  for (const m of ws.mergeRanges as readonly CellRange[]) {
    if (m.left < m.right) {
      skip ??= new Set();
      skip.add(m.top * MAXC + m.left);
    }
  }
  const rows = ws.rowData;
  const from = Math.max(0, (opts.rows?.from ?? 1) - 1);
  const to = Math.min(rows.length - 1, (opts.rows?.to ?? Number.POSITIVE_INFINITY) - 1);
  const header = opts.includeHeader !== false;
  let left = opts.sample ?? Number.POSITIVE_INFINITY;
  const infos: (Info | undefined)[] = [];
  const fmts = new Map<string, Fmt>();
  for (let r = from; r <= to && left > 0; r++) {
    const row = rows[r];
    if (row === undefined || row.n === 0) continue;
    left--;
    if (!header && r === from) continue;
    const { n, col, kind, num, style, obj } = row;
    for (let i = 0; i < n; i++) {
      const c = col[i]!;
      if (want !== undefined && want[c] === 0) continue;
      const kd = kind[i]!;
      if (kd === CellKind.Empty) continue;
      if (skip?.has(r * MAXC + c)) continue;
      const xf = style[i]!;
      let info = infos[xf];
      if (info === undefined) {
        const s = styles.get(xf);
        info = infos[xf] = makeInfo(
          s.font,
          def,
          s.numFmt,
          s.alignment?.textRotation as number | undefined, // interned: always numeric
          s.alignment?.indent,
          fmts,
        );
      }
      const w = measure(kd, num[i]!, obj?.[i], info, def, best[c]!);
      if (w > best[c]!) best[c] = w;
    }
  }
  const out: (number | undefined)[] = [];
  const apply = opts.apply !== false;
  const set = (c: number, w: number): void => {
    out[c] = w;
    if (apply) ws.getColumn(c + 1).width = w;
  };
  if (requested !== undefined) {
    for (const c of requested) {
      if (best[c]! >= 0) set(c, finish(best[c]!, min, max, padding));
      else if (min > 0) set(c, min);
    }
  } else {
    for (let c = 0; c < MAXC; c++) if (best[c]! >= 0) set(c, finish(best[c]!, min, max, padding));
  }
  return out;
}

/**
 * Estimates column widths from rows before writing them (for the streaming writer, which cannot
 * auto-fit after rows are flushed). Pass a sample of the data, then set the columns:
 *
 * ```ts
 * const widths = estimateColumnWidths(sampleRows, { headerBold: true });
 * const sheet = w.addWorksheet('Data', { columns: widths.map((width) => ({ width })) });
 * ```
 * Returns widths indexed by 0-based column; unmeasured (all empty) columns are `undefined`.
 */
export function estimateColumnWidths(
  rows: Iterable<readonly CellValue[]>,
  opts: EstimateWidthOptions = {},
): (number | undefined)[] {
  const min = opts.min ?? 0;
  const max = opts.max ?? 100;
  const padding = opts.padding ?? 1.5;
  const def = opts.font ?? DEFAULT_FONT;
  const fmts = new Map<string, Fmt>();
  const body = makeInfo(def, def, undefined, 0, 0, fmts);
  const bold = makeInfo({ ...def, bold: true }, def, undefined, 0, 0, fmts);
  const dateFmt = compileFmt('yyyy-mm-dd');
  const dateTimeFmt = compileFmt('yyyy-mm-dd hh:mm:ss');
  const colFmt: (Fmt | undefined)[] = [];
  const best: number[] = [];
  let left = opts.sample ?? Number.POSITIVE_INFINITY;
  let first = true;
  for (const row of rows) {
    if (left-- <= 0) break;
    const isHeader = first;
    first = false;
    if (isHeader && opts.includeHeader === false) continue;
    const base = isHeader && opts.headerBold ? bold : body;
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined) continue;
      let f = colFmt[c];
      if (f === undefined) {
        const code = opts.numFmts?.[c];
        f = colFmt[c] = code === undefined ? GENERAL : compileFmt(code);
      }
      const info: Info = f === GENERAL ? base : { ...base, fmt: f };
      const b = best[c] ?? -1;
      let w: number;
      if (typeof v === 'string') w = measure(CellKind.String, 0, v, info, def, b);
      else if (typeof v === 'number') w = measure(CellKind.Number, v, undefined, info, def, b);
      else if (typeof v === 'boolean')
        w = measure(CellKind.Boolean, v ? 1 : 0, undefined, info, def, b);
      else if (v instanceof Date) {
        const fmt = f.t === 1 ? f : v.getTime() % 86_400_000 !== 0 ? dateTimeFmt : dateFmt;
        w = measure(CellKind.Date, 0, undefined, { ...info, fmt }, def, b);
      } else if ('error' in v) w = measure(CellKind.Error, 0, v.error, info, def, b);
      else if ('formula' in v) w = measure(CellKind.Formula, 0, v, info, def, b);
      else if ('richText' in v) w = measure(CellKind.RichText, 0, v, info, def, b);
      else w = measure(CellKind.Hyperlink, 0, v, info, def, b);
      if (w > b) best[c] = w;
    }
  }
  const out: (number | undefined)[] = [];
  for (let c = 0; c < best.length; c++) {
    const b = best[c];
    if (b !== undefined && b >= 0) out[c] = finish(b, min, max, padding);
  }
  return out;
}
