// Runtime values of the formula evaluator and Excel's conversion / comparison rules.

import { dateToSerial } from '../core/dates.js';
import type { CellError, ErrorCode } from '../core/types.js';
import type { Worksheet } from '../core/worksheet.js';

/** A single value. `null` = blank (an empty cell or an omitted argument). */
export type FormulaScalar = number | string | boolean | CellError | null;

/** What an expression evaluates to: a scalar, a reference or an array. */
export type FormulaValue = FormulaScalar | FormulaRange | FormulaArray;

/** Reads a cell through the evaluator (formula cells give their computed result). */
export interface CellReader {
  /** Value at 0-based (row, col) of `ws`. */
  cell(ws: Worksheet, row: number, col: number): FormulaScalar;
  /** Visit the non-empty cells of a rectangle, row by row (blank cells are skipped). */
  each(ref: FormulaRange, cb: (v: FormulaScalar, row: number, col: number) => void): void;
  /**
   * Dense values of a single-column range (indexed by absolute row; filled for the range's
   * rows up to the sheet's last used row), or undefined. Cached per evaluation run.
   */
  column?(ref: FormulaRange): readonly FormulaScalar[] | undefined;
}

/** A rectangular reference into a worksheet (0-based, inclusive). */
export class FormulaRange {
  readonly sheet: Worksheet;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
  /** @internal */
  readonly r: CellReader;
  constructor(
    reader: CellReader,
    sheet: Worksheet,
    top: number,
    left: number,
    bottom: number,
    right: number,
  ) {
    this.r = reader;
    this.sheet = sheet;
    this.top = top;
    this.left = left;
    this.bottom = bottom;
    this.right = right;
  }
  get rows(): number {
    return this.bottom - this.top + 1;
  }
  get cols(): number {
    return this.right - this.left + 1;
  }
  /** Value at offset (i, j) from the top-left corner. */
  at(i: number, j: number): FormulaScalar {
    return this.r.cell(this.sheet, this.top + i, this.left + j);
  }
}

/** An in-memory 2-D array (array constants, array formula intermediates), row-major. */
export class FormulaArray {
  readonly rows: number;
  readonly cols: number;
  readonly data: FormulaScalar[];
  constructor(rows: number, cols: number, data: FormulaScalar[]) {
    this.rows = rows;
    this.cols = cols;
    this.data = data;
  }
  at(i: number, j: number): FormulaScalar {
    return this.data[i * this.cols + j]!;
  }
}

const e = (error: ErrorCode): CellError => Object.freeze({ error });
export const DIV0: CellError = /* @__PURE__ */ e('#DIV/0!');
export const VALUE: CellError = /* @__PURE__ */ e('#VALUE!');
export const REF: CellError = /* @__PURE__ */ e('#REF!');
export const NAME: CellError = /* @__PURE__ */ e('#NAME?');
export const NUM: CellError = /* @__PURE__ */ e('#NUM!');
export const NA: CellError = /* @__PURE__ */ e('#N/A');
export const SPILL: CellError = /* @__PURE__ */ e('#SPILL!');
export const CALC: CellError = /* @__PURE__ */ e('#CALC!');

/** True for a CellError ({ error: '#N/A' } ...). */
export function isError(v: unknown): v is CellError {
  return typeof v === 'object' && v !== null && typeof (v as CellError).error === 'string';
}

export function isRef(v: unknown): v is FormulaRange {
  return v instanceof FormulaRange;
}

export function isArr(v: unknown): v is FormulaArray {
  return v instanceof FormulaArray;
}

// ---- date system (set per evaluation run; runs are synchronous) ---------------------------

let D1904 = false;
/** @internal */
export function setDate1904(v: boolean): boolean {
  const old = D1904;
  D1904 = v;
  return old;
}
/** @internal */
export function date1904(): boolean {
  return D1904;
}

// ---- numbers ----------------------------------------------------------------------------

/** Round to 15 significant digits, as Excel stores results. */
export function round15(n: number): number {
  if (Number.isInteger(n) && n < 1e15 && n > -1e15) return n === 0 ? 0 : n;
  const r = +n.toPrecision(15);
  return r === 0 ? 0 : r;
}

/** Number → text the way Excel converts in "&", CONCAT, TEXTJOIN, ... (General, 15 digits). */
export function numToStr(n: number): string {
  if (n === 0) return '0';
  const a = Math.abs(n);
  if (a >= 1e15 || a < 1e-6) {
    const [m, x] = n.toExponential(14).split('e') as [string, string];
    const mant = m.includes('.') ? m.replace(/\.?0+$/, '') : m;
    const ex = Number(x);
    return `${mant}E${ex < 0 ? '-' : '+'}${String(Math.abs(ex)).padStart(2, '0')}`;
  }
  return String(round15(n));
}

const MONTHS = 'janfebmaraprmayjunjulaugsepoctnovdec';

function monthOf(s: string): number {
  if (s.length < 3) return -1;
  const i = MONTHS.indexOf(s.slice(0, 3).toLowerCase());
  return i < 0 || i % 3 ? -1 : i / 3 + 1;
}

function ymd(y: number, m: number, d: number): number | undefined {
  if (y < 100) y += y < 30 ? 2000 : 1900;
  if (m < 1 || m > 12 || d < 1 || d > 31) return undefined;
  const t = Date.UTC(y, m - 1, d);
  const dt = new Date(t);
  if (dt.getUTCDate() !== d) return undefined;
  if (y < 1900 || y > 9999) return undefined;
  return Math.round(dateToSerial(dt, D1904));
}

const NUM_TEXT = /^(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const TIME = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([AaPp])[Mm]?)?$/;

function parseTime(s: string): number | undefined {
  const m = TIME.exec(s);
  if (!m) return undefined;
  let h = +m[1]!;
  const mi = +m[2]!;
  const sec = m[3] ? +m[3] : 0;
  if (m[4]) {
    if (h < 1 || h > 12) return undefined;
    h = (h % 12) + (m[4].toLowerCase() === 'p' ? 12 : 0);
  }
  if (mi > 59 || sec >= 60) return undefined;
  return (h * 3600 + mi * 60 + sec) / 86400;
}

function parseDate(s: string): number | undefined {
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s);
  if (m) return ymd(+m[1]!, +m[2]!, +m[3]!);
  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/.exec(s);
  if (m) return ymd(+m[3]!, +m[1]!, +m[2]!);
  m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ,]+(\d{2,4})$/.exec(s);
  if (m) return ymd(+m[3]!, monthOf(m[2]!), +m[1]!);
  m = /^([A-Za-z]{3,9})[- ](\d{1,2}),?[- ](\d{2,4})$/.exec(s);
  if (m) return ymd(+m[3]!, monthOf(m[1]!), +m[2]!);
  return undefined;
}

/**
 * Text → number as Excel coerces it: "1,234.5", " 12 ", "50%", "$5", "(3)", "1e3", dates
 * ("2024-01-31", "1/31/2024", "31-Jan-2024", "Jan 31, 2024") and times ("13:30", "1:30 PM").
 * Undefined when the text is not numeric.
 */
export function parseNumber(text: string): number | undefined {
  let s = text.trim();
  if (s === '') return undefined;
  let sign = 1;
  if (s[0] === '(' && s[s.length - 1] === ')') {
    sign = -1;
    s = s.slice(1, -1).trim();
  }
  if (s[0] === '-' || s[0] === '+') {
    if (s[0] === '-') sign = -sign;
    s = s.slice(1).trim();
  }
  if (s[0] === '$') s = s.slice(1).trim();
  let pct = false;
  if (s[s.length - 1] === '%') {
    pct = true;
    s = s.slice(0, -1).trim();
  }
  if (NUM_TEXT.test(s)) {
    const n = Number(s.replace(/,/g, ''));
    if (Number.isFinite(n)) return sign * (pct ? n / 100 : n);
  }
  if (pct || sign < 0) return undefined;
  const t = parseTime(s);
  if (t !== undefined) return t;
  const sp = s.search(/\s+\d{1,2}:/);
  const d = parseDate(sp > 0 ? s.slice(0, sp) : s);
  if (d === undefined) return undefined;
  if (sp > 0) {
    const tt = parseTime(s.slice(sp).trim());
    return tt === undefined ? undefined : d + tt;
  }
  return d;
}

// ---- coercion (non-throwing) ------------------------------------------------------------

/** Scalar → number or the error to return. */
export function toNum(v: FormulaScalar): number | CellError {
  switch (typeof v) {
    case 'number':
      return v;
    case 'boolean':
      return v ? 1 : 0;
    case 'string': {
      const n = parseNumber(v);
      return n === undefined ? VALUE : n;
    }
    default:
      return v === null ? 0 : v;
  }
}

/** Scalar → text or the error to return. */
export function toStr(v: FormulaScalar): string | CellError {
  switch (typeof v) {
    case 'string':
      return v;
    case 'number':
      return numToStr(v);
    case 'boolean':
      return v ? 'TRUE' : 'FALSE';
    default:
      return v === null ? '' : v;
  }
}

/** Scalar → boolean or the error to return. */
export function toBool(v: FormulaScalar): boolean | CellError {
  switch (typeof v) {
    case 'boolean':
      return v;
    case 'number':
      return v !== 0;
    case 'string': {
      const u = v.toUpperCase();
      return u === 'TRUE' ? true : u === 'FALSE' ? false : VALUE;
    }
    default:
      return v === null ? false : v;
  }
}

// ---- throwing coercion, for function implementations --------------------------------------
// They throw the CellError; the evaluator returns it as the function's result.

/** A single-cell range or an array gives its (first) value; a larger range is #VALUE!. */
function one(v: FormulaValue): FormulaScalar {
  if (isRef(v)) {
    if (v.top !== v.bottom || v.left !== v.right) throw VALUE;
    return v.at(0, 0);
  }
  return isArr(v) ? v.data[0]! : v;
}

/** Value → number; throws a CellError (#VALUE! for non-numeric text, or the error itself). */
export function toNumber(v: FormulaValue): number {
  const n = toNum(one(v));
  if (typeof n !== 'number') throw n;
  return n;
}

/** Value → text; throws the error a CellError argument carries. */
export function toText(v: FormulaValue): string {
  const s = toStr(one(v));
  if (typeof s !== 'string') throw s;
  return s;
}

/** Value → boolean; throws #VALUE! for text other than TRUE/FALSE, or the error itself. */
export function toBoolean(v: FormulaValue): boolean {
  const b = toBool(one(v));
  if (typeof b !== 'boolean') throw b;
  return b;
}

/** Throw the scalar if it is an error. */
export function check<T extends FormulaScalar>(v: T): Exclude<T, CellError> {
  if (isError(v)) throw v;
  return v as Exclude<T, CellError>;
}

/** Every scalar of a value, row-major (blank cells of a range included as null). */
export function scalarValues(v: FormulaValue): FormulaScalar[] {
  if (isRef(v)) {
    const out: FormulaScalar[] = [];
    for (let i = 0; i < v.rows; i++) for (let j = 0; j < v.cols; j++) out.push(v.at(i, j));
    return out;
  }
  if (isArr(v)) return v.data.slice();
  return [v];
}

// ---- comparison --------------------------------------------------------------------------

function rank(v: number | string | boolean): number {
  return typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
}

/**
 * Excel ordering of two non-error scalars: numbers < text < logicals, text case-insensitive,
 * numbers equal when they agree to 15 significant digits. Blank compares as 0 / "" / FALSE.
 */
export function compare(a: FormulaScalar, b: FormulaScalar): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const x = a as number | string | boolean;
  const y = b as number | string | boolean;
  const ra = rank(x);
  const rb = rank(y);
  if (ra !== rb) return ra - rb;
  if (ra === 0) {
    if (x === y) return 0;
    if (round15(x as number) === round15(y as number)) return 0;
    return (x as number) < (y as number) ? -1 : 1;
  }
  if (ra === 1) {
    const p = (x as string).toLowerCase();
    const q = (y as string).toLowerCase();
    return p === q ? 0 : p < q ? -1 : 1;
  }
  return x === y ? 0 : x ? 1 : -1;
}
