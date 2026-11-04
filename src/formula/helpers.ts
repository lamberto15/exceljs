// Shared helpers for the evaluator and the built-in functions.

import type { Workbook } from '../core/workbook.js';
import type { Worksheet } from '../core/worksheet.js';
import {
  type FormulaArray as Arr,
  FormulaArray,
  type FormulaRange,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  NUM,
  toNumber,
  VALUE,
} from './values.js';

/** What a function implementation receives besides its arguments. */
export interface FunctionContext {
  readonly workbook: Workbook;
  /** Sheet of the cell being evaluated. */
  readonly sheet: Worksheet;
  /** 0-based row / column of the cell being evaluated; -1 for ad-hoc formulas. */
  readonly row: number;
  readonly col: number;
  /** True inside an array formula (and SUMPRODUCT arguments): operators work element-wise. */
  readonly array: boolean;
  readonly date1904: boolean;
  /** The evaluation clock (TODAY, NOW). */
  now(): Date;
}

/**
 * A function implementation. Arguments are evaluated first; a reference argument arrives as a
 * FormulaRange (not dereferenced), an omitted one as null. Return (or throw) a CellError for
 * Excel errors.
 */
export type FormulaFunction = (args: FormulaValue[], ctx: FunctionContext) => FormulaValue;

const MAX_ARRAY = 1 << 24;

/** Implicit intersection: a multi-cell reference in a single-value position (legacy Excel). */
export function intersect(ref: FormulaRange, ctx: FunctionContext): FormulaScalar {
  if (ref.top === ref.bottom && ref.left === ref.right) return ref.at(0, 0);
  if (ref.left === ref.right && ctx.row >= ref.top && ctx.row <= ref.bottom) {
    return ref.at(ctx.row - ref.top, 0);
  }
  if (ref.top === ref.bottom && ctx.col >= ref.left && ctx.col <= ref.right) {
    return ref.at(0, ctx.col - ref.left);
  }
  return VALUE;
}

/** Materialize a reference as an array. */
export function toArr(ref: FormulaRange): Arr {
  const rows = ref.rows;
  const cols = ref.cols;
  if (rows * cols > MAX_ARRAY) throw NUM;
  const data: FormulaScalar[] = new Array(rows * cols).fill(null);
  ref.r.each(ref, (v, r, c) => {
    data[(r - ref.top) * cols + (c - ref.left)] = v;
  });
  return new FormulaArray(rows, cols, data);
}

/** A value in a single-value position: references intersect, arrays give their first item. */
export function sc(v: FormulaValue, ctx: FunctionContext): FormulaScalar {
  if (isRef(v)) return intersect(v, ctx);
  if (isArr(v)) return v.data[0]!;
  return v;
}

/** Rows / columns of a value (1 x 1 for scalars). */
export function dims(v: FormulaValue): [number, number] {
  return isRef(v) || isArr(v) ? [v.rows, v.cols] : [1, 1];
}

/** Item (i, j) of a value; scalars repeat. */
export function at(v: FormulaValue, i: number, j: number): FormulaScalar {
  return isRef(v) || isArr(v) ? v.at(i, j) : v;
}

function big(v: FormulaValue): boolean {
  return isArr(v) || (isRef(v) && (v.top !== v.bottom || v.left !== v.right));
}

/** Element-wise application over arrays, with Excel broadcasting (1-wide dims repeat; else #N/A). */
export function broadcast(
  args: FormulaValue[],
  f: (xs: FormulaScalar[]) => FormulaValue,
  ctx: FunctionContext,
): Arr {
  const vs = args.map((a) => (isRef(a) ? toArr(a) : a));
  let rows = 1;
  let cols = 1;
  for (const a of vs) {
    if (isArr(a)) {
      rows = Math.max(rows, a.rows);
      cols = Math.max(cols, a.cols);
    }
  }
  const data: FormulaScalar[] = new Array(rows * cols);
  const xs: FormulaScalar[] = new Array(vs.length);
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      for (let k = 0; k < vs.length; k++) {
        const a = vs[k]!;
        if (!isArr(a)) xs[k] = a;
        else {
          const ii = a.rows === 1 ? 0 : i;
          const jj = a.cols === 1 ? 0 : j;
          xs[k] = ii < a.rows && jj < a.cols ? a.at(ii, jj) : NA;
        }
      }
      let r: FormulaValue;
      try {
        r = f(xs);
      } catch (e) {
        if (!isError(e)) throw e;
        r = e;
      }
      data[i * cols + j] = sc(r, ctx);
    }
  }
  return new FormulaArray(rows, cols, data);
}

/**
 * Wrap a function of scalars: references intersect and arrays give their first item, except in
 * array mode where the function is applied element-wise.
 */
export function S(f: (...a: FormulaScalar[]) => FormulaValue): FormulaFunction {
  return (args, ctx) => {
    if (ctx.array && args.some(big)) return broadcast(args, (xs) => f(...xs), ctx);
    const xs: FormulaScalar[] = new Array(args.length);
    for (let i = 0; i < args.length; i++) xs[i] = sc(args[i]!, ctx);
    return f(...xs);
  };
}

/** Visit every scalar of a value (ranges: non-blank cells only). */
export function each(v: FormulaValue, cb: (x: FormulaScalar) => void): void {
  if (isRef(v)) v.r.each(v, cb);
  else if (isArr(v)) for (const x of v.data) cb(x);
  else cb(v);
}

/**
 * SUM-style numbers: direct arguments are coerced (TRUE = 1, "5" = 5, "a" → #VALUE!), while in
 * references and arrays only numbers count (text and logicals are ignored). Errors propagate.
 */
export function nums(args: FormulaValue[], cb: (n: number) => void): void {
  for (const a of args) {
    if (isRef(a) || isArr(a)) {
      each(a, (x) => {
        if (typeof x === 'number') cb(x);
        else if (isError(x)) throw x;
      });
    } else if (a !== null) cb(toNumber(a));
  }
}

/** Number of arguments check. */
export function argc(args: readonly unknown[], min: number, max: number = min): void {
  if (args.length < min || args.length > max) throw VALUE;
}
