// Logical functions.

import type { Builtin, EvalContext } from '../evaluator.js';
import { argc, each, type FormulaFunction, S } from '../helpers.js';
import type { FormulaNode } from '../parser.js';
import {
  compare,
  FormulaArray,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  toBoolean,
  VALUE,
} from '../values.js';

function logic(
  op: (acc: number, b: boolean) => number,
  init: number,
  fin: (n: number) => boolean,
): FormulaFunction {
  return (args) => {
    let acc = init;
    let any = false;
    for (const a of args) {
      if (isRef(a) || isArr(a)) {
        each(a, (x) => {
          if (isError(x)) throw x;
          if (typeof x === 'boolean' || typeof x === 'number') {
            any = true;
            acc = op(acc, x === true || (typeof x === 'number' && x !== 0));
          }
        });
      } else if (a !== null) {
        any = true;
        acc = op(acc, toBoolean(a));
      }
    }
    if (!any) throw VALUE;
    return fin(acc);
  };
}

/** Pick element-wise between two values by an array condition (IF in array formulas). */
function pick(cond: FormulaArray, t: FormulaValue, f: FormulaValue): FormulaArray {
  const get = (v: FormulaValue, i: number, j: number): FormulaScalar =>
    isArr(v)
      ? (i < v.rows || v.rows === 1) && (j < v.cols || v.cols === 1)
        ? v.at(v.rows === 1 ? 0 : i, v.cols === 1 ? 0 : j)
        : NA
      : (v as FormulaScalar);
  const data: FormulaScalar[] = [];
  for (let i = 0; i < cond.rows; i++) {
    for (let j = 0; j < cond.cols; j++) {
      const c = cond.at(i, j);
      let b: boolean | FormulaScalar;
      try {
        b = toBoolean(c);
      } catch (e) {
        b = e as FormulaScalar;
      }
      data.push(typeof b === 'boolean' ? get(b ? t : f, i, j) : b);
    }
  }
  return new FormulaArray(cond.rows, cond.cols, data);
}

const IF = (args: readonly FormulaNode[], c: EvalContext): FormulaValue => {
  argc(args, 1, 3);
  const cond = (c.array ? c.operand(args[0]!) : c.scalar(args[0]!)) as FormulaScalar | FormulaArray;
  if (isArr(cond)) {
    const t = args[1] ? c.operand(args[1]) : true;
    const f = args[2] ? c.operand(args[2]) : false;
    return pick(cond, t, f);
  }
  if (isError(cond)) return cond;
  const b = toBoolean(cond);
  const branch = b ? args[1] : args[2];
  if (branch === undefined) return b;
  return c.ev(branch);
};

const iferror =
  (na: boolean) =>
  (args: readonly FormulaNode[], c: EvalContext): FormulaValue => {
    argc(args, 2);
    const hit = (x: FormulaScalar): boolean => isError(x) && (!na || x.error === '#N/A');
    const v = c.operand(args[0]!) as FormulaScalar | FormulaArray;
    if (isArr(v)) {
      if (!v.data.some(hit)) return v;
      const alt = c.scalar(args[1]!);
      return new FormulaArray(
        v.rows,
        v.cols,
        v.data.map((x) => (hit(x) ? alt : x)),
      );
    }
    return hit(v as FormulaScalar) ? c.ev(args[1]!) : v;
  };

const IFS = (args: readonly FormulaNode[], c: EvalContext): FormulaValue => {
  if (args.length < 2 || args.length % 2) throw VALUE;
  for (let i = 0; i < args.length; i += 2) {
    const v = c.scalar(args[i]!);
    if (isError(v)) return v;
    if (toBoolean(v)) return c.ev(args[i + 1]!);
  }
  return NA;
};

const SWITCH = (args: readonly FormulaNode[], c: EvalContext): FormulaValue => {
  if (args.length < 3) throw VALUE;
  const v = c.scalar(args[0]!);
  if (isError(v)) return v;
  let i = 1;
  for (; i + 1 < args.length; i += 2) {
    const x = c.scalar(args[i]!);
    if (isError(x)) return x;
    if (typeof x === typeof v && compare(x, v) === 0) return c.ev(args[i + 1]!);
  }
  return i < args.length ? c.ev(args[i]!) : NA;
};

export const logical: Record<string, Builtin> = {
  IF: { lazy: IF },
  IFS: { lazy: IFS },
  IFERROR: { lazy: iferror(false) },
  IFNA: { lazy: iferror(true) },
  SWITCH: { lazy: SWITCH },
  AND: logic(
    (a, b) => (b ? a : 0),
    1,
    (n) => n === 1,
  ),
  OR: logic(
    (a, b) => (b ? 1 : a),
    0,
    (n) => n === 1,
  ),
  XOR: logic(
    (a, b) => (b ? a ^ 1 : a),
    0,
    (n) => n === 1,
  ),
  NOT: S((x) => !toBoolean(x)),
  TRUE: () => true,
  FALSE: () => false,
};
