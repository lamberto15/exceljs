// Statistical functions, and AGGREGATE (which dispatches to them).

import type { Builtin } from '../evaluator.js';
import { argc, at, dims, each, type FormulaFunction, nums, sc } from '../helpers.js';
import {
  DIV0,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  NUM,
  toNumber,
  VALUE,
} from '../values.js';
import { ifs, pairsOf, rng } from './math.js';

/** Numbers of the arguments, SUM-style (see nums). */
function list(args: FormulaValue[]): number[] {
  const xs: number[] = [];
  nums(args, (n) => {
    xs.push(n);
  });
  return xs;
}

/**
 * The "A" variants (AVERAGEA, MAXA, STDEVA, ...): in references and arrays text and FALSE count
 * as 0 and TRUE as 1 (blanks are skipped); direct arguments are coerced.
 */
function listA(args: FormulaValue[]): number[] {
  const xs: number[] = [];
  for (const a of args) {
    if (isRef(a) || isArr(a)) {
      each(a, (x) => {
        if (isError(x)) throw x;
        if (typeof x === 'number') xs.push(x);
        else if (typeof x === 'boolean') xs.push(x ? 1 : 0);
        else if (typeof x === 'string') xs.push(0);
      });
    } else if (a !== null) xs.push(toNumber(a));
  }
  return xs;
}

const maxOf = (xs: readonly number[]): number => {
  let m = -Infinity;
  for (const x of xs) if (x > m) m = x;
  return xs.length ? m : 0;
};
const minOf = (xs: readonly number[]): number => {
  let m = Infinity;
  for (const x of xs) if (x < m) m = x;
  return xs.length ? m : 0;
};

const sum = (xs: readonly number[]): number => {
  let s = 0;
  for (const x of xs) s += x;
  return s;
};

/** Variance (sample or population), two-pass. */
export function variance(xs: readonly number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) throw DIV0;
  const m = sum(xs) / n;
  let s = 0;
  for (const x of xs) s += (x - m) ** 2;
  return s / (sample ? n - 1 : n);
}

const sorted = (xs: readonly number[]): number[] => [...xs].sort((a, b) => a - b);

export function median(xs: readonly number[]): number {
  if (xs.length === 0) throw NUM;
  const s = sorted(xs);
  const h = s.length >> 1;
  return s.length % 2 ? s[h]! : (s[h - 1]! + s[h]!) / 2;
}

/** Most frequent value (the first to reach the top count, in argument order); none repeated → #N/A. */
export function mode(xs: readonly number[]): number {
  const count = new Map<number, number>();
  let best: number | undefined;
  let top = 1;
  for (const x of xs) {
    const c = (count.get(x) ?? 0) + 1;
    count.set(x, c);
    if (c > top) {
      top = c;
      best = x;
    }
  }
  // ties: the value that occurs first in the data wins
  if (best !== undefined) {
    for (const x of xs) if (count.get(x) === top) return x;
  }
  throw NA;
}

export function kth(xs: readonly number[], k: number, largest: boolean): number {
  const i = Math.ceil(k);
  if (xs.length === 0 || i < 1 || i > xs.length) throw NUM;
  const s = sorted(xs);
  return largest ? s[s.length - i]! : s[i - 1]!;
}

/** PERCENTILE.INC (exc: PERCENTILE.EXC) with linear interpolation. */
export function percentile(xs: readonly number[], k: number, exc: boolean): number {
  const n = xs.length;
  if (n === 0) throw NUM;
  const s = sorted(xs);
  let h: number;
  if (exc) {
    if (k <= 0 || k >= 1) throw NUM;
    h = (n + 1) * k - 1;
    if (h < 0 || h > n - 1) throw NUM;
  } else {
    if (k < 0 || k > 1) throw NUM;
    h = (n - 1) * k;
  }
  const lo = Math.floor(h);
  const f = h - lo;
  return lo + 1 < n ? s[lo]! + f * (s[lo + 1]! - s[lo]!) : s[lo]!;
}

export function quartile(xs: readonly number[], q: number, exc: boolean): number {
  const k = Math.trunc(q);
  if (exc ? k < 1 || k > 3 : k < 0 || k > 4) throw NUM;
  return percentile(xs, k / 4, exc);
}

const num1 = (v: FormulaValue | undefined, ctx: Parameters<FormulaFunction>[1]): number => {
  if (v === undefined) throw VALUE;
  return toNumber(sc(v, ctx));
};

const of =
  (f: (xs: number[]) => number, a = false): FormulaFunction =>
  (args) =>
    f(a ? listA(args) : list(args));

/** f(array, k): LARGE, SMALL, PERCENTILE, QUARTILE. */
const withK =
  (f: (xs: number[], k: number) => number): FormulaFunction =>
  (args, ctx) => {
    argc(args, 2);
    return f(list([args[0]!]), num1(args[1], ctx));
  };

function rank(avg: boolean): FormulaFunction {
  return (args, ctx) => {
    argc(args, 2, 3);
    const x = num1(args[0], ctx);
    if (!isRef(args[1]) && !isArr(args[1])) throw VALUE;
    const asc = args[2] != null && num1(args[2], ctx) !== 0;
    const xs = list([args[1]!]);
    let before = 0;
    let same = 0;
    for (const y of xs) {
      if (y === x) same++;
      else if (asc ? y < x : y > x) before++;
    }
    if (same === 0) throw NA;
    return avg ? before + (same + 1) / 2 : before + 1;
  };
}

/** MAXIFS / MINIFS(target, criteria_range1, criterion1, ...): 0 when nothing matches. */
function extremeIfs(max: boolean): FormulaFunction {
  return (args, ctx) => {
    const t = rng(args[0]);
    const pairs = pairsOf(args, 1, ctx);
    const [R, C] = dims(t);
    const [a, b] = dims(pairs[0]![0]);
    if (a !== R || b !== C) throw VALUE;
    let m = max ? -Infinity : Infinity;
    ifs(pairs, (i, j) => {
      const x = at(t, i, j);
      if (typeof x === 'number') {
        if (max ? x > m : x < m) m = x;
      } else if (isError(x)) throw x;
    });
    return Number.isFinite(m) ? m : 0;
  };
}

const CORREL: FormulaFunction = (args) => {
  argc(args, 2);
  const a = rng(args[0]);
  const b = rng(args[1]);
  const [R, C] = dims(a);
  const [R2, C2] = dims(b);
  if (R * C !== R2 * C2) throw NA;
  const xs: number[] = [];
  const ys: number[] = [];
  for (let k = 0; k < R * C; k++) {
    const x = at(a, Math.floor(k / C), k % C);
    const y = at(b, Math.floor(k / C2), k % C2);
    if (isError(x)) throw x;
    if (isError(y)) throw y;
    if (typeof x === 'number' && typeof y === 'number') {
      xs.push(x);
      ys.push(y);
    }
  }
  const n = xs.length;
  if (n < 2) throw DIV0;
  const mx = sum(xs) / n;
  const my = sum(ys) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let k = 0; k < n; k++) {
    sxy += (xs[k]! - mx) * (ys[k]! - my);
    sxx += (xs[k]! - mx) ** 2;
    syy += (ys[k]! - my) ** 2;
  }
  if (sxx === 0 || syy === 0) throw DIV0;
  return sxy / Math.sqrt(sxx * syy);
};

// ---- AGGREGATE ------------------------------------------------------------------------------

/** Formula text that is itself a SUBTOTAL / AGGREGATE. */
const NESTED = /^\s*(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i;

/**
 * AGGREGATE(function, options, ref1, ...) for functions 1-13 and AGGREGATE(function, options,
 * array, k) for 14-19. Options: 0-3 skip nested SUBTOTAL / AGGREGATE cells, 1 / 3 / 5 / 7 skip
 * hidden rows, 2 / 3 / 6 / 7 skip error values.
 */
const AGGREGATE: FormulaFunction = (args, ctx) => {
  if (args.length < 3) throw VALUE;
  const fn = Math.trunc(num1(args[0], ctx));
  const opt = args[1] == null ? 0 : Math.trunc(num1(args[1], ctx));
  if (fn < 1 || fn > 19 || opt < 0 || opt > 7) throw VALUE;
  const nested = opt <= 3;
  const hidden = opt % 2 === 1;
  const errors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
  const xs: number[] = [];
  let counta = 0;
  const take = (v: FormulaScalar): void => {
    if (isError(v)) {
      if (errors) return;
      throw v;
    }
    if (v !== null) counta++;
    if (typeof v === 'number') xs.push(v);
  };
  const refs = fn >= 14 ? [args[2]!] : args.slice(2);
  if (fn >= 14 && args.length !== 4) throw VALUE;
  for (const a of refs) {
    if (isRef(a)) {
      const rows = a.sheet.rowData;
      a.r.each(a, (v, r, c) => {
        const row = rows[r];
        if (hidden && row?.hidden) return;
        if (nested) {
          const i = row ? row.find(c) : -1;
          if (
            i >= 0 &&
            row!.kind[i] === 6 &&
            NESTED.test((row!.obj![i] as { formula: string }).formula)
          ) {
            return;
          }
        }
        take(v);
      });
    } else if (isArr(a)) for (const v of a.data) take(v);
    else take(a);
  }
  const k = fn >= 14 ? num1(args[3], ctx) : 0;
  switch (fn) {
    case 1:
      if (!xs.length) throw DIV0;
      return sum(xs) / xs.length;
    case 2:
      return xs.length;
    case 3:
      return counta;
    case 4:
      return maxOf(xs);
    case 5:
      return minOf(xs);
    case 6:
      return xs.length ? xs.reduce((p, x) => p * x, 1) : 0;
    case 7:
      return Math.sqrt(variance(xs, true));
    case 8:
      return Math.sqrt(variance(xs, false));
    case 9:
      return sum(xs);
    case 10:
      return variance(xs, true);
    case 11:
      return variance(xs, false);
    case 12:
      return median(xs);
    case 13:
      return mode(xs);
    case 14:
      return kth(xs, k, true);
    case 15:
      return kth(xs, k, false);
    case 16:
      return percentile(xs, k, false);
    case 17:
      return quartile(xs, k, false);
    case 18:
      return percentile(xs, k, true);
    default:
      return quartile(xs, k, true);
  }
};

const stdev = (sample: boolean, a = false) => of((xs) => Math.sqrt(variance(xs, sample)), a);
const vari = (sample: boolean, a = false) => of((xs) => variance(xs, sample), a);

export const stats: Record<string, Builtin> = {
  MEDIAN: of(median),
  MODE: of(mode),
  'MODE.SNGL': of(mode),
  LARGE: withK((xs, k) => kth(xs, k, true)),
  SMALL: withK((xs, k) => kth(xs, k, false)),
  RANK: rank(false),
  'RANK.EQ': rank(false),
  'RANK.AVG': rank(true),
  PERCENTILE: withK((xs, k) => percentile(xs, k, false)),
  'PERCENTILE.INC': withK((xs, k) => percentile(xs, k, false)),
  'PERCENTILE.EXC': withK((xs, k) => percentile(xs, k, true)),
  QUARTILE: withK((xs, k) => quartile(xs, k, false)),
  'QUARTILE.INC': withK((xs, k) => quartile(xs, k, false)),
  'QUARTILE.EXC': withK((xs, k) => quartile(xs, k, true)),
  STDEV: stdev(true),
  'STDEV.S': stdev(true),
  STDEVP: stdev(false),
  'STDEV.P': stdev(false),
  STDEVA: stdev(true, true),
  STDEVPA: stdev(false, true),
  VAR: vari(true),
  'VAR.S': vari(true),
  VARP: vari(false),
  'VAR.P': vari(false),
  VARA: vari(true, true),
  VARPA: vari(false, true),
  AVERAGEA: of((xs) => {
    if (!xs.length) throw DIV0;
    return sum(xs) / xs.length;
  }, true),
  MAXA: of(maxOf, true),
  MINA: of(minOf, true),
  MAXIFS: extremeIfs(true),
  MINIFS: extremeIfs(false),
  CORREL,
  PEARSON: CORREL,
  AGGREGATE,
};
