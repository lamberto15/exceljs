// More math: logarithms, rounding to multiples, integer functions, combinatorics, random
// numbers, trigonometry.

import type { Builtin } from '../evaluator.js';
import { type FormulaFunction, nums, S } from '../helpers.js';
import { DIV0, type FormulaScalar, NUM, round15, toNumber, VALUE } from '../values.js';

const n = (x: FormulaScalar): number => toNumber(x);

function fin(v: number): number {
  if (!Number.isFinite(v)) throw NUM;
  return v;
}

/** x / y with the quotient snapped to 15 significant digits (so 0.3/0.1 is 3, not 2.9999…). */
const q15 = (x: number, y: number): number => round15(x / y);

/** Round half away from zero. */
const roundAway = (x: number): number => Math.sign(x) * Math.round(Math.abs(x));

/** Integers of the arguments for GCD / LCM (truncated; negatives → #NUM!). */
function ints(args: Parameters<FormulaFunction>[0]): number[] {
  const out: number[] = [];
  nums(args, (x) => {
    if (x < 0) throw NUM;
    out.push(Math.trunc(x));
  });
  if (out.length === 0) throw VALUE;
  return out;
}

const gcd2 = (a: number, b: number): number => {
  while (b) [a, b] = [b, a % b];
  return a;
};

function fact(x: number): number {
  const k = Math.trunc(x);
  if (k < 0 || k > 170) throw NUM;
  let r = 1;
  for (let i = 2; i <= k; i++) r *= i;
  return r;
}

function combin(a: number, b: number): number {
  const N = Math.trunc(a);
  const K = Math.trunc(b);
  if (N < 0 || K < 0 || K > N) throw NUM;
  let r = 1;
  const k = Math.min(K, N - K);
  for (let i = 1; i <= k; i++) r = (r * (N - k + i)) / i;
  return fin(Math.round(r));
}

export const math2: Record<string, Builtin> = {
  LN: S((x) => {
    const v = n(x);
    if (v <= 0) throw NUM;
    return Math.log(v);
  }),
  LOG10: S((x) => {
    const v = n(x);
    if (v <= 0) throw NUM;
    return Math.log10(v);
  }),
  LOG: S((x, b) => {
    const v = n(x);
    const base = b === undefined || b === null ? 10 : n(b);
    if (v <= 0 || base <= 0) throw NUM;
    if (base === 1) throw DIV0;
    return base === 10 ? Math.log10(v) : Math.log(v) / Math.log(base);
  }),
  EXP: S((x) => fin(Math.exp(n(x)))),
  CEILING: S((x, s) => {
    const v = n(x);
    const sig = n(s ?? null);
    if (sig === 0 || v === 0) return 0;
    if (v > 0 && sig < 0) throw NUM;
    return Math.ceil(q15(v, sig)) * sig;
  }),
  FLOOR: S((x, s) => {
    const v = n(x);
    const sig = n(s ?? null);
    if (sig === 0) return v === 0 ? 0 : DIV0;
    if (v > 0 && sig < 0) throw NUM;
    return Math.floor(q15(v, sig)) * sig;
  }),
  'CEILING.MATH': S((x, s, m) => {
    const v = n(x);
    const sig = Math.abs(s === undefined || s === null ? 1 : n(s));
    if (sig === 0) return 0;
    if (v < 0 && m !== undefined && m !== null && n(m) !== 0) return -Math.ceil(q15(-v, sig)) * sig;
    return Math.ceil(q15(v, sig)) * sig;
  }),
  'FLOOR.MATH': S((x, s, m) => {
    const v = n(x);
    const sig = Math.abs(s === undefined || s === null ? 1 : n(s));
    if (sig === 0) return 0;
    if (v < 0 && m !== undefined && m !== null && n(m) !== 0)
      return -Math.floor(q15(-v, sig)) * sig;
    return Math.floor(q15(v, sig)) * sig;
  }),
  MROUND: S((x, m) => {
    const v = n(x);
    const mul = n(m);
    if (mul === 0) return 0;
    if (v !== 0 && Math.sign(v) !== Math.sign(mul)) throw NUM;
    return roundAway(q15(v, mul)) * mul;
  }),
  TRUNC: S((x, d) => {
    const v = n(x);
    const k = d === undefined || d === null ? 0 : Math.trunc(n(d));
    const f = 10 ** k;
    return Math.trunc(round15(v * f)) / f;
  }),
  EVEN: S((x) => {
    const v = n(x);
    const r = Math.ceil(Math.abs(v) / 2) * 2;
    return v < 0 ? -r : r;
  }),
  ODD: S((x) => {
    const v = n(x);
    let r = Math.ceil(Math.abs(v));
    if (r % 2 === 0) r++;
    return v < 0 ? -r : r;
  }),
  QUOTIENT: S((x, y) => {
    const d = n(y);
    if (d === 0) throw DIV0;
    return Math.trunc(q15(n(x), d));
  }),
  GCD: (args) => ints(args).reduce(gcd2),
  LCM: (args) => {
    const xs = ints(args);
    if (xs.includes(0)) return 0;
    return fin(xs.reduce((a, b) => (a / gcd2(a, b)) * b));
  },
  FACT: S((x) => fact(n(x))),
  COMBIN: S((a, b) => combin(n(a), n(b))),
  PERMUT: S((a, b) => {
    const N = Math.trunc(n(a));
    const K = Math.trunc(n(b));
    if (N < 0 || K < 0 || K > N) throw NUM;
    let r = 1;
    for (let i = N - K + 1; i <= N; i++) r *= i;
    return fin(r);
  }),
  RAND: () => Math.random(),
  RANDBETWEEN: S((a, b) => {
    const lo = Math.ceil(n(a));
    const hi = Math.floor(n(b));
    if (lo > hi) throw NUM;
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  }),
  SUMSQ: (args) => {
    let s = 0;
    nums(args, (x) => {
      s += x * x;
    });
    return fin(s);
  },
  RADIANS: S((x) => (n(x) * Math.PI) / 180),
  DEGREES: S((x) => (n(x) * 180) / Math.PI),
  SIN: S((x) => Math.sin(n(x))),
  COS: S((x) => Math.cos(n(x))),
  TAN: S((x) => fin(Math.tan(n(x)))),
  ASIN: S((x) => {
    const v = n(x);
    if (v < -1 || v > 1) throw NUM;
    return Math.asin(v);
  }),
  ACOS: S((x) => {
    const v = n(x);
    if (v < -1 || v > 1) throw NUM;
    return Math.acos(v);
  }),
  ATAN: S((x) => Math.atan(n(x))),
  // Excel's argument order is (x, y)
  ATAN2: S((x, y) => {
    const a = n(x);
    const b = n(y);
    if (a === 0 && b === 0) throw DIV0;
    return Math.atan2(b, a);
  }),
};
