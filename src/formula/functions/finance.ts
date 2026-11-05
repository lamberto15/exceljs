// Financial functions. Sign convention as in Excel: money paid out is negative.

import type { Builtin } from '../evaluator.js';
import { type FormulaFunction, nums, S, sc } from '../helpers.js';
import {
  DIV0,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NUM,
  toNumber,
  VALUE,
} from '../values.js';

const n = (x: FormulaScalar | undefined, dflt?: number): number =>
  x === undefined || x === null ? (dflt ?? 0) : toNumber(x);

function fin(v: number): number {
  if (!Number.isFinite(v)) throw NUM;
  return v;
}

function pmt(r: number, nper: number, pv: number, fv: number, type: number): number {
  if (nper === 0) throw NUM;
  if (r === 0) return -(pv + fv) / nper;
  const f = (1 + r) ** nper;
  return (-(fv + pv * f) * r) / ((1 + r * type) * (f - 1));
}

function fv(r: number, nper: number, p: number, pv: number, type: number): number {
  if (r === 0) return -(pv + p * nper);
  const f = (1 + r) ** nper;
  return -(pv * f + (p * (1 + r * type) * (f - 1)) / r);
}

function pv(r: number, nper: number, p: number, fvv: number, type: number): number {
  if (r === 0) return -(fvv + p * nper);
  const f = (1 + r) ** nper;
  return -(fvv + (p * (1 + r * type) * (f - 1)) / r) / f;
}

function ipmt(
  r: number,
  per: number,
  nper: number,
  pvv: number,
  fvv: number,
  type: number,
): number {
  if (per < 1 || per > nper) throw NUM;
  const p = pmt(r, nper, pvv, fvv, type);
  if (type !== 0) {
    if (per === 1) return 0;
    return (fv(r, per - 2, p, pvv, 1) - p) * r;
  }
  return fv(r, per - 1, p, pvv, 0) * r;
}

const typ = (x: FormulaScalar | undefined): number => (n(x) !== 0 ? 1 : 0);

/**
 * Newton's method from `guess`; #NUM! when it does not converge within `iters` steps (Excel:
 * 20 for IRR / RATE, 100 for XIRR).
 */
function newton(
  f: (x: number) => number,
  df: (x: number) => number,
  guess: number,
  iters: number,
  tol: number,
  small = 0,
): number {
  let x = guess;
  for (let i = 0; i < iters; i++) {
    const y = f(x);
    // close enough to a root (RATE whose answer is 0: the step size never settles)
    if (Math.abs(y) <= small) return x;
    const d = df(x);
    if (!Number.isFinite(y) || !Number.isFinite(d) || d === 0) throw NUM;
    const next = x - y / d;
    if (Math.abs(next - x) <= tol * Math.max(1, Math.abs(next))) return next;
    x = next;
  }
  throw NUM;
}

/** Cash flows of a range / array argument: numbers only (text, logicals, blanks skipped). */
function flows(v: FormulaValue): number[] {
  const out: number[] = [];
  nums([v], (x) => {
    out.push(x);
  });
  return out;
}

const IRR: FormulaFunction = (args, ctx) => {
  if (args.length < 1 || args.length > 2) throw VALUE;
  const vs = flows(args[0]!);
  if (!vs.some((x) => x > 0) || !vs.some((x) => x < 0)) throw NUM;
  const guess = args[1] == null ? 0.1 : toNumber(sc(args[1], ctx));
  return newton(
    (r) => vs.reduce((s, v, i) => s + v / (1 + r) ** i, 0),
    (r) => vs.reduce((s, v, i) => s - (i * v) / (1 + r) ** (i + 1), 0),
    guess,
    20,
    1e-10,
  );
};

const NPV: FormulaFunction = (args, ctx) => {
  if (args.length < 2) throw VALUE;
  const r = toNumber(sc(args[0]!, ctx));
  if (r === -1) throw DIV0;
  let s = 0;
  let i = 1;
  nums(args.slice(1), (x) => {
    s += x / (1 + r) ** i++;
  });
  return fin(s);
};

/** Values and dates of XNPV / XIRR (same count, dates not before the first). */
function dated(vArg: FormulaValue, dArg: FormulaValue): [number[], number[]] {
  const grab = (v: FormulaValue): FormulaScalar[] =>
    isRef(v) || isArr(v)
      ? Array.from({ length: v.rows * v.cols }, (_, k) => v.at(Math.floor(k / v.cols), k % v.cols))
      : [v as FormulaScalar];
  const vs = grab(vArg);
  const ds = grab(dArg);
  if (vs.length !== ds.length) throw NUM;
  const values: number[] = [];
  const days: number[] = [];
  for (let k = 0; k < vs.length; k++) {
    const v = vs[k]!;
    const d = ds[k]!;
    if (isError(v)) throw v;
    if (isError(d)) throw d;
    if (typeof v !== 'number' || typeof d !== 'number') throw VALUE;
    values.push(v);
    days.push(Math.floor(d));
  }
  for (const d of days) if (d < days[0]!) throw NUM;
  return [values, days];
}

const xnpv = (r: number, vs: number[], ds: number[]): number =>
  vs.reduce((s, v, k) => s + v / (1 + r) ** ((ds[k]! - ds[0]!) / 365), 0);

const XNPV: FormulaFunction = (args, ctx) => {
  if (args.length !== 3) throw VALUE;
  const r = toNumber(sc(args[0]!, ctx));
  if (r <= -1) throw NUM;
  const [vs, ds] = dated(args[1]!, args[2]!);
  return fin(xnpv(r, vs, ds));
};

const XIRR: FormulaFunction = (args, ctx) => {
  if (args.length < 2 || args.length > 3) throw VALUE;
  const [vs, ds] = dated(args[0]!, args[1]!);
  if (!vs.some((x) => x > 0) || !vs.some((x) => x < 0)) throw NUM;
  const guess = args[2] == null ? 0.1 : toNumber(sc(args[2], ctx));
  return newton(
    (r) => xnpv(r, vs, ds),
    (r) =>
      vs.reduce((s, v, k) => {
        const t = (ds[k]! - ds[0]!) / 365;
        return s - (t * v) / (1 + r) ** (t + 1);
      }, 0),
    guess,
    100,
    1e-10,
  );
};

const RATE: FormulaFunction = (args, ctx) => {
  if (args.length < 3 || args.length > 6) throw VALUE;
  const [nper, p, pvv, fvv, t, g] = args.map((a) => (a == null ? undefined : sc(a, ctx)));
  const N = n(nper);
  const P = n(p);
  const PV = n(pvv);
  const FV = n(fvv);
  const T = typ(t);
  // expm1 / log1p keep ((1+r)^N - 1) / r accurate near r = 0
  const f = (r: number): number => {
    if (r === 0) return PV + P * N + FV;
    const e = N * Math.log1p(r);
    return PV * Math.exp(e) + P * (1 + r * T) * (Math.expm1(e) / r) + FV;
  };
  const h = 1e-7;
  const scale = Math.abs(PV) + Math.abs(P * N) + Math.abs(FV);
  return newton(f, (r) => (f(r + h) - f(r - h)) / (2 * h), n(g, 0.1), 20, 1e-10, 1e-10 * scale);
};

/** DB's fixed rate, rounded to three decimals as Excel does. */
function db(cost: number, salvage: number, life: number, period: number, month: number): number {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12) throw NUM;
  if (period > life + (month < 12 ? 1 : 0)) throw NUM;
  if (cost === 0) return 0;
  const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
  let total = (cost * rate * month) / 12;
  if (period === 1) return total;
  for (let p = 2; p <= period; p++) {
    const d = p === life + 1 ? ((cost - total) * rate * (12 - month)) / 12 : (cost - total) * rate;
    if (p === period) return d;
    total += d;
  }
  return 0;
}

function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0 || period > life) {
    throw NUM;
  }
  let acc = 0;
  let d = 0;
  for (let p = 1; p <= Math.ceil(period); p++) {
    d = Math.min(((cost - acc) * factor) / life, Math.max(0, cost - salvage - acc));
    acc += d;
  }
  return d;
}

export const finance: Record<string, Builtin> = {
  PMT: S((r, np, p, f, t) => fin(pmt(n(r), n(np), n(p), n(f), typ(t)))),
  FV: S((r, np, p, v, t) => fin(fv(n(r), n(np), n(p), n(v), typ(t)))),
  PV: S((r, np, p, f, t) => fin(pv(n(r), n(np), n(p), n(f), typ(t)))),
  IPMT: S((r, per, np, v, f, t) => fin(ipmt(n(r), n(per), n(np), n(v), n(f), typ(t)))),
  PPMT: S((r, per, np, v, f, t) => {
    const R = n(r);
    const N = n(np);
    const T = typ(t);
    return fin(pmt(R, N, n(v), n(f), T) - ipmt(R, n(per), N, n(v), n(f), T));
  }),
  NPER: S((r, p, v, f, t) => {
    const R = n(r);
    const P = n(p);
    const PV = n(v);
    const FV = n(f);
    const T = typ(t);
    if (R === 0) {
      if (P === 0) throw NUM;
      return -(PV + FV) / P;
    }
    const a = P * (1 + R * T) - FV * R;
    const b = P * (1 + R * T) + PV * R;
    if (a / b <= 0) throw NUM;
    return fin(Math.log(a / b) / Math.log(1 + R));
  }),
  NPV,
  IRR,
  XNPV,
  XIRR,
  RATE,
  SLN: S((c, s, l) => {
    const life = n(l);
    if (life === 0) throw DIV0;
    return (n(c) - n(s)) / life;
  }),
  DB: S((c, s, l, p, m) => db(n(c), n(s), n(l), Math.trunc(n(p)), Math.trunc(n(m, 12)))),
  DDB: S((c, s, l, p, f) => ddb(n(c), n(s), n(l), n(p), n(f, 2))),
  EFFECT: S((r, p) => {
    const R = n(r);
    const P = Math.trunc(n(p));
    if (R <= 0 || P < 1) throw NUM;
    return fin((1 + R / P) ** P - 1);
  }),
  NOMINAL: S((r, p) => {
    const R = n(r);
    const P = Math.trunc(n(p));
    if (R <= 0 || P < 1) throw NUM;
    return fin(P * ((1 + R) ** (1 / P) - 1));
  }),
};
