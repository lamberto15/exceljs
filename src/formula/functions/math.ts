// Math and aggregation functions.

import { type Criterion, criterion } from '../criteria.js';
import type { Builtin } from '../evaluator.js';
import { argc, at, dims, each, type FormulaFunction, nums, S, sc, toArr } from '../helpers.js';
import {
  check,
  DIV0,
  FormulaArray,
  FormulaRange,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NUM,
  round15,
  toNumber,
  VALUE,
} from '../values.js';

export type Rng = FormulaRange | FormulaArray;

export function rng(v: FormulaValue | undefined): Rng {
  if (isRef(v) || isArr(v)) return v;
  if (isError(v)) throw v;
  return new FormulaArray(1, 1, [v ?? null]);
}

/** Same top-left, given shape (SUMIF's sum_range follows the criteria range's shape). */
function resize(v: Rng, rows: number, cols: number): Rng {
  if (isArr(v)) return v;
  return new FormulaRange(v.r, v.sheet, v.top, v.left, v.top + rows - 1, v.left + cols - 1);
}

/**
 * Visit the offsets where every (range, criterion) pair matches. Returns how many matching
 * positions lie past the used rows of every sheet involved (all blank) and were not visited.
 */
export function ifs(pairs: [Rng, Criterion][], cb: (i: number, j: number) => void): number {
  const [R, C] = dims(pairs[0]![0]);
  for (const [r] of pairs) {
    const [a, b] = dims(r);
    if (a !== R || b !== C) throw VALUE;
  }
  const ok = (i: number, j: number, skip: number): boolean => {
    for (let k = 0; k < pairs.length; k++) {
      if (k !== skip && !pairs[k]![1].test(at(pairs[k]![0], i, j))) return false;
    }
    return true;
  };
  const d = pairs.findIndex(([r, c]) => !c.blank && isRef(r));
  if (d >= 0) {
    const [r, c] = pairs[d]! as [FormulaRange, Criterion];
    r.r.each(r, (v, row, col) => {
      if (c.test(v) && ok(row - r.top, col - r.left, d)) cb(row - r.top, col - r.left);
    });
    return 0;
  }
  let lim = R;
  if (pairs.every(([r]) => isRef(r))) {
    lim = 0;
    for (const [r] of pairs) {
      const f = r as FormulaRange;
      lim = Math.max(lim, Math.min(R, f.sheet.rowData.length - f.top));
    }
  }
  for (let i = 0; i < lim; i++) for (let j = 0; j < C; j++) if (ok(i, j, -1)) cb(i, j);
  return (R - lim) * C;
}

export function pairsOf(
  args: FormulaValue[],
  from: number,
  ctx: Parameters<FormulaFunction>[1],
): [Rng, Criterion][] {
  if (args.length < from + 2 || (args.length - from) % 2) throw VALUE;
  const out: [Rng, Criterion][] = [];
  for (let k = from; k < args.length; k += 2)
    out.push([rng(args[k]), criterion(sc(args[k + 1]!, ctx))]);
  return out;
}

/** Sum / count of the numbers of `target` at the matching offsets. */
function sumAt(target: Rng, pairs: [Rng, Criterion][]): [number, number] {
  let s = 0;
  let n = 0;
  ifs(pairs, (i, j) => {
    const x = at(target, i, j);
    if (typeof x === 'number') {
      s += x;
      n++;
    } else if (isError(x)) throw x;
  });
  return [s, n];
}

const SUMIF: FormulaFunction = (args, ctx) => {
  argc(args, 2, 3);
  const r = rng(args[0]);
  const [R, C] = dims(r);
  const t = args[2] == null ? r : resize(rng(args[2]), R, C);
  return sumAt(t, [[r, criterion(sc(args[1]!, ctx))]])[0];
};

const AVERAGEIF: FormulaFunction = (args, ctx) => {
  argc(args, 2, 3);
  const r = rng(args[0]);
  const [R, C] = dims(r);
  const t = args[2] == null ? r : resize(rng(args[2]), R, C);
  const [s, n] = sumAt(t, [[r, criterion(sc(args[1]!, ctx))]]);
  return n ? s / n : DIV0;
};

const COUNTIFS: FormulaFunction = (args, ctx) => {
  let n = 0;
  const outside = ifs(pairsOf(args, 0, ctx), () => {
    n++;
  });
  return n + outside;
};

const sumifs =
  (avg: boolean): FormulaFunction =>
  (args, ctx) => {
    const t = rng(args[0]);
    const pairs = pairsOf(args, 1, ctx);
    const [R, C] = dims(t);
    const [a, b] = dims(pairs[0]![0]);
    if (a !== R || b !== C) throw VALUE;
    const [s, n] = sumAt(t, pairs);
    return avg ? (n ? s / n : DIV0) : s;
  };

const SUMPRODUCT: FormulaFunction = (args) => {
  if (args.length === 0) throw VALUE;
  const arrs = args.map((a) =>
    isRef(a) ? toArr(a) : isArr(a) ? a : new FormulaArray(1, 1, [check(a)]),
  );
  const { rows, cols } = arrs[0]!;
  for (const a of arrs) if (a.rows !== rows || a.cols !== cols) throw VALUE;
  let s = 0;
  for (let k = 0; k < rows * cols; k++) {
    let p = 1;
    for (const a of arrs) {
      const x = a.data[k]!;
      if (isError(x)) throw x;
      p *= typeof x === 'number' ? x : 0;
    }
    s += p;
  }
  return s;
};

/** Round |x| at d decimals (mode 0 half away from zero, 1 away from zero, -1 toward zero). */
export function roundTo(x: number, d: number, mode: number): number {
  d = Math.trunc(d);
  if (d > 15) return x;
  if (d < -308) return 0;
  const sign = x < 0 ? -1 : 1;
  const p = 10 ** Math.abs(d);
  let y = round15(d >= 0 ? Math.abs(x) * p : Math.abs(x) / p);
  y = mode === 0 ? Math.floor(y + 0.5) : mode > 0 ? Math.ceil(y) : Math.floor(y);
  const r = sign * (d >= 0 ? y / p : y * p);
  return r === 0 ? 0 : r;
}

const round = (mode: number): FormulaFunction =>
  S((x, d) => roundTo(toNumber(x), toNumber(d ?? 0), mode));

function minmax(max: boolean): FormulaFunction {
  return (args) => {
    let m = max ? -Infinity : Infinity;
    let any = false;
    nums(args, (n) => {
      any = true;
      if (max ? n > m : n < m) m = n;
    });
    return any ? m : 0;
  };
}

function finite(n: number): number {
  if (!Number.isFinite(n)) throw NUM;
  return n;
}

/** Formula text that is itself a SUBTOTAL / AGGREGATE (SUBTOTAL skips such cells). */
const NESTED = /^\s*(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i;

/**
 * SUBTOTAL(function_num, ref1, ...): 1-11 (AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, STDEV,
 * STDEVP, SUM, VAR, VARP) over references, skipping cells that hold SUBTOTAL / AGGREGATE
 * formulas; 101-111 also skip hidden rows (as Excel's table totals row does).
 */
const SUBTOTAL: FormulaFunction = (args) => {
  if (args.length < 2) return VALUE;
  const fn = toNumber(sc0(args[0]!));
  const k = Math.trunc(fn) % 100;
  if (k < 1 || k > 11 || (fn >= 12 && fn <= 100) || fn > 111 || fn < 1) return VALUE;
  const hidden = fn > 100;
  const xs: number[] = [];
  let counta = 0;
  for (const a of args.slice(1)) {
    if (!isRef(a)) return VALUE;
    const rows = a.sheet.rowData;
    a.r.each(a, (v, r, c) => {
      const row = rows[r];
      if (hidden && row?.hidden) return;
      const i = row ? row.find(c) : -1;
      if (i >= 0 && row!.kind[i] === 6) {
        const f = (row!.obj![i] as { formula: string }).formula;
        if (NESTED.test(f)) return;
      }
      if (isError(v) && k !== 2 && k !== 3) throw v;
      if (v !== null) counta++;
      if (typeof v === 'number') xs.push(v);
    });
  }
  const n = xs.length;
  const sum = xs.reduce((p, x) => p + x, 0);
  const variance = (sample: boolean): FormulaScalar => {
    if (n < (sample ? 2 : 1)) return DIV0;
    const m = sum / n;
    return xs.reduce((p, x) => p + (x - m) ** 2, 0) / (sample ? n - 1 : n);
  };
  switch (k) {
    case 1:
      return n ? sum / n : DIV0;
    case 2:
      return n;
    case 3:
      return counta;
    case 4:
      return n ? Math.max(...xs) : 0;
    case 5:
      return n ? Math.min(...xs) : 0;
    case 6:
      return n ? xs.reduce((p, x) => p * x, 1) : 0;
    case 7:
    case 8: {
      const v = variance(k === 7);
      return typeof v === 'number' ? Math.sqrt(v) : v;
    }
    case 9:
      return sum;
    default:
      return variance(k === 10);
  }
};

const sc0 = (v: FormulaValue): FormulaScalar =>
  isRef(v) ? v.r.cell(v.sheet, v.top, v.left) : isArr(v) ? v.data[0]! : (v as FormulaScalar);

export const math: Record<string, Builtin> = {
  SUBTOTAL,
  SUM: (args) => {
    let s = 0;
    nums(args, (n) => {
      s += n;
    });
    return s;
  },
  PRODUCT: (args) => {
    let p = 1;
    let any = false;
    nums(args, (n) => {
      p *= n;
      any = true;
    });
    return any ? p : 0;
  },
  AVERAGE: (args) => {
    let s = 0;
    let c = 0;
    nums(args, (n) => {
      s += n;
      c++;
    });
    return c ? s / c : DIV0;
  },
  MIN: minmax(false),
  MAX: minmax(true),
  COUNT: (args) => {
    let c = 0;
    for (const a of args) {
      if (isRef(a) || isArr(a)) {
        each(a, (x) => {
          if (typeof x === 'number') c++;
        });
      } else if (typeof a === 'number' || typeof a === 'boolean') c++;
      else if (typeof a === 'string' && !isError(scNum(a))) c++;
    }
    return c;
  },
  COUNTA: (args) => {
    let c = 0;
    for (const a of args) {
      each(a, (x) => {
        if (x !== null) c++;
      });
    }
    return c;
  },
  COUNTBLANK: (args) => {
    argc(args, 1);
    const r = rng(args[0]);
    let c = r.rows * r.cols;
    each(r, (x) => {
      if (x !== '' && x !== null) c--;
    });
    return c;
  },
  SUMIF,
  SUMIFS: sumifs(false),
  AVERAGEIF,
  AVERAGEIFS: sumifs(true),
  COUNTIF: (args, ctx) => {
    argc(args, 2);
    return COUNTIFS(args, ctx);
  },
  COUNTIFS,
  SUMPRODUCT: { array: SUMPRODUCT },
  ROUND: round(0),
  ROUNDUP: round(1),
  ROUNDDOWN: round(-1),
  INT: S((x) => Math.floor(toNumber(x))),
  MOD: S((x, y) => {
    const n = toNumber(x);
    const d = toNumber(y);
    if (d === 0) return DIV0;
    const r = n - d * Math.floor(n / d);
    return Math.abs(r) < 1e-15 * Math.abs(n) ? 0 : r;
  }),
  ABS: S((x) => Math.abs(toNumber(x))),
  POWER: S((x, y) => {
    const a = toNumber(x);
    const b = toNumber(y);
    if (a === 0) return b === 0 ? NUM : b < 0 ? DIV0 : 0;
    return finite(a ** b);
  }),
  SQRT: S((x) => {
    const n = toNumber(x);
    if (n < 0) return NUM;
    return Math.sqrt(n);
  }),
  PI: () => Math.PI,
  SIGN: S((x) => Math.sign(toNumber(x))),
};

function scNum(s: string): FormulaScalar {
  try {
    return toNumber(s);
  } catch (e) {
    return e as FormulaScalar;
  }
}
