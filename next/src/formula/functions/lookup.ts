// Lookup and reference functions.

import { wildcard } from '../criteria.js';
import type { Builtin, EvalContext } from '../evaluator.js';
import { argc, dims, type FormulaFunction, sc } from '../helpers.js';
import type { FormulaNode } from '../parser.js';
import {
  check,
  compare,
  FormulaArray,
  FormulaRange,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  REF,
  round15,
  toBoolean,
  toNumber,
  VALUE,
} from '../values.js';

type Get = (k: number) => FormulaScalar;

/** One row or column to search: getter, length, and the dense column cache when available. */
interface Lane {
  get: Get;
  n: number;
  vec?: readonly FormulaScalar[] | undefined;
  off?: number;
}

function sameType(a: FormulaScalar, b: FormulaScalar): boolean {
  return a !== null && b !== null && typeof a === typeof b && !isError(a) && !isError(b);
}

const keyOf = (x: FormulaScalar): unknown =>
  typeof x === 'number'
    ? round15(x)
    : typeof x === 'string'
      ? `s${x.toLowerCase()}`
      : typeof x === 'boolean'
        ? `${x}`
        : undefined;

/** Exact-match indexes of cached columns, built on the second lookup into the same lane. */
const INDEXES = /* @__PURE__ */ new WeakMap<object, Map<string, number | Map<unknown, number>>>();

/** Exact match (case-insensitive; wildcards when `wild` and the value is text). */
function exact(v: FormulaScalar, l: Lane, wild: boolean, rev = false): number {
  const re = wild && typeof v === 'string' ? wildcard(v) : undefined;
  const { get, n, vec } = l;
  if (!re && !rev && vec) {
    let byLane = INDEXES.get(vec);
    if (!byLane) {
      byLane = new Map();
      INDEXES.set(vec, byLane);
    }
    const key = `${l.off}:${n}`;
    const m = byLane.get(key);
    if (m instanceof Map) return m.get(keyOf(v)) ?? -1;
    if (m) {
      const idx = new Map<unknown, number>();
      for (let k = n - 1; k >= 0; k--) {
        const kk = keyOf(get(k));
        if (kk !== undefined) idx.set(kk, k);
      }
      byLane.set(key, idx);
      return idx.get(keyOf(v)) ?? -1;
    }
    byLane.set(key, 1);
  }
  for (let s = 0; s < n; s++) {
    const k = rev ? n - 1 - s : s;
    const x = get(k);
    if (re ? typeof x === 'string' && re.test(x) : sameType(x, v) && compare(x, v) === 0) return k;
  }
  return -1;
}

/** Largest value <= v in ascending data (binary search, as Excel); -1 when none. */
function approx(v: FormulaScalar, { get, n }: Lane): number {
  let lo = 0;
  let hi = n - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    // blanks / errors: use the nearest typed value below
    let k = mid;
    while (k >= lo && (get(k) === null || isError(get(k)))) k--;
    if (k < lo) {
      lo = mid + 1;
      continue;
    }
    const x = get(k);
    if (compare(x, v) <= 0) {
      if (sameType(x, v)) found = k;
      lo = mid + 1;
    } else hi = k - 1;
  }
  return found;
}

/** Smallest value >= v in descending data (MATCH type -1). */
function approxDesc(v: FormulaScalar, { get, n }: Lane): number {
  let found = -1;
  for (let k = 0; k < n; k++) {
    const x = get(k);
    if (!sameType(x, v)) continue;
    if (compare(x, v) >= 0) found = k;
    else break;
  }
  return found;
}

/** Column k (vertical) or row k of a value; ranges are clipped to the sheet's used rows. */
function lane(v: FormulaValue, k: number, vertical: boolean): Lane {
  if (isRef(v)) {
    if (!vertical) return { get: (i) => v.at(k, i), n: v.cols };
    const n = Math.max(0, Math.min(v.rows, v.sheet.rowData.length - v.top));
    const c = v.left + k;
    const vec = v.r.column?.(new FormulaRange(v.r, v.sheet, v.top, c, v.bottom, c));
    if (vec) return { get: (i) => vec[v.top + i]!, n, vec, off: v.top };
    return { get: (i) => v.at(i, k), n };
  }
  if (isArr(v)) {
    return vertical ? { get: (i) => v.at(i, k), n: v.rows } : { get: (i) => v.at(k, i), n: v.cols };
  }
  return { get: () => v, n: 1 };
}

/** A 1-D lookup vector (MATCH, XLOOKUP): its lane and orientation. */
function vector(v: FormulaValue): [Lane, boolean] {
  if (isError(v)) throw v;
  const [R, C] = dims(v);
  if (R > 1 && C > 1) throw NA;
  const vertical = C === 1 && R > 1;
  return [lane(v, 0, vertical), vertical];
}

const lookupValue = (v: FormulaValue, ctx: EvalContext): FormulaScalar => {
  const x = sc(v, ctx);
  if (isError(x)) throw x;
  return x;
};

function vhlookup(horizontal: boolean): FormulaFunction {
  return (args, ctx) => {
    argc(args, 3, 4);
    const v = lookupValue(args[0]!, ctx as EvalContext);
    const t = args[1]!;
    if (isError(t)) throw t;
    const [R, C] = dims(t);
    const idx = Math.trunc(toNumber(sc(args[2]!, ctx)));
    if (idx < 1) throw VALUE;
    if (idx > (horizontal ? R : C)) throw REF;
    const isApprox = args[3] == null ? true : toBoolean(sc(args[3], ctx));
    const l = lane(t, 0, !horizontal);
    const k = isApprox ? approx(v, l) : exact(v, l, true);
    if (k < 0) throw NA;
    const r = horizontal ? at2(t, idx - 1, k) : at2(t, k, idx - 1);
    return r ?? 0;
  };
}

function at2(v: FormulaValue, i: number, j: number): FormulaScalar {
  return isRef(v) || isArr(v) ? v.at(i, j) : v;
}

/** XMATCH core: match_mode 0 exact, -1 next smaller, 1 next larger, 2 wildcard. */
function xfind(v: FormulaScalar, l: Lane, mode: number, search: number): number {
  const rev = search === -1 || search === -2;
  if (mode === 0 || mode === 2) return exact(v, l, mode === 2, rev);
  const { get, n } = l;
  if (mode !== 1 && mode !== -1) throw VALUE;
  let best = -1;
  for (let s = 0; s < n; s++) {
    const k = rev ? n - 1 - s : s;
    const x = get(k);
    if (!sameType(x, v)) continue;
    const c = compare(x, v);
    if (c === 0) return k;
    if (mode === -1 ? c < 0 : c > 0) {
      if (best < 0) best = k;
      else {
        const d = compare(x, get(best));
        if (mode === -1 ? d > 0 : d < 0) best = k;
      }
    }
  }
  return best;
}

const MATCH: FormulaFunction = (args, ctx) => {
  argc(args, 2, 3);
  const v = lookupValue(args[0]!, ctx as EvalContext);
  const [l] = vector(args[1]!);
  const type = args[2] == null ? 1 : Math.sign(Math.trunc(toNumber(sc(args[2], ctx))));
  const k = type === 0 ? exact(v, l, true) : type > 0 ? approx(v, l) : approxDesc(v, l);
  if (k < 0) throw NA;
  return k + 1;
};

const num = (v: FormulaValue | undefined, ctx: EvalContext, d: number): number =>
  v == null ? d : Math.trunc(toNumber(sc(v, ctx)));

const XMATCH: FormulaFunction = (args, ctx) => {
  argc(args, 2, 4);
  const c = ctx as EvalContext;
  const v = lookupValue(args[0]!, c);
  const [l] = vector(args[1]!);
  const k = xfind(v, l, num(args[2], c, 0), num(args[3], c, 1));
  if (k < 0) throw NA;
  return k + 1;
};

/** Row k (vertical) or column k of a value, as a reference when possible. */
function slice(v: FormulaValue, k: number, vertical: boolean): FormulaValue {
  if (isRef(v)) {
    return vertical
      ? new FormulaRange(v.r, v.sheet, v.top + k, v.left, v.top + k, v.right)
      : new FormulaRange(v.r, v.sheet, v.top, v.left + k, v.bottom, v.left + k);
  }
  if (isArr(v)) {
    const data: FormulaScalar[] = [];
    if (vertical) for (let j = 0; j < v.cols; j++) data.push(v.at(k, j));
    else for (let i = 0; i < v.rows; i++) data.push(v.at(i, k));
    return vertical ? new FormulaArray(1, v.cols, data) : new FormulaArray(v.rows, 1, data);
  }
  return v;
}

const XLOOKUP: FormulaFunction = (args, ctx) => {
  argc(args, 3, 6);
  const c = ctx as EvalContext;
  const v = lookupValue(args[0]!, c);
  const [l, vertical] = vector(args[1]!);
  const ret = args[2]!;
  if (isError(ret)) throw ret;
  const [R, C] = dims(args[1]!);
  const [RR, RC] = dims(ret);
  if (vertical ? RR !== R : RC !== C) throw VALUE;
  const k = xfind(v, l, num(args[4], c, 0), num(args[5], c, 1));
  if (k < 0) {
    if (args[3] != null) return args[3];
    throw NA;
  }
  return slice(ret, k, vertical);
};

const INDEX: FormulaFunction = (args, ctx) => {
  argc(args, 2, 4);
  const a = args[0]!;
  if (isError(a)) throw a;
  const c = ctx as EvalContext;
  const [R, C] = dims(a);
  let r = num(args[1], c, 0);
  let k = num(args[2], c, 0);
  if (args.length < 3 && R === 1) {
    k = r;
    r = 0;
  }
  if (r < 0 || k < 0) throw VALUE;
  if (r > R || k > C) throw REF;
  if (isRef(a)) {
    const top = r ? a.top + r - 1 : a.top;
    const bottom = r ? top : a.bottom;
    const left = k ? a.left + k - 1 : a.left;
    const right = k ? left : a.right;
    return new FormulaRange(a.r, a.sheet, top, left, bottom, right);
  }
  if (isArr(a)) {
    if (r && k) return a.at(r - 1, k - 1);
    if (r) return slice(a, r - 1, true);
    if (k) return slice(a, k - 1, false);
    return a;
  }
  return a;
};

const CHOOSE = (args: readonly FormulaNode[], c: EvalContext): FormulaValue => {
  if (args.length < 2) throw VALUE;
  const i = Math.trunc(toNumber(check(c.scalar(args[0]!))));
  if (i < 1 || i >= args.length) throw VALUE;
  return c.ev(args[i]!);
};

function rowcol(row: boolean): FormulaFunction {
  return (args, ctx) => {
    argc(args, 0, 1);
    const a = args[0];
    if (a == null) {
      const p = row ? ctx.row : ctx.col;
      if (p < 0) throw VALUE;
      return p + 1;
    }
    if (!isRef(a)) throw isError(a) ? a : VALUE;
    const first = (row ? a.top : a.left) + 1;
    const n = row ? a.rows : a.cols;
    if (!ctx.array || n === 1) return first;
    const data: FormulaScalar[] = [];
    for (let i = 0; i < n; i++) data.push(first + i);
    return row ? new FormulaArray(n, 1, data) : new FormulaArray(1, n, data);
  };
}

function count(row: boolean): FormulaFunction {
  return (args) => {
    argc(args, 1);
    const a = args[0]!;
    if (isError(a)) throw a;
    return dims(a)[row ? 0 : 1];
  };
}

export const lookup: Record<string, Builtin> = {
  VLOOKUP: vhlookup(false),
  HLOOKUP: vhlookup(true),
  MATCH,
  XMATCH,
  XLOOKUP,
  INDEX,
  CHOOSE: { lazy: CHOOSE },
  ROW: rowcol(true),
  COLUMN: rowcol(false),
  ROWS: count(true),
  COLUMNS: count(false),
};
