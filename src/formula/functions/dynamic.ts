// Dynamic-array functions (Excel 365): FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TRANSPOSE,
// RANDARRAY. Their results spill (see the evaluator's spill handling).

import type { Builtin } from '../evaluator.js';
import { type FormulaFunction, toArr } from '../helpers.js';
import {
  CALC,
  compare,
  FormulaArray,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NUM,
  toNumber,
  VALUE,
} from '../values.js';

const MAX_CELLS = 1 << 22;

/** Any value as an array (scalars are 1 x 1). */
function arr(v: FormulaValue | undefined): FormulaArray {
  if (isRef(v)) return toArr(v);
  if (isArr(v)) return v;
  return new FormulaArray(1, 1, [v ?? null]);
}

/** A scalar argument (first item of an array / reference); missing → `dflt`. */
function scalar(v: FormulaValue | undefined, dflt: FormulaScalar): FormulaScalar {
  if (v === undefined || v === null) return dflt;
  const x = isRef(v) || isArr(v) ? v.at(0, 0) : v;
  if (isError(x)) throw x;
  return x ?? dflt;
}

const int = (v: FormulaValue | undefined, dflt: number): number =>
  Math.trunc(toNumber(scalar(v, dflt)));

const bool = (v: FormulaValue | undefined): boolean => {
  const x = scalar(v, false);
  if (typeof x === 'string') throw VALUE;
  return x !== 0 && x !== false;
};

function rowsOf(a: FormulaArray): FormulaScalar[][] {
  const out: FormulaScalar[][] = [];
  for (let i = 0; i < a.rows; i++) out.push(a.data.slice(i * a.cols, (i + 1) * a.cols));
  return out;
}

function fromRows(rows: FormulaScalar[][], cols: number): FormulaArray {
  const data: FormulaScalar[] = [];
  for (const r of rows) data.push(...r);
  return new FormulaArray(rows.length, cols, data);
}

function transpose(a: FormulaArray): FormulaArray {
  const data: FormulaScalar[] = new Array(a.rows * a.cols);
  for (let i = 0; i < a.rows; i++) for (let j = 0; j < a.cols; j++) data[j * a.rows + i] = a.at(i, j);
  return new FormulaArray(a.cols, a.rows, data);
}

/** Excel's sort order: numbers < text < logicals < errors; blanks always last. */
function sortCmp(x: FormulaScalar, y: FormulaScalar, desc: boolean): number {
  if (x === null || y === null) return x === y ? 0 : x === null ? 1 : -1;
  const ex = isError(x);
  const ey = isError(y);
  let c: number;
  if (ex || ey) c = ex && ey ? 0 : ex ? 1 : -1;
  else c = compare(x, y);
  return desc ? -c : c;
}

/** Same row for UNIQUE: text compares case-insensitively, as in Excel. */
function key(row: FormulaScalar[]): string {
  return row
    .map((v) =>
      v === null
        ? 'e'
        : typeof v === 'string'
          ? `s${v.toLowerCase()}`
          : typeof v === 'number'
            ? `n${v}`
            : typeof v === 'boolean'
              ? `b${v}`
              : `x${v.error}`,
    )
    .join('\u0000');
}

const FILTER: FormulaFunction = (args) => {
  if (args.length < 2 || args.length > 3) return VALUE;
  const a = arr(args[0]);
  const inc = arr(args[1]);
  const keep = (v: FormulaScalar): boolean => {
    if (isError(v)) throw v;
    if (typeof v === 'string') throw VALUE;
    return v !== null && v !== 0 && v !== false;
  };
  let out: FormulaArray | undefined;
  if (inc.cols === 1 && inc.rows === a.rows) {
    const rows = rowsOf(a).filter((_, i) => keep(inc.at(i, 0)));
    if (rows.length) out = fromRows(rows, a.cols);
  } else if (inc.rows === 1 && inc.cols === a.cols) {
    const t = rowsOf(transpose(a)).filter((_, j) => keep(inc.at(0, j)));
    if (t.length) out = transpose(fromRows(t, a.rows));
  } else return VALUE;
  if (out) return out;
  return args.length === 3 && args[2] !== null ? args[2]! : CALC;
};

/** Sort the rows (or, by_col, the columns) of `a` by key vectors. */
function sortBy(
  a: FormulaArray,
  keys: { get: (i: number) => FormulaScalar; desc: boolean }[],
  byCol: boolean,
): FormulaArray {
  const m = byCol ? transpose(a) : a;
  const rows = rowsOf(m).map((r, i) => ({ r, i }));
  rows.sort((p, q) => {
    for (const k of keys) {
      const c = sortCmp(k.get(p.i), k.get(q.i), k.desc);
      if (c !== 0) return c;
    }
    return p.i - q.i;
  });
  const out = fromRows(
    rows.map((x) => x.r),
    m.cols,
  );
  return byCol ? transpose(out) : out;
}

const order = (v: FormulaValue | undefined): boolean => {
  const o = int(v, 1);
  if (o !== 1 && o !== -1) throw VALUE;
  return o === -1;
};

const SORT: FormulaFunction = (args) => {
  if (args.length < 1 || args.length > 4) return VALUE;
  const a = arr(args[0]);
  const byCol = bool(args[3]);
  const idx = args[1] === undefined || args[1] === null ? arr(1) : arr(args[1]);
  const ord = args[2] === undefined || args[2] === null ? arr(1) : arr(args[2]);
  const width = byCol ? a.rows : a.cols;
  const keys = idx.data.map((x, k) => {
    const n = Math.trunc(toNumber(x));
    if (n < 1 || n > width) throw VALUE;
    const desc = order(ord.data.length === 1 ? ord.data[0] : ord.data[k]);
    return {
      get: (i: number) => (byCol ? a.at(n - 1, i) : a.at(i, n - 1)),
      desc,
    };
  });
  return sortBy(a, keys, byCol);
};

const SORTBY: FormulaFunction = (args) => {
  if (args.length < 2) return VALUE;
  const a = arr(args[0]);
  const keys: { get: (i: number) => FormulaScalar; desc: boolean }[] = [];
  let byCol: boolean | undefined;
  for (let k = 1; k < args.length; k += 2) {
    const by = arr(args[k]);
    const desc = order(args[k + 1]);
    const col = by.rows === 1 && by.cols === a.cols && a.cols > 1;
    if (!col && !(by.cols === 1 && by.rows === a.rows)) return VALUE;
    if (byCol !== undefined && byCol !== col) return VALUE;
    byCol = col;
    keys.push({ get: (i) => (col ? by.at(0, i) : by.at(i, 0)), desc });
  }
  return sortBy(a, keys, byCol === true);
};

const UNIQUE: FormulaFunction = (args) => {
  if (args.length < 1 || args.length > 3) return VALUE;
  const a = arr(args[0]);
  const byCol = bool(args[1]);
  const once = bool(args[2]);
  const m = byCol ? transpose(a) : a;
  const rows = rowsOf(m);
  const count = new Map<string, number>();
  for (const r of rows) {
    const k = key(r);
    count.set(k, (count.get(k) ?? 0) + 1);
  }
  const seen = new Set<string>();
  const out: FormulaScalar[][] = [];
  for (const r of rows) {
    const k = key(r);
    if (seen.has(k) || (once && count.get(k)! > 1)) continue;
    seen.add(k);
    out.push(r);
  }
  if (out.length === 0) return CALC;
  const res = fromRows(out, m.cols);
  return byCol ? transpose(res) : res;
};

const SEQUENCE: FormulaFunction = (args) => {
  if (args.length < 1 || args.length > 4) return VALUE;
  const rows = int(args[0], 1);
  const cols = int(args[1], 1);
  const start = toNumber(scalar(args[2], 1));
  const step = toNumber(scalar(args[3], 1));
  if (rows < 1 || cols < 1) return CALC;
  if (rows * cols > MAX_CELLS) return NUM;
  const data: FormulaScalar[] = new Array(rows * cols);
  for (let k = 0; k < rows * cols; k++) data[k] = start + k * step;
  return new FormulaArray(rows, cols, data);
};

const TRANSPOSE: FormulaFunction = (args) => {
  if (args.length !== 1) return VALUE;
  return transpose(arr(args[0]));
};

const RANDARRAY: FormulaFunction = (args) => {
  if (args.length > 5) return VALUE;
  const rows = int(args[0], 1);
  const cols = int(args[1], 1);
  const min = toNumber(scalar(args[2], 0));
  const max = toNumber(scalar(args[3], 1));
  const whole = bool(args[4]);
  if (rows < 1 || cols < 1) return CALC;
  if (min > max || (whole && Math.ceil(min) > Math.floor(max))) return VALUE;
  if (rows * cols > MAX_CELLS) return NUM;
  const data: FormulaScalar[] = new Array(rows * cols);
  for (let k = 0; k < rows * cols; k++) {
    data[k] = whole
      ? Math.ceil(min) + Math.floor(Math.random() * (Math.floor(max) - Math.ceil(min) + 1))
      : min + Math.random() * (max - min);
  }
  return new FormulaArray(rows, cols, data);
};

export const dynamic: Record<string, Builtin> = {
  FILTER: { array: FILTER },
  SORT: { array: SORT },
  SORTBY: { array: SORTBY },
  UNIQUE: { array: UNIQUE },
  SEQUENCE: { array: SEQUENCE },
  TRANSPOSE: { array: TRANSPOSE },
  RANDARRAY: { array: RANDARRAY },
};
