// Formula evaluation engine: dependency graph over the formula cells of a workbook, Tarjan's
// strongly-connected components (iterative) to get a dependencies-first order and detect
// circular references, and a tree-walking interpreter over cached ASTs.

import type {
  CellError,
  DefinedName,
  ErrorCode,
  Formula,
  Hyperlink,
  RichText,
} from '../core/types.js';
import type { Workbook } from '../core/workbook.js';
import type { RowData, Worksheet } from '../core/worksheet.js';
import {
  broadcast,
  type FormulaFunction,
  type FunctionContext,
  intersect,
  toArr,
} from './helpers.js';
import { type BinaryOp, CELL_RE, colIndex, type FormulaNode, parseFormula } from './parser.js';
import {
  type CellReader,
  compare,
  DIV0,
  FormulaArray,
  FormulaRange,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  NAME,
  NUM,
  REF,
  round15,
  setDate1904,
  toNum,
  toStr,
  VALUE,
} from './values.js';

// CellKind values (core/worksheet.ts), inlined so this module does not import the Worksheet code.
const K_EMPTY = 0;
const K_NUM = 1;
const K_STR = 2;
const K_BOOL = 3;
const K_DATE = 4;
const K_ERR = 5;
const K_FORMULA = 6;
const K_RICH = 7;
const K_LINK = 8;

/** A computed formula result (what `Formula.result` holds). */
export type FormulaResult = number | string | boolean | CellError;

/** Lazy built-in: receives the argument ASTs (IF, IFERROR, CHOOSE, ...). */
export type LazyFunction = (args: readonly FormulaNode[], ctx: EvalContext) => FormulaValue;

/** @internal Built-in table entry. */
export type Builtin =
  | FormulaFunction
  | { readonly lazy: LazyFunction }
  | { readonly array: FormulaFunction };

export interface EvaluatorOptions {
  /** Extra / overriding functions, by name (case-insensitive). */
  functions?: Record<string, FormulaFunction> | undefined;
  /** Clock for TODAY() / NOW() (default: `() => new Date()`). */
  now?: (() => Date) | undefined;
}

export interface EvaluateOptions {
  /** Only evaluate the formulas of these sheets (their dependencies are evaluated too). */
  sheets?: readonly (string | Worksheet)[] | undefined;
  /** Store results into `Formula.result` (default true for `evaluate`, false for `evaluateCell`). */
  writeResults?: boolean | undefined;
  /**
   * Iterative calculation for circular references (Excel's "Enable iterative calculation"):
   * at most this many passes over each cycle. 0 (default): circular cells get #REF!.
   */
  maxIterations?: number | undefined;
  /** Iteration stops when no value changes by more than this (default 0.001). */
  maxChange?: number | undefined;
  /** Called once per unsupported function name, with the first cell using it. */
  onUnsupported?: ((name: string, address: string) => void) | undefined;
  /** Clock for TODAY() / NOW() (overrides the evaluator's). */
  now?: (() => Date) | undefined;
}

export interface EvaluationReport {
  /** Number of formula cells evaluated. */
  evaluated: number;
  /** Cells on a circular reference ("Sheet1!A1"). */
  circular: string[];
  /** Unsupported function names (sorted); their calls evaluate to #NAME?. */
  unsupported: string[];
  /** Formulas that could not be parsed (they evaluate to #NAME?), and internal failures. */
  errors: { address: string; message: string }[];
  /** Results by address ("Sheet1!A1"), when `writeResults` is false. */
  results?: Map<string, FormulaResult> | undefined;
}

export interface EvaluateFormulaOptions {
  /** Sheet for unqualified references (default: the first sheet). */
  sheet?: string | Worksheet | undefined;
  /** Host cell for ROW(), COLUMN() and implicit intersection, e.g. "C3". */
  address?: string | undefined;
  /** Evaluate as an array formula. */
  array?: boolean | undefined;
  now?: (() => Date) | undefined;
}

interface Parsed {
  ast: FormulaNode | undefined;
  error: Error | undefined;
  /** Static references: [sheet, top, left, bottom, right] or a defined name [sheet, name]. */
  refs:
    | (
        | readonly [string | undefined, number, number, number, number]
        | readonly [string | undefined, string]
      )[]
    | undefined;
  calls: string[] | undefined;
}

const errOf = (error: ErrorCode): CellError => ({ error });

/** Quote a sheet name for an address when needed. */
export function sheetPrefix(name: string): string {
  return /^[A-Za-z_][\w.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? `${name}!`
    : `'${name.replace(/'/g, "''")}'!`;
}

function collect(n: FormulaNode, p: Parsed): void {
  switch (n.type) {
    case 'cell':
      p.refs!.push([n.sheet, n.row, n.col, n.row, n.col]);
      return;
    case 'range':
      p.refs!.push([n.sheet, n.top, n.left, n.bottom, n.right]);
      return;
    case 'name':
      p.refs!.push([n.sheet, n.name]);
      return;
    case 'unary':
    case 'percent':
      collect(n.operand, p);
      return;
    case 'binary':
      collect(n.left, p);
      collect(n.right, p);
      return;
    case 'call':
      p.calls!.push(n.name);
      for (const a of n.args) collect(a, p);
      return;
    default:
  }
}

class SheetIndex {
  readonly ws: Worksheet;
  readonly i: number;
  /** Rows holding formulas (ascending) and the first node id of each (+ end sentinel). */
  readonly rows: number[] = [];
  readonly first: number[] = [];
  /** Array formula areas [top, left, bottom, right, masterId]. */
  arrays: number[][] | undefined = undefined;
  constructor(ws: Worksheet, i: number) {
    this.ws = ws;
    this.i = i;
  }
}

/** A reusable evaluator holding the function table and the parsed-formula cache. */
export class Evaluator {
  /** @internal */
  readonly fns: Map<string, Builtin>;
  /** @internal */
  readonly now: () => Date;
  private readonly cache = new Map<string, Parsed>();

  constructor(builtins: Record<string, Builtin>, options: EvaluatorOptions = {}) {
    this.fns = new Map(Object.entries(builtins));
    for (const [k, f] of Object.entries(options.functions ?? {})) this.fns.set(k.toUpperCase(), f);
    this.now = options.now ?? (() => new Date());
  }

  /** Add or replace a function on this evaluator. */
  registerFunction(name: string, impl: FormulaFunction): this {
    this.fns.set(name.toUpperCase(), impl);
    return this;
  }

  /** @internal Parse with caching (never throws). */
  parsed(text: string): Parsed {
    let p = this.cache.get(text);
    if (p === undefined) {
      try {
        p = { ast: parseFormula(text), error: undefined, refs: undefined, calls: undefined };
      } catch (e) {
        p = { ast: undefined, error: e as Error, refs: [], calls: [] };
      }
      if (this.cache.size > 100_000) this.cache.clear();
      this.cache.set(text, p);
    }
    if (p.refs === undefined && p.ast) {
      p.refs = [];
      p.calls = [];
      collect(p.ast, p);
    }
    return p;
  }

  /** Evaluate every formula of the workbook (or of `options.sheets`). */
  evaluate(wb: Workbook, options: EvaluateOptions = {}): EvaluationReport {
    const run = new Run(this, wb, options);
    return run.go(() => {
      const want = options.sheets?.map((s) => (typeof s === 'string' ? wb.getWorksheet(s) : s));
      const roots: number[] = [];
      for (const si of run.sheets) {
        if (want && !want.includes(si.ws)) continue;
        const end = si.first[si.rows.length] ?? 0;
        for (let id = si.first[0] ?? 0; id < end; id++) roots.push(id);
      }
      run.solve(roots);
      return run.report(options.writeResults !== false);
    });
  }

  /** Evaluate one cell (and what it depends on); returns its value. */
  evaluateCell(
    wb: Workbook,
    sheet: string | Worksheet,
    address: string,
    options: EvaluateOptions = {},
  ): FormulaResult | null {
    const run = new Run(this, wb, options);
    return run.go(() => {
      const si = run.sheetOf(sheet);
      if (!si) throw new RangeError(`Unknown worksheet: ${String(sheet)}`);
      const [row, col] = cellAt(address);
      const id = run.find(si, row, col);
      if (id >= 0) run.solve([id]);
      run.report(options.writeResults === true);
      const v = run.cell(si.ws, row, col);
      return v;
    });
  }

  /** Evaluate formula text against the workbook (nothing is written). */
  evaluateFormula(
    wb: Workbook,
    text: string,
    options: EvaluateFormulaOptions = {},
  ): FormulaResult | FormulaResult[][] {
    const run = new Run(this, wb, options);
    return run.go(() => {
      const si = options.sheet !== undefined ? run.sheetOf(options.sheet) : run.sheets[0];
      if (!si) throw new RangeError(`Unknown worksheet: ${String(options.sheet)}`);
      const p = this.parsed(text.replace(/^=/, ''));
      if (!p.ast) throw p.error;
      const roots: number[] = [];
      run.depsOf(p, si, roots, 0);
      run.solve(roots);
      const [ar, ac] = options.address ? cellAt(options.address) : [-1, -1];
      const ctx = new EvalContext(run, si, ar, ac, options.array === true);
      let v = ctx.safe(p.ast);
      if (isRef(v) && (v.rows > 1 || v.cols > 1) && ar < 0) v = toArr(v);
      if (isArr(v) && (v.rows > 1 || v.cols > 1)) {
        const out: FormulaResult[][] = [];
        for (let i = 0; i < v.rows; i++) {
          const r: FormulaResult[] = [];
          for (let j = 0; j < v.cols; j++) r.push(final(v.at(i, j)));
          out.push(r);
        }
        return out;
      }
      return final(isRef(v) ? intersect(v, ctx) : isArr(v) ? v.data[0]! : v);
    });
  }
}

function final(v: FormulaScalar): FormulaResult {
  if (v === null) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? round15(v) : NUM;
  return v;
}

function richText(rt: RichText): string {
  let s = '';
  for (const r of rt.richText) s += r.text;
  return s;
}

/** One evaluation pass over a workbook. */
class Run implements CellReader {
  readonly ev: Evaluator;
  readonly wb: Workbook;
  readonly #opts: EvaluateOptions;
  readonly now: () => Date;
  readonly sheets: SheetIndex[] = [];
  readonly #byName = new Map<string, SheetIndex>();
  readonly #byWs = new Map<Worksheet, SheetIndex>();
  #names: Map<string, DefinedName[]> | undefined;
  #hasArrays = false;
  // nodes
  readonly #nS: SheetIndex[] = [];
  readonly #nR: number[] = [];
  readonly #nC: number[] = [];
  readonly #nF: Formula[] = [];
  readonly #nP: Parsed[] = [];
  /** Dense column caches (see column()); bypassed while iterating circular references. */
  readonly #cols: Map<string, FormulaScalar[]> = new Map();
  #iter = false;
  #val: FormulaValue[] = [];
  #state!: Uint8Array; // 0 pending, 1 evaluating, 2 done
  readonly #order: number[] = [];
  readonly #circ: number[] = [];
  readonly #unsupported: Map<string, string> = new Map();
  readonly #errors: { address: string; message: string }[] = [];

  constructor(ev: Evaluator, wb: Workbook, opts: EvaluateOptions | EvaluateFormulaOptions) {
    this.ev = ev;
    this.wb = wb;
    this.#opts = opts as EvaluateOptions;
    this.now = opts.now ?? ev.now;
    const sheets = wb.worksheets;
    for (let s = 0; s < sheets.length; s++) {
      const ws = sheets[s]!;
      const si = new SheetIndex(ws, s);
      this.sheets.push(si);
      this.#byName.set(ws.name.toLowerCase(), si);
      this.#byWs.set(ws, si);
      const rows = ws.rowData;
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (row === undefined) continue;
        const { n, kind } = row;
        let seen = false;
        for (let i = 0; i < n; i++) {
          if (kind[i] !== K_FORMULA) continue;
          if (!seen) {
            seen = true;
            si.rows.push(r);
            si.first.push(this.#nF.length);
          }
          const f = row.obj![i] as Formula;
          const c = row.col[i]!;
          if (f.array) {
            const a = decodeRangeSafe(f.array);
            if (a && (a[2] > a[0] || a[3] > a[1])) {
              si.arrays = si.arrays ?? [];
              si.arrays.push([a[0], a[1], a[2], a[3], this.#nF.length]);
              this.#hasArrays = true;
            }
          }
          this.#nS.push(si);
          this.#nR.push(r);
          this.#nC.push(c);
          this.#nF.push(f);
        }
      }
      si.first.push(this.#nF.length);
    }
    this.#state = new Uint8Array(this.#nF.length);
  }

  go<T>(f: () => T): T {
    const old = setDate1904(this.wb.date1904);
    try {
      return f();
    } finally {
      setDate1904(old);
    }
  }

  sheetOf(s: string | Worksheet): SheetIndex | undefined {
    return typeof s === 'string' ? this.#byName.get(s.toLowerCase()) : this.#byWs.get(s);
  }

  /** Sheet for a reference: explicit name, or the host sheet. */
  sheetFor(name: string | undefined, host: SheetIndex): SheetIndex | undefined {
    return name === undefined ? host : this.#byName.get(name.toLowerCase());
  }

  /** Node id of the formula cell at (r, c), or -1. */
  find(si: SheetIndex, r: number, c: number): number {
    const rows = si.rows;
    let lo = 0;
    let hi = rows.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      const v = rows[mid]!;
      if (v < r) lo = mid + 1;
      else if (v > r) hi = mid - 1;
      else {
        let a = si.first[mid]!;
        let b = si.first[mid + 1]! - 1;
        const nC = this.#nC;
        while (a <= b) {
          const m = (a + b) >>> 1;
          const x = nC[m]!;
          if (x < c) a = m + 1;
          else if (x > c) b = m - 1;
          else return m;
        }
        return -1;
      }
    }
    return -1;
  }

  #addr(id: number): string {
    return sheetPrefix(this.#nS[id]!.ws.name) + colName(this.#nC[id]!) + (this.#nR[id]! + 1);
  }

  definedName(name: string, sheet: string | undefined, host: SheetIndex): DefinedName | undefined {
    if (!this.#names) {
      this.#names = new Map();
      for (const d of this.wb.definedNames) {
        const k = d.name.toLowerCase();
        const list = this.#names.get(k);
        if (list) list.push(d);
        else this.#names.set(k, [d]);
      }
    }
    const list = this.#names.get(name.toLowerCase());
    if (!list) return undefined;
    const scope = sheet === undefined ? host : this.#byName.get(sheet.toLowerCase());
    let global: DefinedName | undefined;
    for (const d of list) {
      if (d.localSheetId === undefined) global = d;
      else if (scope && d.localSheetId === scope.i) return d;
    }
    return sheet === undefined ? global : undefined;
  }

  /** Push the formula nodes `p` statically depends on (evaluated on host sheet `si`). */
  depsOf(p: Parsed, si: SheetIndex, out: number[], depth: number): void {
    for (const ref of p.refs!) {
      if (ref.length === 2) {
        if (depth > 16) continue;
        const d = this.definedName(ref[1], ref[0], si);
        if (d) this.depsOf(this.ev.parsed(d.value), si, out, depth + 1);
        continue;
      }
      const t = this.sheetFor(ref[0], si);
      if (!t) continue;
      const [, top, left, bottom, right] = ref;
      const rows = t.rows;
      let lo = 0;
      let hi = rows.length;
      while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (rows[mid]! < top) lo = mid + 1;
        else hi = mid;
      }
      const nC = this.#nC;
      for (let k = lo; k < rows.length && rows[k]! <= bottom; k++) {
        const end = t.first[k + 1]!;
        for (let id = t.first[k]!; id < end; id++) {
          const c = nC[id]!;
          if (c > right) break;
          if (c >= left && this.#state[id] !== 2) out.push(id);
        }
      }
      if (t.arrays) {
        for (const a of t.arrays) {
          if (a[0]! <= bottom && a[2]! >= top && a[1]! <= right && a[3]! >= left) out.push(a[4]!);
        }
      }
    }
  }

  /** Evaluate the given nodes and everything they depend on, dependencies first. */
  solve(roots: number[]): void {
    const N = this.#nF.length;
    if (N === 0) return;
    const index = new Int32Array(N).fill(-1);
    const low = new Int32Array(N);
    const onStack = new Uint8Array(N);
    const self = new Uint8Array(N);
    const stack: number[] = [];
    const fNode: number[] = [];
    const fDeps: number[][] = [];
    const fPos: number[] = [];
    let counter = 0;
    const push = (v: number): void => {
      index[v] = low[v] = counter++;
      stack.push(v);
      onStack[v] = 1;
      const deps: number[] = [];
      if (this.#state[v] === 0) {
        const si = this.#nS[v]!;
        const p = this.ev.parsed(this.#nF[v]!.formula);
        this.#nP[v] = p;
        this.depsOf(p, si, deps, 0);
        for (const name of p.calls!) {
          if (!this.ev.fns.has(name) && !this.#unsupported.has(name)) {
            this.#unsupported.set(name, this.#addr(v));
            this.#opts.onUnsupported?.(name, this.#addr(v));
          }
        }
        for (const d of deps) if (d === v) self[v] = 1;
      }
      fNode.push(v);
      fDeps.push(deps);
      fPos.push(0);
    };
    for (const root of roots) {
      if (index[root]! >= 0 || this.#state[root] === 2) continue;
      push(root);
      while (fNode.length > 0) {
        const top = fNode.length - 1;
        const v = fNode[top]!;
        const deps = fDeps[top]!;
        const pos = fPos[top]!;
        if (pos < deps.length) {
          fPos[top] = pos + 1;
          const w = deps[pos]!;
          if (index[w]! < 0) {
            if (this.#state[w] === 2) continue;
            push(w);
          } else if (onStack[w]) {
            if (index[w]! < low[v]!) low[v] = index[w]!;
          }
          continue;
        }
        fNode.pop();
        fDeps.pop();
        fPos.pop();
        if (fNode.length > 0) {
          const parent = fNode[fNode.length - 1]!;
          if (low[v]! < low[parent]!) low[parent] = low[v]!;
        }
        if (low[v] === index[v]) {
          // pop one strongly-connected component
          let w: number;
          const scc: number[] = [];
          do {
            w = stack.pop()!;
            onStack[w] = 0;
            scc.push(w);
          } while (w !== v);
          if (scc.length === 1 && !self[v]) {
            if (this.#state[v] === 0) this.#evalNode(v);
          } else this.#cycle(scc.reverse());
        }
      }
    }
  }

  #cycle(scc: number[]): void {
    const iters = this.#opts.maxIterations ?? 0;
    for (const id of scc) {
      this.#circ.push(id);
      this.#state[id] = 2;
      this.#order.push(id);
      const r = this.#nF[id]!.result;
      this.#val[id] = iters > 0 ? (typeof r === 'number' ? r : 0) : REF;
    }
    if (iters <= 0) return;
    const maxChange = this.#opts.maxChange ?? 0.001;
    this.#iter = true;
    for (let k = 0; k < iters; k++) {
      let delta = 0;
      for (const id of scc) {
        const old = this.#val[id];
        const v = this.#compute(id);
        this.#val[id] = v;
        if (typeof v === 'number' && typeof old === 'number')
          delta = Math.max(delta, Math.abs(v - old));
        else if (v !== old) delta = Infinity;
      }
      if (delta <= maxChange) break;
    }
    this.#iter = false;
  }

  #evalNode(id: number): void {
    this.#state[id] = 1;
    const v = this.#compute(id);
    this.#val[id] = v;
    this.#state[id] = 2;
    this.#order.push(id);
  }

  #compute(id: number): FormulaValue {
    const f = this.#nF[id]!;
    const p = this.#nP[id] ?? this.ev.parsed(f.formula);
    if (!p.ast) {
      this.#errors.push({ address: this.#addr(id), message: p.error!.message });
      return NAME;
    }
    const ctx = new EvalContext(
      this,
      this.#nS[id]!,
      this.#nR[id]!,
      this.#nC[id]!,
      f.array !== undefined,
    );
    let v: FormulaValue;
    try {
      v = ctx.ev(p.ast);
    } catch (e) {
      if (isError(e)) v = e;
      else {
        this.#errors.push({ address: this.#addr(id), message: String((e as Error)?.message ?? e) });
        return VALUE;
      }
    }
    if (isRef(v)) v = ctx.array ? toArr(v) : intersect(v, ctx);
    if (isArr(v)) {
      if (!ctx.array) return final(v.data[0]!);
      return new FormulaArray(v.rows, v.cols, v.data.map(final));
    }
    return final(v);
  }

  /** Current value of formula node `id` (evaluating it on demand). */
  #nodeValue(id: number): FormulaScalar {
    const st = this.#state[id];
    if (st === 1) {
      // reached a node being evaluated through a dynamic reference: circular
      this.#circ.push(id);
      return REF;
    }
    if (st === 0) this.#evalNode(id);
    const v = this.#val[id];
    return isArr(v) ? v.data[0]! : (v as FormulaScalar);
  }

  #slot(ws: Worksheet, row: RowData, i: number): FormulaScalar {
    switch (row.kind[i]) {
      case K_NUM:
      case K_DATE:
        return row.num[i]!;
      case K_STR:
        return row.obj![i] as string;
      case K_BOOL:
        return row.num[i] !== 0;
      case K_ERR:
        return errOf(row.obj![i] as ErrorCode);
      case K_FORMULA: {
        const si = this.#byWs.get(ws);
        const id = si ? this.find(si, row.index, row.col[i]!) : -1;
        if (id >= 0) return this.#nodeValue(id);
        const r = (row.obj![i] as Formula).result;
        return r === undefined ? 0 : r;
      }
      case K_RICH:
        return richText(row.obj![i] as RichText);
      case K_LINK:
        return (row.obj![i] as Hyperlink).text;
      default:
        return null;
    }
  }

  /** Value from an array formula's result when (r, c) lies in its area (non-master cell). */
  #spill(ws: Worksheet, r: number, c: number): FormulaScalar | undefined {
    const si = this.#byWs.get(ws);
    if (!si?.arrays) return undefined;
    for (const a of si.arrays) {
      if (r >= a[0]! && r <= a[2]! && c >= a[1]! && c <= a[3]!) {
        const id = a[4]!;
        if (this.#nR[id] === r && this.#nC[id] === c) return undefined;
        this.#nodeValue(id);
        const v = this.#val[id];
        if (!isArr(v)) return r === a[0] && c === a[1] ? (v as FormulaScalar) : NA;
        const i = r - a[0]!;
        const j = c - a[1]!;
        // a smaller result repeats a 1-wide dimension, else #N/A (as Excel fills array areas)
        const ii = v.rows === 1 ? 0 : i;
        const jj = v.cols === 1 ? 0 : j;
        return ii < v.rows && jj < v.cols ? v.at(ii, jj) : NA;
      }
    }
    return undefined;
  }

  cell(ws: Worksheet, r: number, c: number): FormulaScalar {
    if (this.#hasArrays) {
      const s = this.#spill(ws, r, c);
      if (s !== undefined) return s;
    }
    const row = ws.rowData[r];
    if (row === undefined) return null;
    const i = row.find(c);
    return i < 0 ? null : this.#slot(ws, row, i);
  }

  column(ref: FormulaRange): readonly FormulaScalar[] | undefined {
    const ws = ref.sheet;
    const si = this.#byWs.get(ws);
    const end = Math.min(ref.bottom, ws.rowData.length - 1);
    if (ref.left !== ref.right || this.#iter || !si || end - ref.top < 16) return undefined;
    const key = `${si.i}:${ref.left}`;
    let vec = this.#cols.get(key);
    if (!vec) {
      vec = [];
      this.#cols.set(key, vec);
    }
    let missing = false;
    for (let r = ref.top; r <= end; r++) {
      if (vec[r] === undefined) {
        missing = true;
        break;
      }
    }
    if (missing) {
      const v = vec;
      this.#sparse(new FormulaRange(this, ws, ref.top, ref.left, end, ref.left), (x, r) => {
        v[r] = x;
      });
      for (let r = ref.top; r <= end; r++) if (v[r] === undefined) v[r] = null;
    }
    return vec as FormulaScalar[];
  }

  each(ref: FormulaRange, cb: (v: FormulaScalar, row: number, col: number) => void): void {
    const vec = this.column(ref);
    if (vec) {
      const end = Math.min(ref.bottom, ref.sheet.rowData.length - 1);
      for (let r = ref.top; r <= end; r++) {
        const x = vec[r]!;
        if (x !== null) cb(x, r, ref.left);
      }
      return;
    }
    this.#sparse(ref, cb);
  }

  #sparse(ref: FormulaRange, cb: (v: FormulaScalar, row: number, col: number) => void): void {
    const ws = ref.sheet;
    const si = this.#byWs.get(ws);
    // cells of array formula areas go through cell() to see the recomputed values
    const slow = si?.arrays !== undefined;
    const rows = ws.rowData;
    const end = Math.min(ref.bottom, rows.length - 1);
    const left = ref.left;
    const right = ref.right;
    // formula cells: walk the sheet's node ids alongside the slots (same order)
    const fr = si ? si.rows : [];
    const first = si ? si.first : [];
    const nC = this.#nC;
    let k = 0;
    let hi = fr.length;
    while (k < hi) {
      const mid = (k + hi) >>> 1;
      if (fr[mid]! < ref.top) k = mid + 1;
      else hi = mid;
    }
    for (let r = ref.top; r <= end; r++) {
      const row = rows[r];
      if (row === undefined) continue;
      const n = row.n;
      if (n === 0) continue;
      let id = -1;
      while (k < fr.length && fr[k]! < r) k++;
      if (fr[k] === r) {
        id = first[k]!;
        while (nC[id]! < left) id++;
      }
      const col = row.col;
      let i = left <= col[0]! ? 0 : row.find(left);
      if (i < 0) i = -i - 1;
      const kind = row.kind;
      for (; i < n; i++) {
        const c = col[i]!;
        if (c > right) break;
        const t = kind[i];
        if (t === K_EMPTY) continue;
        cb(
          slow
            ? this.cell(ws, r, c)
            : t === K_FORMULA && id >= 0
              ? this.#nodeValue(id++)
              : this.#slot(ws, row, i),
          r,
          c,
        );
      }
    }
  }

  report(write: boolean): EvaluationReport {
    const results = write ? undefined : new Map<string, FormulaResult>();
    for (const id of this.#order) {
      const v = this.#val[id];
      const res = isArr(v) ? (v.data[0] as FormulaResult) : (v as FormulaResult);
      if (results) {
        results.set(this.#addr(id), res);
        continue;
      }
      const ws = this.#nS[id]!.ws;
      const r = this.#nR[id]!;
      const c = this.#nC[id]!;
      const row = ws.rowData[r]!;
      const i = row.find(c);
      if (i < 0 || row.kind[i] !== K_FORMULA) continue;
      const f = row.obj![i] as Formula;
      row.obj![i] = { ...f, result: res };
      if (isArr(v) && f.array) this.#writeSpill(ws, id, v);
    }
    const circ = [...new Set(this.#circ)].map((id) => this.#addr(id));
    return {
      evaluated: this.#order.length,
      circular: circ,
      unsupported: [...this.#unsupported.keys()].sort(),
      errors: this.#errors,
      ...(results ? { results } : {}),
    };
  }

  #writeSpill(ws: Worksheet, id: number, v: FormulaArray): void {
    const a = decodeRangeSafe(this.#nF[id]!.array!);
    if (!a || (a[2] - a[0] + 1) * (a[3] - a[1] + 1) > 1e6) return;
    for (let r: number = a[0]; r <= a[2]; r++) {
      for (let c: number = a[1]; c <= a[3]; c++) {
        if (r === this.#nR[id] && c === this.#nC[id]) continue;
        const ii = v.rows === 1 ? 0 : r - a[0];
        const jj = v.cols === 1 ? 0 : c - a[1];
        const x = ii < v.rows && jj < v.cols ? v.at(ii, jj) : NA;
        const row = ws.rowData[r];
        const i = row ? row.find(c) : -1;
        if (row && i >= 0 && row.kind[i] === K_FORMULA) continue;
        const xf = row && i >= 0 ? row.style[i]! : 0;
        if (typeof x === 'number') ws.setCellRaw(r, c, K_NUM, x, undefined, xf);
        else if (typeof x === 'string') ws.setCellRaw(r, c, K_STR, 0, x, xf);
        else if (typeof x === 'boolean') ws.setCellRaw(r, c, K_BOOL, x ? 1 : 0, undefined, xf);
        else if (isError(x)) ws.setCellRaw(r, c, K_ERR, 0, x.error, xf);
      }
    }
  }
}

/** "B3" → [2, 1] (0-based row, col). */
function cellAt(a: string): [number, number] {
  const m = CELL_RE.exec(a);
  if (!m) throw new RangeError(`Invalid cell address: "${a}"`);
  return [+m[2]! - 1, colIndex(m[1]!)];
}

function colName(c: number): string {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function decodeRangeSafe(ref: string): [number, number, number, number] | undefined {
  try {
    const [p, q = p] = ref.split(':') as [string, string?];
    const [t, l] = cellAt(p);
    const [b, r] = cellAt(q);
    return [Math.min(t, b), Math.min(l, r), Math.max(t, b), Math.max(l, r)];
  } catch {
    return undefined;
  }
}

// ---- interpreter ----------------------------------------------------------------------------

function pow(a: number, b: number): FormulaScalar {
  if (a === 0) return b === 0 ? NUM : b < 0 ? DIV0 : 0;
  const r = a ** b;
  return Number.isFinite(r) ? r : NUM;
}

function op2(op: BinaryOp, x: FormulaScalar, y: FormulaScalar): FormulaScalar {
  switch (op) {
    case '&': {
      const a = toStr(x);
      if (typeof a !== 'string') return a;
      const b = toStr(y);
      if (typeof b !== 'string') return b;
      return a.length + b.length > 32767 ? VALUE : a + b;
    }
    case '=':
    case '<>':
    case '<':
    case '>':
    case '<=':
    case '>=': {
      if (isError(x)) return x;
      if (isError(y)) return y;
      const c = compare(x, y);
      return op === '='
        ? c === 0
        : op === '<>'
          ? c !== 0
          : op === '<'
            ? c < 0
            : op === '>'
              ? c > 0
              : op === '<='
                ? c <= 0
                : c >= 0;
    }
    default: {
      const a = toNum(x);
      if (typeof a !== 'number') return a;
      const b = toNum(y);
      if (typeof b !== 'number') return b;
      let r: number;
      switch (op) {
        case '+':
          r = a + b;
          break;
        case '-':
          r = a - b;
          break;
        case '*':
          r = a * b;
          break;
        case '/':
          if (b === 0) return DIV0;
          r = a / b;
          break;
        default:
          return pow(a, b);
      }
      return Number.isFinite(r) ? r : NUM;
    }
  }
}

const neg = (x: FormulaScalar): FormulaScalar => {
  const n = toNum(x);
  return typeof n === 'number' ? -n : n;
};
const pct = (x: FormulaScalar): FormulaScalar => {
  const n = toNum(x);
  return typeof n === 'number' ? n / 100 : n;
};

/** Evaluation context of one formula (implements FunctionContext for the functions). */
export class EvalContext implements FunctionContext {
  /** @internal */
  readonly #run: Run;
  /** @internal */
  readonly #si: SheetIndex;
  readonly row: number;
  readonly col: number;
  readonly array: boolean;
  #depth = 0;

  constructor(run: Run, si: SheetIndex, row: number, col: number, array: boolean) {
    this.#run = run;
    this.#si = si;
    this.row = row;
    this.col = col;
    this.array = array;
  }

  get workbook(): Workbook {
    return this.#run.wb;
  }
  get sheet(): Worksheet {
    return this.#si.ws;
  }
  get date1904(): boolean {
    return this.#run.wb.date1904;
  }
  now(): Date {
    return this.#run.now();
  }
  /** The same context in array mode. */
  arrayMode(): EvalContext {
    if (this.array) return this;
    const c = new EvalContext(this.#run, this.#si, this.row, this.col, true);
    c.#depth = this.#depth;
    return c;
  }

  /** Evaluate, turning thrown CellErrors into values. */
  safe(n: FormulaNode): FormulaValue {
    try {
      return this.ev(n);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  }

  #ref(sheet: string | undefined, t: number, l: number, b: number, r: number): FormulaValue {
    const si = this.#run.sheetFor(sheet, this.#si);
    return si ? new FormulaRange(this.#run, si.ws, t, l, b, r) : REF;
  }

  /** Evaluate a node; references stay references. */
  ev(n: FormulaNode): FormulaValue {
    switch (n.type) {
      case 'number':
      case 'string':
      case 'boolean':
        return n.value;
      case 'error':
        return errOf(n.value);
      case 'cell':
        return this.#ref(n.sheet, n.row, n.col, n.row, n.col);
      case 'range':
        return this.#ref(n.sheet, n.top, n.left, n.bottom, n.right);
      case 'missing':
        return null;
      case 'name':
        return this.#name(n.name, n.sheet);
      case 'array': {
        const data: FormulaScalar[] = [];
        for (const row of n.rows) for (const x of row) data.push(isError(x) ? errOf(x.error) : x);
        return new FormulaArray(n.rows.length, n.rows[0]!.length, data);
      }
      case 'unary': {
        if (n.op === '+') return this.ev(n.operand);
        if (n.op === '@') {
          const v = this.ev(n.operand);
          return isRef(v) ? intersect(v, this) : isArr(v) ? v.data[0]! : v;
        }
        return this.#map(this.operand(n.operand), neg);
      }
      case 'percent':
        return this.#map(this.operand(n.operand), pct);
      case 'binary': {
        if (n.op === ':') return this.#union(n.left, n.right);
        const op = n.op;
        const a = this.operand(n.left);
        const b = this.operand(n.right);
        if (!isArr(a) && !isArr(b)) return op2(op, a as FormulaScalar, b as FormulaScalar);
        return broadcast([a, b], (xs) => op2(op, xs[0]!, xs[1]!), this);
      }
      case 'call':
        return this.#call(n.name, n.args);
    }
  }

  #map(v: FormulaValue, f: (x: FormulaScalar) => FormulaScalar): FormulaValue {
    return isArr(v) ? broadcast([v], (xs) => f(xs[0]!), this) : f(v as FormulaScalar);
  }

  /** Evaluate an operator operand: references are dereferenced (intersected, or arrays). */
  operand(n: FormulaNode): FormulaValue {
    if (n.type === 'cell') {
      const si = n.sheet === undefined ? this.#si : this.#run.sheetFor(n.sheet, this.#si);
      return si ? this.#run.cell(si.ws, n.row, n.col) : REF;
    }
    const v = this.ev(n);
    if (isRef(v)) return this.array ? toArr(v) : intersect(v, this);
    return v;
  }

  /** Evaluate to a single value (references intersect, arrays give their first item). */
  scalar(n: FormulaNode): FormulaScalar {
    const v = this.operand(n);
    return isArr(v) ? v.data[0]! : (v as FormulaScalar);
  }

  #union(a: FormulaNode, b: FormulaNode): FormulaValue {
    const x = this.ev(a);
    const y = this.ev(b);
    if (isError(x)) return x;
    if (isError(y)) return y;
    if (!isRef(x) || !isRef(y) || x.sheet !== y.sheet) return VALUE;
    return new FormulaRange(
      this.#run,
      x.sheet,
      Math.min(x.top, y.top),
      Math.min(x.left, y.left),
      Math.max(x.bottom, y.bottom),
      Math.max(x.right, y.right),
    );
  }

  #name(name: string, sheet: string | undefined): FormulaValue {
    const d = this.#run.definedName(name, sheet, this.#si);
    if (!d) return NAME;
    if (this.#depth > 16) return NAME;
    const p = this.#run.ev.parsed(d.value);
    if (!p.ast) return /#REF!/i.test(d.value) ? REF : NAME;
    this.#depth++;
    try {
      return this.ev(p.ast);
    } finally {
      this.#depth--;
    }
  }

  #call(name: string, args: readonly FormulaNode[]): FormulaValue {
    const fn = this.#run.ev.fns.get(name);
    if (fn === undefined) return NAME;
    try {
      if (typeof fn === 'function') {
        const vs: FormulaValue[] = new Array(args.length);
        for (let i = 0; i < args.length; i++) vs[i] = this.ev(args[i]!);
        const r = fn(vs, this);
        return r === undefined ? null : r;
      }
      if ('lazy' in fn) return fn.lazy(args, this);
      const c: EvalContext = this.arrayMode();
      const vs = args.map((a) => c.ev(a));
      return fn.array(vs, c);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  }
}
