// Formula evaluation engine: dependency graph over the formula cells of a workbook, Tarjan's
// strongly-connected components (iterative) to get a dependencies-first order and detect
// circular references, and a tree-walking interpreter over cached ASTs.

import type {
  CellError,
  CellValueInput,
  DefinedName,
  ErrorCode,
  Formula,
  Hyperlink,
  RichText,
} from '../core/types.js';
import type { Workbook } from '../core/workbook.js';
import type { RowData, Worksheet } from '../core/worksheet.js';
import { isDynamic } from '../xlsx/common/spill.js';
import type { Table } from '../xlsx/features/tables.js';
import {
  broadcast,
  type FormulaFunction,
  type FunctionContext,
  intersect,
  toArr,
} from './helpers.js';
import { type BinaryOp, CELL_RE, colIndex, type FormulaNode, parseFormula } from './parser.js';

/** A structured reference node. */
type StructNode = Extract<FormulaNode, { type: 'structured' }>;

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
  SPILL,
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

/** A built-in function table entry (eager function, lazy, or array-mode). */
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
        | readonly [StructNode]
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
    case 'structured':
      p.refs!.push([n]);
      return;
    case 'spill':
      p.refs!.push([n.sheet, n.row, n.col, n.row, n.col]);
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
  /** Previous spill areas of dynamic formulas: their stored values are stale. */
  stale: number[][] | undefined = undefined;
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

  /**
   * @internal Run passes until the spill areas of dynamic-array formulas are stable: a pass assumes the
   * areas found by the previous one (first: the areas stored in the file), so formulas reading
   * spilled cells see the right values. Usually one or two passes.
   */
  passes(wb: Workbook, options: EvaluateOptions, roots: (run: Run) => number[]): Run {
    let hints: Map<string, number[]> | undefined;
    for (let pass = 0; ; pass++) {
      const run = new Run(this, wb, options, hints);
      const next = new Map(hints);
      const changed = run.go(() => {
        run.solve(roots(run));
        return run.settle(next);
      });
      if (!changed || pass >= 4) return run;
      hints = next;
    }
  }

  /**
   * Evaluate the workbook (storing results) and keep the dependency graph for incremental
   * recalculation: `track(wb).set('Sheet1', 'B2', 5)` recalculates only what depends on B2.
   */
  track(wb: Workbook, options: Omit<EvaluateOptions, 'sheets' | 'writeResults'> = {}): Recalc {
    return new Recalc(this, wb, options);
  }

  /** Evaluate every formula of the workbook (or of `options.sheets`). */
  evaluate(wb: Workbook, options: EvaluateOptions = {}): EvaluationReport {
    const want = options.sheets?.map((s) => (typeof s === 'string' ? wb.getWorksheet(s) : s));
    const run = this.passes(wb, options, (run) => {
      const roots: number[] = [];
      for (const si of run.sheets) {
        if (want && !want.includes(si.ws)) continue;
        const end = si.first[si.rows.length] ?? 0;
        for (let id = si.first[0] ?? 0; id < end; id++) roots.push(id);
      }
      return roots;
    });
    return run.go(() => run.report(options.writeResults !== false));
  }

  /** Evaluate one cell (and what it depends on); returns its value. */
  evaluateCell(
    wb: Workbook,
    sheet: string | Worksheet,
    address: string,
    options: EvaluateOptions = {},
  ): FormulaResult | null {
    const [row, col] = cellAt(address);
    if (!new Run(this, wb, options).sheetOf(sheet)) {
      throw new RangeError(`Unknown worksheet: ${String(sheet)}`);
    }
    const run = this.passes(wb, options, (run) => {
      const id = run.find(run.sheetOf(sheet)!, row, col);
      return id >= 0 ? [id] : [];
    });
    return run.go(() => {
      run.report(options.writeResults === true);
      return run.cell(run.sheetOf(sheet)!.ws, row, col);
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

/**
 * A workbook kept calculated: `set()` / `update()` recalculate only the formulas downstream of
 * the changed cells (see Evaluator.track). Structural changes — a formula added, removed or
 * edited, cells moved — need `recalc()` (set() does it by itself when it writes or replaces a
 * formula).
 */
export class Recalc {
  readonly #ev: Evaluator;
  readonly #wb: Workbook;
  readonly #opts: EvaluateOptions;
  #run!: Run;
  /** Report of the last full evaluation. */
  report!: EvaluationReport;

  constructor(ev: Evaluator, wb: Workbook, options: EvaluateOptions) {
    this.#ev = ev;
    this.#wb = wb;
    this.#opts = { ...options, writeResults: true };
    this.recalc();
  }

  /** Evaluate the whole workbook again and rebuild the dependency graph. */
  recalc(): EvaluationReport {
    const run = this.#ev.passes(this.#wb, this.#opts, (r) => {
      const roots: number[] = [];
      for (const si of r.sheets) {
        const end = si.first[si.rows.length] ?? 0;
        for (let id = si.first[0] ?? 0; id < end; id++) roots.push(id);
      }
      return roots;
    });
    this.#run = run;
    this.report = run.go(() => run.report(true));
    return this.report;
  }

  /**
   * Set a cell and recalculate what depends on it. Returns the addresses of the formula cells
   * recalculated ("Sheet1!C3"); writing or replacing a formula recalculates everything.
   */
  set(sheet: string | Worksheet, address: string, value: CellValueInput): string[] {
    const ws = typeof sheet === 'string' ? this.#wb.getWorksheet(sheet) : sheet;
    if (!ws) throw new RangeError(`Unknown worksheet: ${String(sheet)}`);
    ws.setCell(address, value);
    return this.update([{ sheet: ws, address }]);
  }

  /** Recalculate after the given cells were changed directly (e.g. with ws.setCell). */
  update(cells: readonly { sheet: string | Worksheet; address: string }[]): string[] {
    const run = this.#run;
    const list: [SheetIndex, number, number][] = [];
    for (const { sheet, address } of cells) {
      const si = run.sheetOf(sheet);
      if (!si) throw new RangeError(`Unknown worksheet: ${String(sheet)}`);
      const [r, c] = cellAt(address);
      list.push([si, r, c]);
    }
    const done = run.go(() => run.update(list));
    if (done) return done;
    this.recalc();
    return this.#run.addresses();
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
  #inOrder!: Uint8Array;
  /** Rectangles each node reads ([sheet index, top, left, bottom, right]...), for update(). */
  #rects: (number[] | undefined)[] = [];
  readonly #order: number[] = [];
  readonly #circ: number[] = [];
  readonly #unsupported: Map<string, string> = new Map();
  readonly #errors: { address: string; message: string }[] = [];
  /** Dynamic-array formula nodes: their spill area in this pass ([t, l, b, r, id]). */
  readonly #dyn = new Map<number, number[]>();
  /** The area a dynamic formula had in the file / model (cells there are its stale values). */
  readonly #old = new Map<number, [number, number, number, number]>();

  constructor(
    ev: Evaluator,
    wb: Workbook,
    opts: EvaluateOptions | EvaluateFormulaOptions,
    hints?: Map<string, number[]>,
  ) {
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
          if (isDynamic(f)) {
            const id = this.#nF.length;
            const old = f.array ? decodeRangeSafe(f.array) : undefined;
            if (old && old[0] === r && old[1] === c) {
              this.#old.set(id, old);
              si.stale = si.stale ?? [];
              si.stale.push(old);
            }
            const h = hints?.get(`${s}:${r}:${c}`) ?? old ?? [r, c, r, c];
            const e = h[0] === r && h[1] === c ? [r, c, h[2]!, h[3]!, id] : [r, c, r, c, id];
            si.arrays = si.arrays ?? [];
            si.arrays.push(e);
            this.#dyn.set(id, e);
            this.#hasArrays = true;
          } else if (f.array) {
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
    this.#inOrder = new Uint8Array(this.#nF.length);
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

  /**
   * Push the formula nodes `p` statically depends on (evaluated on host sheet `si`, in cell
   * (`row`, `col`) when known: `[@Col]` then depends on that row only).
   */
  depsOf(p: Parsed, si: SheetIndex, out: number[], depth: number, row = -1, col = -1): void {
    for (const ref of p.refs!) {
      if (ref.length === 2) {
        if (depth > 16) continue;
        const d = this.definedName(ref[1], ref[0], si);
        if (d) this.depsOf(this.ev.parsed(d.value), si, out, depth + 1, row, col);
        continue;
      }
      if (ref.length === 1) {
        const a = this.tableRef(ref[0], si, row, col, true);
        if (!isError(a)) this.#depsIn(a[0], a[1], a[2], a[3], a[4], out);
        continue;
      }
      const t = this.sheetFor(ref[0], si);
      if (t) this.#depsIn(t, ref[1], ref[2], ref[3], ref[4], out);
    }
  }

  #tables: Map<string, { si: SheetIndex; t: Table }> | undefined;

  /**
   * The range of a structured reference: [sheet, top, left, bottom, right], or an error
   * (#REF! for an unknown table / column / missing header or totals row, #VALUE! for `@`
   * outside the data rows). `deps`: for dependency collection, where `@` without a host row
   * means every data row.
   */
  tableRef(
    n: StructNode,
    host: SheetIndex,
    row: number,
    col: number,
    deps: boolean,
  ): [SheetIndex, number, number, number, number] | CellError {
    if (!this.#tables) {
      this.#tables = new Map();
      for (const si of this.sheets) {
        for (const t of si.ws.tables ?? []) {
          this.#tables.set(t.name.toLowerCase(), { si, t });
          if (t.displayName) this.#tables.set(t.displayName.toLowerCase(), { si, t });
        }
      }
    }
    let hit: { si: SheetIndex; t: Table } | undefined;
    let box: [number, number, number, number] | undefined;
    if (n.table !== undefined) {
      hit = this.#tables.get(n.table.toLowerCase());
      if (hit) box = decodeRangeSafe(hit.t.ref);
    } else {
      for (const t of host.ws.tables ?? []) {
        const b = decodeRangeSafe(t.ref);
        if (b && row >= b[0] && row <= b[2] && col >= b[1] && col <= b[3]) {
          hit = { si: host, t };
          box = b;
          break;
        }
      }
    }
    if (!hit || !box) return REF;
    const t = hit.t;
    const [top, left, bottom, right] = box;
    const header = t.headerRow !== false ? top : -1;
    const totals = t.totalsRow === true ? bottom : -1;
    const first = top + (header >= 0 ? 1 : 0);
    const last = bottom - (totals >= 0 ? 1 : 0);
    let r0 = first;
    let r1 = last;
    const a = n.areas;
    if (a.length > 0) {
      const has = (x: string) => a.includes(x as never);
      if (has('all')) {
        r0 = top;
        r1 = bottom;
      } else if (has('thisRow')) {
        if (hit.si === host && row >= first && row <= last) r0 = r1 = row;
        else if (!deps) return VALUE;
      } else {
        const h = has('headers');
        const d = has('data');
        const tt = has('totals');
        if ((h && header < 0) || (tt && totals < 0 && !d)) return REF;
        r0 = h ? header : d ? first : totals;
        r1 = tt && totals >= 0 ? totals : d ? last : h ? header : totals;
      }
    }
    let c0 = left;
    let c1 = right;
    if (n.from !== undefined) {
      const i = colOf(t, n.from);
      const j = colOf(t, n.to ?? n.from);
      if (i < 0 || j < 0) return REF;
      c0 = left + Math.min(i, j);
      c1 = left + Math.max(i, j);
    }
    return [hit.si, r0, c0, r1, c1];
  }

  #depsIn(
    t: SheetIndex,
    top: number,
    left: number,
    bottom: number,
    right: number,
    out: number[],
  ): void {
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
        this.depsOf(p, si, deps, 0, this.#nR[v]!, this.#nC[v]!);
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
      if (!this.#inOrder[id]) {
        this.#inOrder[id] = 1;
        this.#order.push(id);
      }
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
    if (!this.#inOrder[id]) {
      this.#inOrder[id] = 1;
      this.#order.push(id);
    }
  }

  #compute(id: number): FormulaValue {
    const f = this.#nF[id]!;
    const p = this.#nP[id] ?? this.ev.parsed(f.formula);
    if (!p.ast) {
      this.#errors.push({ address: this.#addr(id), message: p.error!.message });
      return NAME;
    }
    const dyn = this.#dyn.has(id);
    const ctx = new EvalContext(
      this,
      this.#nS[id]!,
      this.#nR[id]!,
      this.#nC[id]!,
      f.array !== undefined || dyn,
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
      if (dyn && (v.rows > 1 || v.cols > 1) && this.#blocked(id, v.rows, v.cols)) return SPILL;
      return new FormulaArray(v.rows, v.cols, v.data.map(final));
    }
    return final(v);
  }

  /**
   * Whether the spill area of dynamic formula `id` (rows x cols from its cell) is blocked: off
   * the sheet, a non-empty cell (other than the formula's own stale spilled values), or the
   * area of an earlier dynamic formula.
   */
  #blocked(id: number, rows: number, cols: number): boolean {
    const r0 = this.#nR[id]!;
    const c0 = this.#nC[id]!;
    const r1 = r0 + rows - 1;
    const c1 = c0 + cols - 1;
    if (r1 > 1048575 || c1 > 16383) return true;
    const si = this.#nS[id]!;
    const old = this.#old.get(id);
    const data = si.ws.rowData;
    for (let r = r0; r <= r1 && r < data.length; r++) {
      const row = data[r];
      if (row === undefined) continue;
      let i = row.find(c0);
      if (i < 0) i = -i - 1;
      for (; i < row.n; i++) {
        const c = row.col[i]!;
        if (c > c1) break;
        const k = row.kind[i];
        if (k === K_EMPTY || (r === r0 && c === c0)) continue;
        const stale =
          old !== undefined &&
          k !== K_FORMULA &&
          r >= old[0] &&
          r <= old[2] &&
          c >= old[1] &&
          c <= old[3];
        if (!stale) return true;
      }
    }
    for (const [other, e] of this.#dyn) {
      if (other >= id || this.#nS[other] !== si) continue;
      if (e[0]! <= r1 && e[2]! >= r0 && e[1]! <= c1 && e[3]! >= c0) return true;
    }
    return false;
  }

  /**
   * After a pass: record each evaluated dynamic formula's spill area in `hints`; true when an
   * area differs from the one this pass assumed (another pass is needed).
   */
  settle(hints: Map<string, number[]>): boolean {
    let changed = false;
    for (const [id, e] of this.#dyn) {
      if (this.#state[id] !== 2) continue;
      const v = this.#val[id];
      const r = this.#nR[id]!;
      const c = this.#nC[id]!;
      const b = isArr(v) ? r + v.rows - 1 : r;
      const rr = isArr(v) ? c + v.cols - 1 : c;
      hints.set(`${this.#nS[id]!.i}:${r}:${c}`, [r, c, b, rr]);
      if (e[2] !== b || e[3] !== rr) changed = true;
    }
    return changed;
  }

  /** The spilled range of the dynamic formula at (r, c) of `si` (`A1#`), or #REF!. */
  spillRef(si: SheetIndex, r: number, c: number): FormulaValue {
    const id = this.find(si, r, c);
    if (id < 0 || !this.#dyn.has(id)) return REF;
    this.#nodeValue(id);
    const v = this.#val[id];
    if (isError(v) && v.error === '#SPILL!') return REF;
    const rows = isArr(v) ? v.rows : 1;
    const cols = isArr(v) ? v.cols : 1;
    return new FormulaRange(this, si.ws, r, c, r + rows - 1, c + cols - 1);
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
    if (i < 0) return null;
    if (row.kind[i] !== K_FORMULA && this.#isStale(ws, r, c)) return null;
    return this.#slot(ws, row, i);
  }

  /** A stored value left from a dynamic formula's previous, larger spill (no longer spilled). */
  #isStale(ws: Worksheet, r: number, c: number): boolean {
    const st = this.#byWs.get(ws)?.stale;
    if (!st) return false;
    for (const a of st) if (r >= a[0]! && r <= a[2]! && c >= a[1]! && c <= a[3]!) return true;
    return false;
  }

  column(ref: FormulaRange): readonly FormulaScalar[] | undefined {
    const ws = ref.sheet;
    const si = this.#byWs.get(ws);
    const end = Math.min(ref.bottom, ws.rowData.length - 1);
    if (ref.left !== ref.right || this.#iter || !si || si.arrays || end - ref.top < 16) {
      return undefined;
    }
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
    if (si?.arrays !== undefined) {
      this.#withAreas(ref, si.arrays, cb);
      return;
    }
    const slow = false;
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

  /**
   * each() on a sheet with array / spill areas: every stored cell and every cell of an area
   * (spilled values need not be stored yet), row by row, through cell().
   */
  #withAreas(
    ref: FormulaRange,
    areas: number[][],
    cb: (v: FormulaScalar, row: number, col: number) => void,
  ): void {
    const ws = ref.sheet;
    const rows = ws.rowData;
    const hit = areas.filter(
      (a) => a[0]! <= ref.bottom && a[2]! >= ref.top && a[1]! <= ref.right && a[3]! >= ref.left,
    );
    let end = Math.min(ref.bottom, rows.length - 1);
    for (const a of hit) end = Math.max(end, Math.min(ref.bottom, a[2]!));
    for (let r = ref.top; r <= end; r++) {
      const cols: number[] = [];
      const row = rows[r];
      if (row) {
        let i = row.find(ref.left);
        if (i < 0) i = -i - 1;
        for (; i < row.n && row.col[i]! <= ref.right; i++) {
          if (row.kind[i] !== K_EMPTY) cols.push(row.col[i]!);
        }
      }
      for (const a of hit) {
        if (r < a[0]! || r > a[2]!) continue;
        for (let c = Math.max(ref.left, a[1]!); c <= Math.min(ref.right, a[3]!); c++) cols.push(c);
      }
      if (cols.length === 0) continue;
      cols.sort((x, y) => x - y);
      let prev = -1;
      for (const c of cols) {
        if (c === prev) continue;
        prev = c;
        const v = this.cell(ws, r, c);
        if (v !== null) cb(v, r, c);
      }
    }
  }

  report(write: boolean): EvaluationReport {
    const results = write ? undefined : new Map<string, FormulaResult>();
    for (const id of this.#order) {
      const v = this.#val[id];
      const res = isArr(v) ? (v.data[0] as FormulaResult) : (v as FormulaResult);
      if (results) results.set(this.#addr(id), res);
      else this.#write(id);
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

  /** Store node `id`'s value into its cell (and its spilled / array values). */
  #write(id: number): void {
    const v = this.#val[id];
    const res = isArr(v) ? (v.data[0] as FormulaResult) : (v as FormulaResult);
    const ws = this.#nS[id]!.ws;
    const r = this.#nR[id]!;
    const c = this.#nC[id]!;
    const row = ws.rowData[r]!;
    const i = row.find(c);
    if (i < 0 || row.kind[i] !== K_FORMULA) return;
    const f = row.obj![i] as Formula;
    if (this.#dyn.has(id)) {
      const b = isArr(v) ? r + v.rows - 1 : r;
      const rr = isArr(v) ? c + v.cols - 1 : c;
      const old = this.#old.get(id);
      if (old) this.#clearStale(ws, old, [r, c, b, rr]);
      row.obj![i] = {
        ...f,
        result: res,
        array: colName(c) + (r + 1) + (b > r || rr > c ? `:${colName(rr)}${b + 1}` : ''),
        dynamic: true,
      };
      if (isArr(v)) this.#writeSpill(ws, id, v, [r, c, b, rr]);
      return;
    }
    row.obj![i] = { ...f, result: res };
    if (isArr(v) && f.array) {
      const a = decodeRangeSafe(f.array);
      if (a) this.#writeSpill(ws, id, v, a);
    }
  }

  /** Addresses of every evaluated formula cell. */
  addresses(): string[] {
    return this.#order.map((id) => this.#addr(id));
  }

  /**
   * Incremental recalculation after the values (not formulas) of `cells` changed: re-evaluate
   * the formulas reading them, directly or through other formulas, in dependency order, and
   * store their results. Returns their addresses, or undefined when a full evaluation is
   * needed instead (a changed cell is or was a formula or lies in an array / spill area, or a
   * cycle / dynamic-array formula is affected).
   */
  update(cells: readonly (readonly [SheetIndex, number, number])[]): string[] | undefined {
    const dirtyCells = new Map<number, Map<number, number[]>>();
    const mark = (s: number, r: number, c: number): void => {
      let cols = dirtyCells.get(s);
      if (!cols) {
        cols = new Map();
        dirtyCells.set(s, cols);
      }
      let rows = cols.get(c);
      if (!rows) {
        rows = [];
        cols.set(c, rows);
      }
      let lo = 0;
      let hi = rows.length;
      while (lo < hi) {
        const m = (lo + hi) >>> 1;
        if (rows[m]! < r) lo = m + 1;
        else hi = m;
      }
      if (rows[lo] !== r) rows.splice(lo, 0, r);
    };
    const hits = (s: number, t: number, l: number, b: number, r: number): boolean => {
      const cols = dirtyCells.get(s);
      if (!cols) return false;
      const check = (rows: number[]): boolean => {
        let lo = 0;
        let hi = rows.length;
        while (lo < hi) {
          const m = (lo + hi) >>> 1;
          if (rows[m]! < t) lo = m + 1;
          else hi = m;
        }
        return lo < rows.length && rows[lo]! <= b;
      };
      if (cols.size <= r - l + 1) {
        for (const [c, rows] of cols) if (c >= l && c <= r && check(rows)) return true;
        return false;
      }
      for (let c = l; c <= r; c++) {
        const rows = cols.get(c);
        if (rows && check(rows)) return true;
      }
      return false;
    };
    for (const [si, r, c] of cells) {
      if (this.find(si, r, c) >= 0) return undefined;
      const row = si.ws.rowData[r];
      const i = row ? row.find(c) : -1;
      if (i >= 0 && row!.kind[i] === K_FORMULA) return undefined;
      for (const a of si.arrays ?? []) {
        if (r >= a[0]! && r <= a[2]! && c >= a[1]! && c <= a[3]!) return undefined;
      }
      mark(si.i, r, c);
    }
    const circ = new Set(this.#circ);
    const dirty: number[] = [];
    for (const id of this.#order) {
      const rc = this.#rects[id] ?? this.#rectsOf(id);
      // rc[0] === -1: the formula uses a volatile function
      let d = rc[0] === -1;
      for (let k = d ? 1 : 0; k < rc.length && !d; k += 5) {
        d = hits(rc[k]!, rc[k + 1]!, rc[k + 2]!, rc[k + 3]!, rc[k + 4]!);
      }
      if (!d) continue;
      if (circ.has(id) || this.#dyn.has(id)) return undefined;
      dirty.push(id);
      const si = this.#nS[id]!;
      mark(si.i, this.#nR[id]!, this.#nC[id]!);
      for (const a of si.arrays ?? []) {
        if (a[4] !== id) continue;
        for (let r = a[0]!; r <= a[2]!; r++) for (let c = a[1]!; c <= a[3]!; c++) mark(si.i, r, c);
      }
    }
    this.#cols.clear();
    for (const id of dirty) this.#state[id] = 0;
    for (const id of dirty) if (this.#state[id] === 0) this.#evalNode(id);
    for (const id of dirty) this.#write(id);
    return dirty.map((id) => this.#addr(id));
  }

  /**
   * The rectangles node `id` reads (cached): sheet index, top, left, bottom, right, ...;
   * starts with -1 when the formula uses a volatile function.
   */
  #rectsOf(id: number): number[] {
    let rc = this.#rects[id];
    if (rc) return rc;
    const p = this.#nP[id] ?? this.ev.parsed(this.#nF[id]!.formula);
    rc = p.calls?.some((n) => VOLATILE.has(n)) ? [-1] : [];
    this.#collectRects(p, this.#nS[id]!, rc, 0, this.#nR[id]!, this.#nC[id]!);
    this.#rects[id] = rc;
    return rc;
  }

  #collectRects(
    p: Parsed,
    si: SheetIndex,
    out: number[],
    depth: number,
    row: number,
    col: number,
  ): void {
    if (!p.refs) return;
    for (const ref of p.refs) {
      if (ref.length === 2) {
        if (depth > 16) continue;
        const d = this.definedName(ref[1], ref[0], si);
        if (d) this.#collectRects(this.ev.parsed(d.value), si, out, depth + 1, row, col);
        continue;
      }
      if (ref.length === 1) {
        const a = this.tableRef(ref[0], si, row, col, true);
        if (!isError(a)) out.push(a[0].i, a[1], a[2], a[3], a[4]);
        continue;
      }
      const t = this.sheetFor(ref[0], si);
      if (t) out.push(t.i, ref[1], ref[2], ref[3], ref[4]);
    }
  }

  /** Remove a dynamic formula's stale spilled values outside its new area. */
  #clearStale(ws: Worksheet, old: number[], now: number[]): void {
    for (let r = old[0]!; r <= old[2]!; r++) {
      const row = ws.rowData[r];
      if (!row) continue;
      for (let c = old[1]!; c <= old[3]!; c++) {
        if (r >= now[0]! && r <= now[2]! && c >= now[1]! && c <= now[3]!) continue;
        const i = row.find(c);
        if (i >= 0 && row.kind[i] !== K_FORMULA && row.kind[i] !== K_EMPTY) {
          ws.setCellRaw(r, c, K_EMPTY, 0, undefined, row.style[i]!);
        }
      }
    }
  }

  #writeSpill(ws: Worksheet, id: number, v: FormulaArray, a: readonly number[]): void {
    if ((a[2]! - a[0]! + 1) * (a[3]! - a[1]! + 1) > 1e6) return;
    for (let r = a[0]!; r <= a[2]!; r++) {
      for (let c = a[1]!; c <= a[3]!; c++) {
        if (r === this.#nR[id] && c === this.#nC[id]) continue;
        const ii = v.rows === 1 ? 0 : r - a[0]!;
        const jj = v.cols === 1 ? 0 : c - a[1]!;
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

/** 0-based index of a table column by name (case-insensitive), or -1. */
function colOf(t: Table, name: string): number {
  const k = name.trim().toLowerCase();
  return t.columns.findIndex((c) => c.name.toLowerCase() === k);
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

/** Functions whose result can change without a referenced cell changing (or that read cells
 * not visible in the formula text): always recalculated by update(). */
const VOLATILE = new Set([
  'INDIRECT',
  'OFFSET',
  'NOW',
  'TODAY',
  'RAND',
  'RANDBETWEEN',
  'RANDARRAY',
  'CELL',
  'INFO',
]);

const localKey = (name: string): string => name.toLowerCase().replace(/^_xlpm\./, '');

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
  /** LET names in scope (lower-case, without `_xlpm.`). */
  #locals: Map<string, FormulaValue> | undefined;

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
    c.#locals = this.#locals;
    return c;
  }

  /** @internal This context with one more LET name bound to `v`. */
  withLocal(name: string, v: FormulaValue): EvalContext {
    const c = new EvalContext(this.#run, this.#si, this.row, this.col, this.array);
    c.#depth = this.#depth;
    c.#locals = new Map(this.#locals);
    c.#locals.set(localKey(name), v);
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
        if (this.#locals && n.sheet === undefined) {
          const v = this.#locals.get(localKey(n.name));
          if (v !== undefined) return v;
        }
        return this.#name(n.name, n.sheet);
      case 'spill': {
        const si = this.#run.sheetFor(n.sheet, this.#si);
        return si ? this.#run.spillRef(si, n.row, n.col) : REF;
      }
      case 'structured': {
        const a = this.#run.tableRef(n, this.#si, this.row, this.col, false);
        return isError(a) ? a : new FormulaRange(this.#run, a[0].ws, a[1], a[2], a[3], a[4]);
      }
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

  /**
   * A reference from text (INDIRECT): "B2", "Sheet2!A1:C3", "'My Sheet'!B:B" or a defined
   * name, evaluated like a formula; anything that is not a reference gives #REF!.
   */
  reference(text: string): FormulaValue {
    const p = this.#run.ev.parsed(text.trim());
    const t = p.ast?.type;
    if (t !== 'cell' && t !== 'range' && t !== 'name' && t !== 'structured') return REF;
    const v = this.ev(p.ast!);
    return isRef(v) ? v : REF;
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
