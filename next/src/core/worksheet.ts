/**
 * Worksheet model with object-free, row-major cell storage.
 *
 * ## Storage
 * Each non-empty row is a {@link RowData} holding `n` cells in parallel typed arrays sorted by
 * column (all four arrays are views over ONE ArrayBuffer per row, 15 bytes per slot):
 *
 *   col[i]   Uint16Array   0-based column index (strictly increasing for i < n)
 *   kind[i]  Uint8Array    a {@link CellKind}
 *   num[i]   Float64Array  numeric payload (see table below)
 *   style[i] Uint32Array   xf index into `workbook.styles` (0 = default style)
 *   obj[i]   unknown[]     sparse payload for non-numeric kinds; `obj` itself is undefined when
 *                          the row has no such cell, and may be shorter than `n`
 *
 *   kind       num                       obj[i]
 *   Empty      0                         –          (style-only cell; style[i] != 0)
 *   Number     the value                 –
 *   String     0                         string     (writer decides shared vs inline)
 *   Boolean    1 | 0                     –
 *   Date       serial (wb.date1904 aware)–          (style always has a date numFmt, see below)
 *   Error      0                         ErrorCode string, e.g. '#N/A'
 *   Formula    0                         Formula object ({ formula, result?, array? })
 *   RichText   0                         RichText object
 *   Hyperlink  0                         Hyperlink object ({ text, hyperlink, tooltip? })
 *
 * Rows live in a sparse JS array indexed by 0-based row number (`ws.rowData`). A RowData may
 * exist with n = 0 when only row properties (height/hidden/outlineLevel/style) are set.
 *
 * ## Writer iteration API (zero allocation)
 * ```ts
 * const rows = ws.rowData;                       // readonly (RowData | undefined)[]
 * for (let r = 0; r < rows.length; r++) {
 *   const row = rows[r];
 *   if (row === undefined) continue;
 *   // row.index === r; row props: row.height (0 = default), row.hidden, row.outlineLevel,
 *   // row.xf (row default style, 0 = none → customFormat when != 0)
 *   const { n, col, kind, num, style, obj } = row;   // read AFTER any mutation; arrays are
 *   for (let i = 0; i < n; i++) {                     // reallocated when a row grows
 *     switch (kind[i]) { case CellKind.Number: write(num[i]); break;
 *                        case CellKind.String: write(obj![i] as string); break; ... }
 *   }
 * }
 * ```
 * or the callback form `row.forEachCell((col, kind, num, obj, xf) => …)` (obj is undefined for
 * numeric kinds). Columns: `ws.columnData` (sparse, 0-based; Column has width/hidden/
 * outlineLevel/styleIndex). Merges: `ws.mergeRanges` (0-based CellRange). Used range:
 * `ws.dimensions`. Styles/date system: `ws.workbook.styles`, `ws.workbook.date1904`.
 * Readers can fill cells without building JS values via `ws.setCellRaw(r0, c0, kind, num, obj, xf)`
 * (fastest when called in ascending column order within a row).
 *
 * ## Value & style semantics (public API is 1-based like Excel; "B3" strings accepted)
 * - number → Number (NaN/±Infinity → Error '#NUM!'); string → String; boolean → Boolean;
 *   Date → Date (invalid Date → '#VALUE!'); { error } → Error; { formula } → Formula (a leading
 *   '=' is stripped); { richText } → RichText; { hyperlink } → Hyperlink.
 * - null/undefined clears the value but keeps the cell's formatting (Excel "Clear Contents");
 *   the slot disappears when it has no style. `clearCell()` removes value and style.
 * - When a NEW cell is created without an explicit style it inherits the row style, else the
 *   column style (as Excel does when typing into a formatted row/column). Setting a value on
 *   an existing cell without a style keeps its style. Changing a row/column style later does
 *   not restyle existing cells; the writer emits row/column default styles as-is.
 * - Date cells whose resulting style has the General number format get 'yyyy-mm-dd', or
 *   'yyyy-mm-dd hh:mm:ss' when the value has a time part, merged into that style.
 */

import {
  type CellRange,
  decodeAddr,
  decodeCol,
  decodeRange,
  encodeAddr,
  encodeCol,
  encodeRange,
  MAX_COLS,
  MAX_ROWS,
} from './address.js';
import { dateToSerial, serialToDate } from './dates.js';
import type { WorksheetFeatureFields } from './features.js';
import type { StyleRegistry } from './styles.js';
import type {
  Alignment,
  Border,
  CellValue,
  Color,
  ErrorCode,
  Fill,
  Font,
  Formula,
  Protection,
  Style,
} from './types.js';
import type { Workbook } from './workbook.js';

export const CellKind = {
  Empty: 0,
  Number: 1,
  String: 2,
  Boolean: 3,
  Date: 4,
  Error: 5,
  Formula: 6,
  RichText: 7,
  Hyperlink: 8,
} as const;
export type CellKind = (typeof CellKind)[keyof typeof CellKind];

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

export interface WorksheetView {
  state?: 'normal' | 'frozen' | 'split' | undefined;
  /** Frozen/split columns (count for frozen panes). */
  xSplit?: number | undefined;
  /** Frozen/split rows (count for frozen panes). */
  ySplit?: number | undefined;
  topLeftCell?: string | undefined;
  activeCell?: string | undefined;
  showGridLines?: boolean | undefined;
  showRowColHeaders?: boolean | undefined;
  zoomScale?: number | undefined;
  rightToLeft?: boolean | undefined;
  tabSelected?: boolean | undefined;
}

export interface ColumnDef {
  header?: string | undefined;
  key?: string | undefined;
  width?: number | undefined;
  style?: Style | undefined;
  hidden?: boolean | undefined;
  outlineLevel?: number | undefined;
}

export interface RowProps {
  /** Height in points. */
  height?: number | undefined;
  hidden?: boolean | undefined;
  outlineLevel?: number | undefined;
  style?: Style | undefined;
  /** Row style as an xf index of `workbook.styles` (wins over `style`; 0 = none). */
  styleIndex?: number | undefined;
}

/** Used range, 1-based inclusive. */
export interface Dimensions {
  top: number;
  left: number;
  bottom: number;
  right: number;
  /** A1 reference, e.g. "A1:T50000". */
  ref: string;
}

export type CellCallback = (
  col: number,
  kind: number,
  num: number,
  obj: unknown,
  xf: number,
) => void;

// ---- RowData -------------------------------------------------------------------------

const E_U16 = /* @__PURE__ */ new Uint16Array(0);
const E_U8 = /* @__PURE__ */ new Uint8Array(0);
const E_F64 = /* @__PURE__ */ new Float64Array(0);
const E_U32 = /* @__PURE__ */ new Uint32Array(0);

/** Internal row record; see the module header for the layout contract. */
export class RowData {
  /** 0-based row index. */
  readonly index: number;
  /** Number of used slots. */
  n = 0;
  col: Uint16Array = E_U16;
  kind: Uint8Array = E_U8;
  num: Float64Array = E_F64;
  style: Uint32Array = E_U32;
  obj: unknown[] | undefined = undefined;
  /** Row height in points; 0 = default. */
  height = 0;
  hidden = false;
  outlineLevel = 0;
  /** Row default style xf; 0 = none. */
  xf = 0;

  constructor(index: number, capacity = 0) {
    this.index = index;
    if (capacity > 0) this._alloc(capacity);
  }

  get capacity(): number {
    return this.col.length;
  }

  /** @internal Reallocate to `cap` slots (one ArrayBuffer, 15 bytes per slot). */
  _alloc(cap: number): void {
    const buf = new ArrayBuffer(cap * 15);
    const num = new Float64Array(buf, 0, cap);
    const style = new Uint32Array(buf, cap * 8, cap);
    const col = new Uint16Array(buf, cap * 12, cap);
    const kind = new Uint8Array(buf, cap * 14, cap);
    const n = this.n;
    if (n > 0) {
      num.set(this.num.subarray(0, n));
      style.set(this.style.subarray(0, n));
      col.set(this.col.subarray(0, n));
      kind.set(this.kind.subarray(0, n));
    }
    this.num = num;
    this.style = style;
    this.col = col;
    this.kind = kind;
  }

  /** Slot index of column c, or -(insertionPoint + 1) when absent. */
  find(c: number): number {
    const n = this.n;
    const col = this.col;
    if (n === 0 || col[n - 1]! < c) return -(n + 1);
    let lo = 0;
    let hi = n - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      const v = col[mid]!;
      if (v < c) lo = mid + 1;
      else if (v > c) hi = mid - 1;
      else return mid;
    }
    return -(lo + 1);
  }

  /** @internal Open an (empty, unstyled) slot at p for column c. */
  _insert(p: number, c: number): void {
    const n = this.n;
    if (n === this.col.length) this._alloc(n < 4 ? 4 : n + (n >> 1));
    if (p < n) {
      this.col.copyWithin(p + 1, p, n);
      this.kind.copyWithin(p + 1, p, n);
      this.num.copyWithin(p + 1, p, n);
      this.style.copyWithin(p + 1, p, n);
      const obj = this.obj;
      if (obj !== undefined && obj.length > p) obj.splice(p, 0, undefined);
    }
    this.col[p] = c;
    this.kind[p] = 0;
    this.num[p] = 0;
    this.style[p] = 0;
    this.n = n + 1;
  }

  /** @internal Remove slot p. */
  _remove(p: number): void {
    const n = this.n;
    this.col.copyWithin(p, p + 1, n);
    this.kind.copyWithin(p, p + 1, n);
    this.num.copyWithin(p, p + 1, n);
    this.style.copyWithin(p, p + 1, n);
    const obj = this.obj;
    if (obj !== undefined && obj.length > p) obj.splice(p, 1);
    this.n = n - 1;
  }

  /** Visit cells in column order. `col` is 0-based; `obj` is undefined for numeric kinds. */
  forEachCell(cb: CellCallback): void {
    const { n, col, kind, num, style, obj } = this;
    if (obj === undefined) {
      for (let i = 0; i < n; i++) cb(col[i]!, kind[i]!, num[i]!, undefined, style[i]!);
    } else {
      for (let i = 0; i < n; i++) cb(col[i]!, kind[i]!, num[i]!, obj[i], style[i]!);
    }
  }
}

// ---- value classification (writes to module scratch to avoid per-cell allocation) -------

let K = 0;
let N = 0;
let O: unknown;

function classify(v: NonNullable<CellValue>, date1904: boolean): void {
  O = undefined;
  N = 0;
  switch (typeof v) {
    case 'number':
      if (v - v === 0) {
        K = CellKind.Number;
        N = v;
      } else {
        K = CellKind.Error;
        O = '#NUM!';
      }
      return;
    case 'string':
      K = CellKind.String;
      O = v;
      return;
    case 'boolean':
      K = CellKind.Boolean;
      N = v ? 1 : 0;
      return;
    case 'object':
      if (v instanceof Date) {
        const t = v.getTime();
        if (t - t === 0) {
          K = CellKind.Date;
          N = dateToSerial(v, date1904);
        } else {
          K = CellKind.Error;
          O = '#VALUE!';
        }
        return;
      }
      if ('formula' in v) {
        K = CellKind.Formula;
        const f = v.formula;
        O = f.charCodeAt(0) === 61 ? ({ ...v, formula: f.slice(1) } as Formula) : v;
        return;
      }
      if ('richText' in v) {
        K = CellKind.RichText;
        O = v;
        return;
      }
      if ('hyperlink' in v) {
        K = CellKind.Hyperlink;
        O = v;
        return;
      }
      if ('error' in v) {
        K = CellKind.Error;
        O = v.error;
        return;
      }
  }
  throw new TypeError(`Unsupported cell value: ${String(v)}`);
}

const DATE_FMT = 'yyyy-mm-dd';
const DATETIME_FMT = 'yyyy-mm-dd hh:mm:ss';

function checkCell(r0: number, c0: number): void {
  if (
    !(r0 >= 0 && r0 < MAX_ROWS && c0 >= 0 && c0 < MAX_COLS) ||
    r0 !== (r0 | 0) ||
    c0 !== (c0 | 0)
  ) {
    throw new RangeError(`Cell out of range: row ${r0 + 1}, col ${c0 + 1}`);
  }
}

function overlaps(a: CellRange, b: CellRange): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** log2 of the rows per merge-index bucket. */
const MERGE_BUCKET = 4;

function indexMerge(idx: Map<number, CellRange[]>, m: CellRange): void {
  for (let b = m.top >> MERGE_BUCKET; b <= m.bottom >> MERGE_BUCKET; b++) {
    const list = idx.get(b);
    if (list) list.push(m);
    else idx.set(b, [m]);
  }
}

type Addr = string | number;

// ---- Column ------------------------------------------------------------------------

export class Column {
  /** 1-based column number. */
  readonly number: number;
  header: string | undefined = undefined;
  /** Width in characters; undefined = default. */
  width: number | undefined = undefined;
  hidden = false;
  outlineLevel = 0;
  private readonly _ws: Worksheet;
  private _key: string | undefined = undefined;
  private _xf = 0;

  constructor(ws: Worksheet, number: number) {
    this._ws = ws;
    this.number = number;
  }

  get letter(): string {
    return encodeCol(this.number - 1);
  }

  get key(): string | undefined {
    return this._key;
  }
  set key(k: string | undefined) {
    this._key = k;
    this._ws._rebuildKeys();
  }

  /** Column default style (applies to cells created afterwards and to empty cells in Excel). */
  get style(): Style {
    return this._ws.workbook.styles.get(this._xf);
  }
  set style(s: Style | undefined) {
    this.styleIndex = this._ws.workbook.styles.intern(s);
  }

  get styleIndex(): number {
    return this._xf;
  }
  set styleIndex(xf: number) {
    this._xf = xf;
    this._ws._setColXf(this.number - 1, xf);
  }
}

// ---- Worksheet ----------------------------------------------------------------------

/** Feature-owned optional fields (pageSetup, protection, ...), see core/features.ts. */
export interface Worksheet extends WorksheetFeatureFields {}

// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: type-only feature fields
export class Worksheet {
  readonly workbook: Workbook;
  /** sheetId (1-based, stable across removals of other sheets). */
  readonly id: number;
  state: SheetState = 'visible';
  tabColor: Color | undefined = undefined;
  views: WorksheetView[] = [];
  /** Auto-filter range, e.g. "A1:D1". */
  autoFilter: string | undefined = undefined;

  private _name: string;
  private readonly _styles: StyleRegistry;
  private readonly _rows: (RowData | undefined)[] = [];
  private _cols: (Column | undefined)[] = [];
  private _colXf: Uint32Array | undefined = undefined;
  private _keyCols: [string, number][] = [];
  private _keyMap = new Map<string, number>();
  private readonly _merges: CellRange[] = [];
  /** Lazily built merge index: row bucket (row >> MERGE_BUCKET) → merges touching it. */
  private _mergeIdx: Map<number, CellRange[]> | undefined = undefined;

  /** @internal Use `workbook.addWorksheet()`. */
  constructor(workbook: Workbook, id: number, name: string) {
    this.workbook = workbook;
    this.id = id;
    this._name = name;
    this._styles = workbook.styles;
  }

  get name(): string {
    return this._name;
  }
  set name(v: string) {
    this.workbook._checkName(v, this);
    this._name = v;
  }

  // ---- low-level access ----

  /** Sparse array of row records indexed by 0-based row (see module header). */
  get rowData(): readonly (RowData | undefined)[] {
    return this._rows;
  }

  /** Sparse array of column records indexed by 0-based column. */
  get columnData(): readonly (Column | undefined)[] {
    return this._cols;
  }

  /** Merged ranges, 0-based. */
  get mergeRanges(): readonly CellRange[] {
    return this._merges;
  }

  /**
   * Reader fast path (0-based, no value conversion, no style inheritance, no date-format
   * defaulting). `obj` must match the kind contract in the module header.
   */
  setCellRaw(r0: number, c0: number, kind: number, num: number, obj: unknown, xf: number): void {
    checkCell(r0, c0);
    const row = this._rows[r0] ?? this._newRow(r0, 4);
    let i = row.find(c0);
    if (i < 0) {
      i = -i - 1;
      row._insert(i, c0);
    }
    row.kind[i] = kind;
    row.num[i] = num;
    row.style[i] = xf;
    if (obj !== undefined) {
      if (row.obj === undefined) row.obj = [];
      row.obj[i] = obj;
    } else if (row.obj !== undefined && i < row.obj.length) row.obj[i] = undefined;
  }

  /**
   * Internal-but-public (readers, features): the row record for 0-based row `r0`, created when
   * missing. `capacity` pre-sizes the cell arrays (grows an existing row's capacity if smaller)
   * to avoid regrowth when the caller knows roughly how many cells will follow.
   */
  _ensureRow(r0: number, capacity = 0): RowData {
    checkCell(r0, 0);
    const row = this._rows[r0];
    if (row === undefined) return this._newRow(r0, capacity);
    if (capacity > row.capacity) row._alloc(capacity);
    return row;
  }

  // ---- cells ----

  /** Set a cell value: `setCell('B3', v, style?)` or `setCell(row, col, v, style?)` (1-based). */
  setCell(addr: string, value: CellValue, style?: Style | number): void;
  setCell(row: number, col: number | string, value: CellValue, style?: Style | number): void;
  setCell(a: Addr, b: unknown, c?: unknown, d?: unknown): void {
    if (typeof a === 'string') {
      const p = decodeAddr(a);
      this._set(p.row, p.col, b as CellValue, c as Style | number | undefined);
    } else {
      this._set(a - 1, this._colIndex(b as Addr), c as CellValue, d as Style | number | undefined);
    }
  }

  /** Read a cell value (Date/objects reconstructed); null when empty. */
  getValue(addr: string): CellValue;
  getValue(row: number, col: number | string): CellValue;
  getValue(a: Addr, b?: Addr): CellValue {
    const r0 = typeof a === 'string' ? decodeAddr(a) : undefined;
    const row = this._rows[r0 ? r0.row : (a as number) - 1];
    if (row === undefined) return null;
    const i = row.find(r0 ? r0.col : this._colIndex(b!));
    return i < 0 ? null : this._valueAt(row, i);
  }

  /** CellKind of a cell (Empty when absent). */
  getKind(addr: string): CellKind;
  getKind(row: number, col: number | string): CellKind;
  getKind(a: Addr, b?: Addr): CellKind {
    const [r0, c0] = this._rc(a, b);
    const row = this._rows[r0];
    if (row === undefined) return CellKind.Empty;
    const i = row.find(c0);
    return i < 0 ? CellKind.Empty : (row.kind[i] as CellKind);
  }

  /** Effective style of a cell (an absent cell reports its row/column default). */
  getStyle(addr: string): Style;
  getStyle(row: number, col: number | string): Style;
  getStyle(a: Addr, b?: Addr): Style {
    const [r0, c0] = this._rc(a, b);
    return this._styles.get(this._xfAt(r0, c0));
  }

  /** xf index of a cell (see getStyle). */
  getStyleIndex(addr: string): number;
  getStyleIndex(row: number, col: number | string): number;
  getStyleIndex(a: Addr, b?: Addr): number {
    const [r0, c0] = this._rc(a, b);
    return this._xfAt(r0, c0);
  }

  /** Replace a cell's style (creates a style-only cell when absent). */
  setStyle(addr: string, style: Style | number | undefined): void;
  setStyle(row: number, col: number | string, style: Style | number | undefined): void;
  setStyle(a: Addr, b: unknown, c?: unknown): void {
    let r0: number;
    let c0: number;
    let s: Style | number | undefined;
    if (typeof a === 'string') {
      const p = decodeAddr(a);
      r0 = p.row;
      c0 = p.col;
      s = b as Style | number | undefined;
    } else {
      r0 = a - 1;
      c0 = this._colIndex(b as Addr);
      s = c as Style | number | undefined;
    }
    checkCell(r0, c0);
    const xf = this._xf(s);
    const row = this._rows[r0];
    let i = row ? row.find(c0) : -1;
    if (row && i >= 0) {
      if (xf === 0 && row.kind[i] === CellKind.Empty) row._remove(i);
      else {
        row.style[i] = row.kind[i] === CellKind.Date ? this._dateXf(xf, row.num[i]!) : xf;
      }
      return;
    }
    if (xf === 0) return;
    const r = row ?? this._newRow(r0, 1);
    i = -r.find(c0) - 1;
    r._insert(i, c0);
    r.style[i] = xf;
  }

  /** Remove a cell entirely (value and style). */
  clearCell(addr: string): void;
  clearCell(row: number, col: number | string): void;
  clearCell(a: Addr, b?: Addr): void {
    const [r0, c0] = this._rc(a, b);
    const row = this._rows[r0];
    if (row === undefined) return;
    const i = row.find(c0);
    if (i >= 0) row._remove(i);
  }

  /** exceljs-style cell handle (allocated per call; prefer setCell/getValue in hot loops). */
  getCell(addr: string): CellRef;
  getCell(row: number, col: number | string): CellRef;
  getCell(a: Addr, b?: Addr): CellRef {
    const [r0, c0] = this._rc(a, b);
    checkCell(r0, c0);
    return new CellRef(this, r0 + 1, c0 + 1);
  }

  // ---- rows ----

  /**
   * Append a row after the last existing row. Arrays map index 0 → column A; objects map
   * keys through `ws.columns[].key` (unknown keys ignored). Returns the 1-based row number.
   */
  addRow(
    values: readonly CellValue[] | Readonly<Record<string, CellValue>>,
    style?: Style | number,
  ): number {
    const r0 = this._rows.length;
    if (r0 >= MAX_ROWS) throw new RangeError('Worksheet row limit reached');
    this._fillRow(r0, values, style);
    return r0 + 1;
  }

  /** Append several rows (see addRow). */
  addRows(
    rows: Iterable<readonly CellValue[] | Readonly<Record<string, CellValue>>>,
    style?: Style | number,
  ): void {
    const xf = style === undefined ? undefined : this._xf(style);
    for (const r of rows) this.addRow(r, xf);
  }

  /** Row handle (allocated per call). */
  getRow(n: number): RowView {
    checkCell(n - 1, 0);
    return new RowView(this, n - 1);
  }

  getRowProps(n: number): RowProps {
    const row = this._rows[n - 1];
    const p: RowProps = {};
    if (row) {
      if (row.height) p.height = row.height;
      if (row.hidden) p.hidden = true;
      if (row.outlineLevel) p.outlineLevel = row.outlineLevel;
      if (row.xf) p.style = this._styles.get(row.xf);
    }
    return p;
  }

  /** Merge row properties (only keys present in `props` change; style undefined = keep). */
  setRowProps(n: number, props: RowProps): void {
    checkCell(n - 1, 0);
    const row = this._rows[n - 1] ?? this._newRow(n - 1, 0);
    if ('height' in props) row.height = props.height ?? 0;
    if ('hidden' in props) row.hidden = props.hidden === true;
    if ('outlineLevel' in props) row.outlineLevel = props.outlineLevel ?? 0;
    if (props.styleIndex !== undefined) row.xf = props.styleIndex;
    else if (props.style !== undefined) row.xf = this._styles.intern(props.style);
  }

  /** Visit rows that contain cells, ascending. */
  eachRow(cb: (row: RowView, rowNumber: number) => void): void {
    const rows = this._rows;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (row !== undefined && row.n > 0) cb(new RowView(this, r), r + 1);
    }
  }

  // ---- columns ----

  /** Column definitions; assigning writes a header row (row 1) when any def has a header. */
  get columns(): Column[] {
    const out: Column[] = [];
    for (let i = 0; i < this._cols.length; i++) out.push(this.getColumn(i + 1));
    return out;
  }
  set columns(defs: readonly ColumnDef[]) {
    this._cols = [];
    this._colXf = undefined;
    let headers = false;
    for (let i = 0; i < defs.length; i++) {
      const d = defs[i]!;
      const c = new Column(this, i + 1);
      this._cols[i] = c;
      c.header = d.header;
      c.width = d.width;
      c.hidden = d.hidden === true;
      c.outlineLevel = d.outlineLevel ?? 0;
      if (d.style) c.style = d.style;
      if (d.key !== undefined) c.key = d.key;
      if (d.header !== undefined) headers = true;
    }
    this._rebuildKeys();
    if (headers)
      this._fillRow(
        0,
        defs.map((d) => d.header ?? null),
        undefined,
      );
  }

  /** Column by 1-based number, letter ("C") or key (keys take precedence over letters). */
  getColumn(c: number | string): Column {
    let i: number;
    if (typeof c === 'number') i = c - 1;
    else i = this._keyMap.get(c) ?? decodeCol(c);
    if (!(i >= 0 && i < MAX_COLS)) throw new RangeError(`Column out of range: ${c}`);
    let col = this._cols[i];
    if (col === undefined) {
      col = new Column(this, i + 1);
      this._cols[i] = col;
    }
    return col;
  }

  // ---- merges ----

  /** Merge a range ("A1:C2" or 1-based top, left, bottom, right). Throws on overlap. */
  mergeCells(range: string): void;
  mergeCells(top: number, left: number, bottom: number, right: number): void;
  mergeCells(a: string | number, l?: number, b?: number, r?: number): void {
    const m =
      typeof a === 'string'
        ? decodeRange(a)
        : {
            top: Math.min(a, b!) - 1,
            left: Math.min(l!, r!) - 1,
            bottom: Math.max(a, b!) - 1,
            right: Math.max(l!, r!) - 1,
          };
    checkCell(m.top, m.left);
    checkCell(m.bottom, m.right);
    if (m.top === m.bottom && m.left === m.right) throw new Error('Cannot merge a single cell');
    const x = this._mergeOverlap(m);
    if (x) throw new Error(`Merge ${encodeRange(m)} overlaps ${encodeRange(x)}`);
    this.addMergeRaw(m);
  }

  /**
   * Add a merge (0-based) WITHOUT validation (no bounds, single-cell or overlap checks) — for
   * readers and features that copy merges from a trusted source. Overlapping merges produce
   * files Excel repairs.
   */
  addMergeRaw(m: CellRange): void {
    this._merges.push(m);
    const idx = this._mergeIdx;
    if (idx) indexMerge(idx, m);
  }

  /** Remove every merge intersecting the range. */
  unmergeCells(range: string): void {
    const m = decodeRange(range);
    const before = this._merges.length;
    for (let i = this._merges.length - 1; i >= 0; i--) {
      if (overlaps(this._merges[i]!, m)) this._merges.splice(i, 1);
    }
    if (this._merges.length !== before) this._mergeIdx = undefined;
  }

  /** First merge intersecting m, if any. Row-bucket index: O(merges near m), not O(all). */
  private _mergeOverlap(m: CellRange): CellRange | undefined {
    const all = this._merges;
    if (all.length < 16) {
      for (const x of all) if (overlaps(x, m)) return x;
      return undefined;
    }
    let idx = this._mergeIdx;
    if (!idx) {
      idx = new Map();
      for (const x of all) indexMerge(idx, x);
      this._mergeIdx = idx;
    }
    for (let b = m.top >> MERGE_BUCKET; b <= m.bottom >> MERGE_BUCKET; b++) {
      const list = idx.get(b);
      if (list) for (const x of list) if (overlaps(x, m)) return x;
    }
    return undefined;
  }

  /** Merged ranges as A1 strings. */
  get merges(): string[] {
    return this._merges.map(encodeRange);
  }

  // ---- views ----

  /** Freeze the first `rows` rows and `cols` columns (freeze(1) = header row). */
  freeze(rows: number, cols = 0): void {
    const v: WorksheetView = { ...this.views[0] };
    v.state = 'frozen';
    v.xSplit = cols > 0 ? cols : undefined;
    v.ySplit = rows > 0 ? rows : undefined;
    v.topLeftCell = encodeAddr(rows, cols);
    this.views[0] = v;
  }

  // ---- extents ----

  /** Last row number with any record (cells or row properties); 0 when empty. */
  get rowCount(): number {
    const rows = this._rows;
    let r = rows.length;
    while (r > 0 && rows[r - 1] === undefined) r--;
    return r;
  }

  /** Last column number containing a cell; 0 when empty. */
  get columnCount(): number {
    const d = this.dimensions;
    return d ? d.right : 0;
  }

  /** Used range of cells (1-based), or undefined when the sheet has no cells. */
  get dimensions(): Dimensions | undefined {
    const rows = this._rows;
    let top = -1;
    let bottom = -1;
    let left = MAX_COLS;
    let right = -1;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (row === undefined || row.n === 0) continue;
      if (top < 0) top = r;
      bottom = r;
      const a = row.col[0]!;
      const b = row.col[row.n - 1]!;
      if (a < left) left = a;
      if (b > right) right = b;
    }
    if (top < 0) return undefined;
    return {
      top: top + 1,
      left: left + 1,
      bottom: bottom + 1,
      right: right + 1,
      ref: encodeRange({ top, left, bottom, right }),
    };
  }

  // ---- internals ----

  /** @internal */
  _rebuildKeys(): void {
    const pairs: [string, number][] = [];
    const map = new Map<string, number>();
    for (let i = 0; i < this._cols.length; i++) {
      const k = this._cols[i]?.key;
      if (k !== undefined) {
        pairs.push([k, i]);
        map.set(k, i);
      }
    }
    this._keyCols = pairs;
    this._keyMap = map;
  }

  /** @internal */
  _setColXf(c0: number, xf: number): void {
    if (this._colXf === undefined) {
      if (xf === 0) return;
      this._colXf = new Uint32Array(MAX_COLS);
    }
    this._colXf[c0] = xf;
  }

  /** @internal Value at slot i. */
  _valueAt(row: RowData, i: number): CellValue {
    switch (row.kind[i]) {
      case CellKind.Number:
        return row.num[i]!;
      case CellKind.Boolean:
        return row.num[i] !== 0;
      case CellKind.Date:
        return serialToDate(row.num[i]!, this.workbook.date1904);
      case CellKind.Error:
        return { error: row.obj![i] as ErrorCode };
      case CellKind.String:
      case CellKind.Formula:
      case CellKind.RichText:
      case CellKind.Hyperlink:
        return row.obj![i] as CellValue;
      default:
        return null;
    }
  }

  private _rc(a: Addr, b: Addr | undefined): [number, number] {
    if (typeof a === 'string') {
      const p = decodeAddr(a);
      return [p.row, p.col];
    }
    return [a - 1, this._colIndex(b!)];
  }

  /** 1-based number or column letter → 0-based column. */
  private _colIndex(c: Addr): number {
    return typeof c === 'number' ? c - 1 : decodeCol(c);
  }

  private _xf(s: Style | number | undefined): number {
    return typeof s === 'number' ? s : this._styles.intern(s);
  }

  private _xfAt(r0: number, c0: number): number {
    const row = this._rows[r0];
    if (row !== undefined) {
      const i = row.find(c0);
      if (i >= 0) return row.style[i]!;
      if (row.xf) return row.xf;
    }
    return this._colXf ? this._colXf[c0]! : 0;
  }

  private _dateXf(xf: number, serial: number): number {
    return this._styles.withDefaultNumFmt(xf, serial % 1 === 0 ? DATE_FMT : DATETIME_FMT);
  }

  private _newRow(r0: number, cap: number): RowData {
    const row = new RowData(r0, cap);
    this._rows[r0] = row;
    return row;
  }

  /** Store the classified value (K/N/O) into slot i with style xf. */
  private _store(row: RowData, i: number, xf: number): void {
    row.kind[i] = K;
    row.num[i] = N;
    row.style[i] = K === CellKind.Date ? this._dateXf(xf, N) : xf;
    if (O !== undefined) {
      if (row.obj === undefined) row.obj = [];
      row.obj[i] = O;
    } else if (row.obj !== undefined && i < row.obj.length) row.obj[i] = undefined;
  }

  private _set(r0: number, c0: number, value: CellValue, style: Style | number | undefined): void {
    checkCell(r0, c0);
    if (value === null || value === undefined) {
      this._clear(r0, c0, style);
      return;
    }
    classify(value, this.workbook.date1904);
    const row = this._rows[r0] ?? this._newRow(r0, 1);
    let i = row.find(c0);
    let xf: number;
    if (i < 0) {
      i = -i - 1;
      row._insert(i, c0);
      xf = style !== undefined ? this._xf(style) : row.xf || (this._colXf ? this._colXf[c0]! : 0);
    } else {
      xf = style !== undefined ? this._xf(style) : row.style[i]!;
    }
    this._store(row, i, xf);
  }

  private _clear(r0: number, c0: number, style: Style | number | undefined): void {
    const row = this._rows[r0];
    const i = row ? row.find(c0) : -1;
    if (row && i >= 0) {
      const xf = style !== undefined ? this._xf(style) : row.style[i]!;
      if (xf === 0) row._remove(i);
      else {
        row.kind[i] = CellKind.Empty;
        row.num[i] = 0;
        row.style[i] = xf;
        if (row.obj !== undefined && i < row.obj.length) row.obj[i] = undefined;
      }
    } else if (style !== undefined) {
      this.setStyle(r0 + 1, c0 + 1, style);
    }
  }

  private _fillRow(
    r0: number,
    values: readonly CellValue[] | Readonly<Record<string, CellValue>>,
    style: Style | number | undefined,
  ): void {
    const existing = this._rows[r0];
    if (existing !== undefined && existing.n > 0) {
      // slow path: merge into an existing row
      if (Array.isArray(values)) {
        const arr = values as readonly CellValue[];
        for (let i = 0; i < arr.length; i++) if (arr[i] != null) this._set(r0, i, arr[i], style);
      } else {
        const obj = values as Readonly<Record<string, CellValue>>;
        for (const [k, c] of this._keyCols) if (obj[k] != null) this._set(r0, c, obj[k], style);
      }
      return;
    }
    const xfArg = style === undefined ? -1 : this._xf(style);
    const d1904 = this.workbook.date1904;
    const colXf = this._colXf;
    if (Array.isArray(values)) {
      const arr = values as readonly CellValue[];
      const len = arr.length;
      if (len > MAX_COLS) throw new RangeError(`Row has more than ${MAX_COLS} columns`);
      const row = existing ?? this._newRow(r0, 0);
      if (row.col.length < len) row._alloc(len);
      const rowXf = row.xf;
      for (let c = 0; c < len; c++) {
        const v = arr[c];
        if (v === null || v === undefined) continue;
        classify(v, d1904);
        const i = row.n;
        row.col[i] = c;
        this._store(row, i, xfArg >= 0 ? xfArg : rowXf || (colXf ? colXf[c]! : 0));
        row.n = i + 1;
      }
    } else {
      const obj = values as Readonly<Record<string, CellValue>>;
      const keys = this._keyCols;
      const row = existing ?? this._newRow(r0, 0);
      if (row.col.length < keys.length) row._alloc(keys.length);
      const rowXf = row.xf;
      for (let k = 0; k < keys.length; k++) {
        const [key, c] = keys[k]!;
        const v = obj[key];
        if (v === null || v === undefined) continue;
        classify(v, d1904);
        const i = row.n;
        row.col[i] = c;
        this._store(row, i, xfArg >= 0 ? xfArg : rowXf || (colXf ? colXf[c]! : 0));
        row.n = i + 1;
      }
    }
  }
}

// ---- handles -------------------------------------------------------------------------

/** Lightweight row handle (1-based column arguments). */
export class RowView {
  readonly worksheet: Worksheet;
  /** 0-based row index. */
  readonly index: number;

  constructor(ws: Worksheet, index: number) {
    this.worksheet = ws;
    this.index = index;
  }

  /** 1-based row number. */
  get number(): number {
    return this.index + 1;
  }

  /** Underlying record (undefined if the row does not exist yet). */
  get data(): RowData | undefined {
    return this.worksheet.rowData[this.index];
  }

  get cellCount(): number {
    return this.data?.n ?? 0;
  }

  /** exceljs-style values: sparse array indexed by 1-based column (index 0 unused). */
  get values(): CellValue[] {
    const out: CellValue[] = [];
    const row = this.data;
    if (row) for (let i = 0; i < row.n; i++) out[row.col[i]! + 1] = this.worksheet._valueAt(row, i);
    return out;
  }

  getValue(col: number | string): CellValue {
    return this.worksheet.getValue(this.index + 1, col);
  }

  setValue(col: number | string, value: CellValue, style?: Style | number): void {
    this.worksheet.setCell(this.index + 1, col, value, style);
  }

  getCell(col: number | string): CellRef {
    return this.worksheet.getCell(this.index + 1, col);
  }

  /** Visit non-empty cells: (value, 1-based column number, xf). */
  eachCell(cb: (value: CellValue, colNumber: number, xf: number) => void): void {
    const row = this.data;
    if (!row) return;
    for (let i = 0; i < row.n; i++) {
      if (row.kind[i] !== CellKind.Empty) {
        cb(this.worksheet._valueAt(row, i), row.col[i]! + 1, row.style[i]!);
      }
    }
  }

  get height(): number | undefined {
    return this.worksheet.getRowProps(this.number).height;
  }
  set height(h: number | undefined) {
    this.worksheet.setRowProps(this.number, { height: h });
  }
  get hidden(): boolean {
    return this.worksheet.getRowProps(this.number).hidden === true;
  }
  set hidden(v: boolean) {
    this.worksheet.setRowProps(this.number, { hidden: v });
  }
  get outlineLevel(): number {
    return this.worksheet.getRowProps(this.number).outlineLevel ?? 0;
  }
  set outlineLevel(v: number) {
    this.worksheet.setRowProps(this.number, { outlineLevel: v });
  }
}

/** exceljs-style cell handle; reads/writes go straight to the worksheet storage. */
export class CellRef {
  readonly worksheet: Worksheet;
  /** 1-based. */
  readonly row: number;
  /** 1-based. */
  readonly col: number;

  constructor(ws: Worksheet, row: number, col: number) {
    this.worksheet = ws;
    this.row = row;
    this.col = col;
  }

  get address(): string {
    return encodeAddr(this.row - 1, this.col - 1);
  }
  get kind(): CellKind {
    return this.worksheet.getKind(this.row, this.col);
  }
  get value(): CellValue {
    return this.worksheet.getValue(this.row, this.col);
  }
  set value(v: CellValue) {
    this.worksheet.setCell(this.row, this.col, v);
  }
  get style(): Style {
    return this.worksheet.getStyle(this.row, this.col);
  }
  set style(s: Style | undefined) {
    this.worksheet.setStyle(this.row, this.col, s);
  }

  private _patch(p: Style): void {
    this.worksheet.setStyle(this.row, this.col, { ...this.style, ...p });
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }
  set numFmt(v: string | undefined) {
    this._patch({ numFmt: v });
  }
  get font(): Font | undefined {
    return this.style.font;
  }
  set font(v: Font | undefined) {
    this._patch({ font: v });
  }
  get fill(): Fill | undefined {
    return this.style.fill;
  }
  set fill(v: Fill | undefined) {
    this._patch({ fill: v });
  }
  get border(): Border | undefined {
    return this.style.border;
  }
  set border(v: Border | undefined) {
    this._patch({ border: v });
  }
  get alignment(): Alignment | undefined {
    return this.style.alignment;
  }
  set alignment(v: Alignment | undefined) {
    this._patch({ alignment: v });
  }
  get protection(): Protection | undefined {
    return this.style.protection;
  }
  set protection(v: Protection | undefined) {
    this._patch({ protection: v });
  }
}
