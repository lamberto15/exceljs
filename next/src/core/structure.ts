// Row / column structure edits: insert, delete, splice and duplicate rows and columns.
//
// References are adjusted the way Excel adjusts them when whole rows / columns are inserted or
// deleted (not the ExcelJS way, which only moves cells):
//
// - Every reference to the edited sheet, anywhere in the workbook, follows its cells: formulas
//   of all sheets (unqualified refs on the edited sheet, `Sheet!A1` / `'My Sheet'!A1` elsewhere),
//   array-formula ranges, defined names (print areas / titles too), chart series.
// - Absolute and relative references move alike (`$` only matters for copy/fill). A range that
//   spans an insertion point grows; a range losing some rows / columns shrinks; a reference
//   whose cells are all deleted becomes `#REF!`.
// - Sheet model data moves with its cells: row / column properties, merges (dropped when fully
//   deleted or reduced to one cell), conditional formatting and data validation ranges and their
//   formulas, tables (rows inserted inside a table extend it and get its calculated columns;
//   columns inserted inside one become new table columns), notes, picture and chart anchors,
//   `autoFilter`, page breaks, ignored-error ranges, formula-cell hyperlinks kept from a read file.
// - Not adjusted: content kept raw from a read file (`ws.preserved`: charts read from a file,
//   pivot tables, sparklines, threaded comments, ...), hyperlink targets ("#Sheet!A1"),
//   `ws.views` (frozen panes, selection), 3D (`Sheet1:Sheet3!A1`) and external references.
//
// Everything is validated before anything changes: an edit that would push cells off the grid
// or break a table throws and leaves the workbook untouched.

import type { Chart } from '../xlsx/features/charts.js';
import { toAnchor } from '../xlsx/features/drawing.js';
import type { CellPosition, ImageAnchor } from '../xlsx/features/images.js';
import type { Table, TableColumn } from '../xlsx/features/tables.js';
import { translateFormula } from '../xlsx/read/formula.js';
import {
  type CellRange,
  decodeCol,
  decodeRange,
  encodeCol,
  encodeRange,
  MAX_COLS,
  MAX_ROWS,
} from './address.js';
import type { CellValue, DefinedName, Formula, Style } from './types.js';
import { CellKind, Column, RowData, type RowObject, type Worksheet } from './worksheet.js';

type RowValues = readonly CellValue[] | RowObject;

/** Private Worksheet / Column state this module rewires (runtime-visible TS privates). */
interface WsState {
  _rows: (RowData | undefined)[];
  _cols: (Column | undefined)[];
  _colXf: Uint32Array | undefined;
  _merges: CellRange[];
  _mergeIdx: unknown;
  _rebuildKeys(): void;
  _fillRow(r0: number, values: RowValues, style: Style | number | undefined): void;
}
interface ColState {
  _key: string | undefined;
  _xf: number;
}
const st = (ws: Worksheet): WsState => ws as unknown as WsState;

// ---------------------------------------------------------------- edit model

/** One structural edit of one sheet: `n` > 0 inserts n rows/cols before `at`, < 0 deletes. */
interface Edit {
  readonly ws: Worksheet;
  /** Lowercased name of the edited sheet. */
  readonly name: string;
  readonly rows: boolean;
  /** 0-based. */
  readonly at: number;
  readonly n: number;
  readonly max: number;
  /** Insert only: ranges ending right above `at` grow too (template rows, see fillTemplate). */
  readonly grow?: boolean | undefined;
}

/** New index of p, or -1 when deleted / pushed off the grid. */
function mapPoint(e: Edit, p: number): number {
  if (p < e.at) return p;
  if (e.n > 0) return p + e.n < e.max ? p + e.n : -1;
  return p < e.at - e.n ? -1 : p + e.n;
}

/** New bounds of the inclusive span [a, b], or undefined when it is wholly deleted. */
function mapSpan(e: Edit, a: number, b: number): [number, number] | undefined {
  if (e.n > 0) {
    const a2 = mapPoint(e, a);
    if (a2 < 0) return undefined;
    const b2 = b >= e.at || (e.grow && b === e.at - 1 && a < e.at) ? b + e.n : b;
    return [a2, b2 < e.max ? b2 : e.max - 1];
  }
  const end = e.at - e.n;
  if (a >= e.at && b < end) return undefined;
  return [a < e.at ? a : a < end ? e.at : a + e.n, b < e.at ? b : b < end ? e.at - 1 : b + e.n];
}

/** Pre-edit index of the cell that ends up first in the span starting at a (it survives). */
function survivor(e: Edit, a: number): number {
  return e.n < 0 && a >= e.at && a < e.at - e.n ? e.at - e.n : a;
}

function mapRange(e: Edit, r: CellRange): CellRange | undefined {
  const s = e.rows ? mapSpan(e, r.top, r.bottom) : mapSpan(e, r.left, r.right);
  if (!s) return undefined;
  return e.rows
    ? { top: s[0], left: r.left, bottom: s[1], right: r.right }
    : { top: r.top, left: s[0], bottom: r.bottom, right: s[1] };
}

// ---------------------------------------------------------------- reference shifting

function isIdent(c: number): boolean {
  return (
    (c >= 65 && c <= 90) ||
    (c >= 97 && c <= 122) ||
    (c >= 48 && c <= 57) ||
    c === 36 /* $ */ ||
    c === 95 /* _ */ ||
    c === 46 /* . */ ||
    c === 92 /* \ */ ||
    c > 127
  );
}

function readIdent(f: string, i: number): number {
  while (i < f.length && isIdent(f.charCodeAt(i))) i++;
  return i;
}

/** Index after the quoted run starting at i (doubled quote = escaped quote). */
function skipQuoted(f: string, i: number): number {
  const q = f[i]!;
  let j = i + 1;
  for (;;) {
    const k = f.indexOf(q, j);
    if (k < 0) return f.length;
    if (f[k + 1] === q) j = k + 2;
    else return k + 1;
  }
}

/** Index after the bracket group starting at i (nested brackets: structured references). */
function skipBracket(f: string, i: number): number {
  let depth = 0;
  for (let j = i; j < f.length; j++) {
    const d = f.charCodeAt(j);
    if (d === 91) depth++;
    else if (d === 93 && --depth === 0) return j + 1;
  }
  return f.length;
}

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;
const COL = /^(\$?)([A-Za-z]{1,3})$/;
const ROW = /^(\$?)([0-9]{1,7})$/;

function colNum(letters: string): number {
  let v = 0;
  for (let i = 0; i < letters.length; i++) v = v * 26 + ((letters.charCodeAt(i) | 0x20) - 96);
  return v - 1;
}

interface Tok {
  /** Column / row index, -1 when the token has none (row-only / column-only). */
  c: number;
  r: number;
  ca: string;
  ra: string;
  letters: string;
}

/** Parse a cell ("$B3"), column ("B") or row ("3") token; undefined when it is none of them. */
function tok(t: string, kind: 0 | 1 | 2): Tok | undefined {
  if (kind === 0) {
    const m = CELL.exec(t);
    if (!m) return undefined;
    const c = colNum(m[2]!);
    const r = +m[4]! - 1;
    if (c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return undefined;
    return { c, r, ca: m[1]!, ra: m[3]!, letters: m[2]! };
  }
  const m = (kind === 1 ? COL : ROW).exec(t);
  if (!m) return undefined;
  if (kind === 1) {
    const c = colNum(m[2]!);
    return c < MAX_COLS ? { c, r: -1, ca: m[1]!, ra: '', letters: m[2]! } : undefined;
  }
  const r = +m[2]! - 1;
  return r >= 0 && r < MAX_ROWS ? { c: -1, r, ca: '', ra: m[1]!, letters: '' } : undefined;
}

function enc(t: Tok, c: number, r: number): string {
  return (
    (t.c >= 0 ? t.ca + (c === t.c ? t.letters : encodeCol(c)) : '') +
    (t.r >= 0 ? t.ra + (r + 1) : '')
  );
}

/** Shifted text of the cell token t, '#REF!' when deleted. */
function shiftCell(e: Edit, t: Tok): string {
  const p = mapPoint(e, e.rows ? t.r : t.c);
  if (p < 0) return '#REF!';
  return e.rows ? enc(t, t.c, p) : enc(t, p, t.r);
}

/** Shifted text of the range a:b (both tokens of one kind), '#REF!' when wholly deleted. */
function shiftPair(e: Edit, a: Tok, b: Tok): string {
  const pa = e.rows ? a.r : a.c;
  const pb = e.rows ? b.r : b.c;
  if (pa < 0) return `${enc(a, a.c, a.r)}:${enc(b, b.c, b.r)}`; // other axis only
  const swap = pa > pb;
  const s = mapSpan(e, swap ? pb : pa, swap ? pa : pb);
  if (!s) return '#REF!';
  const na = swap ? s[1] : s[0];
  const nb = swap ? s[0] : s[1];
  return e.rows ? `${enc(a, a.c, na)}:${enc(b, b.c, nb)}` : `${enc(a, na, a.r)}:${enc(b, nb, b.r)}`;
}

/**
 * Adjust the references to the edited sheet in formula text `f`. `own`: lowercased name of the
 * sheet the formula belongs to (unqualified references point there); undefined for defined
 * names, whose references are always qualified.
 */
function shiftRefs(f: string, own: string | undefined, e: Edit): string {
  const n = f.length;
  let out = '';
  let last = 0;
  // Parse the reference starting at p (its sheet: target); returns the index after it.
  const ref = (p: number, target: string | undefined): number => {
    const a = readIdent(f, p);
    if (a === p) return p;
    const hit = target === e.name;
    const t1 = f.slice(p, a);
    if (f.charCodeAt(a) === 58 /* : */) {
      const b = readIdent(f, a + 1);
      const nb = f.charCodeAt(b);
      if (b > a + 1 && nb !== 40 && nb !== 33 && nb !== 91) {
        const t2 = f.slice(a + 1, b);
        for (const k of [0, 1, 2] as const) {
          const x = tok(t1, k);
          const y = x && tok(t2, k);
          if (x && y) {
            if (hit) {
              const s = shiftPair(e, x, y);
              if (s !== `${t1}:${t2}`) {
                out += f.slice(last, p) + s;
                last = b;
              }
            }
            return b;
          }
        }
      }
    }
    const x = tok(t1, 0);
    if (x && hit) {
      const s = shiftCell(e, x);
      if (s !== t1) {
        out += f.slice(last, p) + s;
        last = a;
      }
    }
    return a;
  };
  let i = 0;
  let ext = false; // the next sheet-qualified reference follows "[1]": another workbook
  while (i < n) {
    const c = f.charCodeAt(i);
    if (c === 34 /* " */) {
      i = skipQuoted(f, i);
      continue;
    }
    if (c === 39 /* ' */) {
      const j = skipQuoted(f, i);
      if (f.charCodeAt(j) === 33 /* ! */) {
        const sheet = f.slice(i + 1, j - 1).replace(/''/g, "'");
        // "[1]Sheet" (external) and "A:B" (3D) names never equal a sheet name
        i = ref(j + 1, ext ? undefined : sheet.toLowerCase());
      } else i = j;
      ext = false;
      continue;
    }
    if (c === 91 /* [ */) {
      i = skipBracket(f, i);
      const d = f.charCodeAt(i);
      ext = isIdent(d) || d === 39;
      continue;
    }
    if (!isIdent(c)) {
      i++;
      continue;
    }
    const e1 = readIdent(f, i);
    const nx = f.charCodeAt(e1);
    if (nx === 40 /* ( */ || nx === 91 /* [ */) i = e1;
    else if (nx === 33 /* ! */) i = ref(e1 + 1, ext ? undefined : f.slice(i, e1).toLowerCase());
    else {
      const e2 = nx === 58 ? readIdent(f, e1 + 1) : e1;
      // 3D reference Sheet1:Sheet3!A1: left as is
      if (e2 > e1 + 1 && f.charCodeAt(e2) === 33) i = ref(e2 + 1, undefined);
      else i = ref(i, own);
    }
    ext = false;
  }
  return last === 0 ? f : out + f.slice(last);
}

/** Split at top-level commas (outside quotes). */
function splitList(f: string): string[] {
  const parts: string[] = [];
  let start = 0;
  for (let i = 0; i < f.length; i++) {
    const c = f.charCodeAt(i);
    if (c === 39 || c === 34) i = skipQuoted(f, i) - 1;
    else if (c === 44) {
      parts.push(f.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(f.slice(start));
  return parts;
}

const isRef = (s: string): boolean => s.endsWith('#REF!');

/** Space-separated ranges of the edited sheet ("A1:B2 D4"); deleted ones dropped. */
function shiftSqref(e: Edit, sqref: string): string {
  const out: string[] = [];
  for (const p of sqref.split(/\s+/)) {
    if (!p) continue;
    const s = shiftRefs(p, e.name, e);
    const k = s.indexOf(':');
    if (!isRef(s)) out.push(k > 0 && s.slice(0, k) === s.slice(k + 1) ? s.slice(0, k) : s);
  }
  return out.join(' ');
}

/**
 * Shift a conditional-formatting / data-validation formula. Its relative references are
 * relative to the top-left cell of the first range of `sqref`; when that cell is deleted the
 * formula is first re-expressed for the cell that becomes the new top-left.
 */
function shiftAnchored(e: Edit, f: string, sqref: string, own: string): string {
  const parts = sqref.split(/\s+/).filter(Boolean);
  if (parts.length === 0) return shiftRefs(f, own, e);
  const a = decodeRange(parts[0]!);
  for (const p of parts) {
    const r = decodeRange(p);
    if (!mapRange(e, r)) continue;
    const dr = e.rows ? survivor(e, r.top) - a.top : r.top - a.top;
    const dc = e.rows ? r.left - a.left : survivor(e, r.left) - a.left;
    return shiftRefs(translateFormula(f, dr, dc), own, e);
  }
  return shiftRefs(f, own, e);
}

// ---------------------------------------------------------------- the edit

function check(rows: boolean, at: number, count: number): void {
  const max = rows ? MAX_ROWS : MAX_COLS;
  if (
    !(Number.isInteger(at) && Number.isInteger(count) && at >= 0 && count >= 0 && at + count <= max)
  ) {
    throw new RangeError(`${rows ? 'Row' : 'Column'} range out of bounds: ${at + 1} (+${count})`);
  }
}

/** Last used index on the axis (cells, row records, column definitions), -1 when none. */
function lastUsed(ws: Worksheet, rows: boolean): number {
  const s = st(ws);
  if (rows) {
    for (let r = s._rows.length - 1; r >= 0; r--) if (s._rows[r] !== undefined) return r;
    return -1;
  }
  let m = -1;
  for (const row of s._rows)
    if (row !== undefined && row.n > 0) m = Math.max(m, row.col[row.n - 1]!);
  return m;
}

interface TablePlan {
  t: Table;
  /** Mapped ref, undefined = table removed. */
  ref: CellRange | undefined;
  totalsLost: boolean;
}

/** Validate table effects of e; throws before anything changes. */
function planTables(e: Edit): TablePlan[] {
  const plans: TablePlan[] = [];
  for (const t of e.ws.tables ?? []) {
    const r = decodeRange(t.ref);
    const ref = mapRange(e, r);
    let totalsLost = false;
    if (ref && e.rows && e.n < 0) {
      const end = e.at - e.n;
      const gone = (y: number): boolean => y >= e.at && y < end;
      if (t.headerRow !== false && gone(r.top)) {
        throw new Error(
          `Deleting rows ${e.at + 1}-${end} would delete the header row of table "${t.name}"`,
        );
      }
      totalsLost = t.totalsRow === true && gone(r.bottom);
      const data =
        ref.bottom -
        ref.top +
        1 -
        (t.headerRow !== false ? 1 : 0) -
        (t.totalsRow && !totalsLost ? 1 : 0);
      if (data < 1) {
        throw new Error(
          `Deleting rows ${e.at + 1}-${end} would leave table "${t.name}" without data rows`,
        );
      }
    }
    plans.push({ t, ref, totalsLost });
  }
  return plans;
}

function edit(ws: Worksheet, rows: boolean, at: number, n: number, grow = false): void {
  if (n === 0) return;
  const max = rows ? MAX_ROWS : MAX_COLS;
  const e: Edit = { ws, name: ws.name.toLowerCase(), rows, at, n, max, grow };
  if (n > 0) {
    const used = lastUsed(ws, rows);
    if (used >= at && used + n >= max) {
      throw new RangeError(
        `Inserting ${n} ${rows ? 'rows' : 'columns'} would push cells off the sheet`,
      );
    }
  }
  const tables = planTables(e);

  // formulas, charts and names everywhere (text only: independent of cell positions)
  for (const sheet of ws.workbook.worksheets) shiftSheetRefs(e, sheet);
  shiftNames(e);

  // cells, row / column properties
  if (rows) moveRows(e);
  else moveCols(e);

  // sheet data
  const s = st(ws);
  const merges: CellRange[] = [];
  for (const m of s._merges) {
    const r = mapRange(e, m);
    if (r && (r.top !== r.bottom || r.left !== r.right)) merges.push(r);
  }
  s._merges.length = 0;
  for (const m of merges) s._merges.push(m);
  s._mergeIdx = undefined;

  if (ws.autoFilter !== undefined) {
    const a = shiftRefs(ws.autoFilter, e.name, e);
    ws.autoFilter = isRef(a) ? undefined : a;
  }
  if (ws.conditionalFormatting) {
    ws.conditionalFormatting = ws.conditionalFormatting.filter((cf) => {
      const old = cf.ref;
      cf.ref = shiftSqref(e, old);
      for (const rule of cf.rules) {
        for (const v of Object.values(rule)) {
          if (!Array.isArray(v)) continue;
          for (let i = 0; i < v.length; i++) {
            const x = v[i];
            if (typeof x === 'string') v[i] = shiftAnchored(e, x, old, e.name);
            else if (
              x &&
              typeof x === 'object' &&
              x.type === 'formula' &&
              typeof x.value === 'string'
            ) {
              x.value = shiftAnchored(e, x.value, old, e.name);
            }
          }
        }
      }
      return cf.ref !== '';
    });
  }
  if (ws.dataValidations) {
    ws.dataValidations = ws.dataValidations.filter((dv) => {
      const old = dv.sqref;
      dv.sqref = shiftSqref(e, old);
      if (typeof dv.formula1 === 'string') dv.formula1 = shiftAnchored(e, dv.formula1, old, e.name);
      if (typeof dv.formula2 === 'string') dv.formula2 = shiftAnchored(e, dv.formula2, old, e.name);
      return dv.sqref !== '';
    });
  }
  const brk = rows ? 'rowBreaks' : 'colBreaks';
  const breaks = ws[brk];
  if (breaks) {
    // a break after row / column N moves with N and goes when N is deleted
    const out: number[] = [];
    for (const b of breaks) {
      const p = mapPoint(e, b - 1);
      if (p >= 0) out.push(p + 1);
    }
    ws[brk] = out;
  }
  if (ws.ignoredErrors) {
    ws.ignoredErrors = ws.ignoredErrors.filter((x) => {
      x.sqref = shiftSqref(e, x.sqref);
      return x.sqref !== '';
    });
  }
  if (ws.comments) {
    ws.comments = ws.comments.filter((c) => {
      c.ref = shiftRefs(c.ref, e.name, e);
      return !isRef(c.ref);
    });
  }
  if (ws.images) for (const p of ws.images) p.anchor = shiftAnchor(e, p.anchor, p.editAs);
  if (ws.charts)
    for (const c of ws.charts) c.anchor = shiftAnchor(e, toAnchor(c.anchor, ws), undefined);
  const links = ws.preserved?.links;
  if (links) {
    const kept = [];
    for (const l of links) {
      const ref = shiftRefs(l.ref, e.name, e);
      if (!isRef(ref)) kept.push({ ...l, ref, formula: remapped.get(l.formula) ?? l.formula });
    }
    ws.preserved = { ...ws.preserved!, links: kept.length ? kept : undefined };
  }
  remapped.clear();
  applyTables(e, tables);
}

/** Old → new Formula objects of the last shiftSheetRefs pass (kept links compare by identity). */
const remapped = /* @__PURE__ */ new Map<unknown, Formula>();

function shiftSheetRefs(e: Edit, sheet: Worksheet): void {
  const own = sheet.name.toLowerCase();
  const self = sheet === e.ws;
  for (const row of st(sheet)._rows) {
    if (row === undefined || row.obj === undefined) continue;
    const { kind, obj } = row;
    for (let i = 0; i < row.n; i++) {
      if (kind[i] !== CellKind.Formula) continue;
      const fo = obj[i] as Formula;
      if (!self && !fo.formula.toLowerCase().includes(e.name)) continue;
      const formula = shiftRefs(fo.formula, own, e);
      let array = fo.array;
      if (self && array !== undefined) {
        array = shiftRefs(array, own, e);
        if (isRef(array)) array = undefined;
      }
      if (formula !== fo.formula || array !== fo.array) {
        const nf: Formula = { ...fo, formula };
        if (array === undefined) delete (nf as { array?: string }).array;
        else (nf as { array?: string }).array = array;
        obj[i] = nf;
        if (self) remapped.set(fo, nf);
      }
    }
  }
  if (sheet.charts) sheet.charts = sheet.charts.filter((c) => shiftChart(e, c, own));
}

/** Shift a chart's series references; false when no series is left. */
function shiftChart(e: Edit, c: Chart, own: string): boolean {
  const ref = (s: string): string | undefined => {
    const r = shiftRefs(s, own, e);
    return isRef(r) ? undefined : r;
  };
  c.series = c.series.filter((s) => {
    const values = ref(s.values);
    if (values === undefined) return false;
    s.values = values;
    if (s.categories !== undefined) s.categories = ref(s.categories);
    if (s.xValues !== undefined) s.xValues = ref(s.xValues);
    if (typeof s.name === 'object') {
      const r = ref(s.name.ref);
      s.name = r === undefined ? undefined : { ref: r };
    } else if (typeof s.name === 'string' && s.name.startsWith('=')) s.name = ref(s.name);
    return true;
  });
  return c.series.length > 0;
}

function shiftNames(e: Edit): void {
  const wb = e.ws.workbook;
  const out: DefinedName[] = [];
  for (const d of wb.definedNames) {
    let value = shiftRefs(d.value, undefined, e);
    if (value !== d.value && d.name.startsWith('_xlnm.')) {
      // print area / titles: drop deleted parts, the name when none is left
      value = splitList(value)
        .filter((p) => !isRef(p))
        .join(',');
      if (!value) continue;
    }
    out.push(value === d.value ? d : { ...d, value });
  }
  wb.definedNames = out;
}

function setIndex(row: RowData, r: number): void {
  (row as { index: number }).index = r;
}

function moveRows(e: Edit): void {
  const R = st(e.ws)._rows;
  const { at, n } = e;
  if (at >= R.length) return;
  if (n > 0) {
    const len = R.length;
    R.length = len + n;
    for (let r = len - 1; r >= at; r--) {
      const row = R[r];
      R[r + n] = row;
      if (row) setIndex(row, r + n);
    }
    for (let r = at; r < at + n; r++) R[r] = undefined;
  } else {
    R.splice(at, -n);
    for (let r = at; r < R.length; r++) {
      const row = R[r];
      if (row) setIndex(row, r);
    }
  }
}

function moveCols(e: Edit): void {
  const s = st(e.ws);
  const { at, n } = e;
  const end = at - n;
  for (const row of s._rows) {
    if (row === undefined || row.n === 0 || row.col[row.n - 1]! < at) continue;
    let i = row.find(at);
    if (i < 0) i = -i - 1;
    const cnt = row.n;
    const col = row.col;
    if (n > 0) {
      for (let k = i; k < cnt; k++) col[k]! += n;
      continue;
    }
    let j = i;
    while (j < cnt && col[j]! < end) j++;
    if (j > i) {
      row.col.copyWithin(i, j, cnt);
      row.kind.copyWithin(i, j, cnt);
      row.num.copyWithin(i, j, cnt);
      row.style.copyWithin(i, j, cnt);
      if (row.obj !== undefined && row.obj.length > i) row.obj.splice(i, j - i);
      row.n = cnt - (j - i);
    }
    for (let k = i; k < row.n; k++) col[k]! += n;
  }
  // column definitions and default styles
  const cols: (Column | undefined)[] = [];
  s._cols.forEach((c, i) => {
    const j = mapPoint(e, i);
    if (c === undefined || j < 0) return;
    const nc = new Column(e.ws, j + 1);
    nc.header = c.header;
    nc.width = c.width;
    nc.hidden = c.hidden;
    nc.outlineLevel = c.outlineLevel;
    nc.collapsed = c.collapsed;
    const a = c as unknown as ColState;
    const b = nc as unknown as ColState;
    b._key = a._key;
    b._xf = a._xf;
    cols[j] = nc;
  });
  s._cols = cols;
  s._rebuildKeys();
  const xf = s._colXf;
  if (xf) {
    if (n > 0) {
      xf.copyWithin(at + n, at, MAX_COLS - n);
      xf.fill(0, at, at + n);
    } else {
      xf.copyWithin(at, end);
      xf.fill(0, MAX_COLS + n);
    }
  }
}

function shiftPos(e: Edit, p: CellPosition, isTo: boolean): CellPosition {
  const v = e.rows ? p.row : p.col;
  const off = (e.rows ? p.rowOff : p.colOff) ?? 0;
  let nv: number;
  let collapse = false;
  if (e.n > 0) {
    // a bottom-right corner exactly on the insertion line (no offset) stays put
    nv = v > e.at || (v === e.at && !(isTo && off === 0)) ? v + e.n : v;
  } else {
    const end = e.at - e.n;
    if (v < e.at) nv = v;
    else if (v >= end) nv = v + e.n;
    else {
      nv = e.at;
      collapse = true;
    }
  }
  if (nv === v) return p;
  const q = { ...p };
  if (e.rows) {
    q.row = nv;
    if (collapse) q.rowOff = 0;
  } else {
    q.col = nv;
    if (collapse) q.colOff = 0;
  }
  return q;
}

/** Shift a drawing anchor: two-cell anchors move and resize, one-cell anchors only move. */
function shiftAnchor(e: Edit, a: ImageAnchor, editAs: string | undefined): ImageAnchor {
  if (!('from' in a)) return a;
  const from = shiftPos(e, a.from, false);
  if (!('to' in a)) return { from, ext: a.ext };
  if (editAs === 'absolute') return a;
  if (editAs === 'oneCell') {
    const d = e.rows ? from.row - a.from.row : from.col - a.from.col;
    const to = e.rows ? { ...a.to, row: a.to.row + d } : { ...a.to, col: a.to.col + d };
    return { from, to };
  }
  return { from, to: shiftPos(e, a.to, true) };
}

/** Unique "ColumnN" name for a new table column. */
function newColumnName(cols: readonly TableColumn[]): string {
  const used = new Set(cols.map((c) => c.name.toLowerCase()));
  let k = cols.length + 1;
  while (used.has(`column${k}`)) k++;
  return `Column${k}`;
}

function applyTables(e: Edit, plans: TablePlan[]): void {
  const ws = e.ws;
  if (plans.length === 0) return;
  const kept: Table[] = [];
  for (const { t, ref, totalsLost } of plans) {
    if (!ref) continue;
    const old = decodeRange(t.ref);
    const header = t.headerRow !== false;
    t.ref = encodeRange(ref);
    if (totalsLost) t.totalsRow = false;
    if (!e.rows) {
      if (e.n < 0) {
        // drop the deleted columns' definitions
        const end = e.at - e.n;
        t.columns = t.columns.filter((_, i) => {
          const c = old.left + i;
          return c < e.at || c >= end;
        });
      } else if (e.at > old.left && e.at <= old.right) {
        const added: TableColumn[] = [];
        for (let k = 0; k < e.n; k++) {
          const c: TableColumn = { name: newColumnName([...t.columns, ...added]) };
          added.push(c);
          if (header) ws.setCell(ref.top + 1, e.at + k + 1, c.name);
        }
        t.columns.splice(e.at - old.left, 0, ...added);
      }
    } else if (e.n > 0) {
      // rows inserted inside the table (they extend its data area): fill calculated columns
      if (e.at > old.top && e.at <= old.bottom) {
        t.columns.forEach((c, i) => {
          const f = c.calculatedColumnFormula;
          if (f)
            for (let r = e.at; r < e.at + e.n; r++)
              ws.setCell(r + 1, ref.left + i + 1, { formula: f });
        });
      }
    }
    kept.push(t);
  }
  ws.tables = kept;
}

// ---------------------------------------------------------------- public API

function colArg(c: number | string): number {
  return typeof c === 'number' ? c : decodeCol(c) + 1;
}

function clearRow(ws: Worksheet, r0: number): void {
  const R = st(ws)._rows;
  if (r0 < R.length) R[r0] = undefined;
}

function fillRow(ws: Worksheet, r0: number, values: RowValues | undefined): void {
  if (values !== undefined) st(ws)._fillRow(r0, values, undefined);
}

export interface InsertRowsOptions {
  /**
   * Format the new rows like the row above or below the insertion point (Excel's "Format Same
   * As Above / Below"): row height and style, and each cell's style (also of empty cells).
   * Default: unformatted rows.
   */
  inheritStyle?: 'above' | 'below' | undefined;
}

/**
 * Insert rows before row `pos` (1-based); later rows move down and every reference follows
 * (see the module notes). `rows`: the new rows' values (arrays from column A, or objects keyed
 * by `ws.columns[].key`, like addRow), or a count of empty rows.
 */
export function insertRows(
  ws: Worksheet,
  pos: number,
  rows: number | readonly RowValues[],
  options: InsertRowsOptions = {},
): void {
  const count = typeof rows === 'number' ? rows : rows.length;
  check(true, pos - 1, count);
  edit(ws, true, pos - 1, count);
  const r0 = pos - 1;
  const inherit = options.inheritStyle;
  if (inherit) {
    const src = st(ws)._rows[inherit === 'above' ? r0 - 1 : r0 + count];
    if (src) for (let i = 0; i < count; i++) copyFormat(ws, src, r0 + i);
  }
  if (typeof rows !== 'number') for (let i = 0; i < count; i++) fillRow(ws, r0 + i, rows[i]);
}

/** Insert one row before row `pos` (see insertRows). */
export function insertRow(
  ws: Worksheet,
  pos: number,
  values?: RowValues,
  options?: InsertRowsOptions,
): void {
  insertRows(ws, pos, [values ?? []], options);
}

/** Delete `count` rows starting at row `start` (1-based); later rows move up. */
export function deleteRows(ws: Worksheet, start: number, count = 1): void {
  check(true, start - 1, count);
  edit(ws, true, start - 1, -count);
}

/**
 * ExcelJS `spliceRows`: at row `start` (1-based), replace `count` rows with `inserts`. The first
 * min(count, inserts.length) rows are overwritten in place (references to them stay); the rest
 * of `count` is deleted, or the remaining inserts are inserted, moving later rows and
 * references as in deleteRows / insertRows.
 */
export function spliceRows(
  ws: Worksheet,
  start: number,
  count: number,
  ...inserts: RowValues[]
): void {
  const r0 = start - 1;
  check(true, r0, Math.max(count, inserts.length));
  const k = Math.min(count, inserts.length);
  if (count > k) edit(ws, true, r0 + k, -(count - k));
  else if (inserts.length > k) edit(ws, true, r0 + k, inserts.length - k);
  for (let i = 0; i < k; i++) clearRow(ws, r0 + i);
  for (let i = 0; i < inserts.length; i++) fillRow(ws, r0 + i, inserts[i]);
}

/**
 * Insert columns before column `pos` (1-based number or letter); later columns move right.
 * `cols`: the new columns' values (each an array from row 1), or a count of empty columns.
 */
export function insertColumns(
  ws: Worksheet,
  pos: number | string,
  cols: number | readonly (readonly CellValue[])[],
): void {
  const c0 = colArg(pos) - 1;
  const count = typeof cols === 'number' ? cols : cols.length;
  check(false, c0, count);
  edit(ws, false, c0, count);
  if (typeof cols !== 'number') fillCols(ws, c0, cols);
}

/** Delete `count` columns starting at column `start` (1-based number or letter). */
export function deleteColumns(ws: Worksheet, start: number | string, count = 1): void {
  const c0 = colArg(start) - 1;
  check(false, c0, count);
  edit(ws, false, c0, -count);
}

/**
 * ExcelJS `spliceColumns`: at column `start`, replace `count` columns with `inserts` (each an
 * array of values from row 1). Same overwrite / delete / insert split as spliceRows.
 */
export function spliceColumns(
  ws: Worksheet,
  start: number | string,
  count: number,
  ...inserts: (readonly CellValue[])[]
): void {
  const c0 = colArg(start) - 1;
  check(false, c0, Math.max(count, inserts.length));
  const k = Math.min(count, inserts.length);
  if (count > k) edit(ws, false, c0 + k, -(count - k));
  else if (inserts.length > k) edit(ws, false, c0 + k, inserts.length - k);
  if (k > 0) {
    // replaced columns start fresh, like replaced rows: no width, style, key or header
    const s = st(ws);
    for (let c = c0; c < c0 + k; c++) s._cols[c] = undefined;
    s._colXf?.fill(0, c0, c0 + k);
    s._rebuildKeys();
    for (const row of s._rows) {
      if (row === undefined) continue;
      for (let c = c0; c < c0 + k; c++) {
        const i = row.find(c);
        if (i >= 0) row._remove(i);
      }
    }
  }
  fillCols(ws, c0, inserts);
}

function fillCols(ws: Worksheet, c0: number, cols: readonly (readonly CellValue[])[]): void {
  cols.forEach((values, k) => {
    for (let r = 0; r < values.length; r++) {
      const v = values[r];
      if (v !== null && v !== undefined) ws.setCell(r + 1, c0 + k + 1, v);
    }
  });
}

/** Copy row props and cell styles (not values) of src to row r0. */
function copyFormat(ws: Worksheet, src: RowData, r0: number): void {
  ws.setRowProps(r0 + 1, { height: src.height || undefined, styleIndex: src.xf || undefined });
  for (let i = 0; i < src.n; i++)
    if (src.style[i]) ws.setStyle(r0 + 1, src.col[i]! + 1, src.style[i]!);
}

/** Copy of src (values, styles, row properties) as row r0; formulas translated like a paste. */
function cloneRow(src: RowData, r0: number): RowData {
  const row = new RowData(r0, src.n);
  row.n = src.n;
  row.col.set(src.col.subarray(0, src.n));
  row.kind.set(src.kind.subarray(0, src.n));
  row.num.set(src.num.subarray(0, src.n));
  row.style.set(src.style.subarray(0, src.n));
  row.height = src.height;
  row.hidden = src.hidden;
  row.outlineLevel = src.outlineLevel;
  row.collapsed = src.collapsed;
  row.xf = src.xf;
  if (src.obj !== undefined) {
    const dr = r0 - src.index;
    row.obj = src.obj.slice(0, src.n);
    for (let i = 0; i < row.obj.length; i++) {
      if (src.kind[i] !== CellKind.Formula) continue;
      const fo = row.obj[i] as Formula;
      const nf: Formula = { ...fo, formula: translateFormula(fo.formula, dr, 0) };
      if (fo.array !== undefined)
        (nf as { array?: string }).array = translateFormula(fo.array, dr, 0);
      row.obj[i] = nf;
    }
  }
  return row;
}

/**
 * ExcelJS `duplicateRow`: copy row `rowNum` (values, styles, height, single-row merges) into
 * the `count` rows below it — inserted (moving later rows) when `insert` is true, else
 * overwriting them. Relative references in formulas are adjusted as when pasting the row.
 */
export function duplicateRow(ws: Worksheet, rowNum: number, count: number, insert = false): void {
  dupRow(ws, rowNum, count, insert, false);
}

/**
 * @internal duplicateRow; `grow`: ranges ending at the source row extend over the inserted
 * copies (fillTemplate: totals below a repeated row cover every copy).
 */
export function dupRow(
  ws: Worksheet,
  rowNum: number,
  count: number,
  insert: boolean,
  grow: boolean,
): void {
  const r0 = rowNum - 1;
  check(true, r0, count + 1);
  if (insert) edit(ws, true, r0 + 1, count, grow);
  const s = st(ws);
  const src = s._rows[r0];
  if (!insert) {
    for (let r = r0 + 1; r <= r0 + count; r++) clearRow(ws, r);
    const hit = s._merges.filter((m) => m.bottom > r0 && m.top <= r0 + count);
    for (const m of hit) ws.unmergeCells(encodeRange(m));
  }
  const rowMerges = s._merges.filter((m) => m.top === r0 && m.bottom === r0);
  for (let k = 1; k <= count; k++) {
    if (src) s._rows[r0 + k] = cloneRow(src, r0 + k);
    for (const m of rowMerges) ws.addMergeRaw({ ...m, top: r0 + k, bottom: r0 + k });
  }
}

export interface GroupOptions {
  /** Collapse the group: hide its rows / columns and flag the summary row / column. */
  collapsed?: boolean | undefined;
}

/**
 * Group rows `from`..`to` (1-based, inclusive) one outline level deeper (max 7), like Excel's
 * Data > Group. The summary row is the one below the group (Excel's default).
 */
export function groupRows(
  ws: Worksheet,
  from: number,
  to: number,
  options: GroupOptions = {},
): void {
  check(true, from - 1, to - from + 2);
  for (let r = from; r <= to; r++) {
    const level = Math.min(7, (ws.rowData[r - 1]?.outlineLevel ?? 0) + 1);
    ws.setRowProps(
      r,
      options.collapsed ? { outlineLevel: level, hidden: true } : { outlineLevel: level },
    );
  }
  if (options.collapsed) ws.setRowProps(to + 1, { collapsed: true });
}

/** Group columns `from`..`to` (numbers or letters) one outline level deeper; see groupRows. */
export function groupColumns(
  ws: Worksheet,
  from: number | string,
  to: number | string,
  options: GroupOptions = {},
): void {
  const a = colArg(from);
  const b = colArg(to);
  check(false, a - 1, b - a + 2);
  for (let c = a; c <= b; c++) {
    const col = ws.getColumn(c);
    col.outlineLevel = Math.min(7, col.outlineLevel + 1);
    if (options.collapsed) col.hidden = true;
  }
  if (options.collapsed) ws.getColumn(b + 1).collapsed = true;
}
