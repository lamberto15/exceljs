// Table editing (ExcelJS `table.addColumn / removeColumns / removeRows / getColumn(i).name /
// headerRow / totalsRow` + `commit()`), as free functions on `ws.tables`.
//
//   addTableColumn(ws, 'Sales', { name: 'Tax', totalsRowFunction: 'sum' }, [1, 2, 3], 1);
//   removeTableRows(ws, 'Sales', 0, 2);
//   renameTableColumn(ws, 'Sales', 'Tax', 'VAT');
//   setTableTotalsRow(ws, 'Sales', true);
//
// As in ExcelJS (and Excel's table commands), the table keeps its top-left cell and only cells
// inside the table move: removing rows / columns / the header or totals row shifts the table's
// remaining cells up or left and clears the cells it no longer covers; adding a column, a header
// row or a totals row grows the table right / down into cells that must be empty. Moved cells
// keep their value and style; formulas are moved as they are (references TO moved cells are not
// adjusted — use structured references, or insertRows / deleteRows for sheet-wide edits).
// Renaming a column also renames it in structured references (`Table[Old]`, `Table[@Old]`,
// `Table[[#Totals],[Old]]`, unqualified `[@Old]` inside the table) in every formula of the
// workbook, in calculated column / totals formulas and in defined names.

import { type CellRange, decodeRange, encodeAddr, encodeRange } from '../../core/address.js';
import type { CellValue, CellValueInput, Formula } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import {
  dataRows,
  getTable,
  type Table,
  type TableColumn,
  totalsCell,
  validateTable,
} from './tables.js';

// ---------------------------------------------------------------- cell moves (0-based)

function moveCell(ws: Worksheet, r1: number, c1: number, r2: number, c2: number): void {
  const v = ws.getValue(r1 + 1, c1 + 1);
  const xf = ws.getStyleIndex(r1 + 1, c1 + 1);
  if (v === null || v === undefined) ws.clearCell(r2 + 1, c2 + 1);
  else ws.setCell(r2 + 1, c2 + 1, v as CellValueInput);
  if (xf) ws.setStyle(r2 + 1, c2 + 1, xf);
}

function clearBlock(ws: Worksheet, top: number, left: number, bottom: number, right: number): void {
  for (let y = top; y <= bottom; y++)
    for (let c = left; c <= right; c++) ws.clearCell(y + 1, c + 1);
}

/** Throws unless `r` (cells the table grows into) holds no values, merges or other tables. */
function ensureFree(ws: Worksheet, t: Table, r: CellRange): void {
  const who = `Table "${t.name}"`;
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++) {
      const v = ws.getValue(y + 1, c + 1);
      if (v !== null && v !== undefined) {
        throw new Error(`${who}: cannot grow into ${encodeAddr(y, c)}, which is not empty`);
      }
    }
  }
  const hit = (a: CellRange) =>
    a.top <= r.bottom && r.top <= a.bottom && a.left <= r.right && r.left <= a.right;
  for (const m of ws.mergeRanges) {
    if (hit(m)) throw new Error(`${who}: cannot grow into merged cells ${encodeRange(m)}`);
  }
  for (const o of ws.tables ?? []) {
    if (o !== t && hit(decodeRange(o.ref))) {
      throw new Error(`${who}: cannot grow into table "${o.name}"`);
    }
  }
  if (r.bottom >= 1048576 || r.right >= 16384) throw new Error(`${who}: no room on the sheet`);
}

/** Shift the table's rows (below `y0`) up by n, inside the table columns; clears the bottom n. */
function shiftUp(ws: Worksheet, r: CellRange, y0: number, n: number): void {
  for (let y = y0; y + n <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++) moveCell(ws, y + n, c, y, c);
  }
  clearBlock(ws, r.bottom - n + 1, r.left, r.bottom, r.right);
}

/** Shift rows y0..r.bottom down by one (into r.bottom + 1), inside the table columns. */
function shiftDown(ws: Worksheet, r: CellRange, y0: number): void {
  for (let y = r.bottom; y >= y0; y--) {
    for (let c = r.left; c <= r.right; c++) moveCell(ws, y, c, y + 1, c);
  }
  for (let c = r.left; c <= r.right; c++) ws.clearCell(y0 + 1, c + 1);
}

function columnIndex(t: Table, col: number | string): number {
  if (typeof col === 'number') {
    if (!Number.isInteger(col) || col < 0 || col >= t.columns.length) {
      throw new RangeError(`Table "${t.name}" has no column ${col} (0-based)`);
    }
    return col;
  }
  const k = col.toLowerCase();
  const i = t.columns.findIndex((c) => c.name.toLowerCase() === k);
  if (i < 0) throw new Error(`Table "${t.name}" has no column "${col}"`);
  return i;
}

// ---------------------------------------------------------------- columns

/**
 * Insert a column into table `name` at `index` (0-based within the table; default: after the
 * last column). The table grows one column to the right (those cells must be empty); its cells
 * from `index` on move right. Writes the header cell, `values` (data rows, top to bottom; a
 * calculated column fills missing ones) and the totals cell.
 */
export function addTableColumn(
  ws: Worksheet,
  name: string,
  column: TableColumn | string,
  values?: readonly CellValue[],
  index?: number,
): void {
  const t = getTable(ws, name);
  const col: TableColumn = typeof column === 'string' ? { name: column } : { ...column };
  const r = decodeRange(t.ref);
  const at = index ?? t.columns.length;
  if (!Number.isInteger(at) || at < 0 || at > t.columns.length) {
    throw new RangeError(`Table "${t.name}": column index ${at} out of range`);
  }
  if (t.columns.some((c) => c.name.toLowerCase() === col.name.toLowerCase())) {
    throw new Error(`Table "${t.name}": duplicate column name "${col.name}"`);
  }
  ensureFree(ws, t, { top: r.top, left: r.right + 1, bottom: r.bottom, right: r.right + 1 });
  const c0 = r.left + at;
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.right; c >= c0; c--) moveCell(ws, y, c, y, c + 1);
    ws.clearCell(y + 1, c0 + 1);
  }
  const old = t.ref;
  t.columns.splice(at, 0, col);
  t.ref = encodeRange({ ...r, right: r.right + 1 });
  try {
    validateTable(ws, t, false);
  } catch (e) {
    t.columns.splice(at, 1);
    t.ref = old;
    throw e;
  }
  const { first, last } = dataRows(t);
  if (t.headerRow !== false) ws.setCell(r.top + 1, c0 + 1, col.name);
  for (let y = first; y <= last; y++) {
    let v: CellValue = values?.[y - first];
    if ((v === undefined || v === null) && col.calculatedColumnFormula) {
      v = { formula: col.calculatedColumnFormula };
    }
    if (v !== undefined && v !== null) ws.setCell(y + 1, c0 + 1, v);
  }
  if (t.totalsRow === true) {
    const v = totalsCell(t.name, col);
    if (v !== undefined) ws.setCell(r.bottom + 1, c0 + 1, v);
  }
}

/**
 * Remove `count` columns of table `name` starting at `index` (0-based; at least one column must
 * remain). The table's cells right of them move left; the table shrinks from the right.
 */
export function removeTableColumns(ws: Worksheet, name: string, index: number, count = 1): void {
  const t = getTable(ws, name);
  if (!Number.isInteger(count) || count < 1) return;
  columnIndex(t, index);
  if (index + count > t.columns.length) {
    throw new RangeError(`Table "${t.name}": columns ${index}..${index + count - 1} out of range`);
  }
  if (count >= t.columns.length) throw new Error(`Table "${t.name}": a table needs a column`);
  const r = decodeRange(t.ref);
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.left + index; c + count <= r.right; c++) moveCell(ws, y, c + count, y, c);
  }
  clearBlock(ws, r.top, r.right - count + 1, r.bottom, r.right);
  const gone = new Set(t.columns.splice(index, count).map((c) => c.name.toLowerCase()));
  t.ref = encodeRange({ ...r, right: r.right - count });
  // references to the removed columns become #REF!, as in Excel
  eachFormula(ws, t, (f, inTable) =>
    mapStructuredRefs(f, [t.name, t.displayName ?? t.name], inTable, (n) =>
      gone.has(n.toLowerCase()) ? null : n,
    ),
  );
}

/**
 * Rename a column of table `name` (by 0-based index or current name): the column, its header
 * cell, and its structured references everywhere in the workbook (see the module notes).
 */
export function renameTableColumn(
  ws: Worksheet,
  name: string,
  column: number | string,
  newName: string,
): void {
  const t = getTable(ws, name);
  const i = columnIndex(t, column);
  const c = t.columns[i]!;
  const from = c.name;
  if (typeof newName !== 'string' || newName === '') {
    throw new Error(`Table "${t.name}": a column needs a name`);
  }
  if (t.columns.some((o, k) => k !== i && o.name.toLowerCase() === newName.toLowerCase())) {
    throw new Error(`Table "${t.name}": duplicate column name "${newName}"`);
  }
  if (from === newName) return;
  c.name = newName;
  const r = decodeRange(t.ref);
  if (t.headerRow !== false) ws.setCell(r.top + 1, r.left + i + 1, newName);
  eachFormula(ws, t, (f, inTable) =>
    renameStructuredRefs(f, [t.name, t.displayName ?? t.name], from, newName, inTable),
  );
}

// ---------------------------------------------------------------- rows

/**
 * Remove `count` data rows of table `name` starting at data row `index` (0-based, header
 * excluded; at least one data row must remain). The table's rows below move up (totals row
 * included); the table shrinks from the bottom.
 */
export function removeTableRows(ws: Worksheet, name: string, index: number, count = 1): void {
  const t = getTable(ws, name);
  if (!Number.isInteger(count) || count < 1) return;
  const { r, first, last } = dataRows(t);
  if (!Number.isInteger(index) || index < 0 || first + index + count - 1 > last) {
    throw new RangeError(
      `Table "${t.name}": data rows ${index}..${index + count - 1} out of range`,
    );
  }
  if (count >= last - first + 1) throw new Error(`Table "${t.name}": a table needs a data row`);
  shiftUp(ws, r, first + index, count);
  t.ref = encodeRange({ ...r, bottom: r.bottom - count });
}

/**
 * Turn the header row of table `name` on or off. Off: the header cells go and the rest of the
 * table moves up one row. On: the table moves down one row (the row below it must be empty) and
 * the column names are written as header cells.
 */
export function setTableHeaderRow(ws: Worksheet, name: string, on: boolean): void {
  const t = getTable(ws, name);
  const had = t.headerRow !== false;
  if (had === on) return;
  const r = decodeRange(t.ref);
  if (on) {
    ensureFree(ws, t, { top: r.bottom + 1, left: r.left, bottom: r.bottom + 1, right: r.right });
    shiftDown(ws, r, r.top);
    t.columns.forEach((c, i) => {
      ws.setCell(r.top + 1, r.left + i + 1, c.name);
    });
    t.ref = encodeRange({ ...r, bottom: r.bottom + 1 });
    delete t.headerRow;
  } else {
    if (r.bottom - r.top < 1 + (t.totalsRow === true ? 1 : 0)) {
      throw new Error(`Table "${t.name}": a table needs a data row`);
    }
    shiftUp(ws, r, r.top, 1);
    t.ref = encodeRange({ ...r, bottom: r.bottom - 1 });
    t.headerRow = false;
  }
}

/**
 * Turn the totals row of table `name` on or off. On: the table grows one row down (it must be
 * empty) holding each column's totals label / SUBTOTAL / custom formula. Off: the totals row's
 * cells are cleared and the table shrinks.
 */
export function setTableTotalsRow(ws: Worksheet, name: string, on: boolean): void {
  const t = getTable(ws, name);
  const had = t.totalsRow === true;
  if (had === on) return;
  const r = decodeRange(t.ref);
  if (on) {
    ensureFree(ws, t, { top: r.bottom + 1, left: r.left, bottom: r.bottom + 1, right: r.right });
    t.columns.forEach((c, i) => {
      const v = totalsCell(t.name, c);
      if (v !== undefined) ws.setCell(r.bottom + 2, r.left + i + 1, v);
    });
    t.ref = encodeRange({ ...r, bottom: r.bottom + 1 });
    t.totalsRow = true;
  } else {
    clearBlock(ws, r.bottom, r.left, r.bottom, r.right);
    t.ref = encodeRange({ ...r, bottom: r.bottom - 1 });
    delete t.totalsRow;
  }
}

// ---------------------------------------------------------------- structured references

const unesc = (s: string): string => s.replace(/'(.)/g, '$1');
const esc = (s: string): string => s.replace(/([[\]#'])/g, "'$1");

/** End index (exclusive) of the bracket group opening at s[i] === '[' (' escapes a char). */
function groupEnd(s: string, i: number): number {
  let depth = 0;
  for (let k = i; k < s.length; k++) {
    const ch = s[k];
    if (ch === "'") k++;
    else if (ch === '[') depth++;
    else if (ch === ']' && --depth === 0) return k + 1;
  }
  return s.length;
}

/**
 * Map the column items inside one structured reference's brackets (without the outer ones);
 * `col` gets each column name (unescaped) and returns its new name, the same one, or null for
 * a column that no longer exists (→ the whole reference becomes null).
 */
function mapInner(inner: string, col: (name: string) => string | null): string | null {
  let gone = false;
  const item = (raw: string) => {
    if (raw.startsWith('#')) return raw;
    const name = unesc(raw).trim();
    const to = col(name);
    if (to === null) gone = true;
    return to === null || to === name ? raw : esc(to);
  };
  const lead = inner.startsWith('@') ? '@' : '';
  const body = inner.slice(lead.length);
  let out = lead;
  if (!body.startsWith('[')) out += body === '' ? '' : item(body);
  else {
    for (let i = 0; i < body.length; i++) {
      if (body[i] !== '[') {
        out += body[i];
        continue;
      }
      let j = i + 1;
      while (j < body.length && body[j] !== ']') j += body[j] === "'" ? 2 : 1;
      out += `[${item(body.slice(i + 1, j))}]`;
      i = j;
    }
  }
  return gone ? null : out;
}

/**
 * @internal Rewrite the structured references to the table named one of `tables` in formula
 * text `f` (`unqualified`: also table-less ones such as `[@Col]`, for formulas inside the
 * table). `col` maps column names (see mapInner); a reference to a removed column becomes #REF!.
 */
export function mapStructuredRefs(
  f: string,
  tables: readonly string[],
  unqualified: boolean,
  col: (name: string) => string | null,
): string {
  const lower = tables.map((n) => n.toLowerCase());
  let out = '';
  let i = 0;
  while (i < f.length) {
    const ch = f[i]!;
    if (ch === '"') {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === '"') {
          if (f[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch !== '[') {
      out += ch;
      i++;
      continue;
    }
    const end = groupEnd(f, i);
    const m = /[\w.\\]*$/.exec(out)![0];
    const group = f.slice(i, end);
    if ((m !== '' && lower.includes(m.toLowerCase())) || (m === '' && unqualified)) {
      const inner = mapInner(group.slice(1, -1), col);
      out = inner === null ? `${out.slice(0, out.length - m.length)}#REF!` : `${out}[${inner}]`;
    } else out += group;
    i = end;
  }
  return out;
}

/** @internal Rename column `from` → `to` in structured references (see mapStructuredRefs). */
export function renameStructuredRefs(
  f: string,
  tables: readonly string[],
  from: string,
  to: string,
  unqualified: boolean,
): string {
  const k = from.toLowerCase();
  return mapStructuredRefs(f, tables, unqualified, (n) => (n.toLowerCase() === k ? to : n));
}

/** Apply `fix` to every formula of the workbook that may hold a structured reference. */
function eachFormula(ws: Worksheet, t: Table, fix: (f: string, inTable: boolean) => string): void {
  const r = decodeRange(t.ref);
  const wb = ws.workbook;
  for (const s of wb.worksheets) {
    for (const o of s.tables ?? []) {
      const own = o === t;
      for (const oc of o.columns) {
        if (oc.calculatedColumnFormula)
          oc.calculatedColumnFormula = fix(oc.calculatedColumnFormula, own);
        if (oc.totalsRowFormula) oc.totalsRowFormula = fix(oc.totalsRowFormula, own);
      }
    }
    const rows = s.rowData;
    for (let y = 0; y < rows.length; y++) {
      const row = rows[y];
      if (row === undefined) continue;
      for (let k = 0; k < row.n; k++) {
        if (row.kind[k] !== 6) continue;
        const f = row.obj![k] as Formula;
        if (!f.formula.includes('[')) continue;
        const c = row.col[k]!;
        const inside = s === ws && y >= r.top && y <= r.bottom && c >= r.left && c <= r.right;
        const nf = fix(f.formula, inside);
        if (nf !== f.formula) row.obj![k] = { ...f, formula: nf };
      }
    }
  }
  wb.definedNames = wb.definedNames.map((d) =>
    d.value.includes('[') ? { ...d, value: fix(d.value, false) } : d,
  );
}
