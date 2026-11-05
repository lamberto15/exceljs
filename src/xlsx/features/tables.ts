// Feature: tables / ListObjects (xl/tables/tableN.xml + <tableParts>), ECMA-376 §18.5.
//
//   addTable(ws, { name: 'Sales', topLeft: 'B2', columns: ['Region', { name: 'Total',
//     totalsRowFunction: 'sum' }], rows: [['North', 10], ['South', 20]], totalsRow: true });
//   ws.tables = [{ name: 'T', ref: 'A1:C5', columns: [{ name: 'a' }, ...] }];   // model only
//
// Model: `ws.tables`. `ref` includes the header row and the totals row. Excel REPAIRS files in
// which a column name differs from its header cell text, so `addTable` writes the header cells
// and the buffered writers check them. Streaming (`createXlsxWriter`): define the table on
// `streamSheet.worksheet` before commit (`ws.tables = [...]`, or `addTable` without `rows`) and
// add the header / data / totals rows yourself: the Worksheet holds no cells there, so the
// cell filling and the header check are unavailable.
// Writer validation: names unique workbook-wide (case-insensitive), valid identifiers, column
// names unique, columns match the ref width, no overlap with other tables or merged cells.
// Round trip: the table part is regenerated from the model, except for a table whose part has
// relationships of its own (query table → queryTable part): while that table is unchanged, its
// original part is written back (see read/preserve.ts).

import { type CellRange, decodeRange, encodeAddr, encodeRange } from '../../core/address.js';
import { insertRows } from '../../core/structure.js';
import type { StyleRegistry } from '../../core/styles.js';
import type { CellValue, Style } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import { parseXml } from '../../xml/index.js';
import { addXlfn, stripXlfn } from '../common/xlfn.js';
import { isValidName } from './defined-names.js';
import { type FeatureReader, type FeatureWriter, SheetSlot } from './types.js';
import { addPRel, attrs, bool, escapeText, num, presTable } from './util.js';

export type TotalsRowFunction =
  | 'none'
  | 'sum'
  | 'average'
  | 'count'
  | 'countNums'
  | 'max'
  | 'min'
  | 'stdDev'
  | 'var'
  | 'custom';

export interface TableColumn {
  /** Column name; must equal the header cell text and be unique within the table. */
  name: string;
  totalsRowFunction?: TotalsRowFunction | undefined;
  /** Text of the totals cell when no function is used (typically the first column). */
  totalsRowLabel?: string | undefined;
  /** Formula (without '=') for `totalsRowFunction: 'custom'`. */
  totalsRowFormula?: string | undefined;
  /** Formula (without '=') every data cell of a calculated column holds. */
  calculatedColumnFormula?: string | undefined;
  /** Formatting of this column's header cell (a dxf: `headerRowDxfId`). */
  headerStyle?: Style | undefined;
  /** Formatting of this column's data cells (`dataDxfId`). */
  dataStyle?: Style | undefined;
  /** Formatting of this column's totals cell (`totalsRowDxfId`). */
  totalsStyle?: Style | undefined;
}

export interface TableStyle {
  /** Table style name, e.g. "TableStyleMedium2" (default), "TableStyleLight1"; '' = none. */
  theme?: string | undefined;
  showFirstColumn?: boolean | undefined;
  showLastColumn?: boolean | undefined;
  /** Default true. */
  showRowStripes?: boolean | undefined;
  showColumnStripes?: boolean | undefined;
}

export interface Table {
  /** Unique (workbook-wide, case-insensitive) identifier-like name: no spaces, not "A1". */
  name: string;
  /** Defaults to `name`. */
  displayName?: string | undefined;
  /** Range including header and totals rows, e.g. "A1:D10". */
  ref: string;
  /** Header row present (default true). */
  headerRow?: boolean | undefined;
  /** Totals row present (last row of `ref`). */
  totalsRow?: boolean | undefined;
  columns: TableColumn[];
  style?: TableStyle | undefined;
  /** Filter buttons in the header (default true; requires a header row). */
  autoFilter?: boolean | undefined;
  /**
   * Differential formats applied by the table to the header row / data area / totals row (written
   * as dxfs, `headerRowDxfId` / `dataDxfId` / `totalsRowDxfId`; per-column ones live on the
   * columns). Only the listed properties override the table style.
   */
  headerStyle?: Style | undefined;
  dataStyle?: Style | undefined;
  totalsStyle?: Style | undefined;
}

/** A column as `addTable` takes it: a TableColumn plus the sheet column width to set. */
export interface AddTableColumn extends TableColumn {
  /** Width (characters) set on the sheet column; not part of the table model. */
  width?: number | undefined;
}

export interface AddTableOptions {
  name: string;
  displayName?: string | undefined;
  /** Full range ("A1:D10"), or give `topLeft` and let `rows` size the table. */
  ref?: string | undefined;
  /** Top-left cell (default "A1"); the range is derived from columns and rows. */
  topLeft?: string | undefined;
  /** Column names or definitions. */
  columns: (string | AddTableColumn)[];
  /** Data rows (arrays) written below the header; a calculated column fills missing cells. */
  rows?: readonly (readonly CellValue[])[] | undefined;
  headerRow?: boolean | undefined;
  totalsRow?: boolean | undefined;
  style?: TableStyle | undefined;
  autoFilter?: boolean | undefined;
  headerStyle?: Style | undefined;
  dataStyle?: Style | undefined;
  totalsStyle?: Style | undefined;
}

/** SUBTOTAL function numbers (ignoring hidden rows, like Excel's totals row). */
export const SUBTOTAL: Record<string, number> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};

/** Header cell text as Excel sees it, or undefined when the cell cannot be compared. */
function cellText(v: CellValue): string | undefined {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (typeof v === 'object' && 'richText' in v) return v.richText.map((r) => r.text).join('');
  return undefined;
}

function overlaps(a: CellRange, b: CellRange): boolean {
  return a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
}

/** @internal Structured reference to a column's data, e.g. `Sales[Total]` (specials escaped with '). */
export function colRef(table: string, col: string): string {
  return `${table}[${col.replace(/([[\]#'])/g, "'$1")}]`;
}

/** @internal Sheet-level validation of one table; throws an Error naming the table. */
export function validateTable(ws: Worksheet, t: Table, checkHeaderCells: boolean): void {
  const who = `Table "${t.name}"`;
  if (!isValidName(t.name)) throw new Error(`${who}: invalid table name`);
  if (t.displayName !== undefined && !isValidName(t.displayName)) {
    throw new Error(`${who}: invalid displayName "${t.displayName}"`);
  }
  const r = decodeRange(t.ref);
  const header = t.headerRow !== false;
  if (r.bottom - r.top + 1 < (header ? 1 : 0) + (t.totalsRow ? 1 : 0) + 1) {
    throw new Error(`${who}: ref ${t.ref} needs at least one data row`);
  }
  if (t.columns.length !== r.right - r.left + 1) {
    throw new Error(
      `${who}: ${t.columns.length} columns do not match the width of ref ${t.ref} (${r.right - r.left + 1})`,
    );
  }
  const seen = new Set<string>();
  for (let i = 0; i < t.columns.length; i++) {
    const c = t.columns[i]!;
    if (typeof c.name !== 'string' || c.name === '') {
      throw new Error(`${who}: column ${i + 1} has no name`);
    }
    const k = c.name.toLowerCase();
    if (seen.has(k)) throw new Error(`${who}: duplicate column name "${c.name}"`);
    seen.add(k);
    if (c.totalsRowFunction === 'custom' && !c.totalsRowFormula) {
      throw new Error(`${who}: column "${c.name}" has a custom totals function but no formula`);
    }
  }
  for (const m of ws.mergeRanges) {
    if (overlaps(r, m)) {
      throw new Error(`${who}: ref ${t.ref} overlaps merged cells ${encodeRange(m)}`);
    }
  }
  for (const o of ws.tables ?? []) {
    if (o !== t && overlaps(r, decodeRange(o.ref))) {
      throw new Error(`${who}: ref ${t.ref} overlaps table "${o.name}" (${o.ref})`);
    }
  }
  if (checkHeaderCells && header) {
    for (let i = 0; i < t.columns.length; i++) {
      const text = cellText(ws.getValue(r.top + 1, r.left + i + 1));
      const name = t.columns[i]!.name;
      if (text !== undefined && text !== name && !/[\r\n]|_x[0-9a-f]{4}_/i.test(name)) {
        throw new Error(
          `${who}: header cell ${encodeAddr(r.top, r.left + i)} is "${text}" but the column is named "${name}" (Excel would repair the file)`,
        );
      }
    }
  }
}

/** @internal Value of a column's totals cell: its SUBTOTAL / custom formula, or its label. */
export function totalsCell(table: string, c: TableColumn): CellValue {
  const fn = c.totalsRowFunction;
  if (fn === 'custom') return { formula: c.totalsRowFormula! };
  if (fn && fn !== 'none') return { formula: `SUBTOTAL(${SUBTOTAL[fn]},${colRef(table, c.name)})` };
  return c.totalsRowLabel;
}

/** Workbook-wide checks: unique table names, no clash with defined names. */
function validateWorkbook(ws: Worksheet, t: Table): void {
  const k = t.name.toLowerCase();
  for (const s of ws.workbook.worksheets) {
    for (const o of s.tables ?? []) {
      if (o !== t && o.name.toLowerCase() === k) {
        throw new Error(`Table name "${t.name}" is used more than once in the workbook`);
      }
    }
  }
  for (const d of ws.workbook.definedNames) {
    if (d.name.toLowerCase() === k) {
      throw new Error(`Table name "${t.name}" clashes with a defined name`);
    }
  }
}

/**
 * Create a table on `ws`, validate it and append it to `ws.tables`. With `rows` it also writes
 * the header cells, the data rows and the totals row cells (SUBTOTAL formulas / labels).
 * Throws on invalid names, overlaps, duplicate columns or rows that do not fit `ref`.
 */
export function addTable(ws: Worksheet, o: AddTableOptions): Table {
  const widths: (number | undefined)[] = [];
  const columns: TableColumn[] = o.columns.map((c, i) => {
    if (typeof c === 'string') return { name: c };
    const { width, ...col } = c;
    widths[i] = width;
    return col;
  });
  const header = o.headerRow !== false;
  const totals = o.totalsRow === true;
  const rows = o.rows;
  let ref = o.ref;
  if (ref === undefined) {
    const tl = decodeRange(`${o.topLeft ?? 'A1'}:${o.topLeft ?? 'A1'}`);
    const dataRows = Math.max(rows?.length ?? 0, 1);
    ref = encodeRange({
      top: tl.top,
      left: tl.left,
      bottom: tl.top + (header ? 1 : 0) + dataRows + (totals ? 1 : 0) - 1,
      right: tl.left + columns.length - 1,
    });
  }
  const t: Table = { name: o.name, ref, columns };
  if (o.displayName !== undefined) t.displayName = o.displayName;
  if (o.headerRow !== undefined) t.headerRow = o.headerRow;
  if (o.totalsRow !== undefined) t.totalsRow = o.totalsRow;
  if (o.style !== undefined) t.style = o.style;
  if (o.autoFilter !== undefined) t.autoFilter = o.autoFilter;
  if (o.headerStyle !== undefined) t.headerStyle = o.headerStyle;
  if (o.dataStyle !== undefined) t.dataStyle = o.dataStyle;
  if (o.totalsStyle !== undefined) t.totalsStyle = o.totalsStyle;
  validateTable(ws, t, false);
  validateWorkbook(ws, t);
  if (rows) {
    const r = decodeRange(ref);
    const first = r.top + (header ? 1 : 0);
    const last = r.bottom - (totals ? 1 : 0);
    if (rows.length > last - first + 1) {
      throw new Error(`Table "${t.name}": ${rows.length} rows do not fit in ref ${ref}`);
    }
    if (header) {
      for (let i = 0; i < columns.length; i++) {
        ws.setCell(r.top + 1, r.left + i + 1, columns[i]!.name);
      }
    }
    for (let y = first; y <= last; y++) {
      const row = rows[y - first];
      for (let i = 0; i < columns.length; i++) {
        let v: CellValue = row?.[i];
        const f = columns[i]!.calculatedColumnFormula;
        if ((v === undefined || v === null) && f) v = { formula: f };
        if (v !== undefined && v !== null) ws.setCell(y + 1, r.left + i + 1, v);
      }
    }
    if (totals) {
      for (let i = 0; i < columns.length; i++) {
        const v = totalsCell(t.name, columns[i]!);
        if (v !== undefined) ws.setCell(r.bottom + 1, r.left + i + 1, v);
      }
    }
  }
  if (ws.tables) ws.tables.push(t);
  else ws.tables = [t];
  const left = decodeRange(ref).left;
  widths.forEach((w, i) => {
    if (w !== undefined) ws.getColumn(left + i + 1).width = w;
  });
  return t;
}

/** The table named `name` on `ws` (case-insensitive); throws when there is none. */
export function getTable(ws: Worksheet, name: string): Table {
  const k = name.toLowerCase();
  const t = ws.tables?.find((x) => x.name.toLowerCase() === k);
  if (!t) throw new Error(`Worksheet '${ws.name}' has no table "${name}"`);
  return t;
}

/** @internal 0-based first and last data rows of a table (last < first: no data rows). */
export function dataRows(t: Table): { r: CellRange; first: number; last: number } {
  const r = decodeRange(t.ref);
  return {
    r,
    first: r.top + (t.headerRow !== false ? 1 : 0),
    last: r.bottom - (t.totalsRow === true ? 1 : 0),
  };
}

/**
 * The data rows of table `name` as objects keyed by column name (header and totals rows
 * excluded). Values are as stored: formula cells give their Formula object (`.result` holds the
 * cached value).
 */
export function tableRows(ws: Worksheet, name: string): Record<string, CellValue>[] {
  const t = getTable(ws, name);
  const { r, first, last } = dataRows(t);
  const out: Record<string, CellValue>[] = [];
  for (let y = first; y <= last; y++) {
    const o: Record<string, CellValue> = {};
    t.columns.forEach((c, i) => {
      o[c.name] = ws.getValue(y + 1, r.left + i + 1);
    });
    out.push(o);
  }
  return out;
}

/**
 * Append data rows to table `name` (above its totals row): sheet rows are inserted so cells
 * below the table move down and every reference follows (see insertRows), the table grows, and
 * calculated columns are filled where a row gives no value. Rows are arrays (index 0 = the
 * table's first column) or objects keyed by column name.
 */
export function addTableRows(
  ws: Worksheet,
  name: string,
  rows: readonly (readonly CellValue[] | Readonly<Record<string, CellValue>>)[],
): void {
  const t = getTable(ws, name);
  if (rows.length === 0) return;
  const { r, last } = dataRows(t);
  const at = last + 1; // 0-based row of the first new row
  insertRows(ws, at + 1, rows.length);
  if (t.totalsRow !== true) {
    // inserted right below the table: not inside it, so extend it here
    t.ref = encodeRange({ ...r, bottom: r.bottom + rows.length });
    t.columns.forEach((c, i) => {
      const f = c.calculatedColumnFormula;
      if (f)
        for (let k = 0; k < rows.length; k++)
          ws.setCell(at + k + 1, r.left + i + 1, { formula: f });
    });
  }
  rows.forEach((row, k) => {
    t.columns.forEach((c, i) => {
      const v = Array.isArray(row)
        ? (row as readonly CellValue[])[i]
        : (row as Record<string, CellValue>)[c.name];
      if (v !== undefined && v !== null) ws.setCell(at + k + 1, r.left + i + 1, v);
    });
  });
}

// ---------------------------------------------------------------- writing

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const TABLE_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

/** `headerRowDxfId` / `dataDxfId` / `totalsRowDxfId` attributes for the given styles. */
function dxfAttrs(
  styles: StyleRegistry,
  h: Style | undefined,
  d: Style | undefined,
  tot: Style | undefined,
  totals: boolean,
): string {
  return attrs({
    headerRowDxfId: h && styles.internDxf(h),
    dataDxfId: d && styles.internDxf(d),
    totalsRowDxfId: totals && tot ? styles.internDxf(tot) : undefined,
  });
}

function tableXml(t: Table, id: number, styles: StyleRegistry): string {
  const r = decodeRange(t.ref);
  const header = t.headerRow !== false;
  const totals = t.totalsRow === true;
  let s = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${NS}"${attrs({
    id,
    name: t.name,
    displayName: t.displayName ?? t.name,
    ref: t.ref,
  })} headerRowCount="${header ? 1 : 0}"${totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}${dxfAttrs(
    styles,
    header ? t.headerStyle : undefined,
    t.dataStyle,
    t.totalsStyle,
    totals,
  )}>`;
  if (header && t.autoFilter !== false) {
    s += `<autoFilter ref="${encodeRange({ ...r, bottom: totals ? r.bottom - 1 : r.bottom })}"/>`;
  }
  s += `<tableColumns count="${t.columns.length}">`;
  for (let i = 0; i < t.columns.length; i++) {
    const c = t.columns[i]!;
    const fn = totals ? c.totalsRowFunction : undefined;
    s += `<tableColumn${attrs({
      id: i + 1,
      name: c.name,
      totalsRowFunction: fn && fn !== 'none' ? fn : undefined,
      totalsRowLabel: totals && (!fn || fn === 'none') ? c.totalsRowLabel : undefined,
    })}${dxfAttrs(styles, header ? c.headerStyle : undefined, c.dataStyle, c.totalsStyle, totals)}`;
    const inner =
      (c.calculatedColumnFormula
        ? `<calculatedColumnFormula>${escapeText(addXlfn(c.calculatedColumnFormula))}</calculatedColumnFormula>`
        : '') +
      (fn === 'custom' && c.totalsRowFormula
        ? `<totalsRowFormula>${escapeText(addXlfn(c.totalsRowFormula))}</totalsRowFormula>`
        : '');
    s += inner ? `>${inner}</tableColumn>` : '/>';
  }
  const st = t.style;
  s += `</tableColumns><tableStyleInfo${attrs({
    name: st?.theme === undefined ? 'TableStyleMedium2' : st.theme || undefined,
  })} showFirstColumn="${st?.showFirstColumn ? 1 : 0}" showLastColumn="${st?.showLastColumn ? 1 : 0}" showRowStripes="${
    st?.showRowStripes === false ? 0 : 1
  }" showColumnStripes="${st?.showColumnStripes ? 1 : 0}"/></table>`;
  return s;
}

/**
 * Write half of the tables (ws.tables) feature. Always on in the package root; in `./lite` pass it
 * in `features` of writeXlsx / xlsxStream / createXlsxWriter.
 */
export const tablesWriter: FeatureWriter = {
  name: 'tables',
  sheet: {
    emitters: [
      {
        slot: SheetSlot.tableParts,
        emit(c) {
          const list = c.ws.tables;
          if (!list || list.length === 0) return '';
          let out = '';
          for (const t of list) {
            validateTable(c.ws, t, !c.streaming);
            validateWorkbook(c.ws, t);
            // unchanged table read with a part of its own relationships (query table): kept
            const kept = presTable(c.ws, t);
            if (kept) {
              out += `<tablePart r:id="${addPRel(kept.rel, c.path, c.addRel)}"/>`;
              continue;
            }
            const n = c.nextIndex('table');
            c.addPart(`xl/tables/table${n}.xml`, tableXml(t, n, c.styles), TABLE_CT);
            out += `<tablePart r:id="${c.addRel('table', `../tables/table${n}.xml`)}"/>`;
          }
          return `<tableParts count="${list.length}">${out}</tableParts>`;
        },
      },
    ],
  },
};

// ---------------------------------------------------------------- reading

function parseTable(xml: string, dxf: (id: string) => Style | undefined): Table | undefined {
  let t: Table | undefined;
  let col: TableColumn | undefined;
  let text = '';
  let hasFilter = false;
  const dxfs = (
    o: {
      headerStyle?: Style | undefined;
      dataStyle?: Style | undefined;
      totalsStyle?: Style | undefined;
    },
    a: Record<string, string>,
  ): void => {
    const hs = a.headerRowDxfId !== undefined ? dxf(a.headerRowDxfId) : undefined;
    const ds = a.dataDxfId !== undefined ? dxf(a.dataDxfId) : undefined;
    const ts = a.totalsRowDxfId !== undefined ? dxf(a.totalsRowDxfId) : undefined;
    if (hs) o.headerStyle = hs;
    if (ds) o.dataStyle = ds;
    if (ts) o.totalsStyle = ts;
  };
  parseXml(xml, {
    open(name, a) {
      switch (name) {
        case 'table':
          t = { name: a.name ?? '', ref: a.ref ?? '', columns: [] };
          if (a.displayName && a.displayName !== a.name) t.displayName = a.displayName;
          if ((num(a.headerRowCount) ?? 1) === 0) t.headerRow = false;
          if ((num(a.totalsRowCount) ?? 0) > 0) t.totalsRow = true;
          dxfs(t, a);
          break;
        case 'autoFilter':
          hasFilter = true;
          break;
        case 'tableColumn':
          col = { name: a.name ?? '' };
          if (a.totalsRowFunction) col.totalsRowFunction = a.totalsRowFunction as TotalsRowFunction;
          if (a.totalsRowLabel !== undefined) col.totalsRowLabel = a.totalsRowLabel;
          dxfs(col, a);
          t?.columns.push(col);
          break;
        case 'tableStyleInfo':
          if (t) {
            t.style = {
              theme: a.name ?? '',
              showFirstColumn: bool(a.showFirstColumn),
              showLastColumn: bool(a.showLastColumn),
              showRowStripes: bool(a.showRowStripes),
              showColumnStripes: bool(a.showColumnStripes),
            };
          }
          break;
      }
      text = '';
    },
    text(s) {
      text += s;
    },
    close(name) {
      if (col && name === 'calculatedColumnFormula') col.calculatedColumnFormula = stripXlfn(text);
      else if (col && name === 'totalsRowFormula') col.totalsRowFormula = stripXlfn(text);
      else if (name === 'tableColumn') col = undefined;
    },
  });
  if (t && !hasFilter) t.autoFilter = false;
  return t;
}

/**
 * Read half of the tables (ws.tables) feature. Always on in the package root; in `./lite` pass it
 * in `features` of readXlsx (without it, the elements are kept raw and written back).
 */
export const tablesReader: FeatureReader = {
  name: 'tables',
  sheet: {
    elements: ['tableParts'],
    open(ctx) {
      const ids: string[] = [];
      ctx.scratch.set(tablesReader, ids);
      return {
        open(name, a) {
          if (name === 'tablePart' && a['r:id']) ids.push(a['r:id']);
        },
      };
    },
    finish(ctx) {
      const ids = ctx.scratch.get(tablesReader) as string[] | undefined;
      if (!ids) return;
      const list: Table[] = [];
      for (const id of ids) {
        const rel = ctx.rels?.get(id);
        const xml = rel && ctx.readText(rel.path);
        const t = xml ? parseTable(xml, (d) => ctx.workbook.styles.dxfs[ctx.dxf(d)]) : undefined;
        if (t) list.push(t);
      }
      if (list.length) ctx.ws.tables = list;
    },
  },
};
