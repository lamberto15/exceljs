// Copy worksheets and ranges, within a workbook or into another one (Excel's Move or Copy /
// copy-paste). Styles are re-registered in the target workbook, pictures re-added, table names
// made unique, date serials converted between the 1900 and 1904 systems.

import { addImage } from '../xlsx/features/images.js';
import { translateFormula } from '../xlsx/read/formula.js';
import { decodeAddr, decodeRange, encodeRange } from './address.js';
import type { Formula } from './types.js';
import type { Workbook } from './workbook.js';
import { CellKind, type RowData, type Worksheet } from './worksheet.js';

/** Days between the 1900 and 1904 date systems. */
const D1904 = 1462;

/** Source xf → target xf (identity within one workbook). */
function styleMap(src: Workbook, dst: Workbook): (xf: number) => number {
  if (src === dst) return (xf) => xf;
  const map = new Map<number, number>();
  return (xf) => {
    if (xf === 0) return 0;
    let out = map.get(xf);
    if (out === undefined) {
      out = dst.styles.intern(src.styles.get(xf));
      map.set(xf, out);
    }
    return out;
  };
}

function uniqueSheetName(wb: Workbook, base: string): string {
  if (!wb.getWorksheet(base)) return base;
  for (let k = 2; ; k++) {
    const suffix = ` (${k})`;
    const n = base.slice(0, 31 - suffix.length) + suffix;
    if (!wb.getWorksheet(n)) return n;
  }
}

function tableNames(wb: Workbook): Set<string> {
  const s = new Set<string>();
  for (const ws of wb.worksheets) for (const t of ws.tables ?? []) s.add(t.name.toLowerCase());
  for (const d of wb.definedNames) s.add(d.name.toLowerCase());
  return s;
}

/** Copy one cell slot of `row` (index i) to (r0, c0) of `dst`. */
function copySlot(
  row: RowData,
  i: number,
  dst: Worksheet,
  r0: number,
  c0: number,
  xf: number,
  dDate: number,
  dr: number,
  dc: number,
): void {
  const kind = row.kind[i]!;
  let num = row.num[i]!;
  let obj = row.obj?.[i];
  if (kind === CellKind.Date) num += dDate;
  else if (kind === CellKind.Formula && (dr !== 0 || dc !== 0)) {
    const f = obj as Formula;
    const nf: Formula = { ...f, formula: translateFormula(f.formula, dr, dc) };
    if (f.array !== undefined) (nf as { array?: string }).array = translateFormula(f.array, dr, dc);
    obj = nf;
  }
  dst.setCellRaw(r0, c0, kind, num, obj, xf);
}

/**
 * Copy a worksheet: into its own workbook (default; Excel's "Create a copy") or into `target`.
 * The copy gets `name` (default "Name (2)", "Name (3)", ...) and is added after the last sheet.
 * Copied: cells (formulas as written), styles, row / column properties and keys, merges, views,
 * auto-filter, conditional formatting, data validation, tables (renamed "Name_2" when the name
 * is taken), notes, pictures, charts, protection, print setup, page breaks, ignored errors.
 * Not copied: content kept raw from a read file (`ws.preserved`) and sheet-scoped defined names.
 * Formulas referring to other sheets keep their text (in another workbook such sheets may not
 * exist).
 */
export function copyWorksheet(
  src: Worksheet,
  target: Workbook = src.workbook,
  name?: string,
): Worksheet {
  const sw = src.workbook;
  const ws = target.addWorksheet(name ?? uniqueSheetName(target, src.name));
  ws.state = src.state;
  ws.tabColor = src.tabColor;
  ws.views = structuredClone(src.views);
  ws.autoFilter = src.autoFilter;
  const xfOf = styleMap(sw, target);
  const dDate = sw.date1904 === target.date1904 ? 0 : sw.date1904 ? D1904 : -D1904;

  // cells and row properties
  src.rowData.forEach((row, r) => {
    if (!row) return;
    for (let i = 0; i < row.n; i++)
      copySlot(row, i, ws, r, row.col[i]!, xfOf(row.style[i]!), dDate, 0, 0);
    ws.setRowProps(r + 1, {
      height: row.height || undefined,
      hidden: row.hidden,
      outlineLevel: row.outlineLevel,
      collapsed: row.collapsed,
      styleIndex: xfOf(row.xf),
    });
  });
  src.columnData.forEach((c, i) => {
    if (!c) return;
    const d = ws.getColumn(i + 1);
    d.header = c.header;
    d.width = c.width;
    d.hidden = c.hidden;
    d.outlineLevel = c.outlineLevel;
    d.collapsed = c.collapsed;
    if (c.styleIndex) d.styleIndex = xfOf(c.styleIndex);
    if (c.key !== undefined) d.key = c.key;
  });
  for (const m of src.mergeRanges) ws.addMergeRaw({ ...m });

  // sheet features (plain data)
  const clone = <T>(v: T): T => (v === undefined ? v : structuredClone(v));
  ws.conditionalFormatting = clone(src.conditionalFormatting);
  ws.dataValidations = clone(src.dataValidations);
  ws.comments = clone(src.comments);
  ws.charts = clone(src.charts);
  ws.protection = clone(src.protection);
  ws.pageSetup = clone(src.pageSetup);
  ws.headerFooter = clone(src.headerFooter);
  ws.rowBreaks = clone(src.rowBreaks);
  ws.colBreaks = clone(src.colBreaks);
  ws.ignoredErrors = clone(src.ignoredErrors);
  if (src.tables) {
    const taken = tableNames(target);
    ws.tables = src.tables.map((t) => {
      const c = structuredClone(t);
      let n = c.name;
      for (let k = 2; taken.has(n.toLowerCase()); k++) n = `${t.name}_${k}`;
      if (n !== c.name) {
        c.name = n;
        if (c.displayName !== undefined) c.displayName = n;
      }
      taken.add(n.toLowerCase());
      return c;
    });
  }
  const imageId = (id: number): number => {
    if (sw === target) return id;
    const img = sw.images?.find((i) => i.id === id);
    return img ? addImage(target, { data: img.data, extension: img.extension }) : id;
  };
  if (src.images) ws.images = src.images.map((p) => ({ ...clone(p), imageId: imageId(p.imageId) }));
  if (src.backgroundImage !== undefined) ws.backgroundImage = imageId(src.backgroundImage);
  return ws;
}

/**
 * Copy the cells of `range` ("A1:D10") of `src` to `dst` (any sheet, any workbook) with their
 * top-left cell at `to` ("F2"), like copy-paste: values, styles and merges inside the range;
 * relative references in formulas are adjusted by the offset. Target cells in the area are
 * replaced (empty source cells clear them).
 */
export function copyRange(src: Worksheet, range: string, dst: Worksheet, to: string): void {
  const r = decodeRange(range);
  const t = decodeAddr(to);
  const dr = t.row - r.top;
  const dc = t.col - r.left;
  const xfOf = styleMap(src.workbook, dst.workbook);
  const sd = src.workbook.date1904;
  const dDate = sd === dst.workbook.date1904 ? 0 : sd ? D1904 : -D1904;
  // snapshot first: src and dst may be the same (overlapping) sheet
  const rows: (RowData | undefined)[] = [];
  for (let y = r.top; y <= r.bottom; y++) {
    const row = src.rowData[y];
    if (!row) continue;
    const n = row.n;
    rows[y] = {
      n,
      col: row.col.slice(0, n),
      kind: row.kind.slice(0, n),
      num: row.num.slice(0, n),
      style: row.style.slice(0, n),
      obj: row.obj?.slice(0, n),
    } as RowData;
  }
  const merges = src.mergeRanges.filter(
    (m) => m.top >= r.top && m.bottom <= r.bottom && m.left >= r.left && m.right <= r.right,
  );
  const area = { top: t.row, left: t.col, bottom: r.bottom + dr, right: r.right + dc };
  for (let y = area.top; y <= area.bottom; y++) {
    for (let c = area.left; c <= area.right; c++) dst.clearCell(y + 1, c + 1);
  }
  dst.unmergeCells(encodeRange(area));
  rows.forEach((row, y) => {
    if (!row) return;
    for (let i = 0; i < row.n; i++) {
      const c = row.col[i]!;
      if (c >= r.left && c <= r.right) {
        copySlot(row, i, dst, y + dr, c + dc, xfOf(row.style[i]!), dDate, dr, dc);
      }
    }
  });
  for (const m of merges) {
    dst.addMergeRaw({
      top: m.top + dr,
      left: m.left + dc,
      bottom: m.bottom + dr,
      right: m.right + dc,
    });
  }
}
