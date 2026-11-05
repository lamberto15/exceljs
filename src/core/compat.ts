// ExcelJS-style read helpers for code migrating from ExcelJS (functions, so they cost nothing
// unless imported).

import { decodeCol } from './address.js';
import type { CellValue } from './types.js';
import { CellKind, RowView, type Worksheet } from './worksheet.js';

/** Number of rows holding at least one value (ExcelJS `ws.actualRowCount`). */
export function actualRowCount(ws: Worksheet): number {
  let n = 0;
  for (const row of ws.rowData) {
    if (!row) continue;
    for (let i = 0; i < row.n; i++) {
      if (row.kind[i] !== CellKind.Empty) {
        n++;
        break;
      }
    }
  }
  return n;
}

/** The last row that contains cells (ExcelJS `ws.lastRow`), undefined for an empty sheet. */
export function lastRow(ws: Worksheet): RowView | undefined {
  const rows = ws.rowData;
  for (let r = rows.length - 1; r >= 0; r--) if (rows[r]?.n) return new RowView(ws, r);
  return undefined;
}

/** `length` row handles from row `start` (1-based; ExcelJS `ws.getRows`), existing or not. */
export function getRows(ws: Worksheet, start: number, length: number): RowView[] {
  const out: RowView[] = [];
  for (let i = 0; i < length; i++) out.push(ws.getRow(start + i));
  return out;
}

/**
 * All values as ExcelJS `ws.getSheetValues()` returns them: a sparse array indexed by 1-based
 * row, each a sparse array indexed by 1-based column (index 0 unused).
 */
export function getSheetValues(ws: Worksheet): CellValue[][] {
  const out: CellValue[][] = [];
  ws.rowData.forEach((row, r) => {
    if (row?.n) out[r + 1] = new RowView(ws, r).values;
  });
  return out;
}

/** 0-based column index of a 1-based number or a letter. */
const colIndex = (c: number | string): number => (typeof c === 'number' ? c - 1 : decodeCol(c));

/**
 * Values of one column as ExcelJS `column.values`: a sparse array indexed by 1-based row
 * (index 0 unused). `col`: 1-based number or letter.
 */
export function columnValues(ws: Worksheet, col: number | string): CellValue[] {
  const c = colIndex(col);
  const out: CellValue[] = [];
  ws.rowData.forEach((row, r) => {
    if (!row) return;
    const i = row.find(c);
    if (i >= 0 && row.kind[i] !== CellKind.Empty) out[r + 1] = ws._valueAt(row, i);
  });
  return out;
}

/** Visit the non-empty cells of one column, top to bottom (ExcelJS `column.eachCell`). */
export function eachColumnCell(
  ws: Worksheet,
  col: number | string,
  cb: (value: CellValue, rowNumber: number) => void,
): void {
  const c = colIndex(col);
  ws.rowData.forEach((row, r) => {
    if (!row) return;
    const i = row.find(c);
    if (i >= 0 && row.kind[i] !== CellKind.Empty) cb(ws._valueAt(row, i), r + 1);
  });
}
