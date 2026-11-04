// Sheet-wide formatting: what selecting all cells (Ctrl+A) and formatting them does in Excel.

import { MAX_COLS } from './address.js';
import type { Style } from './types.js';
import type { Worksheet } from './worksheet.js';

/**
 * Apply `style` to the whole sheet, like selecting all cells in Excel and formatting them: every
 * existing cell, row style and column (all 16,384 get a column style, written as one compact
 * `<col>` range) gets the given properties merged over its own (top-level keys replace:
 * `{ protection: { locked: false } }` unlocks everything, `{ font: {...} }` replaces fonts).
 * Cells created later inherit it through their row / column. Typical use: unlock all cells,
 * then lock a few before protecting the sheet.
 */
export function setSheetStyle(ws: Worksheet, style: Style): void {
  const reg = ws.workbook.styles;
  const cache = new Map<number, number>();
  const patch = (xf: number): number => {
    let out = cache.get(xf);
    if (out === undefined) {
      out = reg.intern({ ...reg.get(xf), ...style });
      cache.set(xf, out);
    }
    return out;
  };
  for (const row of ws.rowData) {
    if (row === undefined) continue;
    for (let i = 0; i < row.n; i++) row.style[i] = patch(row.style[i]!);
    if (row.xf) row.xf = patch(row.xf);
  }
  for (let c = 1; c <= MAX_COLS; c++) {
    const col = ws.getColumn(c);
    col.styleIndex = patch(col.styleIndex);
  }
}
