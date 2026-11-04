// Sheet-wide formatting (what selecting all cells and formatting them does in Excel) and
// in-cell checkboxes.

import { decodeRange, MAX_COLS, MAX_ROWS } from './address.js';
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

/**
 * Make cells in-cell checkboxes (Excel 365, 2024+): each cell of `range` ("A2:A20") gets the
 * checkbox format and a boolean value (`checked`, default false; existing booleans are kept).
 * Excel versions without checkboxes show TRUE / FALSE.
 */
export function setCheckbox(ws: Worksheet, range: string, checked = false): void {
  const r = decodeRange(range);
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++) {
      const cur = ws.getValue(y + 1, c + 1);
      ws.setCell(y + 1, c + 1, typeof cur === 'boolean' ? cur : checked, {
        ...ws.getStyle(y + 1, c + 1),
        checkbox: true,
      });
    }
  }
}

/** Actions users may still perform on a protected sheet (see `protectSheet`). */
export type ProtectAllow = Partial<
  Record<
    | 'sort'
    | 'autoFilter'
    | 'formatCells'
    | 'formatColumns'
    | 'formatRows'
    | 'insertRows'
    | 'insertColumns'
    | 'deleteRows'
    | 'deleteColumns'
    | 'insertHyperlinks'
    | 'pivotTables'
    | 'selectLockedCells'
    | 'selectUnlockedCells',
    boolean
  >
>;

export interface ProtectSheetOptions {
  /** Password (hashed when written); omit for protection without a password. */
  password?: string | undefined;
  /** Ranges users may still edit ("B2:D20", whole columns "B:D" or rows "5:9"). */
  editable?: string | readonly string[] | undefined;
  /** Actions allowed on the protected sheet. Sorting / filtering only work on unlocked cells. */
  allow?: ProtectAllow | undefined;
}

/** Merge `p` into the protection of every cell of `range` (and of whole column / row styles). */
function patchProtection(ws: Worksheet, range: string, locked: boolean): void {
  const r = decodeRange(range);
  const reg = ws.workbook.styles;
  const cache = new Map<number, number>();
  const patch = (xf: number): number => {
    let out = cache.get(xf);
    if (out === undefined) {
      const s = reg.get(xf);
      out = reg.intern({ ...s, protection: { ...s.protection, locked } });
      cache.set(xf, out);
    }
    return out;
  };
  const allRows = r.top === 0 && r.bottom === MAX_ROWS - 1;
  const allCols = r.left === 0 && r.right === MAX_COLS - 1;
  if (allRows)
    for (let c = r.left; c <= r.right; c++)
      ws.getColumn(c + 1).styleIndex = patch(ws.getColumn(c + 1).styleIndex);
  if (allCols) {
    for (let y = r.top; y <= r.bottom; y++) {
      const row = ws.rowData[y];
      ws.setRowProps(y + 1, { styleIndex: patch(row?.xf ?? 0) });
    }
  }
  if (allRows || allCols) {
    // existing cells inside
    ws.rowData.forEach((row, y) => {
      if (!row || y < r.top || y > r.bottom) return;
      for (let i = 0; i < row.n; i++) {
        const c = row.col[i]!;
        if (c >= r.left && c <= r.right) row.style[i] = patch(row.style[i]!);
      }
    });
    return;
  }
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++)
      ws.setStyle(y + 1, c + 1, patch(ws.getStyleIndex(y + 1, c + 1)));
  }
}

/**
 * Protect a sheet in one call: cells of `editable` are unlocked (everything else stays locked,
 * Excel's default), then `ws.protection` is set with the password and the `allow`ed actions.
 */
export function protectSheet(ws: Worksheet, options: ProtectSheetOptions = {}): void {
  const ed = options.editable;
  for (const range of ed === undefined ? [] : typeof ed === 'string' ? [ed] : ed) {
    patchProtection(ws, range, false);
  }
  ws.protection = { sheet: true, ...options.allow, password: options.password };
}
