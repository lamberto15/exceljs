// Fill a formula over a range (ExcelJS `ws.fillFormula`), like dragging Excel's fill handle.

import { translateFormula } from '../xlsx/read/formula.js';
import { decodeRange } from './address.js';
import type { Formula } from './types.js';
import type { Worksheet } from './worksheet.js';

type Result = Formula['result'];

/**
 * Put `formula` (written for the top-left cell of `range`) into every cell of the range with its
 * relative references adjusted per cell, as Excel's fill handle does. `results` (optional): the
 * cached values, as a row-major array or a function of the 1-based (row, col). Cells keep their
 * styles. Each cell holds its own formula text; writers emit the range as an Excel shared
 * formula (`t="shared"`) while the cells still hold the filled formulas.
 */
export function fillFormula(
  ws: Worksheet,
  range: string,
  formula: string,
  results?: readonly Result[] | ((row: number, col: number) => Result),
): void {
  const r = decodeRange(range);
  const f = formula.charCodeAt(0) === 61 ? formula.slice(1) : formula;
  let k = 0;
  for (let y = r.top; y <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++) {
      const v: { formula: string; result?: Result } = {
        formula: translateFormula(f, y - r.top, c - r.left),
      };
      const res = typeof results === 'function' ? results(y + 1, c + 1) : results?.[k];
      if (res !== undefined) v.result = res;
      ws.setCell(y + 1, c + 1, v);
      k++;
    }
  }
  ws._addSharedGroup(r.top, r.left, r.bottom, r.right);
}
