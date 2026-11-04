// Sort the rows of a range, like Excel's Data > Sort on a selection.

import { translateFormula } from '../xlsx/read/formula.js';
import { decodeCol, decodeRange } from './address.js';
import type { CellValue, Formula } from './types.js';
import { CellKind, type RowData, type Worksheet } from './worksheet.js';

export interface SortKey {
  /** Sheet column (1-based number or letter) inside the range. */
  col: number | string;
  /** Descending (default ascending). */
  desc?: boolean | undefined;
}

export interface SortOptions {
  /** The range's first row is a header and stays in place (default false). */
  header?: boolean | undefined;
}

interface Snap {
  r: number;
  kind: number[];
  num: number[];
  obj: unknown[];
  style: number[];
  cols: number[];
  keys: CellValue[];
}

/** Excel's ascending order: numbers (dates) < text < booleans < errors; blanks always last. */
function rank(v: CellValue): number {
  if (v === null || v === undefined || v === '') return 4;
  if (typeof v === 'number' || v instanceof Date) return 0;
  if (typeof v === 'string' || (typeof v === 'object' && ('richText' in v || 'hyperlink' in v)))
    return 1;
  if (typeof v === 'boolean') return 2;
  return 3;
}

function text(v: CellValue): string {
  if (typeof v === 'string') return v;
  if (v && typeof v === 'object') {
    if ('richText' in v) return v.richText.map((r) => r.text).join('');
    if ('hyperlink' in v) return v.text;
  }
  return '';
}

function compare(a: CellValue, b: CellValue, desc: boolean): number {
  const ra = rank(a);
  const rb = rank(b);
  if (ra === 4 || rb === 4) return ra === rb ? 0 : ra === 4 ? 1 : -1; // blanks last either way
  let c = ra - rb;
  if (c === 0) {
    if (ra === 0) c = +(a as number | Date) - +(b as number | Date);
    else if (ra === 1) c = text(a).localeCompare(text(b), undefined, { sensitivity: 'base' });
    else if (ra === 2) c = Number(a) - Number(b);
  }
  return desc ? -c : c;
}

/** A formula cell sorts by its cached result. */
function sortValue(v: CellValue): CellValue {
  if (v && typeof v === 'object' && 'formula' in v) {
    const r = (v as Formula).result;
    return r === undefined ? null : typeof r === 'object' ? null : r;
  }
  return v;
}

/**
 * Sort the rows of `range` ("A2:D100") by `keys` (first key wins, ties keep their order). Only
 * the range's cells move (values and styles), as in Excel; relative references of formulas in
 * moved rows shift with their row (like a cut/paste), references from outside the range keep
 * pointing at positions. Merged cells inside the range are not supported (throws).
 */
export function sortRange(
  ws: Worksheet,
  range: string,
  keys: SortKey | readonly SortKey[],
  options: SortOptions = {},
): void {
  const r = decodeRange(range);
  const top = r.top + (options.header ? 1 : 0);
  for (const m of ws.mergeRanges) {
    if (m.top <= r.bottom && m.bottom >= top && m.left <= r.right && m.right >= r.left) {
      throw new Error(`sortRange: ${range} contains merged cells`);
    }
  }
  const list = (Array.isArray(keys) ? keys : [keys]) as readonly SortKey[];
  const kc = list.map((k) => (typeof k.col === 'number' ? k.col - 1 : decodeCol(k.col)));
  for (const c of kc) {
    if (c < r.left || c > r.right) throw new RangeError(`sortRange: key column outside ${range}`);
  }
  const snaps: Snap[] = [];
  for (let y = top; y <= r.bottom; y++) {
    const row: RowData | undefined = ws.rowData[y];
    const s: Snap = { r: y, kind: [], num: [], obj: [], style: [], cols: [], keys: [] };
    if (row) {
      for (let i = 0; i < row.n; i++) {
        const c = row.col[i]!;
        if (c < r.left || c > r.right) continue;
        s.cols.push(c);
        s.kind.push(row.kind[i]!);
        s.num.push(row.num[i]!);
        s.obj.push(row.obj?.[i]);
        s.style.push(row.style[i]!);
      }
    }
    s.keys = kc.map((c) => sortValue(ws.getValue(y + 1, c + 1)));
    snaps.push(s);
  }
  const order = snaps
    .map((s, i) => ({ s, i }))
    .sort((a, b) => {
      for (let k = 0; k < list.length; k++) {
        const c = compare(a.s.keys[k]!, b.s.keys[k]!, list[k]!.desc === true);
        if (c !== 0) return c;
      }
      return a.i - b.i;
    })
    .map((x) => x.s);
  for (let y = top; y <= r.bottom; y++) {
    for (let c = r.left; c <= r.right; c++) ws.clearCell(y + 1, c + 1);
  }
  order.forEach((s, k) => {
    const y = top + k;
    const dr = y - s.r;
    for (let j = 0; j < s.cols.length; j++) {
      let obj = s.obj[j];
      if (s.kind[j] === CellKind.Formula && dr !== 0) {
        const f = obj as Formula;
        const nf: Formula = { ...f, formula: translateFormula(f.formula, dr, 0) };
        if (f.array !== undefined)
          (nf as { array?: string }).array = translateFormula(f.array, dr, 0);
        obj = nf;
      }
      ws.setCellRaw(y, s.cols[j]!, s.kind[j]!, s.num[j]!, obj, s.style[j]!);
    }
  });
}
