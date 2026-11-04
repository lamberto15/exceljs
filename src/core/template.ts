// Fill a template workbook: replace {{placeholders}} in cell text with data, repeat rows for
// arrays. Report-from-template workflows (design the layout in Excel, fill it in code).

import { deleteRows, dupRow } from './structure.js';
import type { CellValue } from './types.js';
import type { Workbook } from './workbook.js';
import { CellKind, type Worksheet } from './worksheet.js';

const TAG = /\{\{\s*([^{}]+?)\s*\}\}/g;
const ONLY = /^\{\{\s*([^{}]+?)\s*\}\}$/;
/** `{{items[].name}}`: a field of each element of the array `items`. */
const EACH = /^([^[\]]+)\[\]\.?(.*)$/;

export interface FillTemplateOptions {
  /** Only these sheets (names); default all. */
  sheets?: readonly string[] | undefined;
  /** Text for placeholders whose path has no value (default ''). */
  missing?: string | undefined;
}

function get(data: unknown, path: string): unknown {
  let o = data;
  for (const p of path.split('.')) {
    if (p === '') continue;
    if (o === null || typeof o !== 'object') return undefined;
    o = (o as Record<string, unknown>)[p];
  }
  return o;
}

function asCell(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean' || v instanceof Date)
    return v;
  if (typeof v === 'object') return v as CellValue; // formula / hyperlink / rich text objects
  return String(v);
}

function show(v: unknown, missing: string): string {
  if (v === null || v === undefined) return missing;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') return 'text' in v ? String((v as { text: unknown }).text) : '';
  return String(v);
}

/** Replace the placeholders of one text cell; `scope` resolves `items[]` paths for one element. */
function fillCell(
  ws: Worksheet,
  r: number,
  c: number,
  s: string,
  data: unknown,
  scope: [string, unknown] | undefined,
  missing: string,
): void {
  const resolve = (path: string): unknown => {
    const m = EACH.exec(path);
    if (m && scope && m[1] === scope[0]) return get(scope[1], m[2]!);
    return get(data, path);
  };
  const only = ONLY.exec(s);
  if (only) {
    const v = asCell(resolve(only[1]!));
    ws.setCell(r, c, v === null && missing ? missing : v);
    return;
  }
  ws.setCell(
    r,
    c,
    s.replace(TAG, (_, p: string) => show(resolve(p), missing)),
  );
}

/** [0-based column, text] of the row's text cells containing a placeholder (collected first:
 * filling may remove slots). */
function placeholders(ws: Worksheet, r: number): [number, string][] {
  const row = ws.rowData[r];
  const out: [number, string][] = [];
  if (row) {
    for (let i = 0; i < row.n; i++) {
      if (row.kind[i] === CellKind.String && (row.obj![i] as string).includes('{{')) {
        out.push([row.col[i]!, row.obj![i] as string]);
      }
    }
  }
  return out;
}

/**
 * Fill a template: every text cell containing `{{path}}` placeholders (dotted paths into
 * `data`, e.g. `{{customer.name}}`) is replaced, keeping the cell's style. A cell holding only a
 * placeholder gets the value itself (numbers, dates, booleans, formula objects stay typed);
 * mixed text gets the values as text. A row with `{{items[].field}}` placeholders is repeated
 * once per element of the array `items` (inserted below it, so cells and references below
 * move down, formulas in the row adjust like a pasted row, and ranges ending at the row, such
 * as `SUM(B4:B4)` in a totals row, grow over the copies); an empty array removes the row.
 */
export function fillTemplate(wb: Workbook, data: unknown, options: FillTemplateOptions = {}): void {
  const missing = options.missing ?? '';
  const only = options.sheets?.map((s) => s.toLowerCase());
  for (const ws of wb.worksheets) {
    if (only && !only.includes(ws.name.toLowerCase())) continue;
    // repeat rows first, bottom up, so insertions do not move rows still to visit
    for (let r = ws.rowData.length - 1; r >= 0; r--) {
      const row = ws.rowData[r];
      if (!row?.n) continue;
      let arr: string | undefined;
      for (let i = 0; i < row.n && arr === undefined; i++) {
        if (row.kind[i] !== CellKind.String) continue;
        for (const m of (row.obj![i] as string).matchAll(TAG)) {
          const e = EACH.exec(m[1]!);
          if (e) {
            arr = e[1]!;
            break;
          }
        }
      }
      if (arr === undefined) continue;
      const v = get(data, arr);
      const items = Array.isArray(v) ? v : [];
      if (items.length === 0) {
        deleteRows(ws, r + 1);
        continue;
      }
      if (items.length > 1) dupRow(ws, r + 1, items.length - 1, true, true);
      items.forEach((item, k) => {
        for (const [c, s] of placeholders(ws, r + k)) {
          fillCell(ws, r + k + 1, c + 1, s, data, [arr!, item], missing);
        }
      });
    }
    // the remaining placeholders
    for (let r = 0; r < ws.rowData.length; r++) {
      for (const [c, s] of placeholders(ws, r))
        fillCell(ws, r + 1, c + 1, s, data, undefined, missing);
    }
  }
}
