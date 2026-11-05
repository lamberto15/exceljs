// What Excel displays for a value: the number format applied (ExcelJS `cell.text`).

import { dateToSerial } from '../core/dates.js';
import type { CellValue } from '../core/types.js';
import type { Worksheet } from '../core/worksheet.js';
import { formatText } from './format.js';
import { setDate1904 } from './values.js';

/** General format as a cell shows it at standard width: at most 11 characters, like Excel. */
function generalText(v: number): string {
  const a = Math.abs(v);
  const sign = v < 0 ? '-' : '';
  if (a === 0) return '0';
  if (a >= 1e11 || a < 1e-9) {
    let e = a
      .toExponential(5)
      .replace(/\.?0+e/, 'e')
      .replace('e+', 'E+')
      .replace('e-', 'E-');
    if (/E[+-]\d$/.test(e)) e = `${e.slice(0, -1)}0${e.slice(-1)}`;
    return sign + e;
  }
  if (a === Math.floor(a)) return sign + String(a);
  const id = a >= 1 ? String(Math.floor(a)).length : 1;
  let t = a.toFixed(Math.max(0, 10 - id));
  if (t.indexOf('.') >= 0) t = t.replace(/\.?0+$/, '');
  return sign + t;
}

/**
 * Text Excel shows for `value` under number format `numFmt` (default General): dates, times,
 * percentages, thousands separators, decimals, literal / text sections, booleans as TRUE /
 * FALSE, errors as their code, formulas by their cached result, rich text / hyperlinks as their
 * text. General numbers show at most 11 characters, as in a standard-width cell; column width
 * overflow ("####") and locale-specific separators are not simulated.
 */
export function formatValue(
  value: CellValue,
  numFmt = 'General',
  options: { date1904?: boolean | undefined; localDates?: boolean | undefined } = {},
): string {
  const fmt = numFmt || 'General';
  const old = setDate1904(options.date1904 === true);
  try {
    let v: unknown = value;
    if (v !== null && typeof v === 'object' && !(v instanceof Date) && 'formula' in v) {
      v = (v as { result?: unknown }).result;
    }
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return /^general$/i.test(fmt) ? generalText(v) : formatText(v, fmt);
    if (typeof v === 'string') return formatText(v, fmt);
    if (v instanceof Date) {
      return formatText(
        dateToSerial(v, options.date1904 === true, options.localDates === true),
        fmt,
      );
    }
    if (typeof v === 'object') {
      if ('error' in v) return String((v as { error: string }).error);
      if ('richText' in v)
        return (v as { richText: { text: string }[] }).richText.map((r) => r.text).join('');
      if ('text' in v) return String((v as { text: unknown }).text);
    }
    return String(v);
  } finally {
    setDate1904(old);
  }
}

/** Text Excel shows in a cell (its value under its style's number format); see formatValue. */
export function cellText(ws: Worksheet, row: number | string, col?: number | string): string {
  const value = typeof row === 'string' ? ws.getValue(row) : ws.getValue(row, col!);
  const style = typeof row === 'string' ? ws.getStyle(row) : ws.getStyle(row, col!);
  return formatValue(value, style.numFmt, {
    date1904: ws.workbook.date1904,
    localDates: ws.workbook.localDates,
  });
}
