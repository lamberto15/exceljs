// CSV: spec/integration/workbook/workbook.spec.js "csv file", "CSV file and its configuration",
// "throws an error when csv file not found". Issues 991 / 995 are in issues.test.ts.

import { describe, expect, it } from 'vitest';
import { parseCsv, writeCsv } from '../../src/csv/index.js';
import { CellKind, Workbook } from '../../src/index.js';
import { arraysToSheet, objectsToSheet, sheetToArrays } from '../../src/json/index.js';
import { createTestBook, testValues } from './test-book.js';

/** ExcelJS wb.csv.write: the first sheet's values (turbosheet: sheetToArrays + writeCsv). */
function toCsv(wb: Workbook, opts: Parameters<typeof writeCsv>[1] = {}): string {
  return writeCsv(sheetToArrays(wb.getWorksheet('sheet1')!, { raw: true }) as never, opts);
}

describe('exceljs parity: csv', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "csv file"
  it('csv file: the test book values survive CSV write → read', () => {
    const csv = toCsv(createTestBook(new Workbook(), { checkBadAlignments: false }));
    const ws = arraysToSheet(
      new Workbook().addWorksheet('sheet1'),
      parseCsv(csv, { parseNumbers: true, parseDates: true }),
    );
    expect(ws.getValue('A1')).toBe(7);
    expect(ws.getValue('B1')).toBe(testValues.str);
    const c1 = ws.getValue('C1') as Date;
    expect(Math.abs(c1.getTime() - testValues.date.getTime())).toBeLessThan(1000);
    expect(ws.getValue('D1')).toBe(7); // formula → cached result
    expect(ws.getValue('E1')).toBeNull(); // formula without result → empty field → no cell
    expect(ws.getKind('E1')).toBe(CellKind.Empty);
    // ➖ exceljs writes the URL of a hyperlink; turbosheet the display text
    expect(ws.getValue('F1')).toBe(testValues.hyperlink.text);
    expect(ws.getValue('G1')).toBe(testValues.str2);
    expect(ws.getValue('H1')).toBe(testValues.json.string);
    // ➖ CSV parsing does not type booleans or error codes (exceljs: true / { error })
    expect(ws.getValue('I1')).toBe('TRUE');
    expect(ws.getValue('J1')).toBe('FALSE');
    expect(ws.getValue('K1')).toBe('#N/A');
    expect(ws.getValue('L1')).toBe('#VALUE!');
    expect(ws.getValue('A2')).toBe(5);
    for (let c = 1; c <= 5; c++) expect(ws.getValue(9, c)).toBe(c);
  });

  // exceljs: spec/integration/workbook/workbook.spec.js — "CSV file and its configuration"
  it('CSV with a tab delimiter and a custom date format', () => {
    const pad = (n: number) => String(n).padStart(2, '0');
    const fmt = (d: Date) =>
      `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
    const csv = toCsv(createTestBook(new Workbook(), { checkBadAlignments: false }), {
      delimiter: '\t',
      dateFormat: fmt,
    });
    expect(csv.split('\r\n')[0]!.split('\t')[2]).toBe(fmt(testValues.date));
    const rows = parseCsv(csv, { delimiter: '\t', parseNumbers: true });
    expect(rows[0]![0]).toBe(7);
    expect(rows[0]![1]).toBe(testValues.str);
    // 🚫 no custom date formats on read (exceljs: dateFormats: ['DD/MM/YYYY HH:mm:ss']); only
    // ISO 8601 is recognized, so the date stays text
    expect(rows[0]![2]).toBe(fmt(testValues.date));
  });

  it('empty CSV fields leave cells empty; emptyAs: "string" keeps them', () => {
    const rows = parseCsv('a,,c\r\n,x,\r\n');
    const ws = arraysToSheet(new Workbook().addWorksheet('s'), rows);
    expect(ws.getRow(1).values).toEqual([undefined, 'a', undefined, 'c']);
    expect(ws.getRow(2).values).toEqual([undefined, undefined, 'x']);
    expect(ws.dimensions?.ref).toBe('A1:C2');
    const ws2 = arraysToSheet(new Workbook().addWorksheet('s'), rows, { emptyAs: 'string' });
    expect(ws2.getValue('B1')).toBe('');
    const ws3 = objectsToSheet(new Workbook().addWorksheet('s'), [{ a: '', b: 1 }]);
    expect(ws3.getKind('A2')).toBe(CellKind.Empty);
  });

  // exceljs: spec/integration/workbook/workbook.spec.js — "throws an error when csv file not found"
  it.skip('csv.readFile of a missing file — not applicable (no file system API; use your runtime)', () => {});
});
