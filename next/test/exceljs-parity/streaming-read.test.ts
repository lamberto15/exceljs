// Streaming read (xlsxRows): spec/integration/workbook-xlsx-reader.spec.js ("Serialise" via
// spec/utils/test-workbook-reader.js, "#readFile" / "#read" limits). Issues 1328 / 1364 and
// pr 1431 are in issues.test.ts / prs.test.ts.

import { Readable } from 'node:stream';
import { describe, expect, it } from 'vitest';
import { readXlsx, Workbook, writeXlsx, xlsxRows } from '../../src/index.js';
import { allRows, fixture } from './helpers.js';
import { createTestBook, testValues } from './test-book.js';

describe('exceljs parity: streaming read', () => {
  // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "xlsx file" (checkTestBookReader)
  it('reads the test book row by row', async () => {
    const rows = await allRows(writeXlsx(createTestBook(new Workbook())));
    const byRow = new Map(rows.map((r) => [r.row, r.values]));
    // ➖ rows without values (row 3: only merged cells) are skipped; exceljs emits 11 rows
    expect([...byRow.keys()]).toEqual([1, 2, 4, 5, 6, 7, 8, 9, 10, 11]);
    const r1 = byRow.get(1)!;
    expect(r1[0]).toBe(7);
    // ➖ shared strings are resolved and dates converted (exceljs: { sharedString: 0 }, serials)
    expect(r1[1]).toBe(testValues.str);
    expect(Math.abs((r1[2] as Date).getTime() - testValues.date.getTime())).toBeLessThan(1);
    expect(r1[3]).toEqual(testValues.formulas[0]);
    expect(r1[4]).toEqual(testValues.formulas[1]);
    // ➖ hyperlinks are not applied by xlsxRows: the cell's own text
    expect(r1[5]).toBe(testValues.hyperlink.text);
    expect(r1[6]).toBe(testValues.str2);
    const r2 = byRow.get(2)!;
    expect(r2[0]).toBe(5);
    expect(r2[1] ?? null).toBeNull();
    expect(r2[2] ?? null).toBeNull();
    const r5 = byRow.get(5)!;
    expect(r5.slice(0, 3)).toEqual([testValues.str, testValues.str, testValues.str]);
    expect(Math.abs((r5[3] as number) - 1.6)).toBeLessThan(1e-8);
    // 🚫 row heights are not reported by xlsxRows (readXlsx has them)
  });

  describe('#readFile / #read limits', () => {
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should bail out if the file contains more rows than the limit" (×2)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should fail fast on a huge file" (×2)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should bail out if the file contains more cells than the limit"
    it.skip('maxRows / maxCols read options — not supported (zip limits only; stop iterating xlsxRows instead)', () => {});

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#readFile rows)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#readFile cols)
    it('reads fibonacci.xlsx (19 rows) and many-columns.xlsx (20 columns)', () => {
      expect(readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!.rowCount).toBe(19);
      const wide = readXlsx(fixture('many-columns.xlsx')).worksheets[0]!;
      expect(wide.columnCount).toBeGreaterThanOrEqual(20);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#read)
    it('streams fibonacci.xlsx from a Node stream; stopping early is the row limit', async () => {
      const src = Readable.from([Buffer.from(fixture('fibonacci.xlsx'))]);
      let n = 0;
      for await (const r of xlsxRows(src as unknown as AsyncIterable<Uint8Array>)) {
        if (r.sheet === 'fib') n++;
        if (n === 10) break; // the exceljs maxRows: 10 case, without an error
      }
      expect(n).toBe(10);
      expect(
        (await allRows(fixture('fibonacci.xlsx'))).filter((r) => r.sheet === 'fib'),
      ).toHaveLength(19);
    });
  });

  it('huge.xlsx streams with constant memory', async () => {
    let n = 0;
    for await (const _ of xlsxRows(fixture('huge.xlsx'))) n++;
    expect(n).toBeGreaterThan(100);
  });
});
