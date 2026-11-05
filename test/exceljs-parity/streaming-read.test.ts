// Streaming read (xlsxRows): spec/integration/workbook-xlsx-reader.spec.js ("Serialise" via
// spec/utils/test-workbook-reader.js, "#readFile" / "#read" limits). Issues 1328 / 1364 and
// pr 1431 are in issues.test.ts / prs.test.ts.

import { Readable } from 'node:stream';
import { describe, expect, it } from 'vitest';
import { readXlsx, SheetLimitError, Workbook, writeXlsx, xlsxRows } from '../../src/index.js';
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
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should bail out if the file contains more rows than the limit" (#readFile)
    it('readXlsx: maxRows below the row count throws SheetLimitError', () => {
      // the fib sheet has 19 rows
      const read = () => readXlsx(fixture('fibonacci.xlsx'), { maxRows: 10 });
      expect(read).toThrow(SheetLimitError);
      // ➖ the message adds the sheet and the offending row
      expect(read).toThrow('Max row count (10) exceeded');
      expect(readXlsx(fixture('fibonacci.xlsx'), { maxRows: 20 }).worksheets).toHaveLength(1);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should bail out if the file contains more rows than the limit" (#read)
    it('xlsxRows: maxRows below the row count throws SheetLimitError', async () => {
      const src = Readable.from([Buffer.from(fixture('fibonacci.xlsx'))]);
      const rows: number[] = [];
      await expect(async () => {
        for await (const r of xlsxRows(src as unknown as AsyncIterable<Uint8Array>, {
          maxRows: 10,
        })) {
          rows.push(r.row);
        }
      }).rejects.toThrow('Max row count (10) exceeded');
      expect(rows.every((r) => r <= 10)).toBe(true);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should bail out if the file contains more cells than the limit"
    it('readXlsx: maxCols below the column count throws SheetLimitError', () => {
      // many-columns.xlsx has 20 columns in row 2
      let err: unknown;
      try {
        readXlsx(fixture('many-columns.xlsx'), { maxCols: 15 });
      } catch (e) {
        err = e;
      }
      expect(err).toBeInstanceOf(SheetLimitError);
      expect((err as SheetLimitError).message).toContain('Max column count (15) exceeded');
      expect((err as SheetLimitError).limit).toBe('maxCols');
      expect((err as SheetLimitError).max).toBe(15);
      expect(readXlsx(fixture('many-columns.xlsx'), { maxCols: 40 }).worksheets).toHaveLength(1);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should fail fast on a huge file" (maxRows)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should fail fast on a huge file" (maxCols)
    it('fails fast on a huge file (readXlsx and xlsxRows)', async () => {
      const huge = fixture('huge.xlsx');
      expect(() => readXlsx(huge, { maxRows: 100 })).toThrow('Max row count (100) exceeded');
      expect(() => readXlsx(huge, { maxCols: 10 })).toThrow('Max column count (10) exceeded');
      let n = 0;
      await expect(async () => {
        for await (const _ of xlsxRows(huge, { maxRows: 100 })) n++;
      }).rejects.toThrow(SheetLimitError);
      expect(n).toBeLessThanOrEqual(100);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#readFile rows)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#readFile cols)
    it('reads fibonacci.xlsx (19 rows) and many-columns.xlsx (20 columns)', () => {
      expect(readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!.rowCount).toBe(19);
      const wide = readXlsx(fixture('many-columns.xlsx')).worksheets[0]!;
      expect(wide.columnCount).toBeGreaterThanOrEqual(20);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should parse fine if the limit is not exceeded" (#read)
    it('streams fibonacci.xlsx from a Node stream (maxRows: 20 passes)', async () => {
      const src = Readable.from([Buffer.from(fixture('fibonacci.xlsx'))]);
      let n = 0;
      for await (const r of xlsxRows(src as unknown as AsyncIterable<Uint8Array>)) {
        if (r.sheet === 'fib') n++;
        if (n === 10) break; // stopping early also works
      }
      expect(n).toBe(10);
      let all = 0;
      for await (const r of xlsxRows(fixture('fibonacci.xlsx'), { maxRows: 20 })) {
        if (r.sheet === 'fib') all++;
      }
      expect(all).toBe(19);
    });
  });

  it('huge.xlsx streams with constant memory', async () => {
    let n = 0;
    for await (const _ of xlsxRows(fixture('huge.xlsx'))) n++;
    expect(n).toBeGreaterThan(100);
  });
});
