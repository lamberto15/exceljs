// ExcelJS regression tests for GitHub issues (spec/integration/issues/**), ported to turbosheet.

import { describe, expect, it } from 'vitest';
import { csvStream, parseCsv } from '../../src/csv/index.js';
import {
  addDataValidation,
  CellKind,
  readXlsx,
  translateFormula,
  Workbook,
  writeXlsx,
} from '../../src/index.js';
import {
  commentsReader,
  conditionalFormattingReader,
  imagesReader,
  printSetupReader,
  readXlsx as readLite,
  sheetProtectionReader,
  tablesReader,
} from '../../src/lite.js';
import { allRows, collect, exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';

const SPECIAL = [
  'constructor',
  'hasOwnProperty',
  'isPrototypeOf',
  'propertyIsEnumerable',
  'toLocaleString',
  'toString',
  'valueOf',
  '__defineGetter__',
  '__defineSetter__',
  '__lookupGetter__',
  '__lookupSetter__',
  '__proto__',
];

describe('exceljs parity: github issues', () => {
  // exceljs: spec/integration/issues/issue-1027-sheet-defined-data-validate.spec.js — "issue 1027 - Broken due to Cannot set property 'marked' of undefined error"
  it('issue 1027: whole-column list validation (A2:A1048576) writes and reads back', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    addDataValidation(ws, {
      sqref: 'A2:A1048576',
      type: 'list',
      allowBlank: true,
      error: 'Please use the drop down to select a valid value',
      errorTitle: 'Invalid Selection',
      formula1: '"Apples,Bananas,Oranges"',
      showErrorMessage: true,
    });
    const bytes = writeXlsx(wb);
    const dv = readXlsx(bytes).getWorksheet('Sheet1')!.dataValidations!;
    expect(dv).toEqual([
      expect.objectContaining({
        sqref: 'A2:A1048576',
        type: 'list',
        formula1: '"Apples,Bananas,Oranges"',
        errorTitle: 'Invalid Selection',
      }),
    ]);
    // exceljs still loads it
    const x = await exceljsLoad(bytes);
    expect(x.getWorksheet('Sheet1')!.getCell('A5').dataValidation.type).toBe('list');
  });

  // exceljs: spec/integration/issues/issue-1328-xlsx-worksheet-reader-date.spec.js — "issue 1328 - should emit row with Date Object"
  it('issue 1328: streaming reader returns Date objects for date-formatted cells', async () => {
    const rows = await allRows(fixture('dateIssue.xlsx'));
    expect(rows.map((r) => r.values[0])).toEqual(['Date', new Date('2020-11-20T00:00:00.000Z')]);
  });

  // exceljs: spec/integration/issues/issue-1339-speciel-cell-file.spec.js — "issue 1339 - Special cell value results invalid file"
  it('issue 1339: streaming writer handles Object.prototype names as cell strings', async () => {
    const bytes = await streamWrite(async (w) => {
      const s = w.addWorksheet('Sheet1');
      for (const v of SPECIAL) s.addRow([v, v]);
    });
    const ws = readXlsx(bytes).getWorksheet('Sheet1')!;
    SPECIAL.forEach((v, i) => {
      expect(ws.getValue(`A${i + 1}`)).toBe(v);
      expect(ws.getValue(`B${i + 1}`)).toBe(v);
    });
  });

  // exceljs: spec/integration/issues/issue-1364-incorrect-worksheet-name-streaming-xlsx.spec.js — "issue 1364 - Incorrect Worksheet Name on Streaming XLSX Reader"
  it('issue 1364: streaming reader reports the real sheet name', async () => {
    const rows = await allRows(fixture('test-issue-1364.xlsx'));
    expect(rows.length).toBeGreaterThan(0);
    for (const r of rows) expect(r.sheet).toBe('Sum Worksheet');
  });

  // exceljs: spec/integration/issues/issue-163-xlsx-read-file.spec.js — "issue 163 - Error while using xslx readFile method"
  it('issue 163: reads test-issue-163.xlsx', () => {
    expect(readXlsx(fixture('test-issue-163.xlsx')).worksheets.length).toBeGreaterThan(0);
  });

  // exceljs: spec/integration/issues/issue-1669-optional-custom-autofilter-on-table.spec.js — "issue 1669 - optional autofilter and custom autofilter on tables"
  it('issue 1669: reads tables with optional / custom autofilter, and writes them back', () => {
    const wb = readXlsx(fixture('test-issue-1669.xlsx'));
    const tables = wb.worksheets.flatMap((ws) => ws.tables ?? []);
    expect(tables.length).toBeGreaterThan(0);
    expect(() => roundtrip(wb)).not.toThrow();
  });

  // exceljs: spec/integration/issues/issue-176-unexpected-xml-node.spec.js — "issue 176 - Unexpected xml node in parseOpen"
  it('issue 176: reads test-issue-176.xlsx', () => {
    expect(readXlsx(fixture('test-issue-176.xlsx')).worksheets.length).toBeGreaterThan(0);
  });

  describe('issue 1842 - Memory overload when unnecessary dataValidations apply', () => {
    // turbosheet has no `ignoreNodes`; the equivalent is not enabling a feature reader
    // (turbosheet/lite), which keeps that XML raw instead of modelling it.
    const readers = [
      tablesReader,
      commentsReader,
      imagesReader,
      conditionalFormattingReader,
      printSetupReader,
      sheetProtectionReader,
    ];

    // exceljs: spec/integration/issues/issue-1842-dataValidations-memory-overload.spec.js — "when using readFile"
    it('reads the file with the data-validation reader disabled (lite)', () => {
      const wb = readLite(fixture('test-issue-1842.xlsx'), { features: readers });
      expect(wb.worksheets.length).toBeGreaterThan(0);
      for (const ws of wb.worksheets) expect(ws.dataValidations).toBeUndefined();
    });

    // exceljs: spec/integration/issues/issue-1842-dataValidations-memory-overload.spec.js — "when loading an in memory buffer"
    it('reads the file with every feature (validations modelled, not expanded per cell)', () => {
      const wb = readXlsx(fixture('test-issue-1842.xlsx'));
      const n = wb.worksheets.reduce((a, ws) => a + (ws.dataValidations?.length ?? 0), 0);
      expect(n).toBeGreaterThan(0);
    });
  });

  // exceljs: spec/integration/issues/issue-2125-spliceRows-last-row.spec.js — "issue 2125 - spliceRows remove last row"
  it.skip('issue 2125: spliceRows removing the last row — not supported (no spliceRows)', () => {});

  describe('issue 219 - 1904 dates not supported', () => {
    // exceljs: spec/integration/issues/issue-219-1904-dates.spec.js — "Reading 1904.xlsx"
    it('reads 1904.xlsx', () => {
      const wb = readXlsx(fixture('1904.xlsx'));
      expect(wb.date1904).toBe(true);
      const v = wb.getWorksheet('Sheet1')!.getValue('B4') as Date;
      expect(v.toISOString()).toBe('1904-01-01T00:00:00.000Z');
    });

    // exceljs: spec/integration/issues/issue-219-1904-dates.spec.js — "Writing and Reading"
    it('writes and reads 1904 dates', () => {
      const wb = new Workbook({ date1904: true });
      wb.addWorksheet('Sheet1').setCell('B4', new Date('1904-01-01T00:00:00.000Z'));
      const wb2 = roundtrip(wb);
      expect(wb2.date1904).toBe(true);
      const v = wb2.getWorksheet('Sheet1')!.getValue('B4') as Date;
      expect(v.toISOString()).toBe('1904-01-01T00:00:00.000Z');
    });
  });

  // exceljs: spec/integration/issues/issue-234-vertical-tab-char.spec.js — "issue 234 - Broken XLSX because of \"vertical tab\" ascii character in a cell"
  it('issue 234: control characters keep the file valid (kept as _xHHHH_, not stripped)', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Hello, \x01World!';
    ws.getCell('A2').value = 'Hello, \x0bWorld!';
    for (const strings of ['shared', 'inline'] as const) {
      const bytes = writeXlsx(wb, { strings });
      const ws2 = readXlsx(bytes).getWorksheet('Sheet1')!;
      // ➖ exceljs strips them ('Hello, World!'); turbosheet writes Excel's own `_x0001_`
      // escape, so the text round-trips exactly and the XML stays valid.
      expect(ws2.getValue('A1')).toBe('Hello, \x01World!');
      expect(ws2.getValue('A2')).toBe('Hello, \x0bWorld!');
      await exceljsLoad(bytes); // valid package
    }
  });

  // exceljs: spec/integration/issues/issue-257-sheet-order.spec.js — "issue 257 - worksheet order is not respected"
  it('issue 257: worksheet order follows workbook.xml', () => {
    const wb = readXlsx(fixture('test-issue-257.xlsx'));
    expect(wb.worksheets.map((ws) => ws.name)).toEqual(['First', 'Second']);
  });

  // exceljs: spec/integration/issues/issue-275-hyperlink-query-param.spec.js — "issue 275 - hyperlink with query arguments corrupts workbook"
  it('issue 275: hyperlink with query arguments and XML specials (streaming writer)', async () => {
    const hyperlink = {
      text: 'Somewhere with query params',
      hyperlink: 'www.somewhere.com?a=1&b=2&c=<>&d="\'"',
    };
    const bytes = await streamWrite(async (w) => {
      w.addWorksheet('Sheet1').addRow([hyperlink]);
    });
    expect(readXlsx(bytes).getWorksheet('Sheet1')!.getValue('A1')).toEqual(hyperlink);
    const x = await exceljsLoad(bytes);
    expect(x.getWorksheet('Sheet1')!.getCell('A1').value).toEqual(hyperlink);
  });

  // exceljs: spec/integration/issues/issue-539-contentType-element.spec.js — "Reading 1904.xlsx"
  it('issue 539: reads a file with a <contentType/> element', () => {
    const wb = readXlsx(fixture('1519293514-KRISHNAPATNAM_LINE_UP.xlsx'));
    expect(wb.worksheets.length).toBeGreaterThan(0);
  });

  // exceljs: spec/integration/issues/issue-623-borders-for-merged-cell.spec.js — "issue 623 - Issue with borders for merged cell when rewriting an excel workbook"
  it('issue 623: per-cell borders of a merged range are read as-is and survive a rewrite', () => {
    const check = (wb: Workbook) => {
      const ws = wb.worksheets[0]!;
      const has = (ref: string, sides: string[]) => {
        const b = ws.getStyle(ref).border as Record<string, unknown> | undefined;
        for (const s of sides) expect(b?.[s], `${ref}.${s}`).toBeDefined();
      };
      has('B2', ['left', 'top']);
      has('B3', ['left', 'bottom']);
      has('C2', ['right', 'top']);
      has('C3', ['right', 'bottom']);
    };
    const wb = readXlsx(fixture('test-issue-623.xlsx'));
    check(wb);
    check(roundtrip(wb));
  });

  // exceljs: spec/integration/issues/issue-703-speciel-cell-file.spec.js — "issue 703 - Special cell value results invalid file"
  it('issue 703: Object.prototype names as cell strings (in-memory writer)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    SPECIAL.forEach((v, i) => {
      ws.addRow([v]);
      ws.getCell(`B${i + 1}`).value = v;
    });
    const ws2 = roundtrip(wb).getWorksheet('Sheet1')!;
    SPECIAL.forEach((v, i) => {
      expect(ws2.getValue(`A${i + 1}`)).toBe(v);
      expect(ws2.getValue(`B${i + 1}`)).toBe(v);
    });
  });

  // exceljs: spec/integration/issues/issue-771-data-validation-without-type.spec.js — "issue 771 - Issue with dataValidation without type and with formula1 or formula2"
  it('issue 771: data validation without a type (reads as "any") and round-trips', () => {
    const wb = readXlsx(fixture('test-issue-771.xlsx'));
    const dvs = wb.worksheets.flatMap((ws) => ws.dataValidations ?? []);
    expect(dvs.length).toBeGreaterThan(0);
    expect(dvs.some((d) => d.type === 'any')).toBe(true);
    expect(() => roundtrip(wb)).not.toThrow();
  });

  // exceljs: spec/integration/issues/issue-877-hyperlink-no-text.spec.js — "issue 877 - hyperlink without text crashes on write"
  it('issue 877: hyperlink without text does not crash on write', async () => {
    const wb = readXlsx(fixture('test-issue-877.xlsx'));
    const bytes = writeXlsx(wb);
    expect(readXlsx(bytes).worksheets.length).toBe(wb.worksheets.length);
    await exceljsLoad(bytes);
  });

  // exceljs: spec/integration/issues/issue-880-malformed-comment.spec.js — "issue 880 - malformed comment crashes on write"
  it('issue 880: malformed comment does not crash on write', async () => {
    const wb = readXlsx(fixture('test-issue-880.xlsx'));
    const bytes = writeXlsx(wb);
    expect(readXlsx(bytes).worksheets.length).toBe(wb.worksheets.length);
    await exceljsLoad(bytes);
  });

  // exceljs: spec/integration/issues/issue-988-missing-table-autofilter.spec.js — "issue 988 - table without autofilter model"
  it('issue 988: table without autoFilter reads and round-trips', () => {
    const wb = readXlsx(fixture('test-issue-988.xlsx'));
    const tables = wb.worksheets.flatMap((ws) => ws.tables ?? []);
    expect(tables.length).toBeGreaterThan(0);
    const back = roundtrip(wb).worksheets.flatMap((ws) => ws.tables ?? []);
    expect(back.map((t) => t.name)).toEqual(tables.map((t) => t.name));
  });

  // exceljs: spec/integration/issues/issue-991-csv-read-dates.spec.js — "issue 991 - differentiates between strings with leading numbers and dates when reading csv files"
  it('issue 991: CSV date detection does not turn "00210PRG1"-like strings into dates', () => {
    const text = new TextDecoder().decode(fixture('test-issue-991.csv'));
    const col = parseCsv(text, { parseDates: true }).map((r) => r[0]);
    // ISO dates are parsed as UTC (exceljs: local time via dayjs)
    expect(col[0]).toEqual(new Date('2019-11-04T00:00:00Z'));
    // ➖ "11-04-2019" (US format) is not ISO 8601: turbosheet keeps it a string
    expect(col[1]).toBe('11-04-2019');
    expect(col[2]).toEqual(new Date('2019-11-04T10:17:55Z'));
    expect(col[3]).toBe('00210PRG1');
    expect(col[4]).toBe('1234-5thisisnotadate');
  });

  // exceljs: spec/integration/issues/issue-995-encoding.spec.js — "issue 995 - encoding option works fine"
  it('issue 995: non-ASCII text survives a CSV write/read (UTF-8)', async () => {
    const HEBREW = 'משהו שכתוב בעברית';
    const bytes = await collect(csvStream([[HEBREW]]));
    expect(parseCsv(new TextDecoder().decode(bytes))[0]![0]).toBe(HEBREW);
  });

  describe('Shared Formulas - issue xyz - cells copied as a block treat formulas as values', () => {
    // exceljs: spec/integration/issues/shared-formula.spec.js — "copied cells should have the right formulas"
    it('copied cells have the right formulas (shared formulas are expanded on read)', () => {
      const ws = readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!;
      // ➖ exceljs: { formula, shareType: 'shared', ref } / { sharedFormula: 'A4' }
      expect(ws.getValue('A4')).toEqual({ formula: 'A3+1', result: 4 });
      expect(ws.getValue('A5')).toEqual({ formula: 'A4+1', result: 5 });
      expect(translateFormula('A3+1', 1, 0)).toBe('A4+1');
    });

    // exceljs: spec/integration/issues/shared-formula.spec.js — "copied cells should have the right types"
    it('copied cells have the right types', () => {
      const ws = readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!;
      expect(ws.getKind('A4')).toBe(CellKind.Formula);
      expect(ws.getKind('A5')).toBe(CellKind.Formula);
    });

    // exceljs: spec/integration/issues/shared-formula.spec.js — "copied cells should have the same fields"
    it('copied cells have the same fields', () => {
      const ws = readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!;
      const a4 = ws.getValue('A4') as object;
      const a5 = ws.getValue('A5') as object;
      expect(Object.keys(a4).join()).toBe(Object.keys(a5).join());
    });
  });
});
