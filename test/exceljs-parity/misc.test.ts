// Workbook-level behaviour, sheet views, reader robustness, pivot tables and the behaviour-level
// unit specs (spec/unit/utils/col-cache.spec.js, utils.spec.js, spec/unit/doc/workbook.spec.js).

import { describe, expect, it } from 'vitest';
import {
  CellKind,
  dateToSerial,
  decodeAddr,
  decodeCol,
  decodeRange,
  encodeCol,
  isDateFormat,
  readXlsx,
  serialToDate,
  Workbook,
  writeXlsx,
  XlsxParseError,
} from '../../src/index.js';
import { escapeAttr, escapeText } from '../../src/xml/index.js';
import { exceljsLoad, fixture, partNames, roundtrip } from './helpers.js';

describe('exceljs parity: misc', () => {
  describe('Workbook', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "sheets with correct names"
    it('sheets with correct names (incl. "&")', () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet('Hello, World!');
      expect(ws1.name).toBe('Hello, World!');
      ws1.getCell('A1').value = 'Hello, World!';
      const ws2 = wb.addWorksheet();
      expect(ws2.name).toMatch(/^Sheet\d+$/); // ➖ exceljs: "sheetN"
      wb.addWorksheet('This & That');
      const wb2 = roundtrip(wb);
      expect(wb2.getWorksheet('Hello, World!')).toBeDefined();
      expect(wb2.getWorksheet('This & That')).toBeDefined();
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "creator, lastModifiedBy, etc"
    it('creator, lastModifiedBy, created, modified', () => {
      const wb = new Workbook();
      wb.addWorksheet('Hello').getCell('A1').value = 'World!';
      wb.creator = 'Foo';
      wb.lastModifiedBy = 'Bar';
      wb.created = new Date(Date.UTC(2016, 0, 1));
      wb.modified = new Date(Date.UTC(2016, 4, 19));
      const wb2 = roundtrip(wb);
      expect(wb2.creator).toBe('Foo');
      expect(wb2.lastModifiedBy).toBe('Bar');
      expect(wb2.created).toEqual(wb.created);
      expect(wb2.modified).toEqual(wb.modified);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "company, manager, etc"
    it('company, manager', async () => {
      const wb = new Workbook();
      wb.addWorksheet('Hello').getCell('A1').value = 'World!';
      wb.company = 'Cyber Sapiens, Ltd';
      wb.manager = 'Guyon Roche';
      const bytes = writeXlsx(wb);
      const wb2 = readXlsx(bytes);
      expect(wb2.company).toBe(wb.company);
      expect(wb2.manager).toBe(wb.manager);
      const x = await exceljsLoad(bytes);
      expect(x.company).toBe(wb.company);
      expect(x.manager).toBe(wb.manager);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "title, subject, etc"
    it('title, subject, keywords, category, description', () => {
      const wb = new Workbook();
      wb.addWorksheet('Hello').getCell('A1').value = 'World!';
      wb.title = 'the title';
      wb.subject = 'the subject';
      wb.keywords = 'the keywords';
      wb.category = 'the category';
      wb.description = 'the description';
      const wb2 = roundtrip(wb);
      expect(wb2.title).toBe(wb.title);
      expect(wb2.subject).toBe(wb.subject);
      expect(wb2.keywords).toBe(wb.keywords);
      expect(wb2.category).toBe(wb.category);
      expect(wb2.description).toBe(wb.description);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "language, revision and contentStatus"
    it('language, revision and contentStatus', async () => {
      const wb = new Workbook();
      wb.addWorksheet('Hello').getCell('A1').value = 'World!';
      wb.language = 'Klingon';
      wb.revision = 2;
      wb.contentStatus = 'Final'; // the exceljs spec misspells it (contentStauts)
      const bytes = writeXlsx(wb);
      const wb2 = readXlsx(bytes);
      expect(wb2.language).toBe(wb.language);
      expect(wb2.revision).toBe(wb.revision);
      expect(wb2.contentStatus).toBe(wb.contentStatus);
      // exceljs's typings lack language / revision
      const x = (await exceljsLoad(bytes)) as unknown as { language: string; revision: number };
      expect(x.language).toBe('Klingon');
      expect(x.revision).toBe(2);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "auto filter"
    it('auto filter', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.addRows([
        [1, 1],
        [2, 2],
        [3, 3],
      ]);
      ws.autoFilter = 'A1:B1';
      const bytes = writeXlsx(wb);
      expect(readXlsx(bytes).getWorksheet('Hello')!.autoFilter).toBe('A1:B1');
      expect((await exceljsLoad(bytes)).getWorksheet('Hello')!.autoFilter).toBe('A1:B1');
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "a lot of sheets to xlsx file"
    it('a lot of sheets (90)', () => {
      const wb = new Workbook();
      for (let i = 1; i <= 90; i++) wb.addWorksheet(`sheet${i}`).getCell('A1').value = i;
      const wb2 = roundtrip(wb);
      for (let i = 1; i <= 90; i++) expect(wb2.getWorksheet(`sheet${i}`)!.getValue('A1')).toBe(i);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "throws an error when xlsx file not found"
    it.skip('xlsx.readFile of a missing file — not applicable (no file system API)', () => {});

    // exceljs: spec/integration/workbook/workbook.spec.js — "throw an error for wrong data type"
    it('throws for a wrong data type', () => {
      expect(() => readXlsx({} as unknown as Uint8Array)).toThrow();
    });
  });

  describe('Sheet Views', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "frozen panes"
    it('frozen panes', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('frozen');
      ws.views = [
        { state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4', activeCell: 'D5' },
        { state: 'frozen', ySplit: 1 },
        { state: 'frozen', xSplit: 1 },
      ];
      ws.getCell('A1').value = 'Let it Snow!';
      const ws2 = roundtrip(wb).getWorksheet('frozen')!;
      expect(ws2.getValue('A1')).toBe('Let it Snow!');
      // ➖ only modelled fields (no workbookViewId / showRuler / zoomScaleNormal defaults)
      expect(ws2.views[0]).toMatchObject({
        state: 'frozen',
        xSplit: 2,
        ySplit: 3,
        topLeftCell: 'C4',
        activeCell: 'D5',
      });
      // one view per Excel window (workbookViewId = position in ws.views)
      expect(ws2.views).toHaveLength(3);
      expect(ws2.views[1]).toMatchObject({ state: 'frozen', ySplit: 1, topLeftCell: 'A2' });
      expect(ws2.views[2]).toMatchObject({ state: 'frozen', xSplit: 1, topLeftCell: 'B1' });
      // and on sheets of their own:
      const wb3 = new Workbook();
      wb3.addWorksheet('a').views = [{ state: 'frozen', ySplit: 1 }];
      wb3.addWorksheet('b').views = [{ state: 'frozen', xSplit: 1 }];
      const back = roundtrip(wb3);
      expect(back.getWorksheet('a')!.views[0]).toMatchObject({
        state: 'frozen',
        ySplit: 1,
        topLeftCell: 'A2',
      });
      expect(back.getWorksheet('b')!.views[0]).toMatchObject({
        state: 'frozen',
        xSplit: 1,
        topLeftCell: 'B1',
      });
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "serialises split panes"
    it('split panes', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('split');
      ws.views = [
        { state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'C4', activeCell: 'D5' },
        { state: 'split', ySplit: 1500, topLeftCell: 'A10' },
        { state: 'split', xSplit: 1500 },
      ];
      ws.getCell('A1').value = 'Do the splits!';
      const ws2 = roundtrip(wb).getWorksheet('split')!;
      // 🚫 activePane is not modelled
      expect(ws2.views[0]).toMatchObject({
        state: 'split',
        xSplit: 2000,
        ySplit: 3000,
        topLeftCell: 'C4',
        activeCell: 'D5',
      });
      expect(ws2.views).toHaveLength(3);
      expect(ws2.views[1]).toMatchObject({ state: 'split', ySplit: 1500, topLeftCell: 'A10' });
      expect(ws2.views[2]).toMatchObject({ state: 'split', xSplit: 1500 });
      const wb3 = new Workbook();
      wb3.addWorksheet('a').views = [{ state: 'split', ySplit: 1500, topLeftCell: 'A10' }];
      wb3.addWorksheet('b').views = [{ state: 'split', xSplit: 1500 }];
      const back = roundtrip(wb3);
      expect(back.getWorksheet('a')!.views[0]).toMatchObject({
        state: 'split',
        ySplit: 1500,
        topLeftCell: 'A10',
      });
      expect(back.getWorksheet('b')!.views[0]).toMatchObject({ state: 'split', xSplit: 1500 });
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "multiple book views"
    it('multiple book views', async () => {
      const wb = new Workbook();
      // spec/utils/data/views.json
      const book = [
        {
          x: 0,
          y: 0,
          width: 10000,
          height: 20000,
          firstSheet: 0,
          activeTab: 1,
          visibility: 'visible' as const,
        },
        {
          x: 1,
          y: 2,
          width: 10001,
          height: 20002,
          firstSheet: 1,
          activeTab: 1,
          visibility: 'hidden' as const,
        },
      ];
      wb.views = book;
      const ws1 = wb.addWorksheet('one');
      ws1.views = [
        {
          workbookViewId: 0,
          state: 'frozen',
          xSplit: 2,
          ySplit: 3,
          topLeftCell: 'C4',
          activeCell: 'D5',
        },
      ];
      const ws2 = wb.addWorksheet('two');
      ws2.views = [
        {
          workbookViewId: 1,
          state: 'split',
          xSplit: 2000,
          ySplit: 3000,
          topLeftCell: 'C4',
          activeCell: 'D5',
        },
      ];
      const bytes = writeXlsx(wb);
      const wb2 = readXlsx(bytes);
      // ➖ the first window's activeTab is carried by the sheets' tabSelected; 'visible' is the default
      expect(wb2.views).toEqual([
        { x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0 },
        {
          x: 1,
          y: 2,
          width: 10001,
          height: 20002,
          firstSheet: 1,
          activeTab: 1,
          visibility: 'hidden',
        },
      ]);
      // ➖ only fields that differ from the defaults are modelled; workbookViewId only when it is
      // not the view's position in ws.views
      expect(wb2.getWorksheet('one')!.views[0]).toMatchObject({
        state: 'frozen',
        xSplit: 2,
        ySplit: 3,
        topLeftCell: 'C4',
        activeCell: 'D5',
      });
      expect(wb2.getWorksheet('one')!.views[0]!.workbookViewId).toBeUndefined();
      expect(wb2.getWorksheet('two')!.views[0]).toMatchObject({
        workbookViewId: 1,
        state: 'split',
        xSplit: 2000,
        ySplit: 3000,
        topLeftCell: 'C4',
        activeCell: 'D5',
      });
      const x = await exceljsLoad(bytes);
      expect(x.views).toEqual(book);
      expect(x.getWorksheet('two')!.views[0]).toMatchObject({ workbookViewId: 1, state: 'split' });
    });
  });

  describe('reader robustness', () => {
    // exceljs: spec/integration/worksheet.spec.js — "Should not break when importing a .numbers file"
    it('a .numbers file reads as a workbook without worksheets', () => {
      expect(readXlsx(fixture('numbers.numbers')).worksheets).toHaveLength(0);
    });

    // exceljs: spec/integration/worksheet.spec.js — "Should not break when importing an Excel file that contains a chartsheet"
    it('a file with a chartsheet has one worksheet', () => {
      expect(readXlsx(fixture('chart-sheet.xlsx')).worksheets).toHaveLength(1);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should decode the underscore"
    it('decodes an escaped underscore in shared strings', () => {
      expect(
        readXlsx(fixture('shared_string_with_escape.xlsx')).worksheets[0]!.getValue('A1'),
      ).toBe('_x000D_');
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should reject the promise with the sax error"
    it('an XML parse error in a worksheet throws XlsxParseError', () => {
      expect(() => readXlsx(fixture('invalid-xml.xlsx'))).toThrow(XlsxParseError);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should not break" (missing-bits.xlsx)
    it('a package missing some parts still reads', () => {
      expect(readXlsx(fixture('missing-bits.xlsx')).worksheets.length).toBeGreaterThan(0);
    });
  });

  describe('Pivot Tables', () => {
    // exceljs: spec/integration/workbook/pivot-tables.spec.js — "if pivot table added, then certain xml and rels files are added"
    it.skip('addPivotTable — not supported (pivot tables in read files are preserved)', () => {});

    // exceljs: spec/integration/workbook/pivot-tables.spec.js — "if pivot table NOT added, then certain xml and rels files are not added"
    it('no pivot parts are written when there is no pivot table', () => {
      const wb = new Workbook();
      wb.addWorksheet('Sheet1').addRows([
        ['A', 'B'],
        ['a1', 4],
      ]);
      wb.addWorksheet('Sheet2');
      expect(partNames(writeXlsx(wb)).filter((p) => /pivot/i.test(p))).toEqual([]);
    });
  });

  describe('unit: workbook', () => {
    // exceljs: spec/unit/doc/workbook.spec.js — "returns undefined for non-existant sheet"
    it('returns undefined for a missing sheet', () => {
      const wb = new Workbook();
      wb.addWorksheet('first');
      expect(wb.getWorksheet('w00t')).toBeUndefined();
    });

    // exceljs: spec/unit/doc/workbook.spec.js — "returns undefined for sheet 0"
    // exceljs: spec/unit/doc/workbook.spec.js — "returns undefined for sheet 0 after accessing wb.worksheets or wb.eachSheet "
    it('returns undefined for sheet id 0', () => {
      const wb = new Workbook();
      wb.addWorksheet('first');
      expect(wb.getWorksheet(0)).toBeUndefined();
      expect(wb.worksheets).toHaveLength(1);
      expect(wb.getWorksheet(0)).toBeUndefined();
    });

    // exceljs: spec/unit/doc/workbook.spec.js — "stores shared string values properly"
    // exceljs: spec/unit/doc/workbook.spec.js — "assigns cell types properly"
    it('cell kinds of a simple workbook', () => {
      const ws = new Workbook().addWorksheet('blort');
      ws.setCell('A1', 7);
      ws.setCell('B1', 'Hello, World!');
      ws.setCell('D1', new Date());
      ws.setCell('E1', { text: 'www.google.com', hyperlink: 'http://www.google.com' });
      ws.setCell('A2', { formula: 'A1', result: 7 });
      expect(['A1', 'B1', 'D1', 'E1', 'A2'].map((r) => ws.getKind(r))).toEqual([
        CellKind.Number,
        CellKind.String,
        CellKind.Date,
        CellKind.Hyperlink,
        CellKind.Formula,
      ]);
    });
  });

  describe('unit: col-cache (address helpers, 0-based)', () => {
    // exceljs: spec/unit/utils/col-cache.spec.js — "converts numbers to letters"
    it('converts numbers to letters', () => {
      expect([1, 26, 27, 702, 703].map((n) => encodeCol(n - 1))).toEqual([
        'A',
        'Z',
        'AA',
        'ZZ',
        'AAA',
      ]);
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "converts letters to numbers"
    it('converts letters to numbers', () => {
      expect(['A', 'Z', 'AA', 'ZZ', 'AAA'].map((l) => decodeCol(l) + 1)).toEqual([
        1, 26, 27, 702, 703,
      ]);
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "throws when out of bounds"
    it('throws when out of bounds', () => {
      for (const n of [-1, -2, 16384]) expect(() => encodeCol(n)).toThrow(RangeError);
      for (const l of ['', 'AAAA', 'XFE']) expect(() => decodeCol(l)).toThrow(RangeError);
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "validates addresses properly"
    it('validates addresses', () => {
      for (const a of ['A1', 'AA10', 'ABC100000']) expect(() => decodeAddr(a)).not.toThrow();
      for (const a of ['A', '1', '1A', 'A 1', 'A1A', '1A1']) expect(() => decodeAddr(a)).toThrow();
      // ➖ lowercase addresses are accepted ('a1')
      expect(decodeAddr('a1')).toEqual({ row: 0, col: 0 });
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "decodes addresses"
    it('decodes addresses', () => {
      expect(decodeAddr('A1')).toEqual({ row: 0, col: 0 });
      expect(decodeAddr('AA11')).toEqual({ row: 10, col: 26 });
      expect(decodeAddr('$AA$11')).toEqual({ row: 10, col: 26 });
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "tolerates a missing row number"
    // exceljs: spec/unit/utils/col-cache.spec.js — "tolerates a missing column number"
    it('a lone column / row is a whole-column / whole-row range, not an address', () => {
      // ➖ decodeAddr('$B') throws; decodeRange handles 'B:B' / '2:2'
      expect(() => decodeAddr('$B')).toThrow();
      expect(decodeRange('B:B')).toEqual({ top: 0, left: 1, bottom: 1048575, right: 1 });
      expect(decodeRange('2:2')).toEqual({ top: 1, left: 0, bottom: 1, right: 16383 });
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "convert [sheetName!][$]col[$]row[[$]col[$]row] into address or range structures"
    it('sheet-qualified addresses (quoted names with "!", "$", ":")', () => {
      expect(decodeAddr('Sheet1!$H$1')).toEqual({ row: 0, col: 7 });
      expect(decodeAddr("'Sheet 1'!$H$1")).toEqual({ row: 0, col: 7 });
      expect(decodeAddr("'Sheet !$:1'!$H$1")).toEqual({ row: 0, col: 7 });
      // ➖ no sheetName in the result, and '#REF!' is an error
      expect(() => decodeAddr("'Sheet !$:1'!#REF!")).toThrow();
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "decodes addresses and ranges"
    it('decodes ranges (normalizing reversed corners)', () => {
      const r = { top: 0, left: 0, bottom: 1, right: 1 };
      expect(decodeRange('A1:B2')).toEqual(r);
      expect(decodeRange('A2:B1')).toEqual(r);
      expect(decodeRange('B2:A1')).toEqual(r);
      expect(decodeRange('A1')).toEqual({ top: 0, left: 0, bottom: 0, right: 0 });
    });

    // exceljs: spec/unit/utils/col-cache.spec.js — "caches values"
    // exceljs: spec/unit/utils/col-cache.spec.js — "gets address structures (and caches them)"
    it.skip('colCache internals (_l2n / _n2l / getAddress identity) — not applicable', () => {});
  });

  describe('unit: utils', () => {
    // exceljs: spec/unit/utils/utils.spec.js — "encodes xml text"
    it('XML escaping strips invalid control characters', () => {
      expect(escapeText('<')).toBe('&lt;');
      expect(escapeText('>')).toBe('&gt;');
      expect(escapeText('&')).toBe('&amp;');
      expect(escapeAttr('"')).toBe('&quot;');
      expect(escapeText('abc\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0e\x0f\x10\x1f\x20abc')).toBe(
        'abc abc',
      );
      expect(escapeAttr('<a href="www.whatever.com">Talk to the H&</a>')).toBe(
        '&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;',
      );
      expect(escapeText('new\x0aline')).toBe('new\x0aline');
      // ➖ "'" needs no escaping in text or double-quoted attributes; DEL (\x7f) is legal XML
    });

    // exceljs: spec/unit/utils/utils.spec.js — "'yyyy-mm-dd' a date"
    // exceljs: spec/unit/utils/utils.spec.js — "'' is not a date"
    // exceljs: spec/unit/utils/utils.spec.js — "'[Green]#,##0 ;[Red](#,##0)' is not a date"
    it('date format detection', () => {
      expect(isDateFormat('yyyy-mm-dd')).toBe(true);
      expect(isDateFormat('')).toBe(false);
      expect(isDateFormat('[Green]#,##0 ;[Red](#,##0)')).toBe(false);
    });

    // exceljs: spec/unit/utils/utils.spec.js — "should convert date to excel properly"
    it('date to serial', () => {
      expect(dateToSerial(new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0)))).toBe(43084.70833333333);
    });

    // exceljs: spec/unit/utils/utils.spec.js — "should round to the nearest millisecond when parsing excel date"
    // exceljs: spec/unit/utils/utils.spec.js — "should not lost millisecond precision when parsing excel date"
    it('serial to date round-trips to the millisecond', () => {
      const d = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));
      expect(serialToDate(dateToSerial(d))).toEqual(d);
      const ms = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 123));
      expect(serialToDate(dateToSerial(ms))).toEqual(ms);
    });
  });
});
