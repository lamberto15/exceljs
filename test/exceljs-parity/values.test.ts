// Cell values: the ExcelJS test book round trips (spec/integration/workbook/workbook.spec.js
// "Serialise"), worksheet value semantics (spec/integration/worksheet.spec.js "Values") and the
// gold file.

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';
import { checkTestBook, createTestBook, createTestBookStream } from './test-book.js';

describe('exceljs parity: values', () => {
  describe('Workbook Serialise (test book)', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file"
    it('xlsx file', async () => {
      const bytes = writeXlsx(createTestBook(new Workbook()));
      checkTestBook(readXlsx(bytes));
      await exceljsLoad(bytes);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file with best compression"
    // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file with default compression"
    // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file with fast compression"
    // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file with no compression"
    it.each([
      ['best', 9],
      ['default', undefined],
      ['fast', 1],
      ['no', 0],
    ] as const)('xlsx file with %s compression', (_name, compression) => {
      const wb = createTestBook(new Workbook());
      const bytes = writeXlsx(wb, compression === undefined ? {} : { compression });
      checkTestBook(readXlsx(bytes));
    });

    // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "xlsx file"
    // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "Without styles"
    it('streaming writer: xlsx file (row-level styles only)', async () => {
      const bytes = await streamWrite((w) => createTestBookStream(w));
      checkTestBook(readXlsx(bytes), { checkStyles: false });
      await exceljsLoad(bytes);
    });

    // exceljs: spec/end-to-end/express.spec.js — "downloads a workbook"
    it('serves a workbook as a web stream and reads it back (express e2e)', async () => {
      const { xlsxStream } = await import('../../src/index.js');
      const res = new Response(xlsxStream(createTestBook(new Workbook())), {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
      });
      checkTestBook(readXlsx(new Uint8Array(await res.arrayBuffer())));
    });
  });

  describe('Worksheet Values', () => {
    // exceljs: spec/integration/worksheet.spec.js — "stores values properly"
    it('stores values properly', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('blort');
      const now = new Date();
      ws.getCell('A1').value = 7;
      ws.getCell('B1').value = 'Hello, World!';
      ws.getCell('C1').value = 3.14;
      ws.getCell('D1').value = 5;
      ws.getCell('D1').value = now;
      ws.getCell('E1').value = `${['Hello', 'World'].join(', ')}!`;
      ws.getCell('F1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
      ws.getCell('A2').value = { formula: 'A1', result: 7 };
      ws.getCell('B2').value = {
        formula: 'CONCATENATE("Hello", ", ", "World!")',
        result: 'Hello, World!',
      };
      expect(ws.getCell('A1').value).toBe(7);
      expect(ws.getCell('B1').value).toBe('Hello, World!');
      expect(ws.getCell('C1').value).toBe(3.14);
      expect(ws.getCell('D1').value).toEqual(now);
      expect(ws.getCell('E1').value).toBe('Hello, World!');
      expect(ws.getCell('F1').value).toEqual({
        text: 'www.google.com',
        hyperlink: 'http://www.google.com',
      });
      expect(ws.getCell('A2').value).toEqual({ formula: 'A1', result: 7 });
      expect(ws.getCell('B2').value).toEqual({
        formula: 'CONCATENATE("Hello", ", ", "World!")',
        result: 'Hello, World!',
      });
      // ➖ a formula result is a number/string/boolean/error, not a Date: store the serial
      expect(() => {
        ws.getCell('C2').value = { formula: 'D1', result: now as unknown as number };
      }).not.toThrow();
    });

    // exceljs: spec/integration/worksheet.spec.js — "stores shared string values properly"
    it('stores shared string values properly', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('blort');
      ws.getCell('A1').value = 'Hello, World!';
      ws.getCell('A2').value = 'Hello';
      ws.getCell('B2').value = 'World';
      ws.getCell('C2').value = {
        formula: 'CONCATENATE(A2, ", ", B2, "!")',
        result: 'Hello, World!',
      };
      ws.getCell('A3').value = `${['Hello', 'World'].join(', ')}!`;
      expect(ws.getCell('A1').value).toBe(ws.getCell('A3').value);
      expect(ws.getCell('A1').value).toBe((ws.getCell('C2').value as { result: string }).result);
      // and they are written once to sharedStrings.xml
      const ws2 = roundtrip(wb).getWorksheet('blort')!;
      expect(ws2.getValue('A3')).toBe('Hello, World!');
    });

    // exceljs: spec/integration/worksheet.spec.js — "assigns cell types properly"
    it('assigns cell types properly', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('blort');
      ws.getCell('A1').value = 7;
      ws.getCell('B1').value = 'Hello, World!';
      ws.getCell('C1').value = 3.14;
      ws.getCell('D1').value = new Date();
      ws.getCell('E1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
      ws.getCell('A2').value = { formula: 'A1', result: 7 };
      ws.getCell('B2').value = { formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'x' };
      ws.getCell('C2').value = { formula: 'D1' };
      expect(ws.getCell('A1').kind).toBe(CellKind.Number);
      expect(ws.getCell('B1').kind).toBe(CellKind.String);
      expect(ws.getCell('C1').kind).toBe(CellKind.Number);
      expect(ws.getCell('D1').kind).toBe(CellKind.Date);
      expect(ws.getCell('E1').kind).toBe(CellKind.Hyperlink);
      expect(ws.getCell('A2').kind).toBe(CellKind.Formula);
      expect(ws.getCell('B2').kind).toBe(CellKind.Formula);
      expect(ws.getCell('C2').kind).toBe(CellKind.Formula);
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "empty strings"
    it('empty strings round-trip as ""', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'Foo';
      ws.getCell('A2').value = '';
      ws.getCell('A3').value = 'Baz';
      for (const strings of ['shared', 'inline'] as const) {
        const bytes = writeXlsx(wb, { strings });
        const ws2 = readXlsx(bytes).getWorksheet('Hello')!;
        expect(ws2.getValue('A1')).toBe('Foo');
        expect(ws2.getValue('A2')).toBe('');
        expect(ws2.getValue('A3')).toBe('Baz');
        const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
        // exceljs itself reads an empty inline string as null
        expect(x.getCell('A2').value).toBe(strings === 'shared' ? '' : null);
      }
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "empty string"
    it('empty string in a keyed row writes', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();
      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
      ];
      ws.addRow({ id: 1, name: '' });
      const ws2 = roundtrip(wb).worksheets[0]!;
      expect(ws2.getValue('A1')).toBe(1);
      expect(ws2.getValue('B1')).toBe('');
    });

    // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "rich text"
    it('rich text (streaming writer)', async () => {
      const rich = {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red ' },
          { font: { color: { argb: '00FF00' }, bold: true }, text: ' bold green' },
        ],
      };
      const bytes = await streamWrite((w) => {
        w.addWorksheet('Hello').addRow([rich, 'plain text']);
      });
      const ws2 = readXlsx(bytes).getWorksheet('Hello')!;
      // ➖ 6-digit argb colors read back with an alpha byte (exceljs keeps 'FF0000')
      expect(ws2.getValue('A1')).toEqual({
        richText: [
          { font: { color: { argb: 'FFFF0000' } }, text: 'red ' },
          { font: { color: { argb: 'FF00FF00' }, bold: true }, text: ' bold green' },
        ],
      });
      expect(ws2.getValue('B1')).toBe('plain text');
    });

    // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "does not escape special xml characters"
    it('XML special characters round-trip (streaming writer, shared strings)', async () => {
      const s = 'xml characters: & < > "';
      const bytes = await streamWrite(
        (w) => {
          w.addWorksheet('blort').addRow([s]);
        },
        { strings: 'shared' },
      );
      expect(readXlsx(bytes).getWorksheet('blort')!.getValue('A1')).toBe(s);
    });

    // exceljs: spec/unit/doc/workbook.spec.js — "assigns rich text"
    it('assigns rich text (kind RichText)', () => {
      const ws = new Workbook().addWorksheet('blort');
      const runs = [
        { font: { size: 12, color: { theme: 0 }, name: 'Calibri' }, text: 'This is ' },
        { font: { italic: true, size: 12, color: { theme: 0 }, name: 'Calibri' }, text: 'a' },
        { text: ' colorful text' },
      ];
      ws.getCell('A1').value = { richText: runs };
      expect(ws.getCell('A1').kind).toBe(CellKind.RichText);
      // ➖ no `cell.text`; join the runs
      expect(
        (ws.getValue('A1') as { richText: { text: string }[] }).richText
          .map((r) => r.text)
          .join(''),
      ).toBe('This is a colorful text');
    });
  });

  describe('Gold Book', () => {
    // exceljs: spec/integration/gold.spec.js — "Values"
    it('reads the gold file values', () => {
      const ws = readXlsx(fixture('gold.xlsx')).getWorksheet('Values')!;
      expect(ws.getValue('B1')).toBe('I am Text');
      expect(ws.getValue('B2')).toBe(3.14);
      expect(ws.getValue('B3')).toBe(5);
      expect(ws.getValue('B4')).toEqual(new Date('2016-05-17T00:00:00.000Z'));
      expect(ws.getValue('B5')).toEqual({ formula: 'B1', result: 'I am Text' });
      expect(ws.getValue('B6')).toEqual({
        hyperlink: 'https://www.npmjs.com/package/exceljs',
        text: 'exceljs',
      });
      // ➖ no lastColumn / lastRow objects
      expect(ws.columnCount).toBe(2);
      expect(ws.rowCount).toBe(6);
    });

    // exceljs: spec/integration/gold.spec.js — "Styles"
    it('Styles (empty in exceljs)', () => {
      expect(readXlsx(fixture('gold.xlsx')).getWorksheet('Values')).toBeDefined();
    });
  });
});
