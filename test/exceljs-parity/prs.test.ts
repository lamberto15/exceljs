// ExcelJS regression tests for pull requests (spec/integration/pr/**), ported to turbosheet.

import { describe, expect, it } from 'vitest';
import {
  addImage,
  addNote,
  placeImage,
  readXlsx,
  type Style,
  spliceColumns,
  Workbook,
  writeXlsx,
} from '../../src/index.js';
import { allRows, exceljsLoad, fixture, roundtrip, streamWrite, utilData } from './helpers.js';

const noteText = (wb: Workbook, sheet: string, ref: string) =>
  wb.getWorksheet(sheet)!.comments?.find((n) => n.ref === ref)?.text;

describe('exceljs parity: pull requests', () => {
  // exceljs: spec/integration/pr/test-pr-1157.spec.js — "pull request 1204 - Read and write data validation should be successful"
  it('pr 1204: reads textLength validations and writes them back', async () => {
    const wb = readXlsx(fixture('test-pr-1204.xlsx'));
    const dvs = wb.worksheets[0]!.dataValidations!;
    const cells = dvs.flatMap((d) => d.sqref.split(' '));
    expect(cells).toEqual(expect.arrayContaining(['E1', 'E4']));
    for (const d of dvs) {
      // ➖ formulas read back as formula text ('2'), exceljs: formulae: [2]
      expect(d).toMatchObject({
        type: 'textLength',
        operator: 'greaterThan',
        formula1: '2',
        showInputMessage: true,
        showErrorMessage: true,
      });
    }
    const back = roundtrip(wb).worksheets[0]!.dataValidations!;
    expect(back).toEqual(dvs);
    const x = await exceljsLoad(writeXlsx(wb));
    expect(x.worksheets[0]!.getCell('E4').dataValidation).toMatchObject({
      type: 'textLength',
      formulae: [2],
    });
  });

  // exceljs: spec/integration/pr/test-pr-1220.spec.js — "pull request 1220 - The worksheet should not be undefined"
  it('pr 1220: getWorksheet(1) is defined', () => {
    expect(readXlsx(fixture('test-pr-1220.xlsx')).getWorksheet(1)).toBeDefined();
  });

  // exceljs: spec/integration/pr/test-pr-1262.spec.js — "pull request 1262 - protect should work with streaming workbook writer"
  it('pr 1262: sheet protection with the streaming writer', async () => {
    const locked: Style = { protection: { locked: true } };
    const bytes = await streamWrite(async (w) => {
      const s = w.addWorksheet('data');
      s.addRow(['readonly cell'], locked);
      s.worksheet.protection = { password: 'pw1262', spinCount: 1 };
    });
    const ws = readXlsx(bytes).getWorksheet('data')!;
    expect(ws.protection?.spinCount).toBe(1);
    expect(ws.protection?.hashValue).toBeTruthy();
    const x = await exceljsLoad(bytes);
    expect(
      (x.getWorksheet('data') as unknown as { sheetProtection: { spinCount: number } })
        .sheetProtection.spinCount,
    ).toBe(1);
  });

  // exceljs: spec/integration/pr/test-pr-1334.spec.js — "pull request 1334 - Fix the error that comment does not delete at spliceColumn"
  it('pr 1334: spliceColumns moves and deletes notes', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('testSheet');
    ws.addRow(['test1', 'test2', 'test3', 'test4', 'test5', 'test6', 'test7', 'test8']);
    for (const [ref, text] of [
      ['A1', 'test1'],
      ['B1', 'test2'],
      ['C1', 'test3'],
      ['D1', 'test4'],
    ] as const)
      addNote(ws, ref, text);

    spliceColumns(ws, 2, 1);

    // the note of the removed column B is gone, the others moved left
    // (the original runs its assertions in a never-awaited async IIFE, so they cannot fail;
    // here they are checked, and again after a round trip instead of just writing the file)
    const check = (w: Workbook) => {
      expect(noteText(w, 'testSheet', 'A1')).toBe('test1');
      expect(noteText(w, 'testSheet', 'B1')).toBe('test3');
      expect(noteText(w, 'testSheet', 'C1')).toBe('test4');
      expect(noteText(w, 'testSheet', 'D1')).toBeUndefined();
    };
    check(wb);
    check(roundtrip(wb));
  });

  // exceljs: spec/integration/pr/test-pr-1431.spec.js — "pull request 1431 - streaming reader should handle rich text within shared strings"
  it('pr 1431: streaming reader returns rich text stored in shared strings', async () => {
    const rich = {
      richText: [
        { font: { bold: true }, text: 'This should ' },
        { font: { italic: true }, text: 'be one shared string value' },
      ],
    };
    const bytes = await streamWrite(
      async (w) => {
        w.addWorksheet('data').addRow([rich, 'this should be the second shared string']);
      },
      { strings: 'shared' },
    );
    const rows = await allRows(bytes);
    expect(rows).toHaveLength(1);
    expect(rows[0]!.cells[0]).toEqual(rich);
    expect(rows[0]!.cells[1]).toBe('this should be the second shared string');
  });

  // exceljs: spec/integration/pr/test-pr-1487.spec.js — "Reading 1904.xlsx"
  it('pr 1487: last column with an empty column (columnCount)', () => {
    // ➖ no `lastColumn` / `lastRow` Column/Row objects; columnCount / rowCount give the numbers
    const ws = readXlsx(fixture('1904.xlsx')).getWorksheet('Sheet1')!;
    expect(ws.columnCount).toBe(2);
    expect(ws.getColumn(ws.columnCount).number).toBe(2);
  });

  // exceljs: spec/integration/pr/test-pr-1576.spec.js — "Reading test-issue-1575.xlsx"
  it('pr 1576: inlineStr cells', () => {
    const ws = readXlsx(fixture('test-issue-1575.xlsx')).getWorksheet('Sheet1')!;
    const expected: Record<string, string> = {
      A1: 'A',
      B1: 'B',
      C1: 'C',
      A2: '1.0',
      B2: '2.0',
      C2: '3.0',
      A3: '4.0',
      B3: '5.0',
      C3: '6.0',
    };
    for (const [ref, v] of Object.entries(expected)) expect(ws.getValue(ref)).toBe(v);
  });

  // exceljs: spec/integration/pr/test-pr-2244.spec.js — "pull request 2244- Fix xlsx.writeFile() not catching error when error occurs"
  it('pr 2244: a missing picture is reported as an error', () => {
    // ➖ no `filename` source (no file system access): data/base64 is required, and a
    // placement that points at an unknown image id fails the write.
    const wb = new Workbook();
    const ws = wb.addWorksheet('sheet');
    expect(() => addImage(wb, { extension: 'jpeg' })).toThrow(TypeError);
    placeImage(ws, 42, 'B2:D6');
    expect(() => writeXlsx(wb)).toThrow(/Image 42 not found/);
  });

  // exceljs: spec/integration/pr/test-pr-567.spec.js — "Should be able to read this file"
  it('pr 567: whole-column defined names', () => {
    const wb = readXlsx(fixture('test-pr-567.xlsx'));
    expect(wb.worksheets.length).toBeGreaterThan(0);
    expect(() => roundtrip(wb)).not.toThrow();
  });

  // exceljs: spec/integration/pr/test-pr-728.spec.js — "pull request 728 - Read worksheet hidden state"
  it('pr 728: worksheet hidden state', () => {
    const wb = readXlsx(fixture('test-pr-728.xlsx'));
    const expected: Record<number, string> = { 1: 'visible', 2: 'hidden', 3: 'visible' };
    for (const ws of wb.worksheets) expect(ws.state).toBe(expected[ws.id]);
    expect(wb.worksheets).toHaveLength(3);
  });

  describe('pr 896 leading and trailing whitespace', () => {
    // exceljs: spec/integration/pr/test-pr-896.spec.js — "Should preserve leading and trailing whitespace"
    it('preserves leading and trailing whitespace in values and notes', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('foo');
      for (const [ref, v] of [
        ['A1', ' leading'],
        ['B1', 'trailing '],
        ['C1', ' both '],
      ] as const) {
        ws.getCell(ref).value = v;
        addNote(ws, ref, v);
      }
      for (const strings of ['shared', 'inline'] as const) {
        const wb2 = roundtrip(wb, { strings });
        const ws2 = wb2.getWorksheet('foo')!;
        expect(ws2.getValue('A1')).toBe(' leading');
        expect(noteText(wb2, 'foo', 'A1')).toBe(' leading');
        expect(ws2.getValue('B1')).toBe('trailing ');
        expect(noteText(wb2, 'foo', 'B1')).toBe('trailing ');
        expect(ws2.getValue('C1')).toBe(' both ');
        expect(noteText(wb2, 'foo', 'C1')).toBe(' both ');
      }
    });

    // exceljs: spec/integration/pr/test-pr-896.spec.js — "Should preserve newlines"
    it('preserves newlines in values and notes', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('foo');
      ws.getCell('A1').value = 'Hello,\nWorld!';
      addNote(ws, 'A1', 'Later,\nAlligator!');
      ws.getCell('B1').value = ' Hello, \n World! ';
      addNote(ws, 'B1', ' Later, \n Alligator! ');
      const bytes = writeXlsx(wb);
      const wb2 = readXlsx(bytes);
      const ws2 = wb2.getWorksheet('foo')!;
      expect(ws2.getValue('A1')).toBe('Hello,\nWorld!');
      expect(noteText(wb2, 'foo', 'A1')).toBe('Later,\nAlligator!');
      expect(ws2.getValue('B1')).toBe(' Hello, \n World! ');
      expect(noteText(wb2, 'foo', 'B1')).toBe(' Later, \n Alligator! ');
      const x = (await exceljsLoad(bytes)).getWorksheet('foo')!;
      expect(x.getCell('B1').note).toBe(' Later, \n Alligator! ');
    });
  });

  // exceljs: spec/integration/pr/pr-896/test-pr-896.spec.js — "should store cell text and comment with leading new line"
  it('pr 896: rich text value and note with newlines, with page setup and properties', () => {
    const runs = [
      { text: 'First Line:\n', font: { bold: true } },
      { text: 'Second Line\n' },
      { text: 'Third Line\n' },
      { text: 'Last Line' },
    ];
    const ps = utilData('page-setup');
    const props = utilData('sheet-properties');
    const wb = new Workbook();
    const ws = wb.addWorksheet('sheet1', { tabColor: props.tabColor });
    ws.pageSetup = { margins: ps.margins, orientation: ps.orientation, paperSize: ps.paperSize };
    ws.getColumn(1).width = 20;
    ws.getCell('A1').value = { richText: runs };
    addNote(ws, 'A1', { richText: runs });
    ws.getCell('A1').alignment = { wrapText: true };
    const wb2 = roundtrip(wb);
    const ws2 = wb2.getWorksheet('sheet1')!;
    expect(ws2.getValue('A1')).toEqual({ richText: runs });
    expect(noteText(wb2, 'sheet1', 'A1')).toEqual({ richText: runs });
  });
});
