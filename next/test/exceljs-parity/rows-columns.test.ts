// Rows, columns and sheet names: spec/integration/worksheet.spec.js, worksheet-xlsx-writer.spec.js,
// workbook.spec.js (duplicate rows, splice) and the behaviour-level unit specs
// (spec/unit/doc/worksheet.values.spec.js, worksheet.views.spec.js, worksheet.pagebreak.spec.js).

import { describe, expect, it } from 'vitest';
import { type CellValue, readXlsx, Workbook } from '../../src/index.js';
import { sheetToArrays } from '../../src/json/index.js';
import { fixture, roundtrip } from './helpers.js';

const d1 = new Date(Date.UTC(1970, 1, 1));
const d2 = new Date(Date.UTC(1965, 1, 7));

describe('exceljs parity: rows and columns', () => {
  // exceljs: spec/integration/worksheet.spec.js — "adds columns"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds columns"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds columns"
  it('adds columns (by key, letter and number: same Column object)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { key: 'id', width: 10 },
      { key: 'name', width: 32 },
      { key: 'dob', width: 10 },
    ];
    for (const [key, n, letter, width] of [
      ['id', 1, 'A', 10],
      ['name', 2, 'B', 32],
      ['dob', 3, 'C', 10],
    ] as const) {
      expect(ws.getColumn(key).number).toBe(n);
      expect(ws.getColumn(key).width).toBe(width);
      expect(ws.getColumn(letter)).toBe(ws.getColumn(key));
      expect(ws.getColumn(n)).toBe(ws.getColumn(key));
    }
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds column headers"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers"
  it('adds column headers', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { header: 'Id', width: 10 },
      { header: 'Name', width: 32 },
      { header: 'D.O.B.', width: 10 },
    ];
    expect(ws.getCell('A1').value).toBe('Id');
    expect(ws.getCell('B1').value).toBe('Name');
    expect(ws.getCell('C1').value).toBe('D.O.B.');
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds column headers by number"
  // exceljs: spec/integration/worksheet.spec.js — "adds column headers by letter"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers by number"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers by letter"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers by number"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers by letter"
  it('column key / header / width by property (header cells come from ws.columns only)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getColumn(2).key = 'name';
    ws.getColumn('B').header = 'Name';
    ws.getColumn(2).width = 32;
    expect(ws.getColumn('name')).toBe(ws.getColumn(2));
    expect(ws.getColumn('B').key).toBe('name');
    expect(ws.getColumn(2).header).toBe('Name');
    expect(ws.getColumn(2).width).toBe(32);
    // ➖ no `column.defn` / `headers`, and setting `header` does not write row 1
    expect(ws.getCell('B1').value).toBeNull();
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by object"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by object"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by object"
  it('adds rows by object', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'dob', width: 10 },
    ];
    expect(ws.addRow({ id: 1, name: 'John Doe', dob: d1 })).toBe(2);
    ws.addRow({ id: 2, name: 'Jane Doe', dob: d2 });
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('B2').value).toBe('John Doe');
    expect(ws.getCell('C2').value).toEqual(d1);
    expect(ws.getCell('C3').value).toEqual(d2);
    const values: CellValue[][] = [
      [],
      [undefined, 'Id', 'Name', 'D.O.B.'],
      [undefined, 1, 'John Doe', d1],
      [undefined, 2, 'Jane Doe', d2],
    ];
    expect(ws.getRow(2).values).toEqual(values[2]);
    expect(ws.getRow(3).values).toEqual(values[3]);
    let n = 0;
    ws.eachRow((row, rowNumber) => {
      n++;
      expect(row.values).toEqual(values[rowNumber]);
      // ➖ eachCell receives the value, not a Cell
      row.eachCell((v, col) => expect(v).toEqual(values[rowNumber]![col]));
    });
    expect(n).toBe(3);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by contiguous array"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by contiguous array"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by contiguous array"
  it('adds rows by contiguous array', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.addRow([1, 'John Doe', d1]);
    ws.addRow([2, 'Jane Doe', d2]);
    expect(ws.getCell('A1').value).toBe(1);
    expect(ws.getCell('C2').value).toEqual(d2);
    expect(ws.getRow(1).values).toEqual([undefined, 1, 'John Doe', d1]);
    expect(ws.getRow(2).values).toEqual([undefined, 2, 'Jane Doe', d2]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by sparse array"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by sparse array"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by sparse array"
  it('adds rows by sparse array (0-based arrays: index 0 is column A)', () => {
    const ws = new Workbook().addWorksheet('blort');
    // ➖ exceljs treats a sparse array as 1-based ([, 1, 'John Doe']); turbosheet arrays are 0-based
    const rows: CellValue[][] = [
      [1, 'John Doe', undefined, d1],
      [2, 'Jane Doe', undefined, d2],
    ];
    const row3: CellValue[] = [];
    row3[0] = 3;
    row3[2] = 'Sam';
    row3[4] = d1;
    rows.push(row3);
    for (const r of rows) ws.addRow(r);
    expect(ws.getCell('D1').value).toEqual(d1);
    expect(ws.getCell('C3').value).toBe('Sam');
    expect(ws.getCell('E3').value).toEqual(d1);
    ws.eachRow((row, rowNumber) => {
      expect(row.values).toEqual([undefined, ...rows[rowNumber - 1]!]);
    });
  });

  // regression (found porting workbook-xlsx-writer "shared formula"): a rejected value used to
  // leave a half-written row behind, shifting every later addRow down by one
  it('addRow is atomic: an invalid value adds no row', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.addRow([1]);
    expect(() => ws.addRow(['x', { sharedFormula: 'A1' } as never])).toThrow(TypeError);
    expect(ws.rowCount).toBe(1);
    expect(ws.addRow([2])).toBe(2);
    expect(ws.getValue('A2')).toBe(2);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows with style option"
  it.skip('addRow style inheritance codes ("i", "o", "n") — not supported: addRow takes a Style', () => {});

  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by object"
  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by contiguous array"
  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by sparse array"
  // exceljs: spec/integration/worksheet.spec.js — "inserts rows with style option"
  // exceljs: spec/integration/worksheet.spec.js — "should style of the inserted row with inherited style be mutable"
  it.skip('insertRow / insertRows — not supported', () => {});

  // exceljs: spec/integration/worksheet.spec.js — "iterates over rows"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "iterates over rows"
  it('iterates over rows (empty rows skipped)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getCell('A1').value = 1;
    ws.getCell('B2').value = 2;
    ws.getCell('D4').value = 4;
    ws.getCell('F6').value = 6;
    const seen: number[] = [];
    ws.eachRow((_row, n) => seen.push(n));
    expect(seen).toEqual([1, 2, 4, 6]);
    // 🚫 no { includeEmpty: true }: loop over 1..rowCount with getRow
    const all: number[] = [];
    for (let r = 1; r <= ws.rowCount; r++) all.push(ws.getRow(r).number);
    expect(all).toEqual([1, 2, 3, 4, 5, 6]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "iterates over collumn cells"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "iterates over collumn cells"
  it.skip('column.eachCell — not supported (iterate rows and read the column)', () => {});

  // exceljs: spec/integration/worksheet.spec.js — "returns undefined when row range is less than 1"
  it.skip('getRows(start, length) — not supported', () => {});

  describe('worksheet names', () => {
    // exceljs: spec/integration/worksheet.spec.js — "save the original name"
    it('keeps names up to 31 characters', () => {
      const wb = new Workbook();
      let ws = wb.addWorksheet();
      ws.name = 'ThisIsAWorksheetName';
      expect(ws.name).toBe('ThisIsAWorksheetName');
      ws = wb.addWorksheet();
      ws.name = 'ThisIsAWorksheetNameWith31Chars';
      expect(ws.name).toBe('ThisIsAWorksheetNameWith31Chars');
    });

    // exceljs: spec/integration/worksheet.spec.js — "when empty should thrown an error"
    it('throws for an empty name', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = '';
      }).toThrow(/1-31 characters/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "when isn't string should thrown an error"
    it('throws for a non-string name', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = 0 as unknown as string;
      }).toThrow();
    });

    // exceljs: spec/integration/worksheet.spec.js — "thrown an error"
    it('rejects the reserved name History', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = 'History';
      }).toThrow(/reserved/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "keep first 31 characters"
    it('names longer than 31 characters throw (exceljs truncates)', () => {
      const ws = new Workbook().addWorksheet();
      // ➖ silently truncating could create duplicates; turbosheet throws
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLongerThan31';
      }).toThrow(/1-31 characters/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (illegal characters)
    it('throws for illegal characters', () => {
      const wb = new Workbook();
      for (const ch of ['*', '?', ':', '/', '\\', '[', ']']) {
        const ws = wb.addWorksheet();
        expect(() => {
          ws.name = ch;
        }).toThrow(/invalid character/);
      }
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (leading / trailing quote)
    it('throws for a leading or trailing apostrophe', () => {
      const wb = new Workbook();
      for (const name of ["'sheetName", "sheetName'"]) {
        const ws = wb.addWorksheet();
        expect(() => {
          ws.name = name;
        }).toThrow(/apostrophe/);
      }
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (duplicate, case-insensitive)
    it('throws for a duplicate name (case-insensitive)', () => {
      const wb = new Workbook();
      wb.addWorksheet().name = 'thisisaworksheetnameinuppercase';
      const ws = wb.addWorksheet();
      expect(() => {
        ws.name = 'THISISAWORKSHEETNAMEINUPPERCASE';
      }).toThrow(/already exists/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (duplicate after truncation)
    it('throws for a duplicate 31-character name', () => {
      const wb = new Workbook();
      wb.addWorksheet().name = 'ThisIsAWorksheetNameThatIsLonge';
      const ws = wb.addWorksheet();
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLonge';
      }).toThrow(/already exists/);
      // ➖ the >31 variant throws for its length rather than as a duplicate
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLongerThan31';
      }).toThrow();
    });
  });

  // exceljs: spec/integration/worksheet.spec.js — "returns sheet values"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "returns sheet values"
  it('sheet values (sheetToArrays instead of getSheetValues)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 11;
    ws.getCell('C1').value = 'C1';
    ws.getCell('A2').value = 21;
    ws.getCell('B2').value = 'B2';
    ws.getCell('A4').value = 'end';
    // ➖ dense, 0-based rows from turbosheet/json; exceljs returns sparse 1-based arrays
    expect(sheetToArrays(ws)).toEqual([
      [11, null, 'C1'],
      [21, 'B2', null],
      [null, null, null],
      ['end', null, null],
    ]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "puts the lotion in the basket"
  it('puts the lotion in the basket', () => {
    const ws = new Workbook().addWorksheet('basket');
    ws.getCell('A1').value = 'lotion';
    expect(ws.getValue('A1')).toBe('lotion');
  });

  describe('Hidden', () => {
    for (const file of [
      'google-sheets',
      'libre-calc-as-excel-2007-365',
      'libre-calc-as-office-open-xml-spreadsheet',
    ]) {
      // exceljs: spec/integration/worksheet.spec.js — "Should set hidden attribute correctly (${file})"
      it(`reads hidden rows and columns (${file})`, () => {
        const ws = readXlsx(fixture(`hidden-test/${file}.xlsx`)).getWorksheet(1)!;
        expect(ws.getRow(1).hidden).toBe(false);
        expect(ws.getRow(2).hidden).toBe(true);
        expect(ws.getRow(3).hidden).toBe(false);
        expect(ws.getColumn(1).hidden).toBe(false);
        expect(ws.getColumn(2).hidden).toBe(true);
        expect(ws.getColumn(3).hidden).toBe(false);
      });
    }
  });

  // exceljs: spec/unit/doc/worksheet.values.spec.js — "calculates rowCount and actualRowCount"
  it('rowCount (actualRowCount is not supported)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 'A1';
    ws.getCell('C1').value = 'C1';
    ws.getCell('A3').value = 'A3';
    ws.getCell('D3').value = 'D3';
    ws.getCell('A4').value = null;
    ws.getCell('B5').value = 'B5';
    expect(ws.rowCount).toBe(5);
    let actual = 0;
    ws.eachRow(() => actual++);
    expect(actual).toBe(3);
  });

  // exceljs: spec/unit/doc/worksheet.values.spec.js — "calculates columnCount and actualColumnCount"
  it('columnCount (actualColumnCount is not supported)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 'A1';
    ws.getCell('C1').value = 'C1';
    ws.getCell('A3').value = 'A3';
    ws.getCell('D3').value = 'D3';
    ws.getCell('E4').value = null;
    ws.getCell('F5').value = 'F5';
    expect(ws.columnCount).toBe(6);
    expect(ws.dimensions?.ref).toBe('A1:F5');
  });

  // exceljs: spec/unit/doc/worksheet.views.spec.js — "adjusts collapsed property of columns"
  // exceljs: spec/unit/doc/worksheet.views.spec.js — "adjusts collapsed property of row"
  it.skip('row / column `collapsed` and sheet outlineLevelCol/Row properties — not modelled', () => {});

  // exceljs: spec/unit/doc/worksheet.views.spec.js — "sets outline levels via column headers"
  it('sets outline levels via column definitions (and round-trips them)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('sheet1');
    ws.columns = [
      { key: 'id', width: 10, outlineLevel: 1 },
      { key: 'name', width: 32, outlineLevel: 2 },
      { key: 'dob', width: 10, outlineLevel: 3 },
    ];
    ws.setCell('A1', 1);
    expect(ws.getColumn(1).outlineLevel).toBe(1);
    expect(ws.getColumn(2).outlineLevel).toBe(2);
    expect(ws.getColumn(3).outlineLevel).toBe(3);
    const ws2 = roundtrip(wb).getWorksheet('sheet1')!;
    expect([1, 2, 3].map((c) => ws2.getColumn(c).outlineLevel)).toEqual([1, 2, 3]);
  });

  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds multiple row breaks"
  // exceljs: spec/unit/doc/worksheet.pagebreak.spec.js — "adds multiple row breaks"
  it.skip('row page breaks (row.addPageBreak / ws.rowBreaks) — not modelled (kept raw on round trip)', () => {});

  describe('Duplicate Rows', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows with styles properly"
    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows replacing properly"
    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows shifting properly"
    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows with height properly"
    // exceljs: spec/unit/doc/workbook.spec.js — "inserts duplicates"
    // exceljs: spec/unit/doc/workbook.spec.js — "overwrites with duplicates"
    it.skip('duplicateRow — not supported', () => {});
  });

  // exceljs: spec/integration/workbook/workbook.spec.js — "spliced meat and ham"
  // exceljs: spec/integration/issues/issue-2125-spliceRows-last-row.spec.js (see issues)
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "Splice" (Rows ×10, Columns ×15)
  // exceljs: spec/unit/doc/row.spec.js — "Splice" (×7)
  it.skip('spliceRows / spliceColumns / row.splice — not supported', () => {});
});
