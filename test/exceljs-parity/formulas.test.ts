// Formulas: shared formulas (spec/integration/workbook/workbook.spec.js, workbook-xlsx-writer
// "shared formula", spec/unit/doc/worksheet.shared-formula.spec.js, spec/unit/utils/
// shared-formula.spec.js) and formula cells of formulas.xlsx (workbook-xlsx-reader.spec.js).

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx, translateFormula, Workbook } from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';

describe('exceljs parity: formulas', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "shared formula"
  it('a formula filled over a block (plain formulas; no fillFormula)', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Hello');
    // 🚫 no ws.fillFormula / shared-formula authoring: write each cell's formula
    const results = [
      [2, 3],
      [3, 4],
    ];
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) {
        ws.setCell(r + 1, c + 1, { formula: 'ROW()+COLUMN()', result: results[r]![c]! });
      }
    }
    const bytes = (await import('../../src/index.js')).writeXlsx(wb);
    const ws2 = readXlsx(bytes).getWorksheet('Hello')!;
    expect(ws2.getValue('A1')).toEqual({ formula: 'ROW()+COLUMN()', result: 2 });
    expect(ws2.getValue('B2')).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
    const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
    expect(x.getCell('B1').value).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "shared formula"
  it('streaming writer: { sharedFormula } values are written as translated plain formulas', async () => {
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('Hello');
      s.addRow([
        { formula: 'ROW()+COLUMN()', result: 2 },
        { sharedFormula: 'A1', result: 3 },
      ]);
      s.addRow([
        { sharedFormula: 'A1', result: 3 },
        { sharedFormula: 'A1', result: 4 },
      ]);
    });
    const ws = readXlsx(bytes).getWorksheet('Hello')!;
    // ➖ read back as plain formulas (exceljs: { sharedFormula: 'A1', result } / shareType)
    expect(ws.getValue('A1')).toEqual({ formula: 'ROW()+COLUMN()', result: 2 });
    expect(ws.getValue('B1')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
    expect(ws.getValue('A2')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
    expect(ws.getValue('B2')).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
    const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
    expect(x.getCell('B2').value).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
  });

  // exceljs: spec/unit/doc/cell.spec.js — "shares formulas"
  it('{ sharedFormula } takes the master formula with relative references moved', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.getCell('A1').value = 1;
    ws.getCell('B1').value = { formula: 'A1+1', result: 2 };
    ws.getCell('C1').value = { sharedFormula: 'B1', result: 3 };
    expect(ws.getCell('B1').kind).toBe(CellKind.Formula);
    // ➖ stored as a plain formula (no master / shared formula types)
    expect(ws.getCell('C1').kind).toBe(CellKind.Formula);
    expect(ws.getCell('C1').value).toEqual({ formula: 'B1+1', result: 3 });
  });

  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula using 2D array values"
  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Translates formulae to slave cells"
  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula down using 1D array values"
  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula across using 1D array values"
  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula down and across using 1D array values"
  // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula using function"
  it.skip('ws.fillFormula producing shared formulas — not supported (fillFormula writes plain formulas; shared formulas are expanded on read)', () => {});

  it('shared formulas in a file are expanded with translated references', () => {
    const ws = readXlsx(fixture('fibonacci.xlsx')).getWorksheet('fib')!;
    for (let r = 5; r <= 19; r++) {
      expect((ws.getValue(`A${r}`) as { formula: string }).formula).toBe(`A${r - 1}+1`);
    }
  });

  describe('slideFormula (translateFormula)', () => {
    // [formula, from, to, result] from spec/unit/utils/shared-formula.spec.js
    const cases: [string, number, number, string][] = [
      ['A1+1', 1, 0, 'A2+1'], // A2 → A3
      ['A1+1', 0, 1, 'B1+1'], // A2 → B2
      ['SUM(A1:A10)', 0, 1, 'SUM(B1:B10)'],
      ['$A$1+A1', 1, 0, '$A$1+A2'],
      ['$A$1+A1', 0, 1, '$A$1+B1'],
      ['$A1+A1', 1, 0, '$A2+A2'],
      ['$A1+A1', 0, 1, '$A1+B1'],
      ['A$1+A1', 1, 0, 'A$1+A2'],
      ['A$1+A1', 0, 1, 'B$1+B1'],
    ];
    for (const [f, dr, dc, out] of cases) {
      // exceljs: spec/unit/utils/shared-formula.spec.js — "${args[0]} from ${args[1]} to ${args[2]}"
      it(`${f} moved by (${dr}, ${dc}) → ${out}`, () => {
        expect(translateFormula(f, dr, dc)).toBe(out);
      });
    }
  });

  describe('with a spreadsheet that contains formulas', () => {
    const ws = () => readXlsx(fixture('formulas.xlsx')).worksheets[0]!;

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should be classified as a formula cell"
    it('a regular formula cell is a formula cell', () => {
      expect(ws().getKind('A2')).toBe(CellKind.Formula);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should have text corresponding to the evaluated formula result"
    it('has the cached result', () => {
      expect((ws().getValue('A2') as { result: string }).result).toBe('someone@example.com');
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should have the formula source"
    // ExcelJS exposes the raw stored text `_xlfn.CONCAT(...)`; turbosheet strips the future-function
    // prefix on read (and re-adds it on write), so the model shows what Excel's UI shows.
    it('has the formula source (future-function _xlfn. prefix stripped on read)', () => {
      expect((ws().getValue('A2') as { formula: string }).formula).toBe(
        'CONCAT("someone","@example.com")',
      );
    });

    it('formula values round-trip', () => {
      const wb = readXlsx(fixture('formulas.xlsx'));
      expect(roundtrip(wb).worksheets[0]!.getValue('A2')).toEqual(wb.worksheets[0]!.getValue('A2'));
    });
  });
});
