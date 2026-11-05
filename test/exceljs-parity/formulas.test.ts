// Formulas: shared formulas (spec/integration/workbook/workbook.spec.js, workbook-xlsx-writer
// "shared formula", spec/unit/doc/worksheet.shared-formula.spec.js, spec/unit/utils/
// shared-formula.spec.js) and formula cells of formulas.xlsx (workbook-xlsx-reader.spec.js).

import { describe, expect, it } from 'vitest';
import { CellKind, fillFormula, readXlsx, translateFormula, Workbook } from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';

describe('exceljs parity: formulas', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "shared formula"
  it('fillFormula over a block is written as a shared master + followers', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Hello');
    fillFormula(ws, 'A1:B2', 'ROW()+COLUMN()', [2, 3, 3, 4]);
    const bytes = (await import('../../src/index.js')).writeXlsx(wb);
    const ws2 = readXlsx(bytes).getWorksheet('Hello')!;
    expect(ws2.getValue('A1')).toEqual({ formula: 'ROW()+COLUMN()', result: 2 });
    expect(ws2.getValue('B2')).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
    const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
    expect(x.getCell('A1').value).toEqual({
      formula: 'ROW()+COLUMN()',
      result: 2,
      shareType: 'shared',
      ref: 'A1:B2',
    });
    expect(x.getCell('B1').value).toEqual({ sharedFormula: 'A1', result: 3 });
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "shared formula"
  it('streaming writer: { sharedFormula } values are written as Excel shared formulas', async () => {
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
    // ➖ read back expanded to plain formulas (exceljs: { sharedFormula: 'A1', result } / shareType)
    expect(ws.getValue('A1')).toEqual({ formula: 'ROW()+COLUMN()', result: 2 });
    expect(ws.getValue('B1')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
    expect(ws.getValue('A2')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
    expect(ws.getValue('B2')).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
    const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
    expect(x.getCell('B2').value).toEqual({ sharedFormula: 'A1', result: 4 });
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

  describe('fillFormula (➖ each cell reads back as its own formula, not { sharedFormula } slaves)', () => {
    const f = (ws: ReturnType<Workbook['addWorksheet']>, ref: string) => ws.getValue(ref);

    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula using 2D array values"
    it('fills a 2D range; results row-major (➖ a flat array, not nested)', () => {
      const ws = new Workbook().addWorksheet('s');
      fillFormula(ws, 'A1:B2', 'ROW()+COLUMN()', [2, 3, 3, 4]);
      expect(f(ws, 'A1')).toEqual({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(f(ws, 'B1')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
      expect(f(ws, 'A2')).toEqual({ formula: 'ROW()+COLUMN()', result: 3 });
      expect(f(ws, 'B2')).toEqual({ formula: 'ROW()+COLUMN()', result: 4 });
    });

    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Translates formulae to slave cells"
    it('translates relative references per cell', () => {
      const ws = new Workbook().addWorksheet('s');
      ws.setCell('A1', 1);
      ws.setCell('B1', 2);
      fillFormula(ws, 'A2:B3', 'A1', [1, 2, 1, 2]);
      expect(f(ws, 'A2')).toEqual({ formula: 'A1', result: 1 });
      expect(f(ws, 'B2')).toEqual({ formula: 'B1', result: 2 });
      expect(f(ws, 'A3')).toEqual({ formula: 'A2', result: 1 });
      expect(f(ws, 'B3')).toEqual({ formula: 'B2', result: 2 });
    });

    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula down using 1D array values"
    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula across using 1D array values"
    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula down and across using 1D array values"
    it('fills down, across, and down + across from a 1D array', () => {
      const ws = new Workbook().addWorksheet('s');
      fillFormula(ws, 'A1:A4', 'ROW()+COLUMN()', [2, 3, 4, 5]);
      expect(['A1', 'A2', 'A3', 'A4'].map((r) => f(ws, r))).toEqual(
        [2, 3, 4, 5].map((result) => ({ formula: 'ROW()+COLUMN()', result })),
      );
      fillFormula(ws, 'C1:F1', 'ROW()+COLUMN()', [4, 5, 6, 7]);
      expect(['C1', 'D1', 'E1', 'F1'].map((r) => f(ws, r))).toEqual(
        [4, 5, 6, 7].map((result) => ({ formula: 'ROW()+COLUMN()', result })),
      );
      fillFormula(ws, 'H1:I2', 'ROW()+COLUMN()', [9, 10, 10, 11]);
      expect(['H1', 'I1', 'H2', 'I2'].map((r) => f(ws, r))).toEqual(
        [9, 10, 10, 11].map((result) => ({ formula: 'ROW()+COLUMN()', result })),
      );
    });

    // exceljs: spec/unit/doc/worksheet.shared-formula.spec.js — "Fills formula using function"
    it('takes results from a function of (row, col)', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('s');
      fillFormula(ws, 'A1:B2', 'ROW()+COLUMN()', (r, c) => r + c);
      expect(['A1', 'B1', 'A2', 'B2'].map((r) => f(ws, r))).toEqual(
        [2, 3, 3, 4].map((result) => ({ formula: 'ROW()+COLUMN()', result })),
      );
      // exceljs reads the shared followers with their results
      const x = (
        await exceljsLoad((await import('../../src/index.js')).writeXlsx(wb))
      ).getWorksheet('s')!;
      expect(x.getCell('B2').value).toEqual({ sharedFormula: 'A1', result: 4 });
    });
  });

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
