// Interop with exceljs 4.4.0 as the oracle: files written by exceljs read correctly in turbosheet,
// and turbosheet's rewrite of them still loads in exceljs with the same content. Not a port of a
// single ExcelJS spec; it exercises the ExcelJS test-book features end to end.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { CellKind, getPrintArea, readXlsx, writeXlsx } from '../../src/index.js';
import { exceljsBytes, exceljsLoad, fixture } from './helpers.js';
import { styles, testValues } from './test-book.js';

async function exceljsBook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('sheet1', { properties: { tabColor: { argb: 'FF00FF00' } } });
  ws.getCell('A1').value = 7;
  ws.getCell('B1').value = testValues.str;
  ws.getCell('C1').value = testValues.date;
  ws.getCell('D1').value = { formula: 'A1', result: 7 } as ExcelJS.CellFormulaValue;
  ws.getCell('F1').value = testValues.hyperlink;
  ws.getCell('K1').value = { error: '#N/A' } as ExcelJS.CellErrorValue;
  ws.getCell('A2').value = 5;
  ws.mergeCells('A2:B3');
  ws.getCell('A4').value = 1.5;
  ws.getCell('A4').numFmt = testValues.numFmt1;
  ws.getCell('A4').border = styles.borders.thin;
  ws.getCell('B5').value = 'x';
  ws.getCell('B5').font = styles.fonts.comicSansUdB16;
  ws.getCell('C6').value = 'rot';
  ws.getCell('C6').alignment = { textRotation: -45, vertical: 'middle' };
  ws.getCell('A8').value = 'grad';
  ws.getCell('A8').fill = styles.fills.blueWhiteHGrad;
  ws.getCell('A9').value = 1;
  ws.fillFormula('B9:E9', 'A9+1', [2, 3, 4, 5]);
  // a value too: exceljs only attaches notes to cells present in sheetData
  ws.getCell('A10').value = 'noted';
  ws.getCell('A10').note = 'a note';
  ws.getCell('A11').dataValidation = { type: 'list', formulae: ['"a,b"'], allowBlank: true };
  ws.pageSetup.printArea = 'A1:E11';
  ws.getRow(6).height = 42;
  return exceljsBytes(wb);
}

describe('exceljs parity: interop with exceljs-written files', () => {
  it('turbosheet reads an exceljs workbook', async () => {
    const ws = readXlsx(await exceljsBook()).getWorksheet('sheet1')!;
    expect(ws.getValue('A1')).toBe(7);
    expect(ws.getValue('B1')).toBe(testValues.str);
    expect(ws.getValue('C1')).toEqual(testValues.date);
    expect(ws.getValue('D1')).toEqual({ formula: 'A1', result: 7 });
    expect(ws.getValue('F1')).toEqual(testValues.hyperlink);
    expect(ws.getValue('K1')).toEqual({ error: '#N/A' });
    expect(ws.merges).toEqual(['A2:B3']);
    expect(ws.getStyle('A4')).toMatchObject({
      numFmt: testValues.numFmt1,
      border: styles.borders.thin,
    });
    expect(ws.getStyle('B5').font).toMatchObject({
      name: 'Comic Sans MS',
      bold: true,
      underline: 'double',
    });
    expect(ws.getStyle('C6').alignment).toEqual({ textRotation: 135, vertical: 'center' });
    // shared formulas written by exceljs are expanded
    expect(ws.getValue('C9')).toEqual({ formula: 'B9+1', result: 3 });
    expect(ws.getValue('E9')).toEqual({ formula: 'D9+1', result: 5 });
    expect(ws.comments?.[0]).toMatchObject({ ref: 'A10', text: 'a note' });
    expect(ws.dataValidations?.[0]).toMatchObject({
      sqref: 'A11',
      type: 'list',
      formula1: '"a,b"',
    });
    expect(getPrintArea(ws)).toEqual(['A1:E11']);
    expect(ws.getRow(6).height).toBe(42);
    expect(ws.tabColor).toEqual({ argb: 'FF00FF00' });
    expect(ws.getKind('A8')).toBe(CellKind.String);
  });

  it('turbosheet rewrite of an exceljs workbook loads in exceljs with the same content', async () => {
    const x = (await exceljsLoad(writeXlsx(readXlsx(await exceljsBook())))).getWorksheet('sheet1')!;
    expect(x.getCell('A1').value).toBe(7);
    expect(x.getCell('C1').value).toEqual(testValues.date);
    expect(x.getCell('F1').value).toEqual(testValues.hyperlink);
    expect(x.getCell('A4').numFmt).toBe(testValues.numFmt1);
    expect(x.getCell('C6').alignment).toMatchObject({ textRotation: -45, vertical: 'middle' });
    expect(x.getCell('A10').note).toBe('a note');
    expect(x.getCell('A11').dataValidation).toMatchObject({ type: 'list', formulae: ['"a,b"'] });
    expect((x.getCell('E9').value as { formula: string }).formula).toBe('D9+1');
    // 🚫 gradient fills are not modelled: the rewrite drops A8's gradient
    expect(x.getCell('A8').fill).toBeUndefined();
  });

  it('gold.xlsx (written by Excel) survives turbosheet → exceljs', async () => {
    const x = (await exceljsLoad(writeXlsx(readXlsx(fixture('gold.xlsx'))))).getWorksheet(
      'Values',
    )!;
    expect(x.getCell('B1').value).toBe('I am Text');
    expect(x.getCell('B6').value).toEqual({
      hyperlink: 'https://www.npmjs.com/package/exceljs',
      text: 'exceljs',
    });
  });
});
