/// <reference types="node" />
// fillFormula, calcProperties (fullCalcOnLoad), OFFSET / INDIRECT / LOOKUP.

import { describe, expect, it } from 'vitest';
import { evaluate, evaluateFormula, supportedFunctions } from '../src/formula/index.js';
import { fillFormula, readXlsx, Workbook, writeXlsx } from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const workbookXml = (buf: Uint8Array): string => dec.decode(unzipSync(buf).get('xl/workbook.xml')!);

describe('fillFormula', () => {
  it('fills relative references per cell, keeps styles, takes results', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 4; r++) ws.addRow([r, r * 10]);
    ws.setStyle('C2', { font: { bold: true } });
    fillFormula(ws, 'C1:C4', '=A1+B1*$A$1', [11, 22, 33, 44]);
    expect(ws.getValue('C1')).toEqual({ formula: 'A1+B1*$A$1', result: 11 });
    expect(ws.getValue('C4')).toEqual({ formula: 'A4+B4*$A$1', result: 44 });
    expect(ws.getStyle('C2').font?.bold).toBe(true);
    fillFormula(ws, 'D1:E2', 'A1', (r, c) => r * 100 + c);
    expect(ws.getValue('E2')).toEqual({ formula: 'B2', result: 205 });
    evaluate(wb);
    expect((ws.getValue('C3') as { result: number }).result).toBe(33);
  });
});

describe('calcProperties', () => {
  it('fullCalcOnLoad: on with formulas by default; forced on/off; read back', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', { formula: '1+1', result: 2 });
    expect(workbookXml(writeXlsx(wb))).toContain('fullCalcOnLoad="1"');
    wb.calcProperties = { fullCalcOnLoad: false };
    expect(workbookXml(writeXlsx(wb))).not.toContain('fullCalcOnLoad');
    const plain = new Workbook();
    plain.addWorksheet('S').setCell('A1', 1);
    expect(workbookXml(writeXlsx(plain))).not.toContain('fullCalcOnLoad');
    plain.calcProperties = { fullCalcOnLoad: true };
    const buf = writeXlsx(plain);
    expect(workbookXml(buf)).toContain('fullCalcOnLoad="1"');
    expect(readXlsx(buf).calcProperties).toEqual({ fullCalcOnLoad: true });
  });
});

describe('OFFSET / INDIRECT / LOOKUP', () => {
  function book(): Workbook {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRows([
      [1, 'one', 10],
      [2, 'two', 20],
      [3, 'three', 30],
      [5, 'five', 50],
    ]);
    const o = wb.addWorksheet('My Sheet');
    o.setCell('B2', 42);
    wb.definedNames = [{ name: 'Nums', value: 'Data!$A$1:$A$4' }];
    return wb;
  }

  it('are listed as supported', () => {
    expect(supportedFunctions()).toEqual(expect.arrayContaining(['OFFSET', 'INDIRECT', 'LOOKUP']));
  });

  it('OFFSET moves and resizes references', () => {
    const wb = book();
    const v = (f: string) => evaluateFormula(wb, f);
    expect(v('OFFSET(A1,2,1)')).toBe('three');
    expect(v('SUM(OFFSET(A1,1,0,3,1))')).toBe(10);
    expect(v('SUM(OFFSET(C4,0,0,-2,1))')).toBe(80);
    expect(v('ROWS(OFFSET(A1,0,0,4,3))')).toBe(4);
    expect(v('OFFSET(A1,-1,0)')).toEqual({ error: '#REF!' });
  });

  it('INDIRECT resolves text references, other sheets and names', () => {
    const wb = book();
    const v = (f: string) => evaluateFormula(wb, f);
    expect(v('INDIRECT("C2")')).toBe(20);
    expect(v('INDIRECT("B"&3)')).toBe('three');
    expect(v(`INDIRECT("'My Sheet'!B2")`)).toBe(42);
    expect(v('SUM(INDIRECT("Nums"))')).toBe(11);
    expect(v('SUM(INDIRECT("A1:A2"))')).toBe(3);
    expect(v('INDIRECT("not a ref")')).toEqual({ error: '#REF!' });
    expect(v('INDIRECT("R1C1",FALSE)')).toEqual({ error: '#REF!' });
  });

  it('INDIRECT sees formula cells evaluated in the same run', () => {
    const wb = book();
    const ws = wb.getWorksheet('Data')!;
    ws.setCell('E1', { formula: 'INDIRECT("F1")*2' });
    ws.setCell('F1', { formula: 'C1+1' });
    evaluate(wb);
    expect((ws.getValue('E1') as { result: number }).result).toBe(22);
  });

  it('LOOKUP: vector form (approximate, sorted) and array form', () => {
    const wb = book();
    const v = (f: string) => evaluateFormula(wb, f);
    expect(v('LOOKUP(3,A1:A4,B1:B4)')).toBe('three');
    expect(v('LOOKUP(4,A1:A4,B1:B4)')).toBe('three');
    expect(v('LOOKUP(99,A1:A4,C1:C4)')).toBe(50);
    expect(v('LOOKUP(2,A1:A4)')).toBe(2);
    expect(v('LOOKUP(0,A1:A4,B1:B4)')).toEqual({ error: '#N/A' });
    expect(v('LOOKUP(2,A1:C4)')).toBe(20); // array form: first column, last column
    expect(v('LOOKUP(3,{1,2,3;"a","b","c"})')).toBe('c'); // wider: first row, last row
  });
});
