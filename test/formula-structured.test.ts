// Structured references and LET in the evaluator (#105).

import { describe, expect, it } from 'vitest';
import {
  evaluate,
  evaluateCell,
  evaluateFormula,
  FormulaParseError,
  parseFormula,
} from '../src/formula/index.js';
import { addTable, Workbook, type Worksheet } from '../src/index.js';

function book(): { wb: Workbook; ws: Worksheet } {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  addTable(ws, {
    name: 'Sales',
    topLeft: 'B2',
    totalsRow: true,
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Qty', totalsRowFunction: 'sum' },
      { name: 'Unit Price' },
      { name: 'Amount', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*[@[Unit Price]]' },
    ],
    rows: [
      ['N', 1, 10],
      ['S', 2, 20],
      ['E', 3, 30],
    ],
  });
  return { wb, ws };
}

const ev = (wb: Workbook, f: string, address?: string) =>
  evaluateFormula(wb, f, { sheet: 'Data', ...(address ? { address } : {}) });

describe('structured references', () => {
  it('parses the forms Excel uses', () => {
    expect(parseFormula('Sales[Qty]')).toEqual({
      type: 'structured',
      table: 'Sales',
      areas: [],
      from: 'Qty',
      to: 'Qty',
    });
    expect(parseFormula('[@Qty]')).toMatchObject({
      table: undefined,
      areas: ['thisRow'],
      from: 'Qty',
    });
    expect(parseFormula('Sales[[#Totals],[Qty]:[Amount]]')).toMatchObject({
      areas: ['totals'],
      from: 'Qty',
      to: 'Amount',
    });
    expect(parseFormula('Sales[[#Headers],[#Data],[Region]]')).toMatchObject({
      areas: ['headers', 'data'],
    });
    expect(parseFormula('Sales[#All]')).toMatchObject({ areas: ['all'], from: undefined });
    expect(parseFormula('Sales[]')).toMatchObject({ areas: [], from: undefined });
    expect(parseFormula("T[a'#b]")).toMatchObject({ from: 'a#b' });
    expect(() => parseFormula('Sales[#Nope]')).toThrow(FormulaParseError);
    expect(() => parseFormula('Sales[Qty')).toThrow(FormulaParseError);
  });

  it('evaluates data, totals, headers, all and column ranges', () => {
    const { wb } = book();
    evaluate(wb);
    expect(ev(wb, 'SUM(Sales[Qty])')).toBe(6);
    expect(ev(wb, 'SUM(Sales[Amount])')).toBe(140);
    expect(ev(wb, 'sales[[#Totals],[Qty]]')).toBe(6);
    expect(ev(wb, 'Sales[[#Headers],[Unit Price]]')).toBe('Unit Price');
    expect(ev(wb, 'ROWS(Sales[#All])')).toBe(5);
    expect(ev(wb, 'ROWS(Sales[[#Headers],[#Data],[Qty]])')).toBe(4);
    expect(ev(wb, 'ROWS(Sales[[#Data],[#Totals],[Qty]])')).toBe(4);
    expect(ev(wb, 'COLUMNS(Sales[[Qty]:[Amount]])')).toBe(3);
    expect(ev(wb, 'COLUMNS(Sales[])')).toBe(4);
    expect(ev(wb, 'SUM(Sales[Nope])')).toEqual({ error: '#REF!' });
    expect(ev(wb, 'SUM(Nope[Qty])')).toEqual({ error: '#REF!' });
    expect(ev(wb, 'INDIRECT("Sales[Region]")', 'H4')).toBe('S'); // implicit intersection
  });

  it('calculated columns and [@Col] use the host row; rows outside the data → #VALUE!', () => {
    const { wb, ws } = book();
    const r = evaluate(wb);
    expect(r.errors).toEqual([]);
    expect([3, 4, 5].map((n) => (ws.getValue(`E${n}`) as { result: number }).result)).toEqual([
      10, 40, 90,
    ]);
    expect((ws.getValue('C6') as { result: number }).result).toBe(6); // SUBTOTAL(109,Sales[Qty])
    expect(ev(wb, 'Sales[@Qty]', 'H4')).toBe(2); // beside the table: that row, as in Excel
    expect(ev(wb, 'Sales[@Qty]', 'H10')).toEqual({ error: '#VALUE!' }); // no data row there
    expect(ev(wb, '[@Qty]', 'E4')).toBe(2);
  });

  it('dependencies follow structured references (incl. another sheet)', () => {
    const { wb, ws } = book();
    const rep = wb.addWorksheet('Report');
    rep.setCell('A1', { formula: 'SUM(Sales[Amount])*2' });
    ws.setCell('C3', 5); // Qty of N: Amount 50
    evaluate(wb);
    expect((rep.getValue('A1') as { result: number }).result).toBe(360);
    expect(evaluateCell(wb, 'Report', 'A1')).toBe(360);
  });

  it('a whole-column reference to its own column is circular', () => {
    const { wb, ws } = book();
    ws.setCell('E4', { formula: 'SUM(Sales[Amount])' });
    expect(evaluate(wb).circular).toContain('Data!E4');
  });
});

describe('LET', () => {
  it('binds names left to right, accepts _xlpm. and ranges', () => {
    const { wb } = book();
    evaluate(wb);
    expect(ev(wb, 'LET(x,2,y,x*3,x+y)')).toBe(8);
    expect(ev(wb, 'LET(_xlpm.x,2,_xlpm.x*10)')).toBe(20);
    expect(ev(wb, 'LET(q,Sales[Qty],SUM(q)/ROWS(q))')).toBe(2);
    expect(ev(wb, 'LET(x,1,LET(x,2,x)+x)')).toBe(3);
    expect(ev(wb, 'LET(x,1/0,IFERROR(x,7))')).toBe(7);
    expect(ev(wb, 'LET(x,1,y)')).toEqual({ error: '#NAME?' });
    expect(ev(wb, 'LET(x,1)')).toEqual({ error: '#VALUE!' });
    expect(ev(wb, 'LET(A1,1,A1)')).toEqual({ error: '#VALUE!' });
  });
});
