// Dynamic arrays in the evaluator (#104): spilling, #SPILL!, A1# and the dynamic-array functions.

import { describe, expect, it } from 'vitest';
import { evaluate, evaluateFormula, parseFormula } from '../src/formula/index.js';
import { readXlsx, Workbook, type Worksheet, writeXlsx } from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const res = (ws: Worksheet, a: string) => {
  const v = ws.getValue(a);
  return v && typeof v === 'object' && 'formula' in v ? v.result : v;
};

function data(): { wb: Workbook; ws: Worksheet } {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const rows: [string, number][] = [
    ['b', 3],
    ['a', 1],
    ['c', 2],
    ['a', 5],
  ];
  rows.forEach(([k, n], i) => {
    ws.setCell(i + 1, 1, k);
    ws.setCell(i + 1, 2, n);
  });
  return { wb, ws };
}

describe('dynamic-array functions', () => {
  const wb = data().wb;
  const ev = (f: string) => evaluateFormula(wb, f, { sheet: 'S' });
  it('SEQUENCE / TRANSPOSE / RANDARRAY', () => {
    expect(ev('SEQUENCE(2,3,10,5)')).toEqual([
      [10, 15, 20],
      [25, 30, 35],
    ]);
    expect(ev('SEQUENCE(3)')).toEqual([[1], [2], [3]]);
    expect(ev('SEQUENCE(0)')).toEqual({ error: '#CALC!' });
    expect(ev('TRANSPOSE(A1:B2)')).toEqual([
      ['b', 'a'],
      [3, 1],
    ]);
    const r = ev('RANDARRAY(2,2,1,6,TRUE)') as number[][];
    expect(r.flat().every((x) => Number.isInteger(x) && x >= 1 && x <= 6)).toBe(true);
  });
  it('FILTER', () => {
    expect(ev('FILTER(A1:B4,B1:B4>2)')).toEqual([
      ['b', 3],
      ['a', 5],
    ]);
    expect(ev('FILTER(A1:A4,A1:A4="a")')).toEqual([['a'], ['a']]);
    expect(ev('FILTER(A1:A4,B1:B4>9)')).toEqual({ error: '#CALC!' });
    expect(ev('FILTER(A1:A4,B1:B4>9,"none")')).toBe('none');
    expect(ev('FILTER(A1:B4,{TRUE,FALSE})')).toEqual([['b'], ['a'], ['c'], ['a']]);
    expect(ev('FILTER(A1:B4,B1:B2>0)')).toEqual({ error: '#VALUE!' });
  });
  it('SORT / SORTBY', () => {
    expect(ev('SORT(B1:B4)')).toEqual([[1], [2], [3], [5]]);
    expect(ev('SORT(A1:B4,2,-1)')).toEqual([
      ['a', 5],
      ['b', 3],
      ['c', 2],
      ['a', 1],
    ]);
    expect(ev('SORT(A1:B4,{1,2},{1,-1})')).toEqual([
      ['a', 5],
      ['a', 1],
      ['b', 3],
      ['c', 2],
    ]);
    expect(ev('SORTBY(A1:A4,B1:B4)')).toEqual([['a'], ['c'], ['b'], ['a']]);
    expect(ev('SORT({3,1,2},1,1,TRUE)')).toEqual([[1, 2, 3]]);
    expect(ev('SORT({"b";1;TRUE;"A"})')).toEqual([[1], ['A'], ['b'], [true]]);
    expect(ev('SORT(B1:B4,3)')).toEqual({ error: '#VALUE!' });
  });
  it('UNIQUE', () => {
    expect(ev('UNIQUE(A1:A4)')).toEqual([['b'], ['a'], ['c']]);
    expect(ev('UNIQUE(A1:A4,FALSE,TRUE)')).toEqual([['b'], ['c']]);
    expect(ev('UNIQUE({"x","X","y"},TRUE)')).toEqual([['x', 'y']]);
    expect(ev('ROWS(UNIQUE(A1:A4))')).toBe(3);
  });
});

describe('spilling', () => {
  it('a dynamic formula spills into its neighbours; the file says so', () => {
    const { wb, ws } = data();
    ws.setCell('D1', { formula: 'SORT(B1:B4)' });
    const r = evaluate(wb);
    expect(r.errors).toEqual([]);
    expect(res(ws, 'D1')).toBe(1);
    expect([2, 3, 4].map((n) => ws.getValue(`D${n}`))).toEqual([2, 3, 5]);
    expect(ws.getValue('D1')).toMatchObject({ array: 'D1:D4', dynamic: true });
    const sheet = new TextDecoder().decode(
      unzipSync(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!,
    );
    expect(sheet).toContain(
      '<c r="D1" cm="1"><f t="array" ref="D1:D4">_xlfn._xlws.SORT(B1:B4)</f><v>1</v></c>',
    );
    expect(sheet).toContain('<c r="D4"><v>5</v></c>');
  });

  it('formulas reading spilled cells and A1# see the spilled values', () => {
    const { wb, ws } = data();
    ws.setCell('D1', { formula: 'SEQUENCE(B1)' }); // 3 rows
    ws.setCell('F1', { formula: 'SUM(D1#)' });
    ws.setCell('F2', { formula: 'D3*10' });
    ws.setCell('F3', { formula: 'ROWS(D1#)' });
    ws.setCell('F4', { formula: 'A1#' }); // not a dynamic formula
    evaluate(wb);
    expect(['F1', 'F2', 'F3', 'F4'].map((a) => res(ws, a))).toEqual([6, 30, 3, { error: '#REF!' }]);
    // the source shrinks: stale spilled values are cleared, dependents follow
    ws.setCell('B1', 2);
    evaluate(wb);
    expect(ws.getValue('D3')).toBeNull();
    expect(['F1', 'F2', 'F3'].map((a) => res(ws, a))).toEqual([3, 0, 2]);
    expect(ws.getValue('D1')).toMatchObject({ array: 'D1:D2' });
  });

  it('a blocked area gives #SPILL!; A1# of it gives #REF!; unblocking spills again', () => {
    const { wb, ws } = data();
    ws.setCell('D1', { formula: 'SEQUENCE(3)' });
    ws.setCell('D3', 'in the way');
    ws.setCell('F1', { formula: 'D1#' });
    evaluate(wb);
    expect(res(ws, 'D1')).toEqual({ error: '#SPILL!' });
    expect(res(ws, 'F1')).toEqual({ error: '#REF!' });
    // files store #SPILL! as #VALUE! (Excel keeps the real code in rich-value metadata)
    expect(
      new TextDecoder().decode(unzipSync(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!),
    ).toContain('t="e" cm="1"><f t="array" ref="D1">_xlfn.SEQUENCE(3)</f><v>#VALUE!</v>');
    expect(ws.getValue('D2')).toBeNull();
    expect(ws.getValue('D1')).toMatchObject({ array: 'D1' });
    ws.clearCell('D3');
    evaluate(wb);
    expect([1, 2, 3].map((n) => res(ws, `D${n}`))).toEqual([1, 2, 3]);
    expect(res(ws, 'F1')).toBe(1);
  });

  it('two overlapping spills: the later formula gets #SPILL!', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('B1', { formula: 'SEQUENCE(3)' });
    ws.setCell('A2', { formula: 'SEQUENCE(1,3)' });
    evaluate(wb);
    expect(res(ws, 'B1')).toBe(1);
    expect(res(ws, 'A2')).toEqual({ error: '#SPILL!' });
  });

  it('read back: the spilled values and the area survive; re-evaluating is stable', () => {
    const { wb, ws } = data();
    ws.setCell('D1', { formula: 'UNIQUE(A1:A4)' });
    evaluate(wb);
    const back = readXlsx(writeXlsx(wb));
    const s2 = back.worksheets[0]!;
    expect(s2.getValue('D1')).toMatchObject({
      formula: 'UNIQUE(A1:A4)',
      array: 'D1:D3',
      dynamic: true,
    });
    expect([2, 3].map((n) => s2.getValue(`D${n}`))).toEqual(['a', 'c']);
    s2.setCell('A4', 'z');
    evaluate(back);
    expect([1, 2, 3, 4].map((n) => res(s2, `D${n}`))).toEqual(['b', 'a', 'c', 'z']);
  });

  it('parses A1#, Sheet!A1# and the file form ANCHORARRAY(A1)', () => {
    expect(parseFormula('_xlfn.ANCHORARRAY(Sheet1!C3)')).toMatchObject({
      type: 'spill',
      sheet: 'Sheet1',
      row: 2,
      col: 2,
    });
    expect(parseFormula('SUM(A1#)')).toMatchObject({ args: [{ type: 'spill', row: 0, col: 0 }] });
    expect(parseFormula("'My S'!B2#")).toMatchObject({
      type: 'spill',
      sheet: 'My S',
      row: 1,
      col: 1,
    });
  });

  it('writeResults: false leaves the workbook alone', () => {
    const { wb, ws } = data();
    ws.setCell('D1', { formula: 'SEQUENCE(2)' });
    const r = evaluate(wb, { writeResults: false });
    expect(r.results!.get('S!D1')).toBe(1);
    expect(ws.getValue('D2')).toBeNull();
    expect(ws.getValue('D1')).toEqual({ formula: 'SEQUENCE(2)' });
  });
});
