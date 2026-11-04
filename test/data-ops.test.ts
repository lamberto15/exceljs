/// <reference types="node" />
// Long dropdown lists, sortRange, fillTemplate, ExcelJS migration helpers.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  actualRowCount,
  addDataValidation,
  fillTemplate,
  getRows,
  getSheetValues,
  lastRow,
  readXlsx,
  sortRange,
  Workbook,
  writeXlsx,
} from '../src/index.js';

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('long dropdown lists', () => {
  it('short lists stay literal; long or comma lists go to the hidden _lists sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const short = addDataValidation(ws, { sqref: 'A1', type: 'list', values: ['Yes', 'No'] });
    expect(short.values).toEqual(['Yes', 'No']);
    const many = Array.from({ length: 80 }, (_, i) => `Option ${i + 1}`);
    const long = addDataValidation(ws, { sqref: 'B1:B10', type: 'list', values: many });
    const comma = addDataValidation(ws, {
      sqref: 'C1',
      type: 'list',
      values: ['Doe, John', 'Roe, Jane'],
    });
    expect(long).toMatchObject({ formula1: "'_lists'!$A$1:$A$80" });
    expect(long.values).toBeUndefined();
    expect(comma.formula1).toBe("'_lists'!$B$1:$B$2");
    const ls = wb.getWorksheet('_lists')!;
    expect(ls.state).toBe('hidden');
    expect(ls.getValue('A80')).toBe('Option 80');
    expect(ls.getValue('B1')).toBe('Doe, John');
    const buf = writeXlsx(wb);
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.dataValidations![1]!.formula1).toBe("'_lists'!$A$1:$A$80");
    await oracle(buf);
  });
});

describe('sortRange', () => {
  function sheet() {
    const ws = new Workbook().addWorksheet('S');
    ws.addRow(['Name', 'Score', 'Double']);
    const rows: [string, number | string | null][] = [
      ['cara', 3],
      ['Ann', 1],
      ['bob', 'n/a'],
      ['dan', null],
      ['eve', 2],
    ];
    rows.forEach(([n, v], i) => {
      ws.setCell(i + 2, 1, n);
      if (v !== null) ws.setCell(i + 2, 2, v);
      ws.setCell(i + 2, 3, { formula: `B${i + 2}*2` }, { font: { italic: true } });
    });
    ws.setCell('E1', { formula: 'B2' }); // outside the range: keeps pointing at B2
    return ws;
  }

  it('ascending: numbers < text, blanks last; formulas move with their row; header kept', () => {
    const ws = sheet();
    sortRange(ws, 'A1:C6', { col: 'B' }, { header: true });
    expect([2, 3, 4, 5, 6].map((r) => ws.getValue(r, 1))).toEqual([
      'Ann',
      'eve',
      'cara',
      'bob',
      'dan',
    ]);
    expect(ws.getValue('A1')).toBe('Name');
    expect(ws.getValue('C2')).toEqual({ formula: 'B2*2' }); // was C3 = B3*2 (Ann)
    expect(ws.getStyle('C2').font?.italic).toBe(true);
    expect(ws.getValue('E1')).toEqual({ formula: 'B2' });
  });

  it('descending (blanks still last), text keys case-insensitive, multiple keys', () => {
    const ws = sheet();
    sortRange(ws, 'A2:C6', [{ col: 2, desc: true }]);
    expect([2, 3, 4, 5, 6].map((r) => ws.getValue(r, 1))).toEqual([
      'bob',
      'cara',
      'eve',
      'Ann',
      'dan',
    ]);
    sortRange(ws, 'A2:C6', { col: 'A' });
    expect([2, 3, 4, 5, 6].map((r) => ws.getValue(r, 1))).toEqual([
      'Ann',
      'bob',
      'cara',
      'dan',
      'eve',
    ]);
    const t = new Workbook().addWorksheet('T');
    t.addRows([
      ['x', 2],
      ['y', 1],
      ['x', 1],
    ]);
    sortRange(t, 'A1:B3', [{ col: 'A' }, { col: 'B' }]);
    expect(
      getSheetValues(t)
        .slice(1)
        .map((r) => r.slice(1)),
    ).toEqual([
      ['x', 1],
      ['x', 2],
      ['y', 1],
    ]);
  });

  it('refuses merges inside the range and key columns outside it', () => {
    const ws = sheet();
    ws.mergeCells('A3:B3');
    expect(() => sortRange(ws, 'A2:C6', { col: 'B' })).toThrow(/merged/);
    expect(() => sortRange(sheet(), 'A2:C6', { col: 'E' })).toThrow(/outside/);
  });
});

describe('fillTemplate', () => {
  it('replaces placeholders (typed when alone), repeats rows, removes rows for empty arrays', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Invoice');
    ws.setCell('A1', 'Invoice for {{customer.name}}', { font: { bold: true } });
    ws.setCell('A2', '{{date}}');
    ws.setCell('A4', '{{items[].name}}');
    ws.setCell('B4', '{{items[].qty}}');
    ws.setCell('C4', { formula: 'B4*10' });
    ws.setCell('A5', '{{notes[].text}}');
    ws.setCell('A6', 'Total');
    ws.setCell('B6', { formula: 'SUM(B4:B4)' });
    ws.setCell('A7', '{{missing.path}}!');
    fillTemplate(wb, {
      customer: { name: 'ACME' },
      date: new Date(Date.UTC(2026, 10, 5)),
      items: [
        { name: 'bolt', qty: 3 },
        { name: 'nut', qty: 5 },
        { name: 'gear', qty: 1 },
      ],
      notes: [],
    });
    expect(ws.getValue('A1')).toBe('Invoice for ACME');
    expect(ws.getStyle('A1').font?.bold).toBe(true);
    expect(ws.getValue('A2')).toEqual(new Date(Date.UTC(2026, 10, 5)));
    expect([4, 5, 6].map((r) => [ws.getValue(r, 1), ws.getValue(r, 2)])).toEqual([
      ['bolt', 3],
      ['nut', 5],
      ['gear', 1],
    ]);
    expect(ws.getValue('C6')).toEqual({ formula: 'B6*10' });
    expect(ws.getValue('A7')).toBe('Total'); // notes row removed, total moved up after +2 rows
    expect(ws.getValue('B7')).toEqual({ formula: 'SUM(B4:B6)' });
    expect(ws.getValue('A8')).toBe('!');
  });
});

describe('ExcelJS migration helpers', () => {
  it('actualRowCount, lastRow, getRows, getSheetValues', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('A1', 1);
    ws.setStyle('A2', { font: { bold: true } }); // style only: not counted
    ws.setCell('C4', 'x');
    ws.setRowProps(6, { height: 20 });
    expect(actualRowCount(ws)).toBe(2);
    expect(lastRow(ws)!.number).toBe(4);
    expect(lastRow(new Workbook().addWorksheet('E'))).toBeUndefined();
    expect(getRows(ws, 3, 2).map((r) => r.number)).toEqual([3, 4]);
    const v = getSheetValues(ws);
    expect(v[1]).toEqual([undefined, 1]);
    expect(v[4]![3]).toBe('x');
    expect(v[3]).toBeUndefined();
  });
});
