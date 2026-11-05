// Tables: spec/unit/doc/worksheet-table.spec.js (the observable parts). Issues 988 / 1669 (reading
// tables) are in issues.test.ts.

import { describe, expect, it } from 'vitest';
import { addTable, addTableRows, Workbook, type Worksheet } from '../../src/index.js';
import { exceljsLoad, roundtrip } from './helpers.js';

const rows = [
  [new Date('2019-08-01'), 1, 'Bird'],
  [new Date('2019-08-02'), 2, 'is'],
  [new Date('2019-08-03'), 3, 'the'],
  [new Date('2019-08-04'), 4, 'Word'],
];

function add(ws: Worksheet, topLeft: string) {
  // ➖ exceljs `ref` is the top-left cell; turbosheet `topLeft` (or `ref` = the full range)
  return addTable(ws, {
    name: 'TestTable',
    topLeft,
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleDark3', showRowStripes: true },
    columns: [
      { name: 'Date', totalsRowLabel: 'Totals' },
      { name: 'Id', totalsRowFunction: 'max' },
      // ➖ per-column `filterButton` and cell `style` are not modelled; dataStyle is a dxf
      { name: 'Word', dataStyle: { font: { bold: true, name: 'Comic Sans MS' } } },
    ],
    rows,
  });
}

function check(ws: Worksheet, r0: number, c0: number): void {
  const at = (r: number, c: number) => ws.getValue(r0 + r, c0 + c);
  expect([at(0, 0), at(0, 1), at(0, 2)]).toEqual(['Date', 'Id', 'Word']);
  rows.forEach((row, i) => {
    expect(at(i + 1, 0)).toEqual(row[0]);
    expect(at(i + 1, 1)).toBe(row[1]);
    expect(at(i + 1, 2)).toBe(row[2]);
  });
  expect(at(5, 0)).toBe('Totals');
  // ➖ no totalsRowResult: the SUBTOTAL formula is written without a cached result
  expect((at(5, 1) as { formula: string }).formula).toBe('SUBTOTAL(104,TestTable[Id])');
  expect(at(5, 2)).toBeNull();
}

describe('exceljs parity: tables', () => {
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "creates a table"
  it('creates a table (header, data and totals cells)', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const t = add(ws, 'A1');
    expect(t.ref).toBe('A1:C6');
    check(ws, 1, 1);
    const ws2 = roundtrip(wb).getWorksheet('blort')!;
    check(ws2, 1, 1);
    expect(ws2.tables![0]).toMatchObject({
      name: 'TestTable',
      ref: 'A1:C6',
      totalsRow: true,
      style: { theme: 'TableStyleDark3', showRowStripes: true },
    });
    const { writeXlsx } = await import('../../src/index.js');
    const x = (await exceljsLoad(writeXlsx(wb))).getWorksheet('blort')!;
    expect(x.getTable('TestTable')).toBeDefined();
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "moves the table"
  it('a table placed elsewhere (C2)', () => {
    const ws = new Workbook().addWorksheet('blort');
    expect(add(ws, 'C2').ref).toBe('C2:E7');
    check(ws, 2, 3);
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "adds a row"
  it('adds a row above the totals row (addTableRows; ➖ no table.addRow + commit)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    addTableRows(ws, 'TestTable', [[new Date('2019-08-05'), 5, 'Bird']]);
    expect(ws.tables![0]!.ref).toBe('A1:C7');
    expect([1, 2, 3].map((c) => ws.getValue(6, c))).toEqual([new Date('2019-08-05'), 5, 'Bird']);
    expect(ws.getValue(7, 1)).toBe('Totals');
    expect((ws.getValue(7, 2) as { formula: string }).formula).toBe('SUBTOTAL(104,TestTable[Id])');
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes header"
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes totals"
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes a row"
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes a column"
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "adds a column"
  // exceljs: spec/unit/doc/worksheet-table.spec.js — "renames a column"
  it.skip('table.commit() editing (rows, columns, header/totals toggles, move) — not supported: tables are plain models; edit cells and ws.tables yourself', () => {});
});
