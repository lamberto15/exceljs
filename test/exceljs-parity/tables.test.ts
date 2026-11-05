// Tables: spec/unit/doc/worksheet-table.spec.js (the observable parts). Issues 988 / 1669 (reading
// tables) are in issues.test.ts.

import { describe, expect, it } from 'vitest';
import {
  addTable,
  addTableColumn,
  addTableRows,
  type CellValue,
  removeTableColumns,
  removeTableRows,
  renameTableColumn,
  setTableHeaderRow,
  setTableTotalsRow,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../../src/index.js';
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

  // expected cells, ExcelJS's `values` grid (header, 4 data rows, totals) from A1
  const grid = (): CellValue[][] => [
    ['Date', 'Id', 'Word'],
    ...rows.map((r) => [...r] as CellValue[]),
    ['Totals', { formula: 'SUBTOTAL(104,TestTable[Id])' }, null],
  ];
  /** exceljs checkTable: the grid from A1, and nothing in the row / column after it. */
  function checkGrid(ws: Worksheet, g: CellValue[][]): void {
    for (let i = 0; i <= g.length; i++) {
      for (let j = 0; j <= g[0]!.length; j++) {
        const want = g[i]?.[j] ?? null;
        expect(ws.getValue(i + 1, j + 1), `${i},${j}`).toEqual(want);
      }
    }
  }

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes header"
  it('removes header (setTableHeaderRow)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    setTableHeaderRow(ws, 'TestTable', false);
    checkGrid(ws, grid().slice(1));
    expect(ws.tables![0]).toMatchObject({ ref: 'A1:C5', headerRow: false });
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes totals"
  it('removes totals (setTableTotalsRow)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    setTableTotalsRow(ws, 'TestTable', false);
    checkGrid(ws, grid().slice(0, 5));
    expect(ws.tables![0]!.ref).toBe('A1:C5');
    expect(ws.tables![0]!.totalsRow).toBeUndefined();
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes a row"
  it('removes a row (removeTableRows)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    removeTableRows(ws, 'TestTable', 1);
    const g = grid();
    g.splice(2, 1);
    checkGrid(ws, g);
    expect(ws.tables![0]!.ref).toBe('A1:C5');
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "removes a column"
  it('removes a column (removeTableColumns)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    removeTableColumns(ws, 'TestTable', 1);
    checkGrid(
      ws,
      grid().map((r) => r.filter((_, j) => j !== 1)),
    );
    expect(ws.tables![0]!.ref).toBe('A1:B6');
    expect(ws.tables![0]!.columns.map((c) => c.name)).toEqual(['Date', 'Word']);
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "adds a column"
  it('adds a column (addTableColumn)', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    add(ws, 'A1');
    addTableColumn(
      ws,
      'TestTable',
      { name: 'Letter', totalsRowFunction: 'custom', totalsRowFormula: 'ROW()' },
      ['a', 'b', 'c', 'd'],
      2,
    );
    const col: CellValue[] = ['Letter', 'a', 'b', 'c', 'd', { formula: 'ROW()' }];
    checkGrid(
      ws,
      grid().map((r, i) => [...r.slice(0, 2), col[i]!, ...r.slice(2)]),
    );
    expect(ws.tables![0]!.ref).toBe('A1:D6');
    const x = (await exceljsLoad(writeXlsx(wb))).getWorksheet('blort')!;
    expect(x.getCell('C1').value).toBe('Letter');
  });

  // exceljs: spec/unit/doc/worksheet-table.spec.js — "renames a column"
  it('renames a column (renameTableColumn)', () => {
    const ws = new Workbook().addWorksheet('blort');
    add(ws, 'A1');
    renameTableColumn(ws, 'TestTable', 1, 'Code');
    const g = grid();
    g[0] = ['Date', 'Code', 'Word'];
    g[5] = ['Totals', { formula: 'SUBTOTAL(104,TestTable[Code])' }, null];
    checkGrid(ws, g);
    expect(roundtrip(ws.workbook).worksheets[0]!.tables![0]!.columns[1]!.name).toBe('Code');
  });
});
