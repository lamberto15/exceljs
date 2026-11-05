// Table editing (#103): addTableColumn / removeTableColumns / renameTableColumn /
// removeTableRows / setTableHeaderRow / setTableTotalsRow.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  addTable,
  addTableColumn,
  readXlsx,
  removeTableColumns,
  removeTableRows,
  renameTableColumn,
  setTableHeaderRow,
  setTableTotalsRow,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../src/index.js';
import { renameStructuredRefs } from '../src/xlsx/features/table-edit.js';

function sales(ws: Worksheet, topLeft = 'B2') {
  return addTable(ws, {
    name: 'Sales',
    topLeft,
    totalsRow: true,
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Units', totalsRowFunction: 'sum' },
      { name: 'Price' },
      { name: 'Amount', calculatedColumnFormula: 'Sales[@Units]*Sales[@Price]' },
    ],
    rows: [
      ['N', 1, 10],
      ['S', 2, 20],
      ['E', 3, 30],
    ],
  });
}

const formula = (ws: Worksheet, a: string) => (ws.getValue(a) as { formula: string }).formula;

describe('renameStructuredRefs', () => {
  const ren = (f: string, inTable = false) =>
    renameStructuredRefs(f, ['Sales'], 'Units', 'Qty', inTable);
  it.each([
    ['SUM(Sales[Units])', 'SUM(Sales[Qty])'],
    ['sales[units]*2', 'sales[Qty]*2'],
    ['Sales[@Units]', 'Sales[@Qty]'],
    ['Sales[@[Units]]', 'Sales[@[Qty]]'],
    ['Sales[[#This Row],[Units]]', 'Sales[[#This Row],[Qty]]'],
    ['SUM(Sales[[#Totals],[Units]:[Price]])', 'SUM(Sales[[#Totals],[Qty]:[Price]])'],
    ['Sales[#All]', 'Sales[#All]'],
    ['Other[Units]+Sales[Price]', 'Other[Units]+Sales[Price]'],
    ['"Sales[Units]"&Sales[Units]', '"Sales[Units]"&Sales[Qty]'],
    ['[@Units]', '[@Units]'],
  ])('%s → %s', (f, want) => {
    expect(ren(f)).toBe(want);
  });
  it('unqualified references only inside the table; escaped names', () => {
    expect(ren('[@Units]*[@Price]', true)).toBe('[@Qty]*[@Price]');
    expect(renameStructuredRefs("T[a'#b]+T[@[a'#b]]", ['T'], 'a#b', 'x]y', false)).toBe(
      "T[x']y]+T[@[x']y]]",
    );
  });
});

describe('table editing', () => {
  it('renameTableColumn updates header, totals, calculated columns, other sheets and names', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    ws.setCell('B4', { formula: 'Sales[@Units]+1' }); // not in a table column formula
    const other = wb.addWorksheet('Report');
    other.setCell('A1', { formula: 'SUM(Sales[Units])', result: 6 });
    addDefinedName(wb, 'TotalUnits', 'SUM(Sales[Units])');
    renameTableColumn(ws, 'Sales', 'Units', 'Qty');
    const t = ws.tables![0]!;
    expect(t.columns[1]!.name).toBe('Qty');
    expect(ws.getValue('C2')).toBe('Qty');
    expect(formula(ws, 'C6')).toBe('SUBTOTAL(109,Sales[Qty])');
    expect(t.columns[3]!.calculatedColumnFormula).toBe('Sales[@Qty]*Sales[@Price]');
    expect(formula(ws, 'E3')).toBe('Sales[@Qty]*Sales[@Price]');
    expect(formula(ws, 'B4')).toBe('Sales[@Qty]+1');
    expect(other.getValue('A1')).toEqual({ formula: 'SUM(Sales[Qty])', result: 6 });
    expect(wb.definedNames.find((d) => d.name === 'TotalUnits')!.value).toBe('SUM(Sales[Qty])');
    expect(() => renameTableColumn(ws, 'Sales', 0, 'price')).toThrow('duplicate column name');
    const back = readXlsx(writeXlsx(wb)).getWorksheet('Data')!.tables![0]!;
    expect(back.columns.map((c) => c.name)).toEqual(['Region', 'Qty', 'Price', 'Amount']);
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(writeXlsx(wb) as unknown as ArrayBuffer);
    expect(x.getWorksheet('Data')!.getCell('C2').value).toBe('Qty');
  });

  it('addTableColumn: fills a calculated column, grows only into empty cells', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    addTableColumn(ws, 'Sales', { name: 'Tax', calculatedColumnFormula: 'Sales[@Amount]*0.1' }, [
      undefined as never,
      5,
    ]);
    expect(ws.tables![0]!.ref).toBe('B2:F6');
    expect(ws.getValue('F2')).toBe('Tax');
    expect(formula(ws, 'F3')).toBe('Sales[@Amount]*0.1');
    expect(ws.getValue('F4')).toBe(5);
    expect(ws.getValue('F6')).toBeNull();
    ws.setCell('G4', 'blocker');
    expect(() => addTableColumn(ws, 'Sales', 'More')).toThrow('G4, which is not empty');
    expect(ws.tables![0]!.columns).toHaveLength(5);
    expect(() => addTableColumn(ws, 'Sales', 'region')).toThrow('duplicate column name');
  });

  it('moved cells keep their styles; vacated cells are cleared', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    ws.setStyle('D3', { numFmt: '0.00' });
    removeTableColumns(ws, 'Sales', 1); // Units goes, Price (D) moves to C
    expect(ws.getValue('C3')).toBe(10);
    expect(ws.getStyle('C3').numFmt).toBe('0.00');
    expect(ws.getValue('E2')).toBeNull();
    expect(ws.getStyle('E3').numFmt).toBeUndefined();
    expect(ws.tables![0]!.ref).toBe('B2:D6');
    expect(() => removeTableColumns(ws, 'Sales', 0, 3)).toThrow('needs a column');
    expect(() => removeTableColumns(ws, 'Sales', 2, 2)).toThrow(RangeError);
  });

  it('removeTableColumns turns references to the removed column into #REF!', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    ws.setCell('H1', { formula: 'SUM(Sales[Price])+SUM(Sales[Units])' });
    ws.setCell('H2', { formula: 'SUM(Sales[[#Totals],[Units]:[Price]])' });
    removeTableColumns(ws, 'Sales', 2); // Price
    expect(formula(ws, 'H1')).toBe('SUM(#REF!)+SUM(Sales[Units])');
    expect(formula(ws, 'H2')).toBe('SUM(#REF!)');
    expect(ws.tables![0]!.columns[2]!.calculatedColumnFormula).toBe('Sales[@Units]*#REF!');
    expect(formula(ws, 'D3')).toBe('Sales[@Units]*#REF!'); // the moved Amount cells
  });

  it('removeTableRows moves rows and the totals row up; cells outside the table stay', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    ws.setCell('A4', 'outside');
    ws.setCell('B8', 'below');
    removeTableRows(ws, 'Sales', 0, 2);
    expect(ws.getValue('B3')).toBe('E');
    expect(ws.getValue('B4')).toBe('Total');
    expect(formula(ws, 'C4')).toBe('SUBTOTAL(109,Sales[Units])');
    expect(ws.getValue('B5')).toBeNull();
    expect(ws.getValue('A4')).toBe('outside');
    expect(ws.getValue('B8')).toBe('below');
    expect(ws.tables![0]!.ref).toBe('B2:E4');
    expect(() => removeTableRows(ws, 'Sales', 0)).toThrow('needs a data row');
    expect(() => removeTableRows(ws, 'Sales', 1)).toThrow(RangeError);
  });

  it('header and totals rows toggle back and forth', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws);
    setTableHeaderRow(ws, 'Sales', false);
    expect(ws.getValue('B2')).toBe('N');
    expect(ws.tables![0]).toMatchObject({ ref: 'B2:E5', headerRow: false });
    setTableTotalsRow(ws, 'Sales', false);
    expect(ws.getValue('B5')).toBeNull();
    expect(ws.tables![0]!.ref).toBe('B2:E4');
    setTableHeaderRow(ws, 'Sales', true);
    setTableTotalsRow(ws, 'Sales', true);
    expect(ws.tables![0]!.ref).toBe('B2:E6');
    expect([2, 3, 4, 5, 6].map((r) => ws.getValue(`B${r}`))).toEqual([
      'Region',
      'N',
      'S',
      'E',
      'Total',
    ]);
    expect(formula(ws, 'C6')).toBe('SUBTOTAL(109,Sales[Units])');
    expect(ws.tables![0]!.headerRow).toBeUndefined();
    ws.setCell('C7', 'x');
    setTableTotalsRow(ws, 'Sales', false);
    expect(() => setTableTotalsRow(ws, 'Sales', true)).not.toThrow(); // row 6 is free again
    setTableTotalsRow(ws, 'Sales', false);
    setTableHeaderRow(ws, 'Sales', false);
    ws.setCell('C5', 'x');
    expect(() => setTableHeaderRow(ws, 'Sales', true)).toThrow('C5, which is not empty');
    ws.clearCell('C5');
    const back = readXlsx(writeXlsx(wb)).worksheets[0]!.tables![0]!;
    expect(back).toMatchObject({ ref: 'B2:E4', headerRow: false });
  });

  it('growing into another table or merged cells throws', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    sales(ws, 'A1');
    addTable(ws, { name: 'Next', topLeft: 'E1', columns: ['x'], rows: [[1]] });
    expect(() => addTableColumn(ws, 'Sales', 'More')).toThrow('E1, which is not empty');
    setTableTotalsRow(ws, 'Sales', false); // A1:D4
    ws.mergeCells('A5:B5');
    expect(() => setTableTotalsRow(ws, 'Sales', true)).toThrow('merged cells A5:B5');
  });
});
