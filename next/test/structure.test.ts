/// <reference types="node" />
// Row / column structure edits (core/structure.ts): cells, references and sheet data follow
// inserted / deleted rows and columns the way Excel moves them.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addChart,
  addConditionalFormatting,
  addDataValidation,
  addNote,
  addTable,
  deleteColumns,
  deleteRows,
  duplicateRow,
  insertColumns,
  insertRow,
  insertRows,
  readXlsx,
  spliceColumns,
  spliceRows,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../src/index.js';

const f = (ws: Worksheet, a: string): string | undefined =>
  (ws.getValue(a) as { formula?: string } | null)?.formula;

function sheet(): { wb: Workbook; ws: Worksheet } {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  for (let r = 1; r <= 6; r++) ws.addRow([r, r * 10, r * 100]);
  return { wb, ws };
}

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('insertRows / deleteRows: cells', () => {
  it('moves cells and row properties down', () => {
    const { ws } = sheet();
    ws.setRowProps(3, { height: 30, hidden: true });
    insertRows(ws, 3, 2);
    expect(ws.getValue('A2')).toBe(2);
    expect(ws.getValue('A3')).toBeNull();
    expect(ws.getValue('A4')).toBeNull();
    expect(ws.getValue('A5')).toBe(3);
    expect(ws.getRowProps(5)).toEqual({ height: 30, hidden: true });
    expect(ws.getRowProps(3)).toEqual({});
    expect(ws.rowData[4]!.index).toBe(4);
    expect(ws.rowCount).toBe(8);
  });

  it('inserts values (arrays and keyed objects)', () => {
    const { ws } = sheet();
    ws.columns = [{ key: 'a' }, { key: 'b' }];
    insertRows(ws, 2, [['x', 'y'], { b: 'kb' }]);
    expect(ws.getRow(2).values).toEqual([undefined, 'x', 'y']);
    expect(ws.getRow(3).values).toEqual([undefined, undefined, 'kb']);
    expect(ws.getValue('A4')).toBe(2);
    insertRow(ws, 1, ['top']);
    expect(ws.getValue('A1')).toBe('top');
    expect(ws.getValue('A2')).toBe(1);
  });

  it('inheritStyle copies the format of the row above / below', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 1, { font: { bold: true } });
    ws.setStyle('B1', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } },
    });
    ws.setRowProps(1, { height: 25 });
    ws.setCell('A2', 2, { font: { italic: true } });
    insertRows(ws, 2, [['new']], { inheritStyle: 'above' });
    expect(ws.getValue('A2')).toBe('new');
    expect(ws.getStyle('A2').font?.bold).toBe(true);
    expect(ws.getStyle('B2').fill).toBeDefined();
    expect(ws.getRowProps(2).height).toBe(25);
    insertRows(ws, 1, 1, { inheritStyle: 'below' });
    expect(ws.getStyle('A1').font?.bold).toBe(true);
    expect(ws.getValue('A1')).toBeNull();
  });

  it('deletes rows and moves later rows up', () => {
    const { ws } = sheet();
    deleteRows(ws, 2, 2);
    expect(ws.getValue('A1')).toBe(1);
    expect(ws.getValue('A2')).toBe(4);
    expect(ws.getValue('C4')).toBe(600);
    expect(ws.rowCount).toBe(4);
    expect(ws.rowData[3]!.index).toBe(3);
    deleteRows(ws, 10); // past the end: no-op
    expect(ws.rowCount).toBe(4);
  });

  it('refuses to push cells off the sheet, without changing anything', () => {
    const { ws } = sheet();
    ws.setCell(1_048_576, 1, 'last');
    expect(() => insertRows(ws, 2, 1)).toThrow(/off the sheet/);
    expect(ws.getValue('A2')).toBe(2);
    insertRows(ws, 1_048_576, 0); // nothing to insert
    ws.clearCell(1_048_576, 1);
    ws.setCell(1, 16_384, 'xfd');
    expect(() => insertColumns(ws, 'B', 1)).toThrow(/off the sheet/);
    expect(() => deleteRows(ws, 0)).toThrow(RangeError);
    expect(() => insertRows(ws, 1_048_576, 2)).toThrow(RangeError);
  });
});

describe('formula references', () => {
  it('follow inserted rows on the same sheet (relative and absolute)', () => {
    const { ws } = sheet();
    ws.setCell('E1', { formula: 'A1+A5+$A$5+A$5+$A5' });
    ws.setCell('E2', { formula: 'SUM(A1:A6)*SUM($B$2:$B$3)+SUM(A4:B6)' });
    ws.setCell('E3', { formula: 'SUM(4:5)+SUM(A:B)+SUM(1:2)' });
    insertRows(ws, 4, 2);
    expect(f(ws, 'E1')).toBe('A1+A7+$A$7+A$7+$A7');
    expect(f(ws, 'E2')).toBe('SUM(A1:A8)*SUM($B$2:$B$3)+SUM(A6:B8)');
    expect(f(ws, 'E3')).toBe('SUM(6:7)+SUM(A:B)+SUM(1:2)');
  });

  it('become #REF! when their cells are deleted; ranges shrink', () => {
    const { ws } = sheet();
    ws.setCell('E1', { formula: 'A2+A3+A4+A6' });
    ws.setCell('E6', { formula: 'SUM(A2:A3)+SUM(A1:A6)+SUM(A3:A5)+SUM(2:3)' });
    deleteRows(ws, 2, 2);
    expect(f(ws, 'E1')).toBe('#REF!+#REF!+A2+A4');
    expect(f(ws, 'E4')).toBe('SUM(#REF!)+SUM(A1:A4)+SUM(A2:A3)+SUM(#REF!)');
  });

  it('a range spanning the insertion point grows; one ending above it does not', () => {
    const { ws } = sheet();
    ws.setCell('E1', { formula: 'SUM(A1:A3)+SUM(A1:A4)+SUM(A4:A5)' });
    insertRows(ws, 4, 1);
    expect(f(ws, 'E1')).toBe('SUM(A1:A3)+SUM(A1:A5)+SUM(A5:A6)');
  });

  it('follow the edited sheet from other sheets; other sheets’ refs stay', () => {
    const { wb, ws } = sheet();
    const other = wb.addWorksheet('My Sheet');
    const o2 = wb.addWorksheet('Other');
    other.setCell('A1', { formula: "Data!A5+'Data'!B5+data!$C$6+A5+Other!A5+SUM(Data!A1:A6)" });
    o2.setCell('A1', { formula: "'My Sheet'!A5+A5" });
    ws.setCell('E1', { formula: "'My Sheet'!A5+Other!A5+A5" });
    insertRows(ws, 2, 1);
    expect(f(other, 'A1')).toBe("Data!A6+'Data'!B6+data!$C$7+A5+Other!A5+SUM(Data!A1:A7)");
    expect(f(o2, 'A1')).toBe("'My Sheet'!A5+A5");
    expect(f(ws, 'E1')).toBe("'My Sheet'!A5+Other!A5+A6");
    deleteRows(ws, 6);
    expect(f(other, 'A1')).toBe("Data!#REF!+'Data'!#REF!+data!$C$6+A5+Other!A5+SUM(Data!A1:A6)");
  });

  it('leaves strings, functions, names, structured, 3D and external refs alone', () => {
    const { wb, ws } = sheet();
    ws.setCell('E1', {
      formula:
        'LOG10(A5)&"A5"&T[[#This Row],[A5]]&SUM(Data:Other!A5)&[1]Data!A5&\'[2]Data\'!A5&Rate+A5',
    });
    wb.addWorksheet('Other');
    insertRows(ws, 1, 1);
    expect(f(ws, 'E2')).toBe(
      'LOG10(A6)&"A5"&T[[#This Row],[A5]]&SUM(Data:Other!A5)&[1]Data!A5&\'[2]Data\'!A5&Rate+A6',
    );
  });

  it('shift columns: cells, whole-column refs, widths, keys and column styles', () => {
    const { ws } = sheet();
    ws.columns = [
      { key: 'a', width: 5 },
      { key: 'b', width: 6 },
      { key: 'c', width: 7 },
    ];
    ws.getColumn('C').style = { font: { bold: true } };
    ws.setCell('E1', { formula: 'SUM(B:C)+B2+$C$3+SUM(A1:C1)+SUM(2:2)' });
    insertColumns(ws, 'B', 2);
    expect(ws.getRow(1).values.slice(0, 6)).toEqual([undefined, 1, undefined, undefined, 10, 100]);
    expect(f(ws, 'G1')).toBe('SUM(D:E)+D2+$E$3+SUM(A1:E1)+SUM(2:2)');
    expect(ws.getColumn(4).width).toBe(6);
    expect(ws.getColumn('b').number).toBe(4);
    expect(ws.getColumn(5).styleIndex).not.toBe(0);
    expect(ws.getColumn(3).styleIndex).toBe(0);
    expect(ws.getColumn(2).width).toBeUndefined();
    deleteColumns(ws, 'D');
    expect(f(ws, 'F1')).toBe('SUM(D:D)+#REF!+$D$3+SUM(A1:D1)+SUM(2:2)');
    expect(ws.getRow(1).values.slice(0, 5)).toEqual([undefined, 1, undefined, undefined, 100]);
    expect(ws.getColumn('c').number).toBe(4);
    expect(ws.getColumn(4).styleIndex).not.toBe(0);
  });

  it('adjusts array-formula ranges and defined names (print area / titles)', () => {
    const { wb, ws } = sheet();
    ws.setCell('E2', { formula: 'A2:A3*2', array: 'E2:E3' });
    wb.definedNames = [
      { name: 'Rate', value: 'Data!$B$4' },
      { name: 'Gone', value: 'Data!$B$2' },
      { name: '_xlnm.Print_Area', value: 'Data!$A$1:$C$6', localSheetId: 0 },
      { name: '_xlnm.Print_Titles', value: 'Data!$2:$2,Data!$A:$A', localSheetId: 0 },
    ];
    insertRows(ws, 1, 1);
    expect(ws.getValue('E3')).toEqual({ formula: 'A3:A4*2', array: 'E3:E4' });
    expect(wb.definedNames.map((d) => d.value)).toEqual([
      'Data!$B$5',
      'Data!$B$3',
      'Data!$A$2:$C$7', // a range starting at the insertion row moves (Excel)
      'Data!$3:$3,Data!$A:$A',
    ]);
    deleteRows(ws, 3);
    expect(wb.definedNames.map((d) => d.value)).toEqual([
      'Data!$B$4',
      'Data!#REF!',
      'Data!$A$2:$C$6',
      'Data!$A:$A',
    ]);
  });
});

describe('sheet data follows', () => {
  it('merges move, grow, shrink and disappear', () => {
    const { ws } = sheet();
    ws.mergeCells('A2:B3');
    ws.mergeCells('D1:D4');
    ws.mergeCells('E5:F5');
    insertRows(ws, 3, 1);
    expect(ws.merges).toEqual(['A2:B4', 'D1:D5', 'E6:F6']);
    deleteRows(ws, 2, 2);
    expect(ws.merges).toEqual(['A2:B2', 'D1:D3', 'E4:F4']);
    deleteRows(ws, 4);
    expect(ws.merges).toEqual(['A2:B2', 'D1:D3']);
    deleteRows(ws, 1, 2); // A2:B2 deleted, D1:D3 → one cell: unmerged
    expect(ws.merges).toEqual([]);
    ws.mergeCells('A1:A2'); // the merge index is still consistent
    expect(() => ws.mergeCells('A2:B2')).toThrow(/overlaps/);
  });

  it('conditional formatting and data validation ranges and formulas', () => {
    const { ws } = sheet();
    const cf = addConditionalFormatting(ws, 'A1:A6 C2:C3', [
      { type: 'expression', formulae: ['A1>$B$2'], style: { font: { bold: true } } },
    ]);
    addConditionalFormatting(ws, 'B3', [{ type: 'expression', formulae: ['B3>0'] }]);
    addDataValidation(ws, { sqref: 'B2:B6', type: 'whole', operator: 'lessThan', formula1: 'C2' });
    deleteRows(ws, 1); // the anchor A1 is deleted: re-anchored on the new top-left (old A2)
    expect(cf.ref).toBe('A1:A5 C1:C2');
    expect(cf.rules[0]).toMatchObject({ formulae: ['A1>$B$1'] });
    expect(ws.dataValidations![0]).toMatchObject({ sqref: 'B1:B5', formula1: 'C1' });
    deleteRows(ws, 2);
    expect(ws.conditionalFormatting!.length).toBe(1);
    insertRows(ws, 1, 2);
    expect(cf.ref).toBe('A3:A6 C3');
    expect(cf.rules[0]).toMatchObject({ formulae: ['A3>$B$3'] });
  });

  it('notes, autoFilter, pictures and charts', () => {
    const { wb, ws } = sheet();
    addNote(ws, 'B2', 'two');
    addNote(ws, 'B4', 'four');
    ws.autoFilter = 'A1:C6';
    ws.images = [
      { imageId: 0, anchor: { from: { col: 4, row: 2 }, to: { col: 6, row: 5 } } },
      { imageId: 0, anchor: { from: { col: 4, row: 2 }, ext: { width: 10, height: 10 } } },
      {
        imageId: 0,
        anchor: { from: { col: 4, row: 2 }, to: { col: 6, row: 5 } },
        editAs: 'oneCell',
      },
    ];
    const other = wb.addWorksheet('Charts');
    addChart(other, {
      type: 'column',
      series: [
        { name: { ref: 'Data!$A$1' }, categories: 'Data!$A$2:$A$6', values: 'Data!$B$2:$B$6' },
      ],
      anchor: 'A1:F10',
    });
    addChart(ws, { type: 'line', series: [{ values: '$C$2:$C$6' }], anchor: 'E2:H8' });
    insertRows(ws, 4, 2); // inside the pictures' rows 3..5
    expect(ws.comments!.map((n) => n.ref)).toEqual(['B2', 'B6']);
    expect(ws.autoFilter).toBe('A1:C8');
    expect(ws.images[0]!.anchor).toEqual({ from: { col: 4, row: 2 }, to: { col: 6, row: 7 } });
    expect(ws.images[1]!.anchor).toMatchObject({ from: { col: 4, row: 2 } });
    expect(ws.images[2]!.anchor).toEqual({ from: { col: 4, row: 2 }, to: { col: 6, row: 5 } });
    expect(other.charts![0]!.series[0]).toMatchObject({
      name: { ref: 'Data!$A$1' },
      categories: 'Data!$A$2:$A$8',
      values: 'Data!$B$2:$B$8',
    });
    expect(ws.charts![0]!.series[0]!.values).toBe('$C$2:$C$8');
    expect(ws.charts![0]!.anchor).toEqual({ from: { col: 4, row: 1 }, to: { col: 8, row: 10 } });
    deleteRows(ws, 2);
    expect(ws.comments!.map((n) => n.ref)).toEqual(['B5']);
    deleteRows(ws, 1, 1); // the chart series name cell
    expect(other.charts![0]!.series[0]!.name).toBeUndefined();
  });

  it('tables grow (with calculated columns), shrink, and gain or lose columns', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('T');
    const t = addTable(ws, {
      name: 'Sales',
      ref: 'B2:D5',
      columns: [
        'Qty',
        'Price',
        { name: 'Total', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*2' },
      ],
      rows: [
        [1, 2],
        [3, 4],
        [5, 6],
      ],
    });
    insertRows(ws, 4, 1);
    expect(t.ref).toBe('B2:D6');
    expect(f(ws, 'D4')).toBe('Sales[[#This Row],[Qty]]*2');
    insertRows(ws, 7, 1); // just below: not part of the table
    expect(t.ref).toBe('B2:D6');
    expect(() => deleteRows(ws, 2)).toThrow(/header row of table "Sales"/);
    expect(t.ref).toBe('B2:D6');
    expect(() => deleteRows(ws, 3, 4)).toThrow(/without data rows/);
    deleteRows(ws, 3, 2);
    expect(t.ref).toBe('B2:D4');
    insertColumns(ws, 'C', 1);
    expect(t.ref).toBe('B2:E4');
    expect(t.columns.map((c) => c.name)).toEqual(['Qty', 'Column4', 'Price', 'Total']);
    expect(ws.getValue('C2')).toBe('Column4');
    deleteColumns(ws, 'D', 2);
    expect(t.ref).toBe('B2:C4');
    expect(t.columns.map((c) => c.name)).toEqual(['Qty', 'Column4']);
    insertColumns(ws, 'A', 1);
    expect(t.ref).toBe('C2:D4');
    deleteColumns(ws, 'C', 2);
    expect(ws.tables).toEqual([]);
  });
});

describe('spliceRows / spliceColumns (ExcelJS semantics)', () => {
  function five(): Worksheet {
    const ws = new Workbook().addWorksheet('S');
    ws.addRow(['1,1', '1,2', '1,3']);
    ws.addRow(['2,1', '2,2', '2,3']);
    ws.setCell('A4', 4.1);
    ws.setCell('C4', 4.3);
    ws.addRow(['5,1', '5,2', '5,3']);
    return ws;
  }

  it('remove only', () => {
    const ws = five();
    spliceRows(ws, 2, 1);
    expect(ws.getRow(2).values).toEqual([]);
    expect(ws.getRow(3).values).toEqual([undefined, 4.1, undefined, 4.3]);
    expect(ws.getRow(4).values).toEqual([undefined, '5,1', '5,2', '5,3']);
    expect(ws.addRow(['5,1b'])).toBe(5);
  });

  it('insert fewer / same / more', () => {
    let ws = five();
    spliceRows(ws, 2, 2, ['one', 'two', 'three']);
    expect(ws.getRow(2).values).toEqual([undefined, 'one', 'two', 'three']);
    expect(ws.getRow(3).values).toEqual([undefined, 4.1, undefined, 4.3]);
    expect(ws.getRow(4).values).toEqual([undefined, '5,1', '5,2', '5,3']);
    ws = five();
    spliceRows(ws, 2, 2, ['one'], ['une']);
    expect(ws.getRow(3).values).toEqual([undefined, 'une']);
    expect(ws.getRow(4).values).toEqual([undefined, 4.1, undefined, 4.3]);
    ws = five();
    spliceRows(ws, 2, 2, ['one'], ['une'], ['uno']);
    expect(ws.getRow(4).values).toEqual([undefined, 'uno']);
    expect(ws.getRow(5).values).toEqual([undefined, 4.1, undefined, 4.3]);
    expect(ws.getRow(6).values).toEqual([undefined, '5,1', '5,2', '5,3']);
  });

  it('removing the last row (exceljs issue 2125)', () => {
    const ws = five();
    spliceRows(ws, 5, 1);
    expect(ws.rowCount).toBe(4);
    expect(ws.getRow(5).values).toEqual([]);
  });

  it('replaced rows keep references; the net change moves the rest', () => {
    const ws = five();
    ws.setCell('E1', { formula: 'A2+A4+A5' });
    spliceRows(ws, 2, 2, ['x']);
    expect(f(ws, 'E1')).toBe('A2+A3+A4');
  });

  it('spliceColumns removes, replaces and inserts columns', () => {
    const ws = five();
    spliceColumns(ws, 'B', 1);
    expect(ws.getRow(1).values).toEqual([undefined, '1,1', '1,3']);
    spliceColumns(ws, 2, 1, ['a', 'b'], ['c', 'd']);
    expect(ws.getRow(1).values).toEqual([undefined, '1,1', 'a', 'c']);
    expect(ws.getRow(2).values).toEqual([undefined, '2,1', 'b', 'd']);
    expect(ws.getRow(4).values).toEqual([undefined, 4.1]);
  });
});

describe('duplicateRow', () => {
  it('inserts copies (values, styles, height, merges, translated formulas)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('D');
    ws.addRow([1, 2]);
    ws.setCell('A2', 'src', { font: { bold: true } });
    ws.setCell('B2', { formula: 'B1*2+$A$1' });
    ws.setCell('C2', 'm');
    ws.mergeCells('C2:D2');
    ws.setRowProps(2, { height: 40 });
    ws.addRow(['after']);
    duplicateRow(ws, 2, 2, true);
    expect(ws.getValue('A3')).toBe('src');
    expect(ws.getValue('A4')).toBe('src');
    expect(ws.getStyle('A4').font?.bold).toBe(true);
    expect(ws.getRowProps(4).height).toBe(40);
    expect(f(ws, 'B3')).toBe('B2*2+$A$1');
    expect(f(ws, 'B4')).toBe('B3*2+$A$1');
    expect(ws.merges).toEqual(['C2:D2', 'C3:D3', 'C4:D4']);
    expect(ws.getValue('A5')).toBe('after');
  });

  it('overwrites the rows below', () => {
    const ws = new Workbook().addWorksheet('D');
    ws.addRows([['a'], ['b'], ['c'], ['d']]);
    ws.mergeCells('B3:C3');
    duplicateRow(ws, 1, 2);
    expect(ws.getRow(2).values).toEqual([undefined, 'a']);
    expect(ws.getRow(3).values).toEqual([undefined, 'a']);
    expect(ws.getValue('A4')).toBe('d');
    expect(ws.merges).toEqual([]);
  });
});

describe('files', () => {
  it('an edited workbook writes, reads back and loads in exceljs', async () => {
    const { wb, ws } = sheet();
    const other = wb.addWorksheet('Other');
    other.setCell('A1', { formula: 'SUM(Data!B1:B6)' });
    ws.mergeCells('A7:C7');
    addTable(ws, {
      name: 'Tbl',
      ref: 'E1:F3',
      columns: ['x', 'y'],
      rows: [
        [1, 2],
        [3, 4],
      ],
    });
    addConditionalFormatting(ws, 'A1:A6', [{ type: 'expression', formulae: ['A1>2'] }]);
    insertRows(ws, 2, 1);
    insertColumns(ws, 'A', 1);
    deleteRows(ws, 5);
    const back = readXlsx(writeXlsx(wb));
    const bs = back.getWorksheet('Data')!;
    expect(bs.getValue('B1')).toBe(1);
    expect(bs.getValue('B3')).toBe(2);
    expect(bs.merges).toEqual(['B7:D7']);
    expect(bs.tables![0]!.ref).toBe('F1:G4');
    expect(bs.conditionalFormatting![0]!.ref).toBe('B1:B6');
    expect(f(back.getWorksheet('Other')!, 'A1')).toBe('SUM(Data!C1:C6)');
    const ex = await oracle(writeXlsx(wb));
    expect(ex.getWorksheet('Data')!.getCell('B3').value).toBe(2);
    expect(ex.getWorksheet('Other')!.getCell('A1').formula).toBe('SUM(Data!C1:C6)');
  });
});
