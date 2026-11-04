/// <reference types="node" />
// Dynamic-array formula metadata, copyWorksheet / copyRange, nested column keys.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addConditionalFormatting,
  addImage,
  addNote,
  addTable,
  copyRange,
  copyWorksheet,
  placeImage,
  readXlsx,
  Workbook,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFWEXHbQSACj/P8Fu7N9hAAAAAElFTkSuQmCC',
  ),
  (c) => c.charCodeAt(0),
);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('dynamic-array formulas', () => {
  it('spilling functions get cm="1" + t="array" and xl/metadata.xml; others do not', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRows([[3], [1], [2]]);
    ws.setCell('C1', { formula: 'SORT(A1:A3)' });
    ws.setCell('D1', { formula: 'SUM(A1:A3)' });
    ws.setCell('E1', { formula: 'A1:A3*2', array: 'E1:E3' }); // legacy CSE array
    ws.setCell('F1', { formula: 'UNIQUE(A1:A3)', dynamic: false });
    ws.setCell('G1', { formula: 'A1:A3+1', dynamic: true });
    const buf = writeXlsx(wb);
    const files = unzipSync(buf);
    const xml = dec.decode(files.get('xl/worksheets/sheet1.xml')!);
    expect(xml).toContain('<c r="C1" cm="1"><f t="array" ref="C1">_xlfn._xlws.SORT(A1:A3)</f></c>');
    expect(xml).toContain('<c r="D1"><f>SUM(A1:A3)</f></c>');
    expect(xml).toContain('<c r="E1"><f t="array" ref="E1:E3">A1:A3*2</f></c>');
    expect(xml).toContain('<c r="F1"><f>_xlfn.UNIQUE(A1:A3)</f></c>');
    expect(xml).toContain('<c r="G1" cm="1"><f t="array" ref="G1">A1:A3+1</f></c>');
    expect(dec.decode(files.get('xl/metadata.xml')!)).toContain('fDynamic="1"');
    expect(dec.decode(files.get('xl/_rels/workbook.xml.rels')!)).toContain(
      'relationships/sheetMetadata',
    );
    expect(dec.decode(files.get('[Content_Types].xml')!)).toContain('sheetMetadata+xml');

    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.getValue('C1')).toEqual({ formula: 'SORT(A1:A3)', array: 'C1', dynamic: true });
    expect(back.getValue('E1')).toEqual({ formula: 'A1:A3*2', array: 'E1:E3' });
    // round trip keeps the read metadata part once, cells keep cm
    const again = unzipSync(writeXlsx(readXlsx(buf)));
    expect([...again.keys()].filter((k) => k === 'xl/metadata.xml').length).toBe(1);
    expect(
      dec.decode(again.get('xl/_rels/workbook.xml.rels')!).match(/sheetMetadata/g)!.length,
    ).toBe(1);
    expect(dec.decode(again.get('xl/worksheets/sheet1.xml')!)).toContain('<c r="C1" cm="1">');
    await oracle(buf);
  });

  it('no dynamic formula: no metadata part', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').setCell('A1', { formula: 'SUM(1,2)' });
    expect(unzipSync(writeXlsx(wb)).has('xl/metadata.xml')).toBe(false);
  });
});

describe('copyWorksheet', () => {
  function source(): Workbook {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data', { tabColor: { argb: 'FF00FF00' } });
    ws.columns = [{ key: 'item', width: 20 }, { key: 'qty' }];
    addTable(ws, {
      name: 'Sales',
      columns: ['Item', 'Qty'],
      rows: [
        ['a', 1],
        ['b', 2],
      ],
    });
    ws.setCell('D1', new Date(Date.UTC(2024, 0, 1)));
    ws.setCell('D2', { formula: 'SUM(B2:B3)' }, { font: { bold: true } });
    ws.mergeCells('F1:G2');
    ws.setRowProps(3, { height: 30 });
    addNote(ws, 'A1', 'note');
    addConditionalFormatting(ws, 'B2:B3', [{ type: 'expression', formulae: ['B2>1'] }]);
    placeImage(ws, addImage(wb, { data: PNG, extension: 'png' }), 'H1:I3');
    ws.rowBreaks = [2];
    return wb;
  }

  it('copies within the workbook (unique sheet and table names, independent data)', () => {
    const wb = source();
    const src = wb.getWorksheet('Data')!;
    const c = copyWorksheet(src);
    expect(c.name).toBe('Data (2)');
    expect(c.tables![0]!.name).toBe('Sales_2');
    expect(c.getValue('D2')).toEqual({ formula: 'SUM(B2:B3)' });
    expect(c.getStyleIndex('D2')).toBe(src.getStyleIndex('D2'));
    expect(c.merges).toEqual(['F1:G2']);
    expect(c.getRowProps(3).height).toBe(30);
    expect(c.getColumn('item').width).toBe(20);
    expect(c.images![0]!.imageId).toBe(src.images![0]!.imageId);
    c.comments![0]!.text = 'changed';
    expect(src.comments![0]!.text).toBe('note');
    expect(copyWorksheet(src).name).toBe('Data (3)');
    const back = readXlsx(writeXlsx(wb));
    expect(back.worksheets.map((s) => s.name)).toEqual(['Data', 'Data (2)', 'Data (3)']);
  });

  it('copies into another workbook: styles re-interned, image re-added, dates converted', async () => {
    const wb = source();
    const dst = new Workbook({ date1904: true });
    dst.addWorksheet('Existing').setCell('A1', 'x', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } },
    });
    const c = copyWorksheet(wb.getWorksheet('Data')!, dst, 'Imported');
    expect(c.getStyle('D2').font?.bold).toBe(true);
    expect(c.getValue('D1')).toEqual(new Date(Date.UTC(2024, 0, 1)));
    expect(dst.images!.length).toBe(1);
    expect(c.images![0]!.imageId).toBe(dst.images![0]!.id);
    expect(c.tables![0]!.name).toBe('Sales');
    const buf = writeXlsx(dst);
    const back = readXlsx(buf).getWorksheet('Imported')!;
    expect(back.getValue('D1')).toEqual(new Date(Date.UTC(2024, 0, 1)));
    expect(back.tables![0]!.ref).toBe('A1:B3');
    await oracle(buf);
  });
});

describe('copyRange', () => {
  it('copies values, styles, merges; translates relative formula refs; overlapping is safe', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRows([
      [1, 2],
      [3, 4],
    ]);
    ws.setCell('C1', { formula: 'A1+B1+$A$1' }, { font: { italic: true } });
    ws.mergeCells('A3:B3');
    ws.setCell('E5', 'old');
    copyRange(ws, 'A1:C3', ws, 'D4');
    expect(ws.getValue('D4')).toBe(1);
    expect(ws.getValue('F4')).toEqual({ formula: 'D4+E4+$A$1' });
    expect(ws.getStyle('F4').font?.italic).toBe(true);
    expect(ws.getValue('E5')).toBe(4); // replaced
    expect(ws.merges).toContain('D6:E6');
    copyRange(ws, 'A1:B2', ws, 'B1'); // overlapping: reads before writing
    expect([ws.getValue('B1'), ws.getValue('C1'), ws.getValue('C2')]).toEqual([1, 2, 4]);
    const other = new Workbook().addWorksheet('T');
    copyRange(ws, 'F4', other, 'A1');
    expect(other.getStyle('A1').font?.italic).toBe(true);
    expect(other.getValue('A1')).toEqual({ formula: '#REF!+#REF!+$A$1' });
  });
});

describe('nested column keys', () => {
  it('addRow resolves dotted keys; flat keys with dots still win', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.columns = [{ key: 'name' }, { key: 'address.city' }, { key: 'a.b' }];
    ws.addRow({ name: 'Ann', address: { city: 'Oslo' }, 'a.b': 'flat', a: { b: 'nested' } });
    ws.addRow({ name: 'Bo', address: null });
    expect(ws.getRow(1).values).toEqual([undefined, 'Ann', 'Oslo', 'flat']);
    expect(ws.getRow(2).values).toEqual([undefined, 'Bo']);
  });
});
