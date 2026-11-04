/// <reference types="node" />
// Tables / ListObjects feature, checked against exceljs (oracle) in every writer.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import {
  addDefinedName,
  addTable,
  addTableRows,
  type Table,
  tableRows,
} from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

async function readAll(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(rs).arrayBuffer());
}

function sales(wb: Workbook, sheet = 'Data'): Table {
  const ws = wb.addWorksheet(sheet);
  return addTable(ws, {
    name: 'Sales',
    topLeft: 'B2',
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Units', totalsRowFunction: 'sum' },
      { name: 'Price', totalsRowFunction: 'average' },
      {
        name: 'Revenue',
        calculatedColumnFormula: 'Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]',
      },
    ],
    rows: [
      ['North', 10, 2.5],
      ['South', 20, 3],
      ['East', 5, 4],
    ],
    totalsRow: true,
    style: { theme: 'TableStyleLight9', showFirstColumn: true },
  });
}

// the table as a streaming user defines it: model on the stream worksheet, cells via addRow
async function streamed(): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  const s = w.addWorksheet('Data');
  s.worksheet.tables = [
    {
      name: 'Sales',
      ref: 'A1:C4',
      columns: [{ name: 'Region' }, { name: 'Units' }, { name: 'Price' }],
      style: { theme: 'TableStyleLight9' },
    },
  ];
  s.addRow(['Region', 'Units', 'Price']);
  s.addRow(['North', 10, 2.5]);
  s.addRow(['South', 20, 3]);
  s.addRow(['East', 5, 4]);
  await w.end();
  return new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
}

describe('addTable', () => {
  it('fills header, data and totals cells', () => {
    const wb = new Workbook();
    const t = sales(wb);
    const ws = wb.worksheets[0]!;
    expect(t.ref).toBe('B2:E6');
    expect(ws.getValue('B2')).toBe('Region');
    expect(ws.getValue('E2')).toBe('Revenue');
    expect(ws.getValue('C4')).toBe(20);
    expect(ws.getValue('E3')).toEqual({
      formula: 'Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]',
    });
    expect(ws.getValue('B6')).toBe('Total');
    expect(ws.getValue('C6')).toEqual({ formula: 'SUBTOTAL(109,Sales[Units])' });
    expect(ws.getValue('D6')).toEqual({ formula: 'SUBTOTAL(101,Sales[Price])' });
    expect(ws.tables).toHaveLength(1);
  });

  it('uses the given ref, supports custom totals and escapes structured refs', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addTable(ws, {
      name: 'Tab_1',
      ref: 'A1:B5',
      columns: [
        'a[1]',
        { name: 'b', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(Tab_1[b])*2' },
      ],
      rows: [[1, 2]],
      totalsRow: true,
    });
    expect(ws.getValue('B5')).toEqual({ formula: 'SUM(Tab_1[b])*2' });
    const ws2 = wb.addWorksheet('S2');
    addTable(ws2, {
      name: 'Tab_2',
      columns: [{ name: "x'y", totalsRowFunction: 'max' }],
      rows: [[1]],
      totalsRow: true,
    });
    expect(ws2.getValue('A3')).toEqual({ formula: "SUBTOTAL(104,Tab_2[x''y])" });
  });
});

describe('validation', () => {
  const cols = ['a', 'b'];
  it('rejects bad names, duplicates and overlaps', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    expect(() => addTable(ws, { name: 'A1', columns: cols })).toThrow(/invalid table name/);
    expect(() => addTable(ws, { name: 'has space', columns: cols })).toThrow(/invalid/);
    expect(() => addTable(ws, { name: 'T', columns: ['a', 'A'] })).toThrow(/duplicate column/);
    expect(() => addTable(ws, { name: 'T', ref: 'A1:C3', columns: cols })).toThrow(/do not match/);
    expect(() => addTable(ws, { name: 'T', ref: 'A1:B1', columns: cols })).toThrow(/data row/);
    expect(() =>
      addTable(ws, { name: 'T', columns: cols, rows: [[1, 2]], totalsRow: true, ref: 'A1:B2' }),
    ).toThrow(/data row/);
    expect(() =>
      addTable(ws, { name: 'T', ref: 'A1:B3', columns: cols, rows: [[1], [2], [3]] }),
    ).toThrow(/do not fit/);
    expect(() =>
      addTable(ws, {
        name: 'T',
        columns: [{ name: 'a', totalsRowFunction: 'custom' }, 'b'],
        totalsRow: true,
      }),
    ).toThrow(/custom totals/);

    addTable(ws, { name: 'T', columns: cols, rows: [[1, 2]] });
    expect(() => addTable(ws, { name: 'U', topLeft: 'B2', columns: cols })).toThrow(
      /overlaps table/,
    );
    expect(() => addTable(ws, { name: 't', topLeft: 'A10', columns: cols })).toThrow(
      /more than once/,
    );

    const ws2 = wb.addWorksheet('S2');
    expect(() => addTable(ws2, { name: 'T', columns: cols })).toThrow(/more than once/);
    ws2.mergeCells('B2:C2');
    expect(() =>
      addTable(ws2, { name: 'M', topLeft: 'A1', columns: cols, rows: [[1, 2]] }),
    ).toThrow(/merged cells/);
    addDefinedName(wb, 'Taken', '1');
    expect(() => addTable(ws2, { name: 'taken', topLeft: 'A10', columns: cols })).toThrow(
      /defined name/,
    );
    expect(() => addDefinedName(wb, 'T', '1')).toThrow(/table name/);
  });

  it('writers re-validate the model and check header cells', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['x', 'y']);
    ws.addRow([1, 2]);
    ws.tables = [{ name: 'T', ref: 'A1:B2', columns: [{ name: 'x' }, { name: 'z' }] }];
    expect(() => writeXlsx(wb)).toThrow(/header cell B1 is "y" but the column is named "z"/);
    ws.tables = [{ name: 'T', ref: 'A1:B2', columns: [{ name: 'x' }, { name: 'y' }] }];
    expect(() => writeXlsx(wb)).not.toThrow();
    const ws2 = wb.addWorksheet('S2');
    ws2.tables = [{ name: 'T', ref: 'A1:B2', columns: [{ name: 'x' }, { name: 'y' }] }];
    expect(() => writeXlsx(wb)).toThrow(/more than once/);
  });
});

describe('table part', () => {
  it('writes the required attributes, part, rel and content type', () => {
    const wb = new Workbook();
    sales(wb);
    const buf = writeXlsx(wb);
    const p = parts(buf);
    const x = p.get('xl/tables/table1.xml')!;
    expect(x).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(x).toMatch(/<table [^>]*id="1"/);
    expect(x).toContain('name="Sales" displayName="Sales" ref="B2:E6"');
    expect(x).toContain('headerRowCount="1" totalsRowCount="1"');
    expect(x).toContain('<autoFilter ref="B2:E5"/>');
    expect(x).toContain('<tableColumns count="4">');
    expect(x).toContain('<tableColumn id="1" name="Region" totalsRowLabel="Total"/>');
    expect(x).toContain('<tableColumn id="2" name="Units" totalsRowFunction="sum"/>');
    expect(x).toContain(
      '<calculatedColumnFormula>Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]</calculatedColumnFormula>',
    );
    expect(x).toContain(
      '<tableStyleInfo name="TableStyleLight9" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    );
    expect(p.get('[Content_Types].xml')).toContain(
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
    );
    expect(p.get('xl/worksheets/_rels/sheet1.xml.rels')).toMatch(
      /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/table" Target="..\/tables\/table1\.xml"/,
    );
    const sheet = p.get('xl/worksheets/sheet1.xml')!;
    expect(sheet).toMatch(/<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/);
    expect(sheet.indexOf('</sheetData>')).toBeLessThan(sheet.indexOf('<tableParts'));
  });

  it('gives workbook-unique ids across sheets and omits optional parts', () => {
    const wb = new Workbook();
    for (const n of ['A', 'B']) {
      const ws = wb.addWorksheet(n);
      addTable(ws, {
        name: `T${n}`,
        columns: ['c'],
        rows: [[1]],
        headerRow: n === 'A',
        autoFilter: n === 'A',
        style: { theme: '', showRowStripes: false },
      });
    }
    const p = parts(writeXlsx(wb));
    expect(p.get('xl/tables/table1.xml')).toMatch(/id="1" name="TA"/);
    expect(p.get('xl/tables/table2.xml')).toMatch(/id="2" name="TB"/);
    const b = p.get('xl/tables/table2.xml')!;
    expect(b).toContain('headerRowCount="0"');
    expect(b).not.toContain('<autoFilter');
    expect(b).toContain(
      '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0"',
    );
    expect(p.get('xl/worksheets/_rels/sheet2.xml.rels')).toContain('table2.xml');
  });
});

describe('dxf styles', () => {
  const HEAD = { font: { bold: true, color: { argb: 'FFFF0000' } } };
  const DATA = { numFmt: '0.00' };
  const TOT = { font: { italic: true } };
  const COL = {
    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } },
  } as const;

  function styled(wb: Workbook): void {
    const ws = wb.addWorksheet('S');
    addTable(ws, {
      name: 'Styled',
      topLeft: 'A1',
      columns: [
        { name: 'A', totalsRowLabel: 'Sum', dataStyle: COL },
        { name: 'B', totalsRowFunction: 'sum', headerStyle: COL, totalsStyle: HEAD },
      ],
      rows: [
        ['x', 1],
        ['y', 2],
      ],
      totalsRow: true,
      headerStyle: HEAD,
      dataStyle: DATA,
      totalsStyle: TOT,
    });
  }

  it('writes dxfIds for the table and its columns, deduping equal styles', () => {
    const wb = new Workbook();
    styled(wb);
    const p = parts(writeXlsx(wb));
    const t = p.get('xl/tables/table1.xml')!;
    // dxf order = order of interning: table header, data, totals, then column styles
    expect(t).toContain('headerRowDxfId="0" dataDxfId="1" totalsRowDxfId="2">');
    expect(t).toContain('<tableColumn id="1" name="A" totalsRowLabel="Sum" dataDxfId="3"/>');
    expect(t).toContain(
      '<tableColumn id="2" name="B" totalsRowFunction="sum" headerRowDxfId="3" totalsRowDxfId="0"/>',
    );
    expect(p.get('xl/styles.xml')).toMatch(/<dxfs count="4">/);
  });

  it('omits header / totals dxfs when the row is absent', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addTable(ws, {
      name: 'NoRows',
      columns: ['A'],
      headerRow: false,
      headerStyle: HEAD,
      totalsStyle: TOT,
      dataStyle: DATA,
    });
    const t = parts(writeXlsx(wb)).get('xl/tables/table1.xml')!;
    expect(t).not.toContain('headerRowDxfId');
    expect(t).not.toContain('totalsRowDxfId');
    expect(t).toContain('dataDxfId="0"');
  });

  it('round-trips ours -> ours and exceljs accepts the file', async () => {
    const wb = new Workbook();
    styled(wb);
    const buf = writeXlsx(wb);
    for (const b of [buf, new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())]) {
      const back = readXlsx(b);
      const t = back.getWorksheet('S')!.tables![0]!;
      expect(t.headerStyle).toEqual(HEAD);
      expect(t.dataStyle).toEqual(DATA);
      expect(t.totalsStyle).toEqual(TOT);
      expect(t.columns[0]!.dataStyle).toEqual(COL);
      expect(t.columns[1]!.headerStyle).toEqual(COL);
      expect(t.columns[1]!.totalsStyle).toEqual(HEAD);
      expect(t.columns[0]!.headerStyle).toBeUndefined();
      // and survives a second pass
      expect(readXlsx(writeXlsx(back)).getWorksheet('S')!.tables![0]).toEqual(t);
    }
    expect((await oracle(buf)).getWorksheet('S')!.getTable('Styled')).toBeDefined();
  });
});

describe('round trips', () => {
  it('ours -> ours (writeXlsx, xlsxStream)', async () => {
    const wb = new Workbook();
    const t = sales(wb);
    for (const buf of [writeXlsx(wb), await readAll(xlsxStream(wb))]) {
      const back = readXlsx(buf);
      const got = back.worksheets[0]!.tables![0]!;
      expect(got).toMatchObject({
        name: 'Sales',
        ref: t.ref,
        totalsRow: true,
        style: { theme: 'TableStyleLight9', showFirstColumn: true, showRowStripes: true },
      });
      expect(got.columns).toEqual([
        { name: 'Region', totalsRowLabel: 'Total' },
        { name: 'Units', totalsRowFunction: 'sum' },
        { name: 'Price', totalsRowFunction: 'average' },
        {
          name: 'Revenue',
          calculatedColumnFormula: 'Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]',
        },
      ]);
      // and re-writes identically
      expect(parts(writeXlsx(back)).get('xl/tables/table1.xml')).toBe(
        parts(writeXlsx(wb)).get('xl/tables/table1.xml'),
      );
    }
  });

  it('streaming writer (model on stream sheet, user adds the cells)', async () => {
    const buf = await streamed();
    expect(parts(buf).get('xl/tables/table1.xml')).toContain('ref="A1:C4"');
    const t = readXlsx(buf).worksheets[0]!.tables![0]!;
    expect(t.name).toBe('Sales');
    expect(t.columns.map((c) => c.name)).toEqual(['Region', 'Units', 'Price']);
    const x = await oracle(buf);
    expect(x.getWorksheet('Data')!.getTable('Sales')).toBeDefined();
  });

  it('ours -> exceljs', async () => {
    const wb = new Workbook();
    sales(wb);
    const second = wb.addWorksheet('Other');
    addTable(second, { name: 'Plain', columns: ['k', 'v'], rows: [['a', 1]], headerRow: true });
    for (const buf of [writeXlsx(wb), await readAll(xlsxStream(wb))]) {
      const x = await oracle(buf);
      const ws = x.getWorksheet('Data')!;
      const t = ws.getTable('Sales');
      // biome-ignore lint/suspicious/noExplicitAny: exceljs Table typings are partial
      const m = (t as any).table as Record<string, any>;
      expect(m.name).toBe('Sales');
      expect(m.tableRef).toBe('B2:E6');
      expect(m.totalsRow).toBe(true);
      expect(m.headerRow).toBe(true);
      expect(m.columns.map((c: { name: string }) => c.name)).toEqual([
        'Region',
        'Units',
        'Price',
        'Revenue',
      ]);
      expect(m.columns[1].totalsRowFunction).toBe('sum');
      expect(m.columns[0].totalsRowLabel).toBe('Total');
      expect(m.style.theme).toBe('TableStyleLight9');
      expect(m.style.showFirstColumn).toBe(true);
      expect(ws.getCell('B2').value).toBe('Region');
      expect(x.getWorksheet('Other')!.getTable('Plain')).toBeDefined();
      expect(x.getWorksheet('Other')!.getTables()).toHaveLength(1);
    }
  });

  it('exceljs -> ours', async () => {
    const x = new ExcelJS.Workbook();
    const ws = x.addWorksheet('Sheet1');
    ws.addTable({
      name: 'MyTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: true,
      style: { theme: 'TableStyleDark3', showRowStripes: true },
      columns: [
        { name: 'Date', totalsRowLabel: 'Totals', filterButton: true },
        { name: 'Amount', totalsRowFunction: 'sum', filterButton: false },
      ],
      rows: [
        [new Date('2019-07-20'), 70.1],
        [new Date('2019-07-21'), 70.6],
      ],
    });
    const back = readXlsx(new Uint8Array(await x.xlsx.writeBuffer()));
    const t = back.worksheets[0]!.tables![0]!;
    expect(t).toMatchObject({
      name: 'MyTable',
      ref: 'A1:B4',
      totalsRow: true,
      style: { theme: 'TableStyleDark3', showRowStripes: true },
    });
    expect(t.columns[0]).toMatchObject({ name: 'Date', totalsRowLabel: 'Totals' });
    expect(t.columns[1]).toMatchObject({ name: 'Amount', totalsRowFunction: 'sum' });
    // header text matches the columns, so our writer accepts the table unchanged
    expect(() => writeXlsx(back)).not.toThrow();
    const again = await oracle(writeXlsx(back));
    expect(again.getWorksheet('Sheet1')!.getTable('MyTable')).toBeDefined();
  });
});

describe('table rows: tableRows / addTableRows / column width', () => {
  function sales(totals: boolean) {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const t = addTable(ws, {
      name: 'Sales',
      topLeft: 'B2',
      columns: [
        { name: 'Item', width: 18, totalsRowLabel: 'Total' },
        { name: 'Qty', totalsRowFunction: 'sum' },
        { name: 'Double', calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*2' },
      ],
      rows: [
        ['a', 1],
        ['b', 2],
      ],
      totalsRow: totals,
    });
    ws.setCell('B10', 'below');
    ws.setCell('F1', { formula: 'SUM(C3:C4)' });
    return { wb, ws, t };
  }

  it('addTable sets column widths without storing them in the table', () => {
    const { ws, t } = sales(false);
    expect(ws.getColumn('B').width).toBe(18);
    expect(t.columns[0]).not.toHaveProperty('width');
  });

  it('tableRows returns data rows keyed by column name', () => {
    const { ws } = sales(true);
    expect(tableRows(ws, 'sales')).toEqual([
      { Item: 'a', Qty: 1, Double: { formula: 'Sales[[#This Row],[Qty]]*2' } },
      { Item: 'b', Qty: 2, Double: { formula: 'Sales[[#This Row],[Qty]]*2' } },
    ]);
    expect(() => tableRows(ws, 'Nope')).toThrow(/no table "Nope"/);
  });

  for (const totals of [false, true]) {
    it(`addTableRows appends rows (totals row: ${totals})`, async () => {
      const { wb, ws, t } = sales(totals);
      addTableRows(ws, 'Sales', [['c', 3], { Qty: 4, Item: 'd' }]);
      expect(t.ref).toBe(totals ? 'B2:D7' : 'B2:D6');
      expect(tableRows(ws, 'Sales').map((r) => [r.Item, r.Qty])).toEqual([
        ['a', 1],
        ['b', 2],
        ['c', 3],
        ['d', 4],
      ]);
      expect(ws.getValue('D6')).toEqual({ formula: 'Sales[[#This Row],[Qty]]*2' });
      expect(ws.getValue('B12')).toBe('below'); // moved down with the inserted rows
      if (totals) expect(ws.getValue('B7')).toBe('Total');
      expect(ws.getValue('F1')).toEqual({ formula: 'SUM(C3:C4)' });
      const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
      expect(back.tables![0]!.ref).toBe(t.ref);
      const ex = await oracle(writeXlsx(wb));
      expect(ex.getWorksheet('S')!.getCell('B6').value).toBe('d');
    });
  }
});
