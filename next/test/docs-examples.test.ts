// Code samples of docs/features.md and docs/migrating-from-exceljs.md (same code, imports from
// '../src'). Together with readme-examples.test.ts this keeps the documentation honest.

import { describe, expect, it } from 'vitest';
import {
  addChart,
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addImage,
  addNote,
  addTable,
  createXlsxWriter,
  deleteRows,
  duplicateRow,
  getPrintArea,
  insertColumns,
  insertRows,
  listPreservedParts,
  placeImage,
  readXlsx,
  setPrintArea,
  setPrintTitles,
  spliceRows,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../src/index.js';

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

describe('docs/migrating-from-exceljs.md', () => {
  it('file round trip and csv', async () => {
    const { mkdtempSync, readFileSync, rmSync, writeFileSync } = await import('node:fs');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const { arraysToSheet, sheetToArrays } = await import('../src/json/index.js');
    const { parseCsv, writeCsv } = await import('../src/csv/index.js');
    const dir = mkdtempSync(join(tmpdir(), 'turbosheet-mig-'));
    try {
      const src = new Workbook();
      src.addWorksheet('S').setCell('A1', 'orig');
      writeFileSync(join(dir, 'in.xlsx'), writeXlsx(src));

      const wb = readXlsx(readFileSync(join(dir, 'in.xlsx')));
      wb.getWorksheet(1)!.getCell('A1').value = 'edited';
      writeFileSync(join(dir, 'out.xlsx'), writeXlsx(wb));
      expect(
        readXlsx(readFileSync(join(dir, 'out.xlsx')))
          .getWorksheet(1)!
          .getValue('A1'),
      ).toBe('edited');

      const rows = parseCsv('a,1\nb,2', { parseNumbers: true });
      const csv = writeCsv(rows, { escapeFormulas: true });
      expect(csv).toBe('a,1\r\nb,2\r\n');
      const ws = wb.addWorksheet('Csv');
      arraysToSheet(ws, rows);
      expect(sheetToArrays(ws, { raw: true })).toEqual([
        ['a', 1],
        ['b', 2],
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('merged cells and styles are values', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 'top-left');
    ws.setCell('B1', 'hidden');
    ws.mergeCells('A1:B1');
    expect(ws.getValue('B1')).toBe('hidden');
    const style = ws.getStyle('A1');
    ws.setStyle('A1', { ...style, font: { ...style.font, bold: true } });
    expect(ws.getStyle('A1').font?.bold).toBe(true);
    expect(() => ws.mergeCells('B1:C2')).toThrow();
  });
});

describe('docs/features.md', () => {
  it('inserting and deleting rows and columns', () => {
    const ws = new Workbook().addWorksheet('Data');
    ws.addRows([
      ['Item', 'Qty'],
      ['a', 1],
      ['b', 2],
      ['c', 3],
      ['d', 4],
    ]);
    ws.setCell('D1', { formula: 'SUM(B2:B5)' });
    insertRows(ws, 3, [['Inserted', 1]], { inheritStyle: 'above' });
    expect(ws.getValue('D1')).toEqual({ formula: 'SUM(B2:B6)' });
    deleteRows(ws, 2);
    expect(ws.getValue('D1')).toEqual({ formula: 'SUM(B2:B5)' });
    insertColumns(ws, 'B', 1);
    expect(ws.getValue('E1')).toEqual({ formula: 'SUM(C2:C5)' });
    spliceRows(ws, 2, 1, ['Replaced', 2]);
    duplicateRow(ws, 2, 2, true);
    expect(ws.getRow(4).values).toEqual([undefined, 'Replaced', 2]);
    expect(ws.getValue('A5')).toBe('b');
  });

  it('values and styles', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    ws.setCell('A1', 'plain text');
    ws.setCell('A2', {
      richText: [{ text: 'bold ', font: { bold: true } }, { text: 'and plain' }],
    });
    ws.setCell('A3', {
      text: 'turbosheet',
      hyperlink: 'https://example.com',
      tooltip: 'Home page',
    });
    ws.setCell('B1', 10);
    ws.setCell('B2', 32);
    ws.setCell('B3', { formula: 'SUM(B1:B2)', result: 42 });
    ws.setCell('C1', new Date(Date.UTC(2025, 5, 1)));
    ws.setCell('C2', { error: '#N/A' });
    ws.setStyle('A1', {
      font: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } },
      alignment: { horizontal: 'center', wrapText: true },
      border: { bottom: { style: 'thin', color: { argb: 'FF000000' } } },
    });
    ws.getCell('B1').numFmt = '#,##0.00';
    ws.getColumn('B').width = 14;
    const xf = wb.styles.intern({ font: { italic: true } });
    ws.setCell(5, 1, 'x', xf);

    const back = readXlsx(writeXlsx(wb)).getWorksheet('Data')!;
    expect(back.getValue('A3')).toMatchObject({ hyperlink: 'https://example.com' });
    expect(back.getValue('A2')).toHaveProperty('richText');
    expect(back.getValue('B3')).toMatchObject({ formula: 'SUM(B1:B2)', result: 42 });
    expect(back.getValue('C1')).toEqual(new Date(Date.UTC(2025, 5, 1)));
    expect(back.getValue('C2')).toEqual({ error: '#N/A' });
    expect(back.getStyle('A1').font?.bold).toBe(true);
    expect(back.getStyle('B1').numFmt).toBe('#,##0.00');
    expect(back.getStyle(5, 1).font?.italic).toBe(true);
  });

  it('charts', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['Month', 'North', 'South']);
    ws.addRow(['Jan', 10, 30]);
    ws.addRow(['Feb', 15, 27]);
    ws.addRow(['Mar', 20, 24]);
    addChart(ws, {
      type: 'column', // 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area'
      title: 'Units per month',
      series: [
        { name: { ref: 'B1' }, categories: 'A2:A4', values: 'B2:B4' }, // the chart's own sheet
        {
          name: 'South',
          categories: "'Data'!$A$2:$A$4",
          values: "'Data'!$C$2:$C$4",
          color: { argb: 'FFC00000' },
        },
      ],
      anchor: 'E2:L16', // same forms as placeImage
      legend: { position: 'b' }, // or false
      axes: { y: { title: 'Units', min: 0 } },
      stacked: 'stacked', // bar / column / line / area
      dataLabels: { showValue: true },
    });
    const out = writeXlsx(wb);
    expect(listPreservedParts(readXlsx(out)).map((p) => p.path)).toContain('xl/charts/chart1.xml');
  });

  it('notes, pictures, conditional formatting, validation, names, print, protection', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    for (let r = 2; r <= 20; r++) ws.addRow(['North', r * 10, r, r]);
    addNote(ws, 'B2', 'Check this value');
    addNote(ws, 'B3', { text: 'Approved', author: 'Ann', visible: true, width: 200, height: 60 });
    addNote(ws, 'B4', { richText: [{ text: 'Important', font: { bold: true } }] });
    const id = addImage(wb, { data: PNG, extension: 'png' });
    placeImage(ws, id, 'D2:F6');
    placeImage(ws, id, { tl: 'H2', ext: { width: 120, height: 80 } }, { description: 'Logo' });
    ws.backgroundImage = id;
    addConditionalFormatting(ws, 'B2:B20', [
      { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } },
    ]);
    addConditionalFormatting(ws, 'C2:C20', [
      { type: 'colorScale', colors: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }] },
    ]);
    addConditionalFormatting(ws, 'D2:D20', [{ type: 'dataBar', color: { argb: 'FF638EC6' } }]);
    addDataValidation(ws, { sqref: 'A2:A20', type: 'list', values: ['North', 'South', 'East'] });
    addDataValidation(ws, {
      sqref: 'B2:B20',
      type: 'whole',
      operator: 'between',
      formula1: 0,
      formula2: 1000,
      showErrorMessage: true,
      errorTitle: 'Units',
      error: 'Enter a whole number from 0 to 1000',
    });
    addDefinedName(wb, 'Rate', '0.2');
    addDefinedName(wb, 'Totals', 'Data!B2:B20');
    setPrintArea(ws, 'A1:D50');
    setPrintTitles(ws, { rows: '1:1' });
    expect(getPrintArea(ws)).toEqual(['A1:D50']);
    ws.pageSetup = {
      orientation: 'landscape',
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    };
    ws.headerFooter = { oddHeader: '&CMonthly report', oddFooter: '&LPage &P of &N' };
    ws.protection = { password: 's3cret', formatCells: false, insertRows: false, spinCount: 1000 };

    const back = readXlsx(writeXlsx(wb)).getWorksheet('Data')!;
    expect(back.comments).toHaveLength(3);
    expect(back.images).toHaveLength(2);
    expect(back.backgroundImage).toBeDefined();
    expect(back.conditionalFormatting).toHaveLength(3);
    expect(back.dataValidations).toHaveLength(2);
    expect(back.headerFooter?.oddHeader).toBe('&CMonthly report');
    expect(back.protection?.sheet).not.toBe(false);
  });

  it('table, preservation, streaming with features', async () => {
    const wb = new Workbook();
    const ws: Worksheet = wb.addWorksheet('Data');
    addTable(ws, {
      name: 'Sales',
      columns: ['Region', { name: 'Units', totalsRowFunction: 'sum' }],
      rows: [
        ['North', 120],
        ['South', 80],
      ],
      totalsRow: true,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
    });
    const bytes = writeXlsx(wb);
    const loaded = readXlsx(bytes);
    loaded.getWorksheet('Data')!.setCell('B2', 999);
    expect(listPreservedParts(loaded)).toBeInstanceOf(Array);
    expect(readXlsx(writeXlsx(loaded)).getWorksheet('Data')!.tables).toHaveLength(1);
    expect(readXlsx(bytes, { preserveUnknown: false }).worksheets).toHaveLength(1);

    const chunks: Uint8Array[] = [];
    const sink = (c: Uint8Array): void => void chunks.push(c.slice());
    const w = createXlsxWriter({ sink });
    const s = w.addWorksheet('Data', { columns: [{ header: 'Region', key: 'region' }], freeze: 1 });
    s.worksheet.dataValidations = [{ sqref: 'A2:A1000', type: 'list', values: ['North', 'South'] }];
    for (const r of [{ region: 'North' }, { region: 'South' }]) s.addRow(r);
    await s.commit();
    await w.end();
    const out = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    expect(readXlsx(out).getWorksheet('Data')!.dataValidations).toHaveLength(1);
  });
});
