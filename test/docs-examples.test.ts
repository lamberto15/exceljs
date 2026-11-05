// Code samples of docs/features.md, docs/migrating-from-exceljs.md and docs/formula.md (same
// code, imports from '../src'). Together with readme-examples.test.ts this keeps the documentation honest.

import { describe, expect, it } from 'vitest';
import {
  addChart,
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addImage,
  addNote,
  addTable,
  addTableRows,
  createXlsxWriter,
  deleteRows,
  duplicateRow,
  getPrintArea,
  insertColumns,
  insertRows,
  listPreservedParts,
  placeImage,
  readXlsx,
  richTextFromHtml,
  setPrintArea,
  setPrintTitles,
  spliceRows,
  tableRows,
  Workbook,
  type Worksheet,
  writeXlsx,
  xlsxRows,
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

  it('csv typing like ExcelJS csv.read', async () => {
    const { arraysToSheet } = await import('../src/json/index.js');
    const { parseCsv } = await import('../src/csv/index.js');
    const text = 'n,ok,err,when\n5,true,#N/A,11-04-2019\n';
    const rows = parseCsv(text, {
      parseNumbers: true,
      parseBooleans: true,
      parseErrors: true,
      dateFormats: ['MM-DD-YYYY'],
      localDates: true,
    });
    const ws = arraysToSheet(new Workbook({ localDates: true }).addWorksheet('Sheet1'), rows);
    expect(ws.getRow(2).values).toEqual([
      undefined,
      5,
      true,
      { error: '#N/A' },
      new Date(2019, 10, 4),
    ]);
  });

  it('streaming write and read', async () => {
    const { createWriteStream, mkdtempSync, openAsBlob, rmSync } = await import('node:fs');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = mkdtempSync(join(tmpdir(), 'turbosheet-mig-'));
    try {
      const rows = [
        ['a', 1],
        ['b', 2],
      ];
      const file = createWriteStream(join(dir, 'big.xlsx'));
      const w = createXlsxWriter({
        sink: (chunk) => new Promise((resolve) => file.write(chunk, () => resolve())),
      });
      const ws = w.addWorksheet('Data');
      for (const r of rows) ws.addRow(r);
      await ws.commit();
      await w.end();
      await new Promise<void>((resolve) => file.end(resolve));

      const seen: unknown[] = [];
      for await (const row of xlsxRows(await openAsBlob(join(dir, 'big.xlsx')))) {
        seen.push([row.sheet, row.row, row.values]);
      }
      expect(seen).toEqual([
        ['Data', 1, ['a', 1]],
        ['Data', 2, ['b', 2]],
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('api map extras', async () => {
    const { columnValues, groupRows } = await import('../src/index.js');
    const ws = new Workbook().addWorksheet('S');
    ws.addRows([['x'], ['y'], ['z']]);
    ws.mergeCells('B1:C2');
    ws.unmergeCells('B1');
    expect(ws.merges).toHaveLength(0);
    expect(columnValues(ws, 'A')).toEqual([undefined, 'x', 'y', 'z']);
    ws.setRowProps(4, { collapsed: true });
    ws.getColumn('E').collapsed = true;
    groupRows(ws, 2, 3, { collapsed: true });
    const back = readXlsx(writeXlsx(ws.workbook)).getWorksheet('S')!;
    expect(back.getRowProps(4).collapsed).toBe(true);
    expect(back.getRowProps(2).hidden).toBe(true);
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
  it('shared formulas, html rich text, notes, lite', async () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('B3', { formula: 'A3*2' });
    ws.setCell('C3', { sharedFormula: 'B3', result: 4 });
    expect(ws.getValue('C3')).toEqual({ formula: 'B3*2', result: 4 });
    ws.setCell('A1', richTextFromHtml('<b>Total</b> due'));
    expect(ws.getValue('A1')).toEqual({
      richText: [{ text: 'Total', font: { bold: true } }, { text: ' due' }],
    });
    addNote(ws, 'B5', {
      text: 'Locked note',
      margins: { insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35] },
      protection: { locked: true, lockText: false },
      editAs: 'twoCells',
    });
    const note = readXlsx(writeXlsx(ws.workbook)).getWorksheet('S')!.comments![0]!;
    expect(note).toMatchObject({
      editAs: 'twoCells',
      protection: { locked: true, lockText: false },
    });

    const lite = await import('../src/lite.js');
    const wb = new lite.Workbook();
    lite.addTable(wb.addWorksheet('Data'), {
      name: 'T',
      columns: ['Region', 'Units'],
      rows: [['N', 1]],
    });
    const bytes = lite.writeXlsx(wb, { features: [lite.tablesWriter] });
    const back = lite.readXlsx(bytes, { features: [lite.tablesReader] });
    expect(back.getWorksheet('Data')!.tables).toHaveLength(1);
    expect(() => lite.writeXlsx(wb)).toThrow(/tablesWriter/);
  });

  it('transform recipe', async () => {
    const { mkdtempSync, openAsBlob, rmSync, writeFileSync } = await import('node:fs');
    const { tmpdir } = await import('node:os');
    const { join } = await import('node:path');
    const dir = mkdtempSync(join(tmpdir(), 'turbosheet-feat-'));
    try {
      const src = new Workbook();
      src.addWorksheet('Orders').addRows([
        ['item', 'qty', 'price'],
        ['bolt', 3, 2],
        ['nut', 0, 1],
      ]);
      writeFileSync(join(dir, 'in.xlsx'), writeXlsx(src));

      const chunks: Uint8Array[] = [];
      const sink = (c: Uint8Array): void => void chunks.push(c.slice());
      const w = createXlsxWriter({ sink });
      const out = w.addWorksheet('Orders');
      for await (const row of xlsxRows(await openAsBlob(join(dir, 'in.xlsx')))) {
        const [item, qty, price] = row.values as [string, number, number];
        if (row.row === 1) out.addRow([item, qty, price, 'total']);
        else if (qty) out.addRow([item, qty, price, qty * price]);
        if (row.row % 5000 === 0) await out.flush();
      }
      await w.end();
      const bytes = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
      expect(readXlsx(bytes).getWorksheet('Orders')!.getRow(2).values).toEqual([
        undefined,
        'bolt',
        3,
        2,
        6,
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('table rows', () => {
    const ws = new Workbook().addWorksheet('T');
    addTable(ws, {
      name: 'Sales',
      columns: ['Region', { name: 'Units', totalsRowFunction: 'sum' }],
      rows: [
        ['North', 120],
        ['South', 80],
      ],
      totalsRow: true,
    });
    addTableRows(ws, 'Sales', [['East', 95], { Region: 'West', Units: 60 }]);
    expect(tableRows(ws, 'Sales')).toEqual([
      { Region: 'North', Units: 120 },
      { Region: 'South', Units: 80 },
      { Region: 'East', Units: 95 },
      { Region: 'West', Units: 60 },
    ]);
    expect(ws.tables![0]!.ref).toBe('A1:B6');
  });

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

describe('docs/formula.md', () => {
  it('evaluate, evaluateCell, evaluateFormula, custom functions', async () => {
    const { createEvaluator, evaluate, evaluateCell, evaluateFormula, toNumber, toText } =
      await import('../src/formula/index.js');
    const src = new Workbook();
    const s1 = src.addWorksheet('Sheet1');
    for (let r = 2; r <= 9; r++) s1.setCell(r, 2, r);
    s1.setCell('D10', { formula: 'SUM(B2:B9)' });
    s1.setCell('E1', { formula: 'BESSELJ(1.9,2)' });
    addDefinedName(src, 'Rate', '0.5');
    const bytes = writeXlsx(src);

    const wb = readXlsx(bytes);
    const report = evaluate(wb);
    expect(report).toMatchObject({
      evaluated: 2,
      circular: [],
      unsupported: ['BESSELJ'],
      errors: [],
    });
    const out = writeXlsx(wb);
    expect(readXlsx(out).getWorksheet('Sheet1')!.getValue('D10')).toEqual({
      formula: 'SUM(B2:B9)',
      result: 44,
    });
    expect(evaluateCell(wb, 'Sheet1', 'D10')).toBe(44);
    expect(evaluateFormula(wb, 'SUM(B2:B9)*Rate', { sheet: 'Sheet1' })).toBe(22);

    const ws = wb.getWorksheet('Sheet1')!;
    ws.setCell('F1', { formula: 'DOUBLE(B2)' });
    ws.setCell('F2', { formula: 'GREET("Ann")' });
    const ev = createEvaluator({
      functions: { DOUBLE: (args) => toNumber(args[0]!) * 2 },
      now: () => new Date(2024, 0, 15),
    });
    ev.registerFunction('GREET', (args) => `Hello ${toText(args[0]!)}`);
    ev.evaluate(wb);
    expect(ws.getValue('F1')).toMatchObject({ result: 4 });
    expect(ws.getValue('F2')).toMatchObject({ result: 'Hello Ann' });
  });
});
