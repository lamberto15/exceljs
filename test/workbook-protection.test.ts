/// <reference types="node" />
// Document properties, default-font metrics, protectSheet.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { fontMetrics } from '../src/core/styles.js';
import {
  addImage,
  fitImage,
  placeImage,
  protectSheet,
  readXlsx,
  Workbook,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('document properties', () => {
  it('are written to core.xml / app.xml and read back; exceljs sees them', async () => {
    const wb = new Workbook();
    wb.addWorksheet('S').setCell('A1', 1);
    wb.creator = 'me';
    wb.title = 'Q3 <report> & more';
    wb.subject = 'Sales';
    wb.keywords = 'sales, q3';
    wb.description = 'Quarterly numbers';
    wb.category = 'Finance';
    wb.contentStatus = 'Draft';
    wb.company = 'ACME & Co';
    wb.manager = 'Boss';
    const buf = writeXlsx(wb);
    const core = dec.decode(unzipSync(buf).get('docProps/core.xml')!);
    expect(core).toContain('<dc:title>Q3 &lt;report&gt; &amp; more</dc:title>');
    expect(core).toContain('<cp:keywords>sales, q3</cp:keywords>');
    const app = dec.decode(unzipSync(buf).get('docProps/app.xml')!);
    expect(app).toContain('<Company>ACME &amp; Co</Company>');
    expect(app).toContain('<Manager>Boss</Manager>');
    const back = readXlsx(buf);
    expect([
      back.title,
      back.subject,
      back.keywords,
      back.description,
      back.category,
      back.contentStatus,
    ]).toEqual([
      'Q3 <report> & more',
      'Sales',
      'sales, q3',
      'Quarterly numbers',
      'Finance',
      'Draft',
    ]);
    expect([back.company, back.manager, back.creator]).toEqual(['ACME & Co', 'Boss', 'me']);
    const ex = await oracle(buf);
    expect(ex.title).toBe('Q3 <report> & more');
    expect(ex.company).toBe('ACME & Co');
  });

  it('nothing set: no empty elements', () => {
    const wb = new Workbook();
    wb.addWorksheet('S');
    const files = unzipSync(writeXlsx(wb));
    expect(dec.decode(files.get('docProps/core.xml')!)).not.toContain('dc:title');
    expect(dec.decode(files.get('docProps/app.xml')!)).not.toContain('<Company>');
  });
});

describe('default font metrics', () => {
  it('Calibri 11 = 7 px / 15 pt; other fonts scale pictures and the default row height', () => {
    expect(fontMetrics({ name: 'Calibri', size: 11 })).toEqual({ mdw: 7, rowPt: 15 });
    expect(fontMetrics({ name: 'Arial', size: 10 })).toEqual({ mdw: 7, rowPt: 12.75 });
    expect(fontMetrics({ name: 'Arial', size: 20 }).mdw).toBe(14);
    const wb = new Workbook({ defaultFont: { name: 'Arial', size: 20 } });
    const ws = wb.addWorksheet('S');
    const xml = dec.decode(unzipSync(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!);
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="25.5"');
    const id = addImage(wb, {
      data: new Uint8Array([
        0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 10,
      ]),
      extension: 'png',
    });
    // a default column is 9.140625 units x 14 px = 128 px wide: half of it = 64 px
    const p = placeImage(ws, id, { tl: { col: 0.5, row: 0 }, ext: { width: 5, height: 5 } });
    expect(p.anchor).toMatchObject({ from: { col: 0, colOff: 64 } });
    expect(fitImage(ws, id, 'A1').anchor).toMatchObject({ ext: { width: 10, height: 10 } });
  });
});

describe('protectSheet', () => {
  it('unlocks editable ranges (cells, whole columns), protects with allowed actions', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRows([
      ['Name', 'Qty'],
      ['a', 1],
    ]);
    ws.setCell('C2', 5, { font: { bold: true } });
    protectSheet(ws, {
      password: 'pw',
      editable: ['B2:B10', 'C:C'],
      allow: { sort: true, autoFilter: true },
    });
    expect(ws.getStyle('A2').protection).toBeUndefined(); // locked (Excel default)
    expect(ws.getStyle('B2').protection).toMatchObject({ locked: false });
    expect(ws.getStyle('B10').protection).toMatchObject({ locked: false });
    expect(ws.getStyle('C2')).toMatchObject({
      font: { bold: true },
      protection: { locked: false },
    });
    expect(ws.getStyle('C500').protection).toMatchObject({ locked: false }); // column style
    expect(ws.protection).toMatchObject({
      sheet: true,
      sort: true,
      autoFilter: true,
      password: 'pw',
    });
    const buf = writeXlsx(wb);
    const xml = dec.decode(unzipSync(buf).get('xl/worksheets/sheet1.xml')!);
    expect(xml).toMatch(/<sheetProtection [^>]*sheet="1"[^>]*sort="0"[^>]*autoFilter="0"/);
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.getStyle('C700').protection).toMatchObject({ locked: false });
    expect(back.protection).toMatchObject({
      algorithmName: 'SHA-512',
      sort: true,
      autoFilter: true,
    });
    await oracle(buf);
  });
});
