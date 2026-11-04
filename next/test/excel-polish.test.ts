/// <reference types="node" />
// Excel-like defaults and sheet XML polish: default font, hyperlink look, wrap for line breaks,
// ignored errors, page breaks, sheet-wide style.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  autoFitColumns,
  deleteRows,
  insertRows,
  readXlsx,
  setSheetStyle,
  Workbook,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const part = (buf: Uint8Array, name: string): string => dec.decode(unzipSync(buf).get(name)!);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('default font', () => {
  it('new Workbook({ defaultFont }) changes font 0 and completes partial fonts', () => {
    const wb = new Workbook({ defaultFont: { name: 'Arial', size: 10 } });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 'x', { font: { bold: true } });
    expect(wb.styles.defaultFont).toMatchObject({ name: 'Arial', size: 10 });
    expect(wb.styles.defaultFont.scheme).toBeUndefined();
    expect(ws.getStyle('A1').font).toMatchObject({ name: 'Arial', size: 10, bold: true });
    const buf = writeXlsx(wb);
    expect(part(buf, 'xl/styles.xml')).toMatch(
      /<fonts count="\d+"><font><sz val="10"\/>.*?<name val="Arial"\/>/,
    );
    expect(readXlsx(buf).styles.defaultFont).toMatchObject({ name: 'Arial', size: 10 });
  });
});

describe('cell defaults', () => {
  it('hyperlinks without a style get the hyperlink look; explicit styles win', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('A1', { text: 'site', hyperlink: 'https://example.com' });
    ws.addRow([{ text: 'row', hyperlink: 'https://example.com' }]);
    ws.setCell(
      'A3',
      { text: 'red', hyperlink: 'https://example.com' },
      { font: { color: { argb: 'FFFF0000' } } },
    );
    expect(ws.getStyle('A1').font).toMatchObject({ underline: 'single', color: { theme: 10 } });
    expect(ws.getStyle('A2').font).toMatchObject({ underline: 'single' });
    expect(ws.getStyle('A3').font?.underline).toBeUndefined();
  });

  it('text with line breaks gets wrapText (merged into its style)', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('A1', 'one\ntwo', { font: { bold: true } });
    ws.addRow(['plain', 'a\nb']);
    expect(ws.getStyle('A1')).toMatchObject({
      font: { bold: true },
      alignment: { wrapText: true },
    });
    expect(ws.getStyle('A2').alignment).toBeUndefined();
    expect(ws.getStyle('B2').alignment).toEqual({ wrapText: true });
  });
});

describe('ignored errors and page breaks', () => {
  it('write, read back, follow structure edits, exceljs loads', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 40; r++) ws.addRow([String(10000 + r), r]);
    ws.ignoredErrors = [
      { sqref: 'A1:A40', numberStoredAsText: true },
      { sqref: 'B5', formula: true, evalError: true },
    ];
    ws.rowBreaks = [20, 10, 10];
    ws.colBreaks = [1];
    const buf = writeXlsx(wb);
    const xml = part(buf, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain(
      '<rowBreaks count="2" manualBreakCount="2"><brk id="10" max="16383" man="1"/><brk id="20" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks><ignoredErrors><ignoredError sqref="A1:A40" numberStoredAsText="1"/><ignoredError sqref="B5" evalError="1" formula="1"/></ignoredErrors>',
    );
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.rowBreaks).toEqual([10, 20]);
    expect(back.colBreaks).toEqual([1]);
    expect(back.ignoredErrors).toEqual(ws.ignoredErrors);
    await oracle(buf);

    insertRows(ws, 5, 2);
    expect(ws.rowBreaks).toEqual([22, 12, 12]);
    expect(ws.ignoredErrors.map((e) => e.sqref)).toEqual(['A1:A42', 'B7']);
    deleteRows(ws, 7);
    expect(ws.ignoredErrors.map((e) => e.sqref)).toEqual(['A1:A41']);
    deleteRows(ws, 11);
    expect(ws.rowBreaks).toEqual([20]);
  });
});

describe('setSheetStyle', () => {
  it('styles every cell, row and column; one <col> range; later cells inherit', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 1, { font: { bold: true } });
    ws.setCell('B2', new Date(Date.UTC(2025, 0, 1)));
    ws.setRowProps(3, {
      style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } },
    });
    setSheetStyle(ws, { protection: { locked: false } });
    expect(ws.getStyle('A1')).toMatchObject({
      font: { bold: true },
      protection: { locked: false },
    });
    expect(ws.getStyle('B2').numFmt).toBe('yyyy-mm-dd');
    expect(ws.getStyle('B2').protection).toMatchObject({ locked: false });
    expect(ws.getRowProps(3).style).toMatchObject({ protection: { locked: false } });
    ws.setCell('Z100', 'later');
    expect(ws.getStyle('Z100').protection).toMatchObject({ locked: false });
    expect(ws.getStyle('XFD1048576').protection).toMatchObject({ locked: false });
    const buf = writeXlsx(wb);
    const cols = part(buf, 'xl/worksheets/sheet1.xml').match(/<cols>.*?<\/cols>/)![0];
    expect(cols).toMatch(/^<cols><col min="1" max="16384" width="[\d.]+" style="\d+"\/><\/cols>$/);
    expect(readXlsx(buf).getWorksheet('S')!.getStyle('C5').protection).toMatchObject({
      locked: false,
    });
    expect(autoFitColumns(ws)[0]).toBeGreaterThan(0); // still works with every column defined
  });
});
