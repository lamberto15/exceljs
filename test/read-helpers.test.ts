/// <reference types="node" />
// formatValue / cellText, column helpers, richTextFromHtml, and the stream-transform recipe.

import { describe, expect, it } from 'vitest';
import { cellText, formatValue } from '../src/formula/index.js';
import {
  columnValues,
  createXlsxWriter,
  eachColumnCell,
  readXlsx,
  richTextFromHtml,
  Workbook,
  writeXlsx,
  xlsxRows,
} from '../src/index.js';

describe('formatValue / cellText', () => {
  it('applies number formats like Excel', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(0.256, '0.0%')).toBe('25.6%');
    expect(formatValue(-3, '0;(0)')).toBe('(3)');
    expect(formatValue(new Date(Date.UTC(2026, 10, 5, 14, 30)), 'yyyy-mm-dd hh:mm')).toBe(
      '2026-11-05 14:30',
    );
    expect(formatValue(true)).toBe('TRUE');
    expect(formatValue({ error: '#N/A' })).toBe('#N/A');
    expect(formatValue({ formula: 'A1', result: 0.5 }, '0%')).toBe('50%');
    expect(formatValue({ richText: [{ text: 'a' }, { text: 'b' }] })).toBe('ab');
    expect(formatValue(null)).toBe('');
    expect(formatValue('text', '@')).toBe('text');
    expect(formatValue(1 / 3)).toBe('0.333333333');
  });

  it('cellText uses the cell style and the workbook date system', () => {
    const wb = new Workbook({ date1904: true });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 0.125, { numFmt: '0.00%' });
    ws.setCell('B1', new Date(Date.UTC(2026, 0, 31)), { numFmt: 'd mmm yyyy' });
    ws.setCell('C1', 42);
    expect(cellText(ws, 'A1')).toBe('12.50%');
    expect(cellText(ws, 1, 2)).toBe('31 Jan 2026');
    expect(cellText(ws, 'C1')).toBe('42');
    expect(cellText(ws, 'D1')).toBe('');
  });
});

describe('column helpers', () => {
  it('columnValues and eachColumnCell (ExcelJS column.values / eachCell)', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('B1', 'h');
    ws.setCell('B3', 3);
    ws.setStyle('B4', { font: { bold: true } }); // style only: skipped
    ws.setCell('C2', 'other column');
    expect(columnValues(ws, 'B')).toEqual([undefined, 'h', undefined, 3]);
    expect(columnValues(ws, 3)[2]).toBe('other column');
    const seen: [unknown, number][] = [];
    eachColumnCell(ws, 'B', (v, r) => seen.push([v, r]));
    expect(seen).toEqual([
      ['h', 1],
      [3, 3],
    ]);
  });
});

describe('richTextFromHtml', () => {
  it('converts simple formatting, colors, breaks and entities', () => {
    const rt = richTextFromHtml(
      '<p>Hello <b>bold <i>both</i></b> &amp; <span style="color:#f00; text-decoration: underline">red</span></p><p>line&nbsp;2<br>x<sup>2</sup></p>',
    );
    expect(rt.richText).toEqual([
      { text: 'Hello ' },
      { text: 'bold ', font: { bold: true } },
      { text: 'both', font: { bold: true, italic: true } },
      { text: ' & ' },
      { text: 'red', font: { color: { argb: 'FFFF0000' }, underline: 'single' } },
      { text: '\nline 2\nx' },
      { text: '2', font: { vertAlign: 'superscript' } },
    ]);
    expect(
      richTextFromHtml('<font color="blue">b</font><!-- c --><unknown>t</unknown>').richText,
    ).toEqual([{ text: 'b', font: { color: { argb: 'FF0000FF' } } }, { text: 't' }]);
    expect(richTextFromHtml('plain').richText).toEqual([{ text: 'plain' }]);
  });

  it('round-trips through a file', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', richTextFromHtml('<strong>x</strong> y'));
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!.getValue('A1');
    expect(back).toMatchObject({ richText: [{ text: 'x', font: { bold: true } }, { text: ' y' }] });
  });
});

describe('stream transform recipe (docs/features.md)', () => {
  it('reads rows, transforms them and writes a new file at constant memory', async () => {
    const src = new Workbook();
    const s = src.addWorksheet('Orders');
    s.addRow(['item', 'qty', 'price']);
    for (let i = 1; i <= 1000; i++) s.addRow([`item ${i}`, i % 7, 1.5]);
    const input = writeXlsx(src);

    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    let out = w.addWorksheet('Orders');
    let current = 'Orders';
    for await (const row of xlsxRows(input)) {
      if (row.sheet !== current) {
        current = row.sheet;
        out = w.addWorksheet(row.sheet);
      }
      const [item, qty, price] = row.values;
      if (row.row === 1) out.addRow([item, qty, price, 'total']);
      else if (qty) out.addRow([item, qty, price, (qty as number) * (price as number)]); // drop qty 0
      if (row.row % 500 === 0) await out.flush(); // backpressure
    }
    await w.end();
    let n = 0;
    for (const c of chunks) n += c.length;
    const buf = new Uint8Array(n);
    let o = 0;
    for (const c of chunks) {
      buf.set(c, o);
      o += c.length;
    }
    const back = readXlsx(buf).getWorksheet('Orders')!;
    expect(back.getRow(1).values).toEqual([undefined, 'item', 'qty', 'price', 'total']);
    expect(back.getValue(2, 4)).toBe(1.5);
    expect(back.rowCount).toBe(1 + 1000 - Math.floor(1000 / 7));
  });
});
