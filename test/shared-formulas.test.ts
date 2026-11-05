// Shared formulas on write (#101): fillFormula ranges, `{ sharedFormula }` values and shared
// formulas read from a file are written as <f t="shared"> masters + followers. Only cells that
// still hold the master's formula moved by their offset are followers; everything else in the
// range is written plain, so the written file always means what the model holds.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  type CellValue,
  createXlsxWriter,
  fillFormula,
  insertRows,
  readXlsx,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';
import { fixture } from './exceljs-parity/helpers.js';

const dec = new TextDecoder();
const sheetXml = (buf: Uint8Array, n = 1): string =>
  dec.decode(unzipSync(buf).get(`xl/worksheets/sheet${n}.xml`)!);
const fcount = (xml: string, re: RegExp): number => xml.match(re)?.length ?? 0;

/** Every cell value of the sheet, keyed by address. */
function cells(ws: Worksheet): Map<string, CellValue> {
  const m = new Map<string, CellValue>();
  ws.eachRow((row, r) => {
    row.eachCell((v, c) => {
      m.set(ws.getCell(r, c).address, v);
    });
  });
  expect(m.size).toBeGreaterThan(0);
  return m;
}

/** Written file reads back to exactly the model's values (ours and exceljs's formulas). */
async function expectSame(wb: Workbook): Promise<Uint8Array> {
  const buf = writeXlsx(wb);
  const back = readXlsx(buf);
  for (let i = 0; i < wb.worksheets.length; i++) {
    expect(cells(back.worksheets[i]!)).toEqual(cells(wb.worksheets[i]!));
  }
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(buf as unknown as ArrayBuffer);
  for (const [addr, v] of cells(wb.worksheets[0]!)) {
    if (v && typeof v === 'object' && 'formula' in v) {
      expect(x.worksheets[0]!.getCell(addr).formula, addr).toBe(v.formula);
    }
  }
  return buf;
}

describe('shared formulas on write', () => {
  it('fillFormula down a column: one master, followers without text', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    for (let r = 1; r <= 100; r++) ws.setCell(r, 1, r);
    fillFormula(ws, 'B1:B100', 'A1*2+$A$1', (r) => r * 2 + 1);
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<c r="B1"><f t="shared" ref="B1:B100" si="0">A1*2+$A$1</f><v>3</v></c>');
    expect(xml).toContain('<c r="B100"><f t="shared" si="0"/><v>201</v></c>');
    expect(fcount(xml, /<f t="shared" si="0"\/>/g)).toBe(99);
    expect(fcount(xml, /<f>/g)).toBe(0);
  });

  it('a 2D block and several groups get distinct si in row order', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    fillFormula(ws, 'C5:D6', 'ROW()*10+COLUMN()');
    fillFormula(ws, 'A1:A3', 'ROW()');
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<f t="shared" ref="A1:A3" si="0">ROW()</f>');
    expect(xml).toContain('<f t="shared" ref="C5:D6" si="1">ROW()*10+COLUMN()</f>');
    expect(fcount(xml, /<f t="shared" si="1"\/>/g)).toBe(3);
  });

  it('overwritten or edited cells inside the range are written plain', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    fillFormula(ws, 'A1:A6', 'B1+1');
    ws.setCell('A3', 'text');
    ws.setCell('A4', { formula: 'B4+2' });
    // the same text a follower would have, at the wrong position: not a follower
    ws.setCell('A5', ws.getValue('A2'));
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<c r="A2"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="A4"><f>B4+2</f></c>');
    expect(xml).toContain('<c r="A5"><f>B2+1</f></c>');
    expect(xml).toContain('<c r="A6"><f t="shared" si="0"/></c>');
  });

  it('followers of a replaced master are written plain', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    fillFormula(ws, 'A1:A3', 'B1');
    ws.setCell('A1', { formula: 'C1' });
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<c r="A2"><f>B2</f></c>');
    expect(xml).toContain('<c r="A3"><f>B3</f></c>');
    ws.setCell('A1', 5);
    expect(sheetXml(await expectSame(wb))).not.toContain('t="shared"');
  });

  it('a later fill over part of an earlier one wins; the earlier group is dropped', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    fillFormula(ws, 'A1:A10', 'B1');
    fillFormula(ws, 'A5:A6', 'C5');
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<f t="shared" ref="A5:A6" si="0">C5</f>');
    expect(xml).toContain('<c r="A2"><f>B2</f></c>');
  });

  it('array and dynamic-array formulas are never shared', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    fillFormula(ws, 'A1:A3', 'SEQUENCE(2)');
    const xml = sheetXml(writeXlsx(wb));
    expect(xml).not.toContain('t="shared"');
    expect(fcount(xml, /cm="1"/g)).toBe(3);
  });

  it('structural edits keep the written file correct', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    for (let r = 1; r <= 6; r++) ws.setCell(r, 1, r);
    fillFormula(ws, 'B1:B6', 'A1+1');
    insertRows(ws, 3, 2);
    await expectSame(wb);
  });

  it('shared formulas read from a file are written back shared', async () => {
    const wb = readXlsx(fixture('fibonacci.xlsx'));
    const before = sheetXml(writeXlsx(readXlsx(fixture('fibonacci.xlsx'))));
    expect(before).toContain('t="shared"');
    const xml = sheetXml(await expectSame(wb));
    expect(fcount(xml, /t="shared"/g)).toBe(fcount(before, /t="shared"/g));
  });

  it('{ sharedFormula } followers above / left of the master are written plain', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    ws.setCell('B2', { formula: 'A1*2' });
    ws.setCell('B3', { sharedFormula: 'B2' });
    ws.setCell('A2', { sharedFormula: 'B2' }); // left of the master: #REF!-free but plain
    const xml = sheetXml(await expectSame(wb));
    expect(xml).toContain('<f t="shared" ref="B2:B3" si="0">A1*2</f>');
  });

  it('streaming writer: { sharedFormula } values within a batch', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c), flushRows: 50 });
    const s = w.addWorksheet('s');
    s.addRow([1, { formula: 'A1*2' }]);
    for (let r = 2; r <= 40; r++) s.addRow([r, { sharedFormula: 'B1' }]);
    // a second sheet's groups get their own si numbering
    const t = w.addWorksheet('t');
    t.addRow([{ formula: 'ROW()' }]);
    t.addRow([{ sharedFormula: 'A1' }]);
    await w.end();
    const buf = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of chunks) {
      buf.set(c, o);
      o += c.length;
    }
    const xml = sheetXml(buf);
    expect(xml).toContain('<f t="shared" ref="B1:B40" si="0">A1*2</f>');
    expect(fcount(xml, /<f t="shared" si="0"\/>/g)).toBe(39);
    expect(sheetXml(buf, 2)).toContain('<f t="shared" ref="A1:A2" si="0">ROW()</f>');
    const back = readXlsx(buf).worksheets[0]!;
    expect(back.getValue('B40')).toEqual({ formula: 'A40*2' });
  });
});
