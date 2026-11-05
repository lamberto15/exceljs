/// <reference types="node" />
// Streaming writer backpressure (#47) and bug sweep 2 (#49): reported ExcelJS streaming and
// file-corruption cases checked against turbosheet. `exceljs#N` names the upstream report.
// .xlsm round trip (exceljs#1196): see preserve.test.ts "macro-enabled workbook".

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addImage,
  addNote,
  addTable,
  createXlsxWriter,
  placeImage,
  readXlsx,
  Workbook,
  writeXlsx,
  type XlsxRow,
  xlsxRows,
} from '../src/index.js';

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

function concat(chunks: Uint8Array[]): Uint8Array {
  let n = 0;
  for (const c of chunks) n += c.length;
  const out = new Uint8Array(n);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

describe('streaming writer backpressure (#47)', () => {
  it('addRows waits for a slow sink: the undelivered backlog stays bounded', async () => {
    const chunks: Uint8Array[] = [];
    let pending = 0;
    let maxPending = 0;
    const w = createXlsxWriter({
      flushRows: 500,
      sink: async (c) => {
        pending++;
        maxPending = Math.max(maxPending, pending);
        await new Promise((r) => setTimeout(r, 1));
        chunks.push(c);
        pending--;
      },
    });
    const s = w.addWorksheet('S');
    let produced = 0;
    async function* rows() {
      for (let i = 0; i < 20_000; i++) {
        produced++;
        yield [i, `row ${i}`, i * 1.5];
      }
    }
    await s.addRows(rows());
    await w.end();
    expect(produced).toBe(20_000);
    expect(maxPending).toBeLessThanOrEqual(2); // sequential delivery, producer waited
    const back = readXlsx(concat(chunks)).getWorksheet('S')!;
    expect(back.getValue(20_000, 2)).toBe('row 19999');
  });

  it('the streaming sheet never accumulates rows (constant memory, exceljs#2916)', async () => {
    const w = createXlsxWriter({ flushRows: 1000, sink: () => {} });
    const s = w.addWorksheet('S');
    for (let i = 0; i < 50_000; i++) s.addRow([i, 'x']);
    expect(s.worksheet.rowData.length).toBe(0); // the metadata sheet holds no cells
    await w.end();
  });
});

describe('bug sweep 2 (exceljs reports)', () => {
  it('exceljs#1709: pictures in the streaming writer', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    placeImage(
      s.worksheet,
      addImage(s.worksheet.workbook, { data: PNG, extension: 'png' }),
      'C2:D4',
    );
    s.addRow(['with a picture']);
    await w.end();
    const buf = concat(chunks);
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.images!.length).toBe(1);
    const ex = await oracle(buf);
    expect(ex.getWorksheet('S')!.getImages().length).toBe(1);
  });

  it('exceljs#1693: number formats in the streaming writer', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    s.addRow([1234.5], { numFmt: '#,##0.00' });
    s.addRow([0.25], { numFmt: '0%' });
    await w.end();
    const back = readXlsx(concat(chunks)).getWorksheet('S')!;
    expect(back.getStyle('A1').numFmt).toBe('#,##0.00');
    expect(back.getStyle('A2').numFmt).toBe('0%');
  });

  it('exceljs#2084 / #2316: multi-byte text split across stream chunks', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('数据');
    const texts = ['中文字符测试', '日本語のテキスト', '한국어 텍스트', 'emoji 😀🎉', 'Ünïcödé'];
    for (let i = 0; i < 300; i++) ws.addRow([texts[i % texts.length]!, i]);
    const buf = writeXlsx(wb, { compression: 0 });
    async function* tiny(): AsyncGenerator<Uint8Array> {
      for (let i = 0; i < buf.length; i += 7) yield buf.subarray(i, i + 7); // splits UTF-8 sequences
    }
    const rows: XlsxRow[] = [];
    for await (const r of xlsxRows(tiny())) rows.push(r);
    expect(rows.length).toBe(300);
    expect(rows[0]!.sheet).toBe('数据');
    for (let i = 0; i < 300; i++) expect(rows[i]!.cells[0]).toBe(texts[i % texts.length]);
  });

  it('exceljs#2234: merging over a table throws instead of corrupting the file', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addTable(ws, { name: 'T', columns: ['a', 'b'], rows: [[1, 2]] });
    ws.mergeCells('A2:B2');
    expect(() => writeXlsx(wb)).toThrow(/overlaps merged cells/);
  });

  it('exceljs#2771: background picture and notes on one sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.backgroundImage = addImage(wb, { data: PNG, extension: 'png' });
    addNote(ws, 'A1', 'a note');
    addNote(ws, 'B3', 'another');
    const buf = writeXlsx(wb);
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.backgroundImage).toBeDefined();
    expect(back.comments!.length).toBe(2);
    await oracle(buf);
  });
});
