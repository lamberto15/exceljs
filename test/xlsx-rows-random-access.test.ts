// #86: xlsxRows reads random-access sources (bytes, Blob) by the central directory, metadata
// first, so files that store worksheets before styles / shared strings (Excel, ExcelJS,
// createXlsxWriter) stream at constant memory, with only the ratio limit on streamed sheets.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/workbook.js';
import { readXlsx, xlsxRows } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx } from '../src/xlsx/write/index.js';
import { unzipBlob, unzipLazy, ZipLimitError } from '../src/zip/index.js';

const EXCEL = new Uint8Array(
  readFileSync(fileURLToPath(new URL('./fixtures/excel-saved-sheets-first.xlsx', import.meta.url))),
);

async function* once(d: Uint8Array): AsyncGenerator<Uint8Array> {
  yield d;
}

async function createWriterFile(rows: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  const s = w.addWorksheet('Data');
  for (let i = 0; i < rows; i++)
    s.addRow([i, `name ${i}`, new Date(Date.UTC(2021, 0, 1 + (i % 400)))]);
  s.mergeCells('E1:F2');
  await w.end();
  return new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
}

async function exceljsFile(rows: number): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let i = 0; i < rows; i++)
    ws.addRow([i, `name ${i}`, new Date(Date.UTC(2021, 0, 1 + (i % 400)))]);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** Sizes of the sheet part and of the largest non-sheet part. */
function sizes(buf: Uint8Array): { sheet: number; meta: number; order: string[] } {
  const m = unzipLazy(buf);
  let sheet = 0;
  let meta = 0;
  for (const e of m.values()) {
    if (e.path.includes('worksheets/')) sheet = Math.max(sheet, e.size);
    else meta = Math.max(meta, e.size);
  }
  return { sheet, meta, order: [...m.keys()] };
}

async function collect(src: Parameters<typeof xlsxRows>[0], opts?: Parameters<typeof xlsxRows>[1]) {
  const rows: { sheet: string; row: number; values: unknown[] }[] = [];
  for await (const r of xlsxRows(src, opts))
    rows.push({ sheet: r.sheet, row: r.row, values: r.values });
  return rows;
}

const cases: [string, () => Promise<Uint8Array>][] = [
  ['Excel-saved', async () => EXCEL],
  ['createXlsxWriter', () => createWriterFile(3000)],
  ['ExcelJS', () => exceljsFile(3000)],
];

describe('xlsxRows random-access sources (#86)', () => {
  it.each(cases)('%s files store worksheets before styles.xml', async (_, make) => {
    const { order } = sizes(await make());
    expect(order.findIndex((p) => p.includes('worksheets/'))).toBeLessThan(
      order.indexOf('xl/styles.xml'),
    );
  });

  it.each(cases)(
    '%s: bytes and Blob stream sheets past maxEntryBytes; streams buffer',
    async (_, make) => {
      const buf = await make();
      const { sheet, meta } = sizes(buf);
      expect(sheet).toBeGreaterThan(meta * 2);
      // every metadata part fits, the sheet does not: only a streamed sheet can pass
      const limits = { maxEntryBytes: meta + 1, maxTotalBytes: meta * 4 + 1 };
      const expected = await collect(buf);
      expect(expected.length).toBeGreaterThan(1000);
      expect(await collect(buf, { limits })).toEqual(expected);
      expect(await collect(new Blob([buf as BlobPart]), { limits })).toEqual(expected);
      await expect(collect(once(buf), { limits })).rejects.toThrow(/too large|total/);
      // same values as the in-memory reader
      const ws = readXlsx(buf).worksheets[0]!;
      const last = expected.filter((r) => r.sheet === ws.name).at(-1)!;
      expect(last.values[1]).toBe(ws.getValue(last.row, 2));
    },
  );

  it('Blob path: sheets filter and onSheetEnd merges', async () => {
    for (const src of [EXCEL, new Blob([EXCEL as BlobPart])]) {
      const ends: { sheet: string; merges: string[] }[] = [];
      const rows = await collect(src, {
        sheets: ['Second'],
        onSheetEnd: (i) => ends.push({ sheet: i.sheet, merges: i.merges }),
      });
      expect(rows).toEqual([{ sheet: 'Second', row: 1, values: ['only', 'row'] }]);
      expect(ends).toEqual([{ sheet: 'Second', merges: [] }]);
      const all: { sheet: string; merges: string[] }[] = [];
      await collect(src, { onSheetEnd: (i) => all.push({ sheet: i.sheet, merges: i.merges }) });
      expect(all).toEqual([
        { sheet: 'Data', merges: ['F1:G2'] },
        { sheet: 'Second', merges: [] },
      ]);
    }
    const rows = await collect(new Blob([EXCEL as BlobPart]), { sheets: ['Data'] });
    expect(rows[0]).toEqual({
      sheet: 'Data',
      row: 1,
      values: ['id', 'name', 'when', 'flag', undefined, 'merged'],
    });
    expect(rows[2]!.values[2]).toBeInstanceOf(Date);
    expect(rows.length).toBe(2001);
  });

  it('metadata parts keep every limit; streamed sheets keep the ratio limit', async () => {
    const buf = EXCEL;
    const sst = unzipLazy(buf).get('xl/sharedStrings.xml')!.size;
    for (const src of [buf, new Blob([buf as BlobPart])]) {
      await expect(collect(src, { limits: { maxEntryBytes: sst - 1 } })).rejects.toThrow(
        /sharedStrings.*too large|too large.*sharedStrings/,
      );
    }
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 30_000; i++) ws.addRow([1, 1, 1, 1, 1, 1]);
    const bomb = writeXlsx(wb, { compression: 9, profile: 'low-memory' });
    for (const src of [bomb, new Blob([bomb as BlobPart])]) {
      await expect(collect(src, { limits: { maxCompressionRatio: 5 } })).rejects.toThrow(/ratio/);
    }
  });

  it('Blob without DecompressionStream falls back to fflate', async () => {
    const g = globalThis as { DecompressionStream?: unknown };
    const saved = g.DecompressionStream;
    const expected = await collect(EXCEL);
    try {
      delete g.DecompressionStream;
      expect(await collect(new Blob([EXCEL as BlobPart]))).toEqual(expected);
    } finally {
      g.DecompressionStream = saved;
    }
  });

  it('unzipBlob matches unzipLazy and rejects non-zip / truncated blobs', async () => {
    const lazy = unzipLazy(EXCEL);
    const blob = await unzipBlob(new Blob([EXCEL as BlobPart]));
    expect([...blob.keys()]).toEqual([...lazy.keys()]);
    for (const [k, e] of blob) {
      const parts: Uint8Array[] = [];
      for await (const c of e.chunks()) parts.push(c);
      expect(new Uint8Array(await new Blob(parts as BlobPart[]).arrayBuffer())).toEqual(
        lazy.get(k)!.bytes(),
      );
    }
    await expect(unzipBlob(new Blob([new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])]))).rejects.toThrow(
      ZipLimitError,
    );
    expect((await unzipBlob(new Blob([]))).size).toBe(0);
    const cut = await unzipBlob(new Blob([EXCEL.subarray(0, EXCEL.length - 30) as BlobPart])).catch(
      (e) => e,
    );
    expect(cut).toBeInstanceOf(ZipLimitError);
  });

  it('bytes without a readable central directory fall back to the stream path', async () => {
    const buf = writeXlsx(readXlsx(EXCEL));
    const cut = buf.subarray(0, buf.length - 40); // central directory gone, entries intact
    const viaStream = await collect(once(cut));
    expect(viaStream.length).toBe(2002);
    expect(await collect(cut)).toEqual(viaStream);
  });
});
