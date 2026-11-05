// #74: turbosheet's own output is stream-readable (metadata parts before worksheets), and
// xlsxRows applies the byte limits only to parts it has to buffer.
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/workbook.js';
import { xlsxRows } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { ZipOut } from '../src/xlsx/write/zipout.js';
import { unzipLazy, unzipStream, zipSync } from '../src/zip/index.js';

const META = [
  '[Content_Types].xml',
  '_rels/.rels',
  'xl/workbook.xml',
  'xl/_rels/workbook.xml.rels',
  'xl/styles.xml',
];

function book(rows: number): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  for (let i = 0; i < rows; i++)
    ws.addRow([i, `row ${i}`, `t${i % 40}`, new Date(Date.UTC(2020, 0, 1 + (i % 365)))]);
  wb.addWorksheet('T').addRow(['second']);
  return wb;
}

async function bytes(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(rs).arrayBuffer());
}

/** Feeds `data` in 16 KiB chunks and records how many bytes were handed out. */
function source(data: Uint8Array): { it: AsyncIterable<Uint8Array>; fed: () => number } {
  let fed = 0;
  return {
    fed: () => fed,
    it: {
      async *[Symbol.asyncIterator]() {
        for (let i = 0; i < data.length; i += 16384) {
          const c = data.subarray(i, i + 16384);
          fed += c.length;
          yield c;
        }
      },
    },
  };
}

describe('metadata before worksheets (#74)', () => {
  const outputs = async (wb: Workbook) => ({
    'writeXlsx balanced': writeXlsx(wb),
    'writeXlsx low-memory': writeXlsx(wb, { profile: 'low-memory' }),
    'xlsxStream balanced': await bytes(xlsxStream(wb)),
    'xlsxStream low-memory': await bytes(xlsxStream(wb, { profile: 'low-memory' })),
  });

  it('writeXlsx / xlsxStream store metadata parts first, in Excel order', async () => {
    for (const [name, buf] of Object.entries(await outputs(book(10)))) {
      const keys = [...unzipLazy(buf).keys()];
      const shared = name.includes('balanced');
      expect(keys.slice(0, META.length + (shared ? 1 : 0)), name).toEqual(
        shared ? [...META, 'xl/sharedStrings.xml'] : META,
      );
      expect(keys.indexOf('xl/worksheets/sheet1.xml')).toBeGreaterThan(
        keys.indexOf('xl/styles.xml'),
      );
    }
  });

  it('xlsxRows emits rows before the whole file has been read', async () => {
    for (const [name, buf] of Object.entries(await outputs(book(40_000)))) {
      const src = source(buf);
      let firstAt = -1;
      let n = 0;
      let last: unknown[] = [];
      for await (const r of xlsxRows(src.it)) {
        if (firstAt < 0) firstAt = src.fed();
        if (r.sheet === 'S') last = r.values;
        n++;
      }
      expect(n, name).toBe(40_001);
      expect(firstAt, name).toBeLessThan(buf.length / 2);
      expect(last[1]).toBe('row 39999');
      expect(last[3]).toBeInstanceOf(Date);
    }
  });

  it('ZipOut refuses to end with a reserved part unwritten', () => {
    const z = new ZipOut(() => {}, 6);
    z.reserve(['a.xml']);
    z.add('b.xml', 'x');
    expect(() => z.end()).toThrow(/a\.xml/);
  });
});

async function* once(d: Uint8Array): AsyncGenerator<Uint8Array> {
  yield d;
}

describe('xlsxRows byte limits apply only to buffered parts (#74)', () => {
  const limits = { maxEntryBytes: 64 * 1024, maxTotalBytes: 128 * 1024 };

  it('a streamed sheet may exceed maxEntryBytes / maxTotalBytes', async () => {
    // inline strings: sheet1.xml is well over 128 KiB, every metadata part is small
    const buf = writeXlsx(book(20_000), { profile: 'low-memory' });
    let n = 0;
    for await (const _ of xlsxRows(buf, { limits })) n++;
    expect(n).toBe(20_001);
  });

  it('a buffered sheet (worksheets before metadata) is still capped', async () => {
    const parts = [...unzipLazy(writeXlsx(book(20_000))).values()];
    const sheetsFirst = zipSync(
      [
        ...parts.filter((p) => p.path.includes('worksheets/')),
        ...parts.filter((p) => !p.path.includes('worksheets/')),
      ].map((p) => ({ path: p.path, data: p.bytes() })),
      6,
    );
    await expect(async () => {
      for await (const _ of xlsxRows(sheetsFirst, { limits })) {
      }
    }).rejects.toThrow(/too large|total/);
    let n = 0;
    for await (const _ of xlsxRows(sheetsFirst)) n++;
    expect(n).toBe(20_001);
  });

  it('createXlsxWriter output (sheets first) streams with default limits, buffering', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    for (let i = 0; i < 1000; i++) s.addRow([i]);
    await w.end();
    let n = 0;
    for await (const _ of xlsxRows(await bytes(new Blob(chunks as BlobPart[]).stream()))) n++;
    expect(n).toBe(1000);
  });

  it('unbounded() exempts only its own entry, also when entries share a slice', async () => {
    const zip = zipSync(
      [
        { path: 'small.xml', data: 'tiny' },
        { path: 'big.xml', data: 'x'.repeat(200_000) }, // compresses into the same 16 KiB slice
      ],
      6,
    );
    const run = async () => {
      for await (const e of unzipStream(once(zip), { maxEntryBytes: 100_000 })) {
        if (e.path === 'small.xml') e.unbounded();
        for await (const _ of e.chunks) {
        }
      }
    };
    await expect(run()).rejects.toThrow(/big\.xml/);
    for await (const e of unzipStream(once(zip), { maxEntryBytes: 100_000 })) {
      e.unbounded();
      let n = 0;
      for await (const c of e.chunks) n += c.length;
      expect(n).toBeGreaterThan(0);
    }
  });

  it('the ratio limit still applies to streamed sheets', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 30_000; i++) ws.addRow(['a', 'a', 'a', 'a', 'a', 'a']);
    const buf = writeXlsx(wb, { compression: 9 });
    await expect(async () => {
      for await (const _ of xlsxRows(buf, { limits: { maxCompressionRatio: 5 } })) {
      }
    }).rejects.toThrow(/ratio/);
  });
});
