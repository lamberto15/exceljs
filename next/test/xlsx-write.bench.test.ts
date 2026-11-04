/// <reference types="node" />
// Write benchmark: ours vs exceljs. Gated: RUN_BENCH=1 NODE_OPTIONS=--expose-gc npx vitest run test/xlsx-write.bench.test.ts
// 100k rows x 20 cols: 8 numbers, 6 short strings, 3 dates, 3 styled numbers (every 10th row bold).
// "heap" = peak sampled heapUsed above the pre-case baseline; "rss" = peak sampled RSS above baseline
// (RSS rarely shrinks, so later cases are under-reported; run with -t to isolate one case).

import { Writable } from 'node:stream';
import ExcelJS from 'exceljs';
import { describe, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import type { Profile } from '../src/core/types.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';

// NO_NATIVE=1 hides CompressionStream so the async writers use fflate (A/B comparison).
if (process.env.NO_NATIVE) delete (globalThis as { CompressionStream?: unknown }).CompressionStream;

const ROWS = Number(process.env.BENCH_ROWS ?? 100_000);
const COLS = 20;
const WORDS = Array.from({ length: 500 }, (_, i) => `word${i}`);
const BASE_DATE = Date.UTC(2020, 0, 1);
const BOLD = { font: { bold: true } };
const MONEY = { numFmt: '#,##0.00' };

function makeRow(r: number, out: unknown[]): unknown[] {
  for (let c = 0; c < 8; c++) out[c] = r * 8 + c + 0.25;
  for (let c = 8; c < 14; c++) out[c] = c === 13 ? `u${r}` : WORDS[(r * 7 + c) % WORDS.length];
  for (let c = 14; c < 17; c++) out[c] = new Date(BASE_DATE + (r + c) * 86_400_000);
  for (let c = 17; c < COLS; c++) out[c] = (r % 1000) * 1.5;
  return out;
}

const g = (globalThis as { gc?: () => void }).gc;
let peakHeap = 0;
let peakRss = 0;
function sample(): void {
  const m = process.memoryUsage();
  if (m.heapUsed > peakHeap) peakHeap = m.heapUsed;
  if (m.rss > peakRss) peakRss = m.rss;
}

interface Result {
  name: string;
  ms: number;
  bytes: number;
  heapMB: number;
  rssMB: number;
}
const results: Result[] = [];

async function measure(name: string, fn: () => Promise<number> | number): Promise<void> {
  g?.();
  g?.();
  const base = process.memoryUsage();
  peakHeap = base.heapUsed;
  peakRss = base.rss;
  const timer = setInterval(sample, 5);
  const t0 = performance.now();
  const bytes = await fn();
  const ms = performance.now() - t0;
  clearInterval(timer);
  sample();
  results.push({
    name,
    ms: Math.round(ms),
    bytes,
    heapMB: Math.round((peakHeap - base.heapUsed) / 1048576),
    rssMB: Math.round((peakRss - base.rss) / 1048576),
  });
}

function oursBuffered(profile: Profile): number {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.columns = Array.from({ length: COLS }, (_, c) => ({
    width: 12,
    style: c >= 17 ? MONEY : undefined,
  }));
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    // biome-ignore lint/suspicious/noExplicitAny: bench data
    ws.addRow(makeRow(r, row) as any, r % 10 === 0 ? BOLD : undefined);
    if (r % 5000 === 0) sample();
  }
  sample();
  const buf = writeXlsx(wb, { profile });
  sample();
  return buf.length;
}

function buildWorkbook(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.columns = Array.from({ length: COLS }, (_, c) => ({
    width: 12,
    style: c >= 17 ? MONEY : undefined,
  }));
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    // biome-ignore lint/suspicious/noExplicitAny: bench data
    ws.addRow(makeRow(r, row) as any, r % 10 === 0 ? BOLD : undefined);
    if (r % 5000 === 0) sample();
  }
  return wb;
}

async function oursXlsxStream(profile: Profile): Promise<number> {
  const wb = buildWorkbook();
  let bytes = 0;
  const reader = xlsxStream(wb, { profile }).getReader();
  for (;;) {
    const r = await reader.read();
    if (r.done) break;
    bytes += r.value.length;
    sample();
  }
  return bytes;
}

async function oursStreamingAsync(profile: Profile): Promise<number> {
  let bytes = 0;
  const w = createXlsxWriter({ profile, sink: (c) => void (bytes += c.length) });
  const s = w.addWorksheet('Data', {
    columns: Array.from({ length: COLS }, (_, c) => ({
      width: 12,
      style: c >= 17 ? MONEY : undefined,
    })),
  });
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    // biome-ignore lint/suspicious/noExplicitAny: bench data
    s.addRow(makeRow(r, row) as any, r % 10 === 0 ? BOLD : undefined);
    if (r % 1000 === 999) {
      sample();
      await s.flush();
    }
  }
  await s.commit();
  await w.end();
  return bytes;
}

async function oursStreaming(profile: Profile): Promise<number> {
  let bytes = 0;
  const w = createXlsxWriter({ profile, sink: (c) => void (bytes += c.length) });
  const s = w.addWorksheet('Data', {
    columns: Array.from({ length: COLS }, (_, c) => ({
      width: 12,
      style: c >= 17 ? MONEY : undefined,
    })),
  });
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    // biome-ignore lint/suspicious/noExplicitAny: bench data
    s.addRow(makeRow(r, row) as any, r % 10 === 0 ? BOLD : undefined);
    if (r % 5000 === 0) sample();
  }
  await s.commit();
  await w.end();
  return bytes;
}

async function exceljsBuffered(): Promise<number> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.columns = Array.from({ length: COLS }, (_, c) => ({ width: 12, style: c >= 17 ? MONEY : {} }));
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    const xr = ws.addRow(makeRow(r, row).slice());
    if (r % 10 === 0) xr.font = BOLD.font;
    if (r % 5000 === 0) sample();
  }
  const buf = await wb.xlsx.writeBuffer();
  return buf.byteLength;
}

async function exceljsStreaming(useSharedStrings: boolean): Promise<number> {
  let bytes = 0;
  const stream = new Writable({
    write(chunk: Uint8Array, _enc: string, cb: () => void) {
      bytes += chunk.length;
      cb();
    },
  });
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useStyles: true, useSharedStrings });
  const ws = wb.addWorksheet('Data');
  ws.columns = Array.from({ length: COLS }, (_, c) => ({ width: 12, style: c >= 17 ? MONEY : {} }));
  const row: unknown[] = new Array(COLS);
  for (let r = 0; r < ROWS; r++) {
    const xr = ws.addRow(makeRow(r, row).slice());
    if (r % 10 === 0) xr.font = BOLD.font;
    xr.commit();
    if (r % 5000 === 0) sample();
  }
  ws.commit();
  await wb.commit();
  return bytes;
}

describe.skipIf(!process.env.RUN_BENCH)('xlsx write benchmark', () => {
  it(`writes ${ROWS} x ${COLS}`, async () => {
    for (const p of ['balanced', 'low-memory', 'max-speed'] as const) {
      await measure(`ours writeXlsx (${p})`, () => oursBuffered(p));
    }
    for (const p of ['balanced', 'low-memory', 'max-speed'] as const) {
      await measure(`ours createXlsxWriter (${p})`, () => oursStreaming(p));
    }
    for (const p of ['balanced', 'low-memory', 'max-speed'] as const) {
      await measure(`ours createXlsxWriter+flush (${p})`, () => oursStreamingAsync(p));
    }
    for (const p of ['balanced', 'low-memory', 'max-speed'] as const) {
      await measure(`ours xlsxStream (${p})`, () => oursXlsxStream(p));
    }
    await measure('exceljs xlsx.writeBuffer', exceljsBuffered);
    await measure('exceljs WorkbookWriter (sharedStrings)', () => exceljsStreaming(true));
    await measure('exceljs WorkbookWriter (inline)', () => exceljsStreaming(false));
    const lines = [
      `| case | time ms | size MB | peak heap MB | peak RSS MB |`,
      `|---|---:|---:|---:|---:|`,
      ...results.map(
        (r) =>
          `| ${r.name} | ${r.ms} | ${(r.bytes / 1048576).toFixed(2)} | ${r.heapMB} | ${r.rssMB} |`,
      ),
    ];
    console.log(`\n${lines.join('\n')}\n`);
  }, 600_000);
});
