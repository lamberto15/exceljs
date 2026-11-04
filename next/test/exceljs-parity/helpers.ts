// Shared helpers for the ExcelJS parity suite (test/exceljs-parity/*.test.ts).
// Each test ports one ExcelJS spec (`// exceljs: <path> — "<title>"`); see docs/exceljs-parity.md.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { unzipSync } from 'fflate';
import {
  type CellValue,
  createXlsxWriter,
  readXlsx,
  type Workbook,
  type WriteOptions,
  writeXlsx,
  type XlsxRow,
  type XlsxWriter,
  type XlsxWriterOptions,
  xlsxRows,
} from '../../src/index.js';

/** spec/integration/data (fileURLToPath, not URL.pathname: Windows paths). */
export const DATA_DIR: string = fileURLToPath(
  new URL('../../../spec/integration/data/', import.meta.url),
);
const UTIL_DIR = fileURLToPath(new URL('../../../spec/utils/data/', import.meta.url));

/** Text of a file in spec/integration/data. */
export const fixtureText = (name: string): string => readFileSync(join(DATA_DIR, name), 'utf8');

/** Bytes of a fixture from spec/integration/data. */
export const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(DATA_DIR, name)));

const DT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[.]\d{3}Z$/;

/** spec/utils/tools.js `fix`: deep clone, ISO date strings become Dates. */
function fix(o: unknown): unknown {
  if (typeof o === 'string') return DT.test(o) ? new Date(o) : o;
  if (Array.isArray(o)) return o.map(fix);
  if (o && typeof o === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(o)) if (v !== undefined) out[k] = fix(v);
    return out;
  }
  return o;
}

/** One of spec/utils/data/*.json (dates revived like the ExcelJS test utils). */
// biome-ignore lint/suspicious/noExplicitAny: test data mirrors untyped ExcelJS JSON fixtures
export const utilData = (name: string): any =>
  fix(JSON.parse(readFileSync(join(UTIL_DIR, `${name}.json`), 'utf8')));

/** write → read. */
export const roundtrip = (wb: Workbook, opts?: WriteOptions): Workbook =>
  readXlsx(writeXlsx(wb, opts));

/** Collect a ReadableStream<Uint8Array> into one buffer. */
export async function collect(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  let n = 0;
  for await (const c of rs as unknown as AsyncIterable<Uint8Array>) {
    parts.push(c);
    n += c.length;
  }
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** Run a streaming-writer script and return the finished zip. */
export async function streamWrite(
  build: (w: XlsxWriter) => Promise<void> | void,
  opts: XlsxWriterOptions = {},
): Promise<Uint8Array> {
  const w = createXlsxWriter(opts);
  const bytes = collect(w.readable!);
  await build(w);
  await w.end();
  return bytes;
}

/** All rows of xlsxRows. */
export async function allRows(src: Uint8Array): Promise<XlsxRow[]> {
  const out: XlsxRow[] = [];
  for await (const r of xlsxRows(src)) out.push(r);
  return out;
}

/** Load bytes with exceljs 4.4.0 (the oracle). */
export async function exceljsLoad(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(buf as unknown as ArrayBuffer);
  return x;
}

/** Serialize an exceljs workbook to bytes. */
export async function exceljsBytes(wb: ExcelJS.Workbook): Promise<Uint8Array> {
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/** Text of one part of a zip package. */
export function partText(zip: Uint8Array, path: string): string | undefined {
  const z = unzipSync(zip);
  const p = z[path];
  return p ? new TextDecoder().decode(p) : undefined;
}

/** Part names of a zip package. */
export const partNames = (zip: Uint8Array): string[] => Object.keys(unzipSync(zip));

export type { CellValue };
