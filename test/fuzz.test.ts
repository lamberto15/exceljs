/// <reference types="node" />
// Seeded mutation fuzzing of the reader (zip, XML, xlsx), the CSV parser and the XML tokenizer.
// Property: every input either succeeds or throws a documented error class (XlsxParseError,
// ZipLimitError/ZipPathError, RangeError), never hangs and never throws a raw TypeError / undefined
// access. `FUZZ_ITER=N` sets the iterations per target (default 300).

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync as fUnzip, zipSync as fZip } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { CsvError, CsvParser } from '../src/csv/parser.js';
import { addNote, addTable } from '../src/xlsx/features/index.js';
import { readXlsx, XlsxParseError, xlsxRows } from '../src/xlsx/read/index.js';
import { writeXlsx } from '../src/xlsx/write/index.js';
import { parseXml, XmlTokenizer } from '../src/xml/index.js';
import { ZipLimitError } from '../src/zip/index.js';

const ITER = Number(process.env.FUZZ_ITER) || 300;
const CASE_MS = 5000;
const SEED = Number(process.env.FUZZ_SEED) * 1000 || 0;
const enc = new TextEncoder();
const dec = new TextDecoder();

// ------------------------------------------------------------------ PRNG + helpers

function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
type Rng = () => number;
const int = (r: Rng, n: number): number => Math.floor(r() * n);
const pick = <T>(r: Rng, a: readonly T[]): T => a[int(r, a.length)] as T;

/** Documented error classes of the reader. */
function isDocumented(e: unknown): boolean {
  return e instanceof XlsxParseError || e instanceof ZipLimitError || e instanceof RangeError;
}
function describeErr(e: unknown): string {
  return e instanceof Error ? `${e.name}: ${e.message}\n${e.stack ?? ''}` : String(e);
}

// ------------------------------------------------------------------ corpus

function generated(): Uint8Array {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.setCell('A1', 'hello');
  ws.setCell('B1', 42);
  ws.setCell('C1', new Date(Date.UTC(2020, 1, 3)));
  ws.setCell('A2', 'world <&> "q"');
  ws.setCell('B2', 3.5);
  ws.setCell('C2', true);
  ws.setCell('D2', { formula: 'B1*2', result: 84 });
  ws.setCell('E2', { richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }] });
  ws.mergeCells('A4:C5');
  ws.setCell('A4', 'merged');
  addNote(ws, 'A1', 'a note');
  const t = wb.addWorksheet('T');
  addTable(t, {
    name: 'Tbl',
    topLeft: 'A1',
    columns: [{ name: 'X' }, { name: 'Y', totalsRowFunction: 'sum' }],
    rows: [
      ['a', 1],
      ['b', 2],
    ],
    totalsRow: true,
  });
  return writeXlsx(wb);
}

function corpus(): Uint8Array[] {
  const dir = join(__dirname, 'fixtures/exceljs/data');
  const out: Uint8Array[] = [generated()];
  for (const f of readdirSync(dir)) {
    if (!f.endsWith('.xlsx')) continue;
    const buf = new Uint8Array(readFileSync(join(dir, f)));
    if (buf.length < 300_000) out.push(buf); // keeps mutation fast; big files add no new shapes
  }
  return out;
}

// ------------------------------------------------------------------ mutators

function mutateBytes(r: Rng, src: Uint8Array): Uint8Array {
  let b = Array.from(src);
  const n = 1 + int(r, 8);
  for (let i = 0; i < n; i++) {
    const p = int(r, b.length);
    switch (int(r, 4)) {
      case 0:
        b[p] = int(r, 256);
        break;
      case 1:
        b[p] = (b[p] as number) ^ (1 << int(r, 8));
        break;
      case 2:
        b.splice(p, 0, int(r, 256));
        break;
      default:
        b.splice(p, 1 + int(r, 4));
    }
  }
  if (r() < 0.1) b = b.slice(0, int(r, b.length));
  return Uint8Array.from(b);
}

const ABSURD = [
  'XFE1048577',
  'XFD1048577',
  'A0',
  'ZZZZ1',
  '-1',
  'NaN',
  '1e400',
  '999999',
  '4294967296',
  '-5',
  '0',
  '1048576',
  '16384',
  'A1:XFD1048576',
  '1:1048576',
  'Infinity',
  '',
];
const ATTRS = ['r', 's', 't', 'min', 'max', 'ref', 'count', 'spans', 'si', 'ht', 'width', 'span'];
const ENTS = ['&amp;', '&lt;', '&#0;', '&#x110000;', '&#xD800;', '&bogus;', '&#99999999999;', '&'];
const FRAGS = [
  '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>',
  '<cols><col min="1" max="16384" width="9"/></cols>',
  '<cols><col min="1" max="99999999" width="9" style="999999"/></cols>',
  '<c r="A1" t="s"><v>99999999</v></c>',
  '<c r="B1" s="999999"><v>1</v></c>',
  '<c r="C1" s="-1"><v>1</v></c>',
  '<c r="D1"><f t="shared" si="77"/><v>1</v></c>',
  '<c r="E1"><f t="shared" si="3" ref="E1:E5">A1</f></c>',
  '<c r="XFE1048577"><v>1</v></c>',
  '<row r="0"><c r="A1"/></row>',
  '<row r="2000000"><c r="A1"/></row>',
  '<row r="-3" ht="NaN" s="abc" customFormat="1"><c/></row>',
  '<row spans="1:99999999"><c><v>1</v></c></row>',
  '<si><t>x</t></si>',
  '<![CDATA[ ]]]]><![CDATA[> ]]>',
  '<!DOCTYPE x [<!ENTITY a "aaaa">]>',
  '<?pi ?>',
  '<c r="A1" t="e"><v>#REF!</v></c>',
  '<c r="A1" t="inlineStr"><is><t>q</t></is></c>',
  '<hyperlinks><hyperlink ref="Z99999999" r:id="rId999"/></hyperlinks>',
  '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
];

function mutateXml(r: Rng, xml: string): string {
  const n = 1 + int(r, 4);
  for (let i = 0; i < n; i++) {
    const p = int(r, xml.length + 1);
    switch (int(r, 12)) {
      case 0:
        xml = xml.slice(0, p);
        break;
      case 1: {
        // duplicate a tag
        const s = xml.indexOf('<', p);
        const e = s < 0 ? -1 : xml.indexOf('>', s);
        if (e > 0) xml = xml.slice(0, e + 1) + xml.slice(s, e + 1) + xml.slice(e + 1);
        break;
      }
      case 2: {
        // drop a closing tag
        const s = xml.indexOf('</', p);
        const e = s < 0 ? -1 : xml.indexOf('>', s);
        if (e > 0) xml = xml.slice(0, s) + xml.slice(e + 1);
        break;
      }
      case 3:
        xml = `${xml.slice(0, p)}${'x'.repeat(100000)}${xml.slice(p)}`;
        break;
      case 4:
        xml = xml.slice(0, p) + pick(r, ENTS) + xml.slice(p);
        break;
      case 5: {
        // replace an attribute value with an absurd one
        const a = pick(r, ATTRS);
        const hits = [...xml.matchAll(new RegExp(`\\b${a}="[^"]*"`, 'g'))];
        if (hits.length > 0) {
          const h = pick(r, hits);
          const idx = h.index as number;
          xml = `${xml.slice(0, idx)}${a}="${pick(r, ABSURD)}"${xml.slice(idx + h[0].length)}`;
        }
        break;
      }
      case 6: {
        // swap two adjacent rows (out of order)
        const a = xml.indexOf('<row', p);
        const b = a < 0 ? -1 : xml.indexOf('</row>', a);
        const c = b < 0 ? -1 : xml.indexOf('<row', b);
        const d = c < 0 ? -1 : xml.indexOf('</row>', c);
        if (d > 0) {
          const A = xml.slice(a, b + 6);
          const B = xml.slice(c, d + 6);
          xml = xml.slice(0, a) + B + xml.slice(b + 6, c) + A + xml.slice(d + 6);
        }
        break;
      }
      case 7: {
        // duplicate a cell
        const a = xml.indexOf('<c ', p);
        const b = a < 0 ? -1 : xml.indexOf('</c>', a);
        if (b > 0) xml = xml.slice(0, b + 4) + xml.slice(a, b + 4) + xml.slice(b + 4);
        break;
      }
      case 8:
        xml = xml.slice(0, p) + xml.slice(p + 1 + int(r, 30));
        break;
      case 9:
        xml = xml.slice(0, p) + pick(r, FRAGS) + xml.slice(p);
        break;
      case 10:
        xml = `${xml.slice(0, p)}${'<a>'.repeat(1000 + int(r, 6000))}${xml.slice(p)}`;
        break;
      default:
        xml = xml.replace(
          /="(\d+)"/,
          (_m, d: string) => `="${Number(d) + pick(r, [-1, 1, 1e6, -1e9])}"`,
        );
    }
  }
  return xml;
}

function mutateZip(r: Rng, src: Uint8Array): Uint8Array {
  let files: Record<string, Uint8Array>;
  try {
    files = fUnzip(src);
  } catch {
    return mutateBytes(r, src);
  }
  const names = Object.keys(files).filter((n) => /\.(xml|rels|vml)$/.test(n));
  const k = 1 + int(r, 2);
  for (let i = 0; i < k && names.length > 0; i++) {
    const name = pick(r, names);
    files[name] = enc.encode(mutateXml(r, dec.decode(files[name])));
  }
  if (r() < 0.05) delete files[pick(r, Object.keys(files))];
  return fZip(files);
}

// ------------------------------------------------------------------ runners

async function runCase(name: string, data: Uint8Array): Promise<void> {
  if (process.env.FUZZ_TRACE) process.stderr.write(`${name} ${data.length}\n`);
  const t0 = performance.now();
  const check = (e: unknown, what: string): void => {
    if (!isDocumented(e)) {
      throw new Error(`[${name}] ${what} threw undocumented error: ${describeErr(e)}`);
    }
  };
  let wb: Workbook | undefined;
  try {
    wb = readXlsx(data);
  } catch (e) {
    check(e, 'readXlsx');
  }
  // whatever was read must be writable again (consistent model: style ids, ranges, ...)
  if (wb !== undefined && wb.worksheets.length > 0) {
    try {
      readXlsx(writeXlsx(wb));
    } catch (e) {
      // the writer validates feature models (tables, ...) with plain Errors; only crashes count
      if (e instanceof TypeError || e instanceof ReferenceError) check(e, 'write after read');
    }
  }
  try {
    for await (const _row of xlsxRows(data)) {
      if (performance.now() - t0 > CASE_MS) throw new Error(`[${name}] xlsxRows too slow`);
    }
  } catch (e) {
    if (e instanceof Error && e.message.endsWith('too slow')) throw e;
    check(e, 'xlsxRows');
  }
  // streaming input split at arbitrary points
  try {
    const cut = 1 + (data.length % 977);
    async function* chunks(): AsyncGenerator<Uint8Array> {
      for (let i = 0; i < data.length; i += cut) yield data.subarray(i, i + cut);
    }
    for await (const _row of xlsxRows(chunks())) {
      if (performance.now() - t0 > CASE_MS) throw new Error(`[${name}] chunked xlsxRows too slow`);
    }
  } catch (e) {
    if (e instanceof Error && e.message.endsWith('too slow')) throw e;
    check(e, 'xlsxRows(chunked)');
  }
  const dt = performance.now() - t0;
  if (dt > CASE_MS) throw new Error(`[${name}] case took ${dt}ms`);
}

describe('fuzz: xlsx reader', () => {
  const files = corpus();

  it('corpus itself reads', () => {
    for (const f of files) {
      try {
        readXlsx(f);
      } catch (e) {
        expect(isDocumented(e)).toBe(true);
      }
    }
  });

  it('zip byte mutations', async () => {
    const r = prng(1 + SEED);
    for (let i = 0; i < ITER; i++) await runCase(`bytes#${i}`, mutateBytes(r, pick(r, files)));
  }, 300000);

  it('xml part mutations', async () => {
    const r = prng(2 + SEED);
    for (let i = 0; i < ITER; i++) await runCase(`xml#${i}`, mutateZip(r, pick(r, files)));
  }, 300000);

  it('deep nesting (100k levels) fails cleanly', () => {
    const base = fUnzip(generated());
    base['xl/worksheets/sheet1.xml'] = enc.encode(
      `<worksheet>${'<a>'.repeat(100000)}${'</a>'.repeat(100000)}</worksheet>`,
    );
    try {
      readXlsx(fZip(base));
    } catch (e) {
      expect(isDocumented(e)).toBe(true);
    }
  });
});

// ------------------------------------------------------------------ resource bounds

/** Rebuild the generated workbook with `sheet1.xml` transformed. */
function withSheet(fn: (xml: string) => string): Uint8Array {
  const base = fUnzip(generated());
  const name = 'xl/worksheets/sheet1.xml';
  base[name] = enc.encode(fn(dec.decode(base[name])));
  return fZip(base);
}

/** Run `fn`, return elapsed ms and heap growth (MiB). */
function measure(fn: () => void): { ms: number; heapMiB: number } {
  const gc = (globalThis as { gc?: () => void }).gc;
  gc?.();
  const h0 = process.memoryUsage().heapUsed;
  const t0 = performance.now();
  fn();
  return { ms: performance.now() - t0, heapMiB: (process.memoryUsage().heapUsed - h0) / 2 ** 20 };
}

describe('regressions: error classes', () => {
  it('truncated [Content_Types].xml is an XlsxParseError, not a raw tokenizer Error', () => {
    const base = fUnzip(generated());
    base['[Content_Types].xml'] = enc.encode('<Types xmlns="x"><Default Extension="xml" Conte');
    expect(() => readXlsx(fZip(base))).toThrow(XlsxParseError);
  });

  it('corrupt deflate data is a ZipLimitError (sync and streaming), not a raw fflate Error', async () => {
    const good = generated();
    // flip bytes inside the first compressed entry (local header is 30 + name bytes)
    const bad = good.slice();
    const nameLen = (bad[26] as number) | ((bad[27] as number) << 8);
    const extraLen = (bad[28] as number) | ((bad[29] as number) << 8);
    const start = 30 + nameLen + extraLen;
    for (let i = start; i < start + 40; i++) bad[i] = 0xff;
    expect(() => readXlsx(bad)).toThrow(ZipLimitError);
    await expect(
      (async () => {
        for await (const _ of xlsxRows(bad)) void _;
      })(),
    ).rejects.toBeInstanceOf(ZipLimitError);
  });
});

describe('fuzz: resource bounds', () => {
  const HEAP_MIB = 64;
  it('merge A1:XFD1048576 does not expand into cells', () => {
    const data = withSheet((x) =>
      x.replace(
        /<mergeCells[\s\S]*?<\/mergeCells>/,
        '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>',
      ),
    );
    let wb: Workbook | undefined;
    const m = measure(() => {
      wb = readXlsx(data);
    });
    expect(m.heapMiB).toBeLessThan(HEAP_MIB);
    expect(m.ms).toBeLessThan(2000);
    expect(wb?.worksheets[0]?.merges).toBeDefined();
  });

  it('<col min=1 max=16384> does not allocate per column', () => {
    const data = withSheet((x) =>
      x.replace(
        /<sheetData/,
        '<cols><col min="1" max="16384" width="9" style="0"/></cols><sheetData',
      ),
    );
    const m = measure(() => void readXlsx(data));
    expect(m.heapMiB).toBeLessThan(HEAP_MIB);
    expect(m.ms).toBeLessThan(2000);
  });

  it('many rows with a cell in the last column stay sparse', () => {
    let rows = '';
    for (let i = 1; i <= 20000; i++) rows += `<row r="${i}"><c r="XFD${i}"><v>1</v></c></row>`;
    const data = withSheet((x) =>
      x.replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${rows}</sheetData>`),
    );
    const m = measure(() => void readXlsx(data));
    expect(m.heapMiB).toBeLessThan(HEAP_MIB);
    expect(m.ms).toBeLessThan(3000);
  });

  it('rows far apart and huge spans stay cheap', () => {
    const data = withSheet((x) =>
      x.replace(
        /<sheetData>/,
        '<sheetData><row r="1048576" spans="1:99999999"><c r="XFD1048576"><v>1</v></c></row>',
      ),
    );
    const m = measure(() => {
      try {
        readXlsx(data);
      } catch (e) {
        expect(isDocumented(e)).toBe(true);
      }
    });
    expect(m.heapMiB).toBeLessThan(HEAP_MIB);
    expect(m.ms).toBeLessThan(2000);
  });
});

// ------------------------------------------------------------------ CSV + XML tokenizer

describe('fuzz: csv parser', () => {
  it('random input only throws CsvError and chunking is irrelevant', () => {
    const r = prng(3 + SEED);
    const alphabet = [
      'a',
      'b',
      ',',
      ';',
      '"',
      '"',
      '\n',
      '\r',
      '\r\n',
      ' ',
      '\t',
      '1',
      '﻿',
      '\u{1F600}',
    ];
    for (let i = 0; i < ITER * 3; i++) {
      let s = '';
      const n = int(r, 60);
      for (let j = 0; j < n; j++) s += pick(r, alphabet);
      const opts = { delimiter: pick(r, [',', ';', '\t']), quote: pick(r, ['"', "'"]) };
      const whole: string[][] = [];
      const chunked: string[][] = [];
      try {
        const p = new CsvParser((row) => void whole.push(row.slice()), opts);
        p.write(s);
        p.end();
        const q = new CsvParser((row) => void chunked.push(row.slice()), opts);
        for (let k = 0; k < s.length; ) {
          const step = 1 + int(r, 5);
          q.write(s.slice(k, k + step));
          k += step;
        }
        q.end();
      } catch (e) {
        expect(e).toBeInstanceOf(CsvError);
        continue;
      }
      expect(chunked).toEqual(whole);
    }
  });

  it('maxFieldBytes bounds memory', () => {
    const p = new CsvParser(() => {}, { maxFieldBytes: 1000 });
    expect(() => {
      p.write(`"${'a'.repeat(5000)}`);
      p.end();
    }).toThrow(CsvError);
  });
});

describe('fuzz: xml tokenizer', () => {
  const merge = (ev: string[]): string[] => {
    const o: string[] = [];
    for (const e of ev) {
      const last = o[o.length - 1];
      if (e[0] === 't' && last !== undefined && last[0] === 't')
        o[o.length - 1] = last + e.slice(1);
      else o.push(e);
    }
    return o;
  };
  const handler = (ev: string[]) => ({
    open: (nm: string, a: Record<string, string>, sc: boolean) =>
      void ev.push(`o${nm}${JSON.stringify(a)}${sc}`),
    close: (nm: string) => void ev.push(`c${nm}`),
    text: (t: string) => void ev.push(`t${t}`),
  });

  it('random markup never throws a TypeError and chunking is irrelevant', () => {
    const r = prng(4 + SEED);
    const toks = [
      '<',
      '>',
      '</',
      '/>',
      'a',
      'b:c',
      ' ',
      '=',
      '"',
      "'",
      '&amp;',
      '&',
      '&#',
      ';',
      '<![CDATA[',
      ']]>',
      '<!--',
      '-->',
      '<?',
      '?>',
      '<!DOCTYPE',
      '\n',
      'x="1"',
      '\u{1F600}',
    ];
    for (let i = 0; i < ITER * 3; i++) {
      let s = '';
      const n = int(r, 40);
      for (let j = 0; j < n; j++) s += pick(r, toks);
      const ev1: string[] = [];
      const ev2: string[] = [];
      let ok = true;
      try {
        parseXml(s, handler(ev1));
        const tk = new XmlTokenizer(handler(ev2));
        for (let k = 0; k < s.length; ) {
          const step = 1 + int(r, 4);
          tk.write(s.slice(k, k + step));
          k += step;
        }
        tk.end();
      } catch (e) {
        ok = false;
        expect(e).not.toBeInstanceOf(TypeError);
      }
      if (ok) expect(merge(ev2)).toEqual(merge(ev1));
    }
  });

  it('100k nested elements', () => {
    let n = 0;
    parseXml('<a>'.repeat(100000), { open: () => void n++ });
    expect(n).toBe(100000);
  });
});
