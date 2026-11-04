/// <reference types="node" />
// Writer gaps: theme part, ST_Xstring escaping, numeric edge cases.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { parseXml } from '../src/xml/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

async function readAll(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(rs).arrayBuffer());
}

async function viaWriter(values: unknown[], opts: { theme?: boolean } = {}): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c), ...opts });
  const s = w.addWorksheet('S');
  // biome-ignore lint/suspicious/noExplicitAny: test values
  s.addRow(values as any);
  s.commit();
  await w.end();
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

function simple(...values: unknown[]): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  values.forEach((v, i) => {
    // biome-ignore lint/suspicious/noExplicitAny: test values
    ws.setCell(1, i + 1, v as any);
  });
  return wb;
}

/** All three writers for the same single row. */
async function allWriters(
  values: unknown[],
  opts: { theme?: boolean } = {},
): Promise<Uint8Array[]> {
  return [
    writeXlsx(simple(...values), opts),
    await readAll(xlsxStream(simple(...values), opts)),
    await viaWriter(values, opts),
  ];
}

describe('theme part', () => {
  it('is written, related and typed in every writer', async () => {
    for (const buf of await allWriters(['x'])) {
      const p = parts(buf);
      const theme = p.get('xl/theme/theme1.xml');
      expect(theme).toContain('<a:clrScheme name="Office">');
      expect(theme).toContain('<a:accent1><a:srgbClr val="4472C4"/>');
      expect(p.get('xl/_rels/workbook.xml.rels')).toMatch(
        /Type="[^"]*\/relationships\/theme" Target="theme\/theme1\.xml"/,
      );
      expect(p.get('[Content_Types].xml')).toContain(
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
      );
      expect(readXlsx(buf).worksheets).toHaveLength(1);
      expect((await oracle(buf)).getWorksheet('S')?.getCell('A1').value).toBe('x');
    }
  });

  it('is well-formed XML with the 12 scheme colors in order', () => {
    const theme = parts(writeXlsx(simple('x'))).get('xl/theme/theme1.xml') as string;
    const stack: string[] = [];
    const colors: string[] = [];
    parseXml(theme, {
      open(name) {
        if (stack[stack.length - 1] === 'a:clrScheme') colors.push(name.slice(2));
        stack.push(name);
      },
      close(name) {
        expect(stack.pop()).toBe(name);
      },
    });
    expect(stack).toEqual([]);
    expect(colors).toEqual([
      'dk1',
      'lt1',
      'dk2',
      'lt2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ]);
  });

  it('theme: false omits part, relationship and content type', async () => {
    for (const buf of await allWriters(['x'], { theme: false })) {
      const p = parts(buf);
      expect(p.has('xl/theme/theme1.xml')).toBe(false);
      expect(p.get('xl/_rels/workbook.xml.rels')).not.toContain('theme');
      expect(p.get('[Content_Types].xml')).not.toContain('theme');
    }
  });
});
