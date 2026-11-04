/// <reference types="node" />
// Writer gaps: theme part, ST_Xstring escaping, numeric edge cases.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { parseXml } from '../src/xml/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

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

const TRICKY = [
  '_x0041_',
  'a_x005F_x0041_b',
  '__x0041_',
  '_x00',
  '_xZZZZ_',
  '_x',
  'ctrl\u0001\u001f!',
  'cr\r\nlf\rx',
  'lone\ud800hi',
  'lone\udc00lo',
  'bad￾￿',
  'pair \u{1F600} ok',
  'tab\tnl\n',
  '<&> "q"',
];

function expectTricky(buf: Uint8Array): void {
  const ws = readXlsx(buf).worksheets[0]!;
  TRICKY.forEach((t, i) => {
    expect(ws.getCell(1, i + 1).value).toBe(t);
  });
}

describe('ST_Xstring escaping', () => {
  it('writes control chars / CR as _xHHHH_ and protects literal patterns', () => {
    const xml = parts(
      writeXlsx(simple('a\u0001b', 'c\rd', '_x0041_', 'lone\ud800', 'ok_x', 'tab\t'), {
        strings: 'inline',
      }),
    ).get('xl/worksheets/sheet1.xml') as string;
    expect(xml).toContain('<t>a_x0001_b</t>');
    expect(xml).toContain('<t>c_x000D_d</t>');
    expect(xml).toContain('<t>_x005F_x0041_</t>');
    expect(xml).toContain('<t>lone_xD800_</t>');
    expect(xml).toContain('<t>ok_x</t>');
    expect(xml).toContain('<t xml:space="preserve">tab\t</t>');
  });

  it('round-trips ours -> ours (shared and inline strings, every writer)', async () => {
    for (const strings of ['shared', 'inline'] as const) {
      expectTricky(writeXlsx(simple(...TRICKY), { strings }));
      expectTricky(await readAll(xlsxStream(simple(...TRICKY), { strings })));
    }
    expectTricky(await viaWriter(TRICKY));
  });

  it('rich text runs use the same encoding', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').setCell('A1', {
      richText: [{ text: 'x\u0002_x0041_' }, { text: 'y', font: { bold: true } }],
    });
    const back = readXlsx(writeXlsx(wb)).worksheets[0]!.getCell(1, 1).value as {
      richText: { text: string }[];
    };
    expect(back.richText.map((r) => r.text)).toEqual(['x\u0002_x0041_', 'y']);
  });

  it('exceljs reads our escapes (oracle)', async () => {
    const x = await oracle(writeXlsx(simple('a_b', 'c\u0001d', '_x0041_', 'e\rf')));
    const row = x.getWorksheet('S')?.getRow(1);
    expect(row?.getCell(1).value).toBe('a_b');
    expect(row?.getCell(2).value).toBe('c\u0001d');
    expect(row?.getCell(3).value).toBe('_x0041_');
    expect(row?.getCell(4).value).toBe('e\rf');
  });

  it('exceljs writes -> we read (exceljs itself strips controls, so only plain text)', async () => {
    const x = new ExcelJS.Workbook();
    const s = x.addWorksheet('S');
    s.getCell(1, 1).value = 'a_b_x';
    s.getCell(1, 2).value = 'line1\nline2';
    // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
    const buf = new Uint8Array((await x.xlsx.writeBuffer()) as any);
    const ws = readXlsx(buf).worksheets[0]!;
    expect(ws.getCell(1, 1).value).toBe('a_b_x');
    expect(ws.getCell(1, 2).value).toBe('line1\nline2');
  });
});

describe('numeric edge cases', () => {
  const NUMS = [
    1e21,
    -0,
    5e-324,
    1.7976931348623157e308,
    -1.5e-7,
    123456789012345680000,
    0.1 + 0.2,
  ];

  it('serialize as JS shortest round-trip text, read back exactly by us and exceljs', async () => {
    const buf = writeXlsx(simple(...NUMS));
    const xml = parts(buf).get('xl/worksheets/sheet1.xml') as string;
    // JS spelling (1e+21) is a valid xsd:double accepted by exceljs and our reader; -0 is written as 0
    expect(xml).toContain('<v>1e+21</v>');
    expect(xml).toContain('<v>5e-324</v>');
    expect(xml).toContain('<v>1.7976931348623157e+308</v>');
    expect(xml).toContain('<c r="B1"><v>0</v></c>');
    const ours = readXlsx(buf).worksheets[0]!;
    const theirs = (await oracle(buf)).getWorksheet('S')!;
    NUMS.forEach((n, i) => {
      expect(ours.getCell(1, i + 1).value).toBe(n === 0 ? 0 : n);
      expect(theirs.getCell(1, i + 1).value).toBe(n === 0 ? 0 : n);
    });
  });

  it('Excel-style uppercase exponents are read', () => {
    const base = unzipSync(writeXlsx(simple(1)));
    const sheet = dec
      .decode(base.get('xl/worksheets/sheet1.xml'))
      .replace('<v>1</v>', '<v>1E+21</v>');
    base.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(sheet));
    const buf = zipSync(
      [...base].map(([path, data]) => ({ path, data })),
      6,
    );
    expect(readXlsx(buf).worksheets[0]!.getCell(1, 1).value).toBe(1e21);
  });

  it('non-finite numbers become #NUM!', () => {
    const ws = readXlsx(writeXlsx(simple(Number.NaN, Number.POSITIVE_INFINITY))).worksheets[0]!;
    expect(ws.getCell(1, 1).value).toEqual({ error: '#NUM!' });
    expect(ws.getCell(1, 2).value).toEqual({ error: '#NUM!' });
  });
});
