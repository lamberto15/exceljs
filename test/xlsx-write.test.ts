/// <reference types="node" />
import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';
import ExcelJS from 'exceljs';
import { unzipSync as rawUnzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { CellKind, Workbook } from '../src/core/index.js';
import { readXlsx, xlsxRows } from '../src/xlsx/read/index.js';
import {
  createXlsxWriter,
  resolveWriteOptions,
  writeXlsx,
  type XlsxWriterOptions,
  xlsxStream,
} from '../src/xlsx/write/index.js';
import { parseXml } from '../src/xml/index.js';
import { nativeDeflateAvailable, unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

/** Every part must be well-formed XML (tokenizer throws on malformed input). */
function assertWellFormed(files: Map<string, string>): void {
  for (const [name, xml] of files) {
    const stack: string[] = [];
    let roots = 0;
    try {
      parseXml(xml, {
        open(n) {
          if (stack.length === 0) roots++;
          stack.push(n);
        },
        close(n) {
          const top = stack.pop();
          if (top !== n) throw new Error(`mismatched </${n}>, expected </${top}>`);
        },
      });
    } catch (e) {
      throw new Error(`${name}: ${(e as Error).message}`);
    }
    expect(stack.length, name).toBe(0);
    expect(roots, name).toBe(1);
  }
}

async function readAll(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const reader = rs.getReader();
  for (;;) {
    const r = await reader.read();
    if (r.done) break;
    chunks.push(r.value);
  }
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

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

const NASTY = `a<b>&"c'\u0001d\u0008 😀 ü`;
const D1 = new Date(Date.UTC(2024, 0, 15));
const D2 = new Date(Date.UTC(2024, 5, 30, 13, 45, 10));

/** Builds the same content through the in-memory model. */
function buildWorkbook(): Workbook {
  const wb = new Workbook();
  wb.creator = 'Tester <&>';
  const ws = wb.addWorksheet('Data & Stuff');
  ws.columns = [
    { header: 'Id', key: 'id', width: 8 },
    { header: 'Name', key: 'name', width: 25 },
    { header: 'Amount', key: 'amount', width: 12, style: { numFmt: '#,##0.00' } },
    { header: 'When', key: 'when', width: 12 },
    { header: 'Flag', key: 'flag' },
  ];
  ws.addRow({ id: 1, name: 'alpha', amount: 10.5, when: D1, flag: true });
  ws.addRow({ id: 2, name: '  padded  ', amount: -3, when: D2, flag: false });
  ws.addRow([3, NASTY, 1e21, 1e-7, { error: '#N/A' }]);
  ws.addRow([
    { formula: 'A2+A3', result: 3 },
    { formula: 'CONCAT("a","b")', result: 'ab' },
    { formula: 'TRUE()', result: true },
    { formula: '1/0', result: { error: '#DIV/0!' } },
    { formula: '=SUM(A2:A3)' },
  ]);
  ws.addRow([
    {
      richText: [
        { text: 'bold ', font: { bold: true, color: { argb: 'FFFF0000' } } },
        { text: 'plain' },
      ],
    },
    { text: 'example', hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'go' },
    { text: 'internal', hyperlink: "#'Other'!A1" },
  ]);
  ws.getCell('A7').value = 'styled';
  ws.getCell('A7').style = {
    font: { bold: true, italic: true, size: 14, name: 'Arial', underline: 'double' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
    border: { top: { style: 'thin', color: { argb: 'FF0000FF' } }, bottom: { style: 'double' } },
    alignment: { horizontal: 'center', vertical: 'top', wrapText: true },
    numFmt: '0.000%',
  };
  ws.getCell('B7').value = 0.25;
  ws.getCell('B7').numFmt = '0.000%';
  ws.mergeCells('C7:E8');
  ws.setRowProps(7, { height: 30 });
  ws.setRowProps(9, { hidden: true, outlineLevel: 1 });
  ws.getCell('A9').value = 'hidden row';
  ws.freeze(1, 1);
  ws.autoFilter = 'A1:E1';
  ws.tabColor = { argb: 'FF00FF00' };

  const other = wb.addWorksheet('Other');
  other.getCell('A1').value = 'second';
  other.getCell('B2').value = 42;
  const hidden = wb.addWorksheet('Secret', { state: 'hidden' });
  hidden.getCell('A1').value = 'hidden sheet';
  return wb;
}

/** Same content via the streaming writer. */
async function buildStreaming(opts: XlsxWriterOptions): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({
    ...opts,
    creator: 'Tester <&>',
    flushRows: 2,
    sink: (c) => void chunks.push(c),
  });
  const s = w.addWorksheet('Data & Stuff', {
    columns: [
      { header: 'Id', key: 'id', width: 8 },
      { header: 'Name', key: 'name', width: 25 },
      { header: 'Amount', key: 'amount', width: 12, style: { numFmt: '#,##0.00' } },
      { header: 'When', key: 'when', width: 12 },
      { header: 'Flag', key: 'flag' },
    ],
    freeze: { rows: 1, cols: 1 },
    autoFilter: 'A1:E1',
    tabColor: { argb: 'FF00FF00' },
  });
  s.addRow({ id: 1, name: 'alpha', amount: 10.5, when: D1, flag: true });
  s.addRow({ id: 2, name: '  padded  ', amount: -3, when: D2, flag: false });
  s.addRow([3, NASTY, 1e21, 1e-7, { error: '#N/A' }]);
  s.addRow([
    { formula: 'A2+A3', result: 3 },
    { formula: 'CONCAT("a","b")', result: 'ab' },
    { formula: 'TRUE()', result: true },
    { formula: '1/0', result: { error: '#DIV/0!' } },
    { formula: '=SUM(A2:A3)' },
  ]);
  s.addRow([
    {
      richText: [
        { text: 'bold ', font: { bold: true, color: { argb: 'FFFF0000' } } },
        { text: 'plain' },
      ],
    },
    { text: 'example', hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'go' },
    { text: 'internal', hyperlink: "#'Other'!A1" },
  ]);
  const styled = {
    font: { bold: true, italic: true, size: 14, name: 'Arial', underline: 'double' as const },
    fill: { type: 'pattern' as const, pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
    border: {
      top: { style: 'thin' as const, color: { argb: 'FF0000FF' } },
      bottom: { style: 'double' as const },
    },
    alignment: { horizontal: 'center' as const, vertical: 'top' as const, wrapText: true },
    numFmt: '0.000%',
  };
  s.addRow(['styled'], styled, { height: 30 });
  // B7 needs its own style: the per-row style applies to all cells, so restyle via a second write
  s.addRow([]);
  s.addRow(['hidden row'], undefined, { hidden: true, outlineLevel: 1 });
  s.mergeCells('C7:E8');
  await s.commit();
  const o = w.addWorksheet('Other');
  o.addRow(['second']);
  o.addRow([null, 42]);
  const h = w.addWorksheet('Secret', { state: 'hidden' });
  h.addRow(['hidden sheet']);
  await w.end();
  let n = 0;
  for (const c of chunks) n += c.length;
  const out = new Uint8Array(n);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

async function checkOracle(buf: Uint8Array, streaming: boolean): Promise<void> {
  const files = parts(buf);
  assertWellFormed(files);
  const x = await oracle(buf);
  expect(x.creator).toBe('Tester <&>');
  expect(x.worksheets.map((w) => w.name)).toEqual(['Data & Stuff', 'Other', 'Secret']);
  const ws = x.getWorksheet('Data & Stuff')!;
  expect(x.getWorksheet('Secret')!.state).toBe('hidden');
  expect(x.getWorksheet('Other')!.state).toBe('visible');
  expect(x.getWorksheet('Other')!.getCell('B2').value).toBe(42);
  expect(x.getWorksheet('Secret')!.getCell('A1').value).toBe('hidden sheet');

  // header + values / types
  expect(ws.getRow(1).values).toEqual([undefined, 'Id', 'Name', 'Amount', 'When', 'Flag']);
  expect(ws.getCell('A2').value).toBe(1);
  expect(ws.getCell('B2').value).toBe('alpha');
  expect(ws.getCell('C2').value).toBe(10.5);
  expect(ws.getCell('C2').numFmt).toBe('#,##0.00');
  expect(ws.getCell('D2').value).toEqual(D1);
  expect(ws.getCell('D3').value).toEqual(D2);
  expect(ws.getCell('E2').value).toBe(true);
  expect(ws.getCell('E3').value).toBe(false);
  expect(ws.getCell('B3').value).toBe('  padded  ');
  // control chars are written as _xHHHH_ (Excel's ST_Xstring); exceljs decodes them in the shared
  // string table only, not in inlineStr
  expect(ws.getCell('B4').value).toBe(
    files.has('xl/sharedStrings.xml') ? NASTY : `a<b>&"c'_x0001_d_x0008_ 😀 ü`,
  );
  expect(ws.getCell('C4').value).toBe(1e21);
  expect(ws.getCell('D4').value).toBe(1e-7);
  expect(ws.getCell('E4').value).toEqual({ error: '#N/A' });

  // formulas
  expect(ws.getCell('A5').value).toEqual({ formula: 'A2+A3', result: 3 });
  // exceljs shows the file text: future functions carry the _xlfn. prefix Excel needs
  expect(ws.getCell('B5').value).toEqual({ formula: '_xlfn.CONCAT("a","b")', result: 'ab' });
  expect(ws.getCell('C5').value).toEqual({ formula: 'TRUE()', result: true });
  expect(ws.getCell('D5').value).toEqual({ formula: '1/0', result: { error: '#DIV/0!' } });
  expect((ws.getCell('E5').value as ExcelJS.CellFormulaValue).formula).toBe('SUM(A2:A3)');

  // rich text + hyperlinks
  const rt = ws.getCell('A6').value as ExcelJS.CellRichTextValue;
  expect(rt.richText.map((r) => r.text)).toEqual(['bold ', 'plain']);
  expect(rt.richText[0]!.font?.bold).toBe(true);
  expect(rt.richText[0]!.font?.color?.argb).toBe('FFFF0000');
  // exceljs 4.4 drops tooltips when reconciling hyperlinks; checked in the raw XML below
  expect(ws.getCell('B6').value).toMatchObject({
    text: 'example',
    hyperlink: 'https://example.com/?a=1&b=2',
  });
  // internal links use location= without a relationship (per spec); exceljs only resolves
  // hyperlinks through rels, so it sees the plain text
  expect(ws.getCell('C6').value).toBe('internal');

  // styles
  const a7 = ws.getCell('A7');
  expect(a7.value).toBe('styled');
  expect(a7.font).toMatchObject({
    bold: true,
    italic: true,
    size: 14,
    name: 'Arial',
    underline: 'double',
  });
  expect(a7.fill).toMatchObject({
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' },
  });
  expect(a7.border.top).toEqual({ style: 'thin', color: { argb: 'FF0000FF' } });
  expect(a7.border.bottom).toMatchObject({ style: 'double' });
  expect(a7.alignment).toMatchObject({ horizontal: 'center', vertical: 'top', wrapText: true });
  expect(a7.numFmt).toBe('0.000%');
  if (!streaming) expect(ws.getCell('B7').numFmt).toBe('0.000%');

  // merges, widths, heights, views, rows
  expect(ws.getCell('C7').isMerged).toBe(true);
  expect(ws.getCell('E8').isMerged).toBe(true);
  expect(ws.getCell('E8').master.address).toBe('C7');
  expect(ws.getColumn(1).width).toBe(8);
  expect(ws.getColumn(2).width).toBe(25);
  expect(ws.getColumn(3).numFmt).toBe('#,##0.00');
  expect(ws.getRow(7).height).toBe(30);
  expect(ws.getRow(9).hidden).toBe(true);
  expect(ws.getRow(9).outlineLevel).toBe(1);
  expect(ws.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' });
  expect(ws.autoFilter).toBe('A1:E1');
  expect(ws.properties.tabColor).toEqual({ argb: 'FF00FF00' });

  // raw XML spot checks
  const sheet1 = files.get('xl/worksheets/sheet1.xml')!;
  expect(sheet1).toContain('<pageMargins left="0.7"');
  expect(sheet1).toContain(
    '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
  );
  expect(sheet1.indexOf('<autoFilter')).toBeLessThan(sheet1.indexOf('<mergeCells'));
  expect(sheet1.indexOf('<mergeCells')).toBeLessThan(sheet1.indexOf('<hyperlinks'));
  expect(sheet1.indexOf('<hyperlinks')).toBeLessThan(sheet1.indexOf('<pageMargins'));
  expect(sheet1).not.toMatch(/ s="0"/);
  expect(sheet1).toMatch(/<hyperlink ref="B6" r:id="rId1" tooltip="go"\/>/);
  expect(sheet1).toContain(`<hyperlink ref="C6" location="'Other'!A1"/>`);
  expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain('TargetMode="External"');
  expect(files.get('xl/workbook.xml')).toContain(
    `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Data &amp; Stuff'!$A$1:$E$1</definedName>`,
  );
  expect(files.get('xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
}

describe('resolveWriteOptions', () => {
  it('applies profiles and explicit overrides', () => {
    expect(resolveWriteOptions()).toMatchObject({
      profile: 'balanced',
      compression: 6,
      strings: 'shared',
    });
    expect(resolveWriteOptions({ profile: 'low-memory' })).toMatchObject({
      compression: 6,
      strings: 'inline',
    });
    expect(resolveWriteOptions({ profile: 'max-speed' })).toMatchObject({
      compression: 1,
      strings: 'inline',
    });
    expect(
      resolveWriteOptions({ profile: 'max-speed', strings: 'shared', compression: 0 }),
    ).toMatchObject({
      compression: 0,
      strings: 'shared',
    });
    expect(() => resolveWriteOptions({ compression: 12 })).toThrow(RangeError);
  });
});

describe('writeXlsx', () => {
  for (const strings of ['shared', 'inline'] as const) {
    it(`round-trips through exceljs (${strings} strings)`, async () => {
      const buf = writeXlsx(buildWorkbook(), { strings });
      await checkOracle(buf, false);
      const files = parts(buf);
      expect(files.has('xl/sharedStrings.xml')).toBe(strings === 'shared');
      if (strings === 'inline') {
        const sheet = files.get('xl/worksheets/sheet1.xml')!;
        expect(sheet).toContain('<is><t xml:space="preserve">  padded  </t></is>');
        expect(sheet).toContain('<t>a&lt;b&gt;&amp;"c\'_x0001_d_x0008_ 😀 ü</t>');
      } else {
        expect(files.get('xl/sharedStrings.xml')).toContain(
          '<si><t xml:space="preserve">  padded  </t></si>',
        );
      }
    });
  }

  it('writes all required parts with consistent rels / content types', () => {
    const files = parts(writeXlsx(buildWorkbook()));
    for (const p of [
      '[Content_Types].xml',
      '_rels/.rels',
      'docProps/core.xml',
      'docProps/app.xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/sharedStrings.xml',
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/sheet2.xml',
      'xl/worksheets/sheet3.xml',
    ]) {
      expect(files.has(p), p).toBe(true);
    }
    const ct = files.get('[Content_Types].xml')!;
    expect(ct).toContain('PartName="/xl/worksheets/sheet3.xml"');
    const styles = files.get('xl/styles.xml')!;
    // stylesheet child order
    const order = [
      '<numFmts',
      '<fonts',
      '<fills',
      '<borders',
      '<cellStyleXfs',
      '<cellXfs',
      '<cellStyles',
    ];
    for (let i = 1; i < order.length; i++) {
      expect(styles.indexOf(order[i - 1]!)).toBeLessThan(styles.indexOf(order[i]!));
    }
    // worksheet child order
    const s1 = files.get('xl/worksheets/sheet1.xml')!;
    const so = ['<sheetPr', '<dimension', '<sheetViews', '<sheetFormatPr', '<cols', '<sheetData'];
    for (let i = 1; i < so.length; i++)
      expect(s1.indexOf(so[i - 1]!)).toBeLessThan(s1.indexOf(so[i]!));
    expect(s1).toContain('<dimension ref="A1:E9"/>');
    expect(s1).toContain('outlineLevelRow="1"');
    expect(files.get('xl/workbook.xml')).toContain(
      '<sheet name="Secret" sheetId="3" state="hidden" r:id="rId3"/>',
    );
  });

  it('merges identical consecutive columns', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let c = 1; c <= 5; c++) ws.getColumn(c).width = c <= 3 ? 20 : 10;
    ws.getColumn(7).hidden = true;
    ws.getCell('A1').value = 1;
    const s = parts(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!;
    expect(s).toContain(
      '<cols><col min="1" max="3" width="20" customWidth="1"/><col min="4" max="5" width="10" customWidth="1"/><col min="7" max="7" width="9.140625" hidden="1"/></cols>',
    );
  });

  it('writes numbers exactly and handles 1904 / empty-styled cells', async () => {
    const wb = new Workbook({ date1904: true });
    const ws = wb.addWorksheet('N');
    const nums = [0, -0.5, 123456789.123, 1e21, 1.5e-300, Number.MAX_SAFE_INTEGER, -1e-7];
    ws.addRow(nums);
    ws.getCell('A2').value = D1;
    ws.setStyle('B2', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF112233' } },
    });
    expect(ws.getKind('B2')).toBe(CellKind.Empty);
    const buf = writeXlsx(wb);
    const s = parts(buf).get('xl/worksheets/sheet1.xml')!;
    expect(s).toContain('<v>1e+21</v>');
    expect(s).toMatch(/<c r="B2" s="\d+"\/>/);
    expect(parts(buf).get('xl/workbook.xml')).toContain('<workbookPr date1904="1"/>');
    const x = await oracle(buf);
    const xs = x.worksheets[0]!;
    expect(xs.getRow(1).values).toEqual([undefined, ...nums]);
    expect(xs.getCell('A2').value).toEqual(D1);
    expect(xs.getCell('B2').fill).toMatchObject({ fgColor: { argb: 'FF112233' } });
  });

  it('escapes sheet names, tooltips, formulas and core props', async () => {
    const wb = new Workbook();
    wb.creator = `"q" & <x>`;
    const ws = wb.addWorksheet("It's <&>");
    ws.getCell('A1').value = { formula: 'IF(1<2,"a&b","c")', result: 'a&b' };
    ws.getCell('A2').value = {
      text: 't',
      hyperlink: 'http://x.test/?q="1"&r=<2>',
      tooltip: '"<tip>"',
    };
    ws.getCell('A3').value = '\ttab lead';
    ws.getCell('A4').value = 'trail\n';
    ws.getCell('A5').value = '';
    ws.autoFilter = 'A1:A5';
    for (const strings of ['shared', 'inline'] as const) {
      const buf = writeXlsx(wb, { strings });
      assertWellFormed(parts(buf));
      const x = await oracle(buf);
      const xs = x.worksheets[0]!;
      expect(xs.name).toBe("It's <&>");
      expect(x.creator).toBe(`"q" & <x>`);
      expect(xs.getCell('A1').value).toEqual({ formula: 'IF(1<2,"a&b","c")', result: 'a&b' });
      expect(xs.getCell('A2').value).toMatchObject({
        text: 't',
        hyperlink: 'http://x.test/?q="1"&r=<2>',
      });
      expect(parts(buf).get('xl/worksheets/sheet1.xml')).toContain(
        'tooltip="&quot;&lt;tip&gt;&quot;"',
      );
      expect(xs.getCell('A3').value).toBe('\ttab lead');
      expect(xs.getCell('A4').value).toBe('trail\n');
    }
  });

  it('rejects workbooks without visible sheets', () => {
    const wb = new Workbook();
    expect(() => writeXlsx(wb)).toThrow(/no worksheets/);
    wb.addWorksheet('a', { state: 'hidden' });
    expect(() => writeXlsx(wb)).toThrow(/visible/);
  });

  it('selects the tabSelected sheet as active tab', () => {
    const wb = new Workbook();
    wb.addWorksheet('a').getCell('A1').value = 1;
    const b = wb.addWorksheet('b');
    b.views = [{ tabSelected: true }];
    const files = parts(writeXlsx(wb));
    expect(files.get('xl/workbook.xml')).toContain('activeTab="1"');
    expect(files.get('xl/worksheets/sheet2.xml')).toContain('tabSelected="1"');
    expect(files.get('xl/worksheets/sheet1.xml')).not.toContain('tabSelected');
  });

  it('writes array formulas and compression level 0', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('A');
    ws.getCell('A1').value = { formula: 'TRANSPOSE(B1:B2)', array: 'A1:B1', result: 1 };
    const buf = writeXlsx(wb, { compression: 0 });
    expect(parts(buf).get('xl/worksheets/sheet1.xml')).toContain(
      '<c r="A1"><f t="array" ref="A1:B1">TRANSPOSE(B1:B2)</f><v>1</v></c>',
    );
    // stored entries: sheet XML visible in the raw bytes
    expect(dec.decode(buf)).toContain('<sheetData>');
  });
});

describe('xlsxStream', () => {
  it('produces the same workbook as writeXlsx', async () => {
    for (const profile of ['balanced', 'low-memory', 'max-speed'] as const) {
      const buf = await readAll(xlsxStream(buildWorkbook(), { profile }));
      await checkOracle(buf, false);
    }
  });

  it('streams large sheets in several chunks with backpressure', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Big');
    for (let r = 0; r < 20000; r++) ws.addRow([r, `row ${r}`, r * 1.5]);
    const rs = xlsxStream(wb, { compression: 1 });
    const reader = rs.getReader();
    let count = 0;
    const chunks: Uint8Array[] = [];
    for (;;) {
      const r = await reader.read();
      if (r.done) break;
      count++;
      chunks.push(r.value);
    }
    expect(count).toBeGreaterThan(3);
    const buf = await readAll(new Blob(chunks as BlobPart[]).stream());
    const x = await oracle(buf);
    expect(x.worksheets[0]!.getCell('B20000').value).toBe('row 19999');
    expect(x.worksheets[0]!.rowCount).toBe(20000);
  });
});

describe('createXlsxWriter', () => {
  for (const strings of ['shared', 'inline'] as const) {
    it(`round-trips through exceljs (${strings} strings, sink)`, async () => {
      await checkOracle(await buildStreaming({ strings }), true);
    });
  }

  it('works with the readable stream and async sinks', async () => {
    const w = createXlsxWriter({ profile: 'max-speed', flushRows: 100 });
    const done = readAll(w.readable!);
    const s = w.addWorksheet('R', {
      columns: [
        { header: 'n', key: 'n', width: 5 },
        { header: 's', key: 's' },
      ],
    });
    async function* gen() {
      for (let i = 0; i < 5000; i++) yield { n: i, s: `v${i}` };
    }
    await s.addRows(gen());
    s.mergeCells(1, 3, 1, 4);
    await s.commit();
    await w.end();
    const x = await oracle(await done);
    const xs = x.getWorksheet('R')!;
    expect(xs.rowCount).toBe(5001);
    expect(xs.getCell('A5001').value).toBe(4999);
    expect(xs.getCell('B5001').value).toBe('v4999');
    expect(xs.getColumn(1).width).toBe(5);
    expect(xs.getCell('D1').isMerged).toBe(true);

    // async sink with delays: output must be delivered in order
    const got: Uint8Array[] = [];
    const w2 = createXlsxWriter({
      sink: async (c) => {
        await new Promise((r) => setTimeout(r, 0));
        got.push(c);
      },
    });
    const s2 = w2.addWorksheet('A');
    for (let i = 0; i < 3000; i++) {
      s2.addRow([i, 'x']);
      if (i % 1000 === 0) await s2.flush();
    }
    await w2.end();
    const x2 = await oracle(await readAll(new Blob(got as BlobPart[]).stream()));
    expect(x2.worksheets[0]!.getCell('A3000').value).toBe(2999);
  });

  it('returns absolute row numbers and enforces sequential sheets', async () => {
    const w = createXlsxWriter({ sink: () => {}, flushRows: 3 });
    const a = w.addWorksheet('a', { columns: [{ header: 'h' }] });
    expect(a.addRow([1])).toBe(2);
    for (let i = 0; i < 10; i++) a.addRow([i]);
    expect(a.addRow([99])).toBe(13);
    const b = w.addWorksheet('b');
    expect(a.committed).toBe(true);
    expect(() => a.addRow([1])).toThrow(/committed/);
    b.addRow([1]);
    await w.end();
    expect(() => w.addWorksheet('c')).toThrow(/ended/);
  });

  it('propagates sink errors', async () => {
    const w = createXlsxWriter({
      sink: async () => {
        throw new Error('disk full');
      },
    });
    const s = w.addWorksheet('a');
    for (let i = 0; i < 100; i++) s.addRow([i]);
    await expect(w.end()).rejects.toThrow('disk full');
  });

  it('keeps heap roughly flat for 300k rows x 10 cols (inline strings)', async () => {
    // expose a GC function even when node was not started with --expose-gc
    setFlagsFromString('--expose-gc');
    const g = runInNewContext('gc') as () => void;
    let bytes = 0;
    const w = createXlsxWriter({
      profile: 'low-memory',
      compression: 1,
      sink: (c) => void (bytes += c.length),
    });
    const s = w.addWorksheet('Big');
    const row: (string | number)[] = new Array(10);
    const samples: number[] = [];
    for (let r = 0; r < 300_000; r++) {
      for (let c = 0; c < 10; c++) row[c] = c % 2 === 0 ? r * 10 + c : `s${r}-${c}`;
      s.addRow(row);
      if (r % 50_000 === 49_999) {
        g();
        samples.push(process.memoryUsage().heapUsed);
      }
    }
    await s.commit();
    await w.end();
    const mb = samples.map((b) => Math.round(b / 1048576));
    console.log(
      `streaming 300k x 10: output ${(bytes / 1048576).toFixed(1)} MB, heapUsed samples MB: ${mb.join(', ')}`,
    );
    expect(bytes).toBeGreaterThan(1_000_000);
    // loose: growth between the first and last sample stays small vs. the ~100MB+ an
    // in-memory workbook of this size would need
    expect(samples[samples.length - 1]! - samples[0]!).toBeLessThan(64 * 1048576);
  }, 60_000);
});

describe('native CompressionStream path', () => {
  it('is used in this runtime', () => {
    expect(nativeDeflateAvailable()).toBe(true);
  });

  it('xlsxStream / createXlsxWriter output opens in our readers, fflate and exceljs', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('N');
    for (let r = 0; r < 3000; r++) ws.addRow([r, `s${r % 50}`, r % 2 === 0, r * 1.5]);
    const a = await readAll(xlsxStream(wb));
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c), flushRows: 100 });
    const s = w.addWorksheet('N');
    for (let r = 0; r < 3000; r++) s.addRow([r, `s${r % 50}`, r % 2 === 0, r * 1.5]);
    await w.end();
    const b = await readAll(new Blob(chunks as BlobPart[]).stream());
    for (const buf of [a, b]) {
      // local headers use data descriptors (flag bit 3), deflate method 8
      const dv = new DataView(buf.buffer, buf.byteOffset);
      expect(dv.getUint32(0, true)).toBe(0x04034b50);
      expect(dv.getUint16(6, true) & 8).toBe(8);
      expect(dv.getUint16(8, true)).toBe(8);
      expect(rawUnzipSync(buf)['xl/worksheets/sheet1.xml']!.length).toBeGreaterThan(50_000);
      const ours = readXlsx(buf).getWorksheet('N')!;
      expect(ours.getValue('B2999')).toBe('s48');
      let n = 0;
      for await (const row of xlsxRows(buf)) n += row.values.length;
      expect(n).toBe(12_000);
      const x = (await oracle(buf)).getWorksheet('N')!;
      expect(x.getCell('A3000').value).toBe(2999);
      expect(x.getCell('D3000').value).toBe(2999 * 1.5);
    }
  });
});
