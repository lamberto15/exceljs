/// <reference types="node" />
// Feature extension points (src/xlsx/features): ordered sheet emitters, extra parts/rels/content
// types in every writer, sheet reader delegation + related-part loading, dxfs, and the two
// reference features (sheet protection, print setup) checked against exceljs.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { hashPassword } from '../src/xlsx/features/index.js';
import { READERS, WRITERS } from '../src/xlsx/features/registry.js';
import { type FeatureReader, type FeatureWriter, SheetSlot } from '../src/xlsx/features/types.js';
import { readXlsx, xlsxRows } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync } from '../src/zip/index.js';

// ---- a test-only feature exercising every hook (registered before first use; vitest gives
// each test file its own module instances, so the registry of other files is untouched)

const DUMMY_URI = '{00000000-TEST-4000-8000-000000000001}';
const seen: { sheets: string[]; book?: string | undefined } = { sheets: [] };
const dummyWriter: FeatureWriter = {
  name: 'dummy',
  sheet: {
    namespaces: ['xmlns:dm="urn:dummy"'],
    emitters: [
      {
        slot: SheetSlot.extLst,
        emit(c) {
          const n = c.nextIndex('thing');
          c.addPart(
            `xl/things/thing${n}.xml`,
            `<thing sheet="${c.ws.name}"/>`,
            'application/x-thing+xml',
          );
          c.addPart(`xl/things/blob${n}.bin`, new Uint8Array([1, 2, 3]));
          c.addDefault('bin', 'application/octet-stream');
          const rid = c.addRel('http://example.com/rel/thing', `../things/thing${n}.xml`);
          return `<ext uri="${DUMMY_URI}"><dm:thing r:id="${rid}"/></ext>`;
        },
      },
    ],
  },
  book: {
    write(c) {
      c.addPart('xl/bookthing.xml', '<bookthing/>', 'application/x-bookthing+xml');
      c.addRel('http://example.com/rel/bookthing', 'bookthing.xml');
    },
  },
};
const dummyReader: FeatureReader = {
  name: 'dummy',
  sheet: {
    elements: ['extLst/ext'],
    open(ctx, _el, a) {
      if (a.uri !== DUMMY_URI) return undefined;
      return {
        open(name, at) {
          if (name === 'thing') ctx.scratch.set(dummyReader, at['r:id']);
        },
      };
    },
    finish(ctx) {
      const rid = ctx.scratch.get(dummyReader) as string | undefined;
      const rel = rid ? ctx.rels?.get(rid) : undefined;
      if (rel) seen.sheets.push(ctx.readText(rel.path)!);
    },
  },
  book(c) {
    const rels = c.relsOf('xl/workbook.xml');
    for (const r of rels?.values() ?? []) {
      if (r.type.endsWith('/bookthing')) seen.book = c.readText(r.path);
    }
  },
};
(WRITERS as FeatureWriter[]).push(dummyWriter);
(READERS as FeatureReader[]).push(dummyReader);

// ---- helpers

const dec = new TextDecoder();
function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

async function readAll(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(rs).arrayBuffer());
}

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

/** Child element names of <worksheet> in document order. */
function worksheetChildren(xml: string): string[] {
  const body = xml.slice(xml.indexOf('<worksheet'));
  const out: string[] = [];
  let depth = 0;
  for (const m of body.matchAll(/<(\/?)([A-Za-z][\w:]*)[^>]*?(\/?)>/g)) {
    if (m[1]) depth--;
    else {
      if (depth === 1) out.push(m[2]!);
      if (!m[3]) depth++;
    }
  }
  return out;
}

const ORDER = Object.keys(SheetSlot);

function fullSheet(wb: Workbook): void {
  const ws = wb.addWorksheet('Full', { tabColor: { argb: 'FF00FF00' } });
  ws.columns = [{ header: 'A', width: 10 }, { header: 'B' }];
  ws.addRow([1, { text: 'link', hyperlink: 'https://example.com' }]);
  ws.mergeCells('C3:D4');
  ws.autoFilter = 'A1:B2';
  ws.freeze(1);
  ws.protection = { password: 'pw', spinCount: 1000, formatColumns: true, sort: true };
  ws.pageSetup = {
    orientation: 'landscape',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.5, right: 0.5 },
    horizontalCentered: true,
    showGridLines: true,
    firstPageNumber: 3,
  };
  ws.headerFooter = { differentFirst: true, oddHeader: '&C&BReport <1>', firstFooter: 'Page &P' };
}

async function allWriters(build: (wb: Workbook) => void): Promise<[string, Uint8Array][]> {
  const wb = new Workbook();
  build(wb);
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  for (const src of wb.worksheets) {
    const s = w.addWorksheet(src.name, { tabColor: src.tabColor, columns: src.columns });
    const ws = s.worksheet;
    ws.views = src.views;
    ws.autoFilter = src.autoFilter;
    for (const m of src.merges) ws.mergeCells(m);
    ws.protection = src.protection;
    ws.pageSetup = src.pageSetup;
    ws.headerFooter = src.headerFooter;
    src.eachRow((row, n) => {
      if (n > 1) s.addRow(row.values.slice(1));
    });
  }
  await w.end();
  return [
    ['writeXlsx', writeXlsx(wb)],
    ['xlsxStream', await readAll(xlsxStream(wb))],
    ['createXlsxWriter', new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())],
  ];
}

// ---- tests

describe('ordered sheet emitters', () => {
  it('emit worksheet children in CT_Worksheet order in every writer', async () => {
    for (const [name, buf] of await allWriters(fullSheet)) {
      const xml = parts(buf).get('xl/worksheets/sheet1.xml')!;
      const kids = worksheetChildren(xml);
      expect(kids, name).toEqual([
        'sheetPr',
        ...(name === 'createXlsxWriter' ? [] : ['dimension']),
        'sheetViews',
        'sheetFormatPr',
        'cols',
        'sheetData',
        'sheetProtection',
        'autoFilter',
        'mergeCells',
        'hyperlinks',
        'printOptions',
        'pageMargins',
        'pageSetup',
        'headerFooter',
        'extLst',
      ]);
      const idx = kids.map((k) => ORDER.indexOf(k));
      expect(
        [...idx].sort((a, b) => a - b),
        name,
      ).toEqual(idx);
      // sheetPr children: tabColor before pageSetUpPr
      expect(xml, name).toMatch(
        /<sheetPr><tabColor rgb="FF00FF00"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/,
      );
      expect(xml, name).toContain('xmlns:dm="urn:dummy"');
    }
  });
});

describe('extra parts, rels and content types', () => {
  it('are written by every writer and read back through the reader hooks', async () => {
    for (const [name, buf] of await allWriters((wb) => {
      fullSheet(wb);
      wb.addWorksheet('Second').setCell('A1', 'x');
    })) {
      const p = parts(buf);
      const names = [...p.keys()];
      expect(p.get('xl/things/thing1.xml'), name).toBe('<thing sheet="Full"/>');
      expect(p.get('xl/things/thing2.xml'), name).toBe('<thing sheet="Second"/>');
      expect(p.has('xl/things/blob1.bin'), name).toBe(true);
      // sheet parts are followed by their feature parts and rels
      const i = names.indexOf('xl/worksheets/sheet1.xml');
      expect(names.slice(i, i + 4), name).toEqual([
        'xl/worksheets/sheet1.xml',
        'xl/things/thing1.xml',
        'xl/things/blob1.bin',
        'xl/worksheets/_rels/sheet1.xml.rels',
      ]);
      const rels = p.get('xl/worksheets/_rels/sheet1.xml.rels')!;
      // hyperlink and feature rels share one id space
      expect(rels).toContain(
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"',
      );
      expect(rels).toContain(
        'Id="rId2" Type="http://example.com/rel/thing" Target="../things/thing1.xml"',
      );
      const ct = p.get('[Content_Types].xml')!;
      expect(ct).toContain('<Default Extension="bin" ContentType="application/octet-stream"/>');
      expect(ct).toContain(
        '<Override PartName="/xl/things/thing2.xml" ContentType="application/x-thing+xml"/>',
      );
      expect(ct).toContain(
        '<Override PartName="/xl/bookthing.xml" ContentType="application/x-bookthing+xml"/>',
      );
      expect(p.get('xl/_rels/workbook.xml.rels')).toMatch(
        /Id="rId\d+" Type="http:\/\/example.com\/rel\/bookthing"/,
      );

      seen.sheets = [];
      seen.book = undefined;
      readXlsx(buf);
      expect(seen.sheets, name).toEqual(['<thing sheet="Full"/>', '<thing sheet="Second"/>']);
      expect(seen.book, name).toBe('<bookthing/>');
      // streaming rows ignore features but still parse
      let n = 0;
      for await (const r of xlsxRows(buf)) n += r.cells.length;
      expect(n, name).toBeGreaterThan(0);
      // and Excel-compatible readers still open the file
      expect((await oracle(buf)).worksheets.length, name).toBe(2);
    }
  });
});

describe('sheet protection', () => {
  it('hash matches exceljs (ECMA-376 agile SHA-512)', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('P');
    await xs.protect('s3cr3t ü', { spinCount: 2000, formatCells: true });
    const sp = (xs as unknown as { sheetProtection: Record<string, string> }).sheetProtection;
    expect(hashPassword('s3cr3t ü', sp.saltValue!, 2000)).toBe(sp.hashValue);
    // exceljs file → ours: hash kept verbatim and written back unchanged
    const buf = new Uint8Array(await x.xlsx.writeBuffer());
    const ws = readXlsx(buf).getWorksheet('P')!;
    expect(ws.protection).toMatchObject({
      algorithmName: 'SHA-512',
      hashValue: sp.hashValue,
      saltValue: sp.saltValue,
      spinCount: 2000,
      formatCells: true,
    });
    const again = parts(writeXlsx(ws.workbook)).get('xl/worksheets/sheet1.xml')!;
    expect(again).toContain(`hashValue="${sp.hashValue}"`);
  });

  it('round-trips through ours and exceljs in every writer', async () => {
    for (const [name, buf] of await allWriters(fullSheet)) {
      const ours = readXlsx(buf).getWorksheet('Full')!;
      const p = ours.protection!;
      expect(p.algorithmName, name).toBe('SHA-512');
      expect(p.spinCount, name).toBe(1000);
      expect(hashPassword('pw', p.saltValue!, 1000), name).toBe(p.hashValue);
      expect(p.formatColumns && p.sort, name).toBe(true);
      expect(p.formatCells, name).toBeUndefined();
      const x = (await oracle(buf)).getWorksheet('Full')!;
      const xp = (x as unknown as { sheetProtection: Record<string, unknown> }).sheetProtection;
      expect(xp, name).toMatchObject({
        sheet: true,
        algorithmName: 'SHA-512',
        hashValue: p.hashValue,
        spinCount: 1000,
        formatColumns: true,
        sort: true,
      });
    }
  });

  it('no password / unprotected', () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    a.protection = { selectLockedCells: false };
    wb.addWorksheet('B');
    const p = parts(writeXlsx(wb));
    expect(p.get('xl/worksheets/sheet1.xml')).toContain(
      '<sheetProtection sheet="1" selectLockedCells="1"/>',
    );
    expect(p.get('xl/worksheets/sheet2.xml')).not.toContain('sheetProtection');
    expect(readXlsx(writeXlsx(wb)).getWorksheet('A')!.protection).toEqual({
      selectLockedCells: false,
    });
  });

  it('salt is stable across writes of the same protection object', () => {
    const wb = new Workbook();
    wb.addWorksheet('A').protection = { password: 'x', spinCount: 10 };
    const salt = (b: Uint8Array) =>
      /saltValue="([^"]+)"/.exec(parts(b).get('xl/worksheets/sheet1.xml')!)![1];
    expect(salt(writeXlsx(wb))).toBe(salt(writeXlsx(wb)));
  });
});

describe('print setup', () => {
  it('round-trips through ours and exceljs in every writer', async () => {
    for (const [name, buf] of await allWriters(fullSheet)) {
      const ws = readXlsx(buf).getWorksheet('Full')!;
      expect(ws.pageSetup, name).toEqual({
        orientation: 'landscape',
        paperSize: 9,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
        horizontalCentered: true,
        showGridLines: true,
        firstPageNumber: 3,
      });
      expect(ws.headerFooter, name).toEqual({
        differentFirst: true,
        oddHeader: '&C&BReport <1>',
        firstFooter: 'Page &P',
      });
      const x = (await oracle(buf)).getWorksheet('Full')!;
      expect(x.pageSetup, name).toMatchObject({
        orientation: 'landscape',
        paperSize: 9,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        horizontalCentered: true,
        showGridLines: true,
        firstPageNumber: 3,
        margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      });
      expect(x.headerFooter, name).toMatchObject({
        differentFirst: true,
        oddHeader: '&C&BReport <1>',
        firstFooter: 'Page &P',
      });
    }
  });

  it('reads exceljs print settings', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('S', {
      pageSetup: {
        orientation: 'landscape',
        scale: 80,
        blackAndWhite: true,
        verticalCentered: true,
      },
      headerFooter: { oddFooter: '&LLeft&RRight', differentOddEven: true, evenFooter: 'even' },
    });
    xs.getCell('A1').value = 1;
    const ws = readXlsx(new Uint8Array(await x.xlsx.writeBuffer())).getWorksheet('S')!;
    expect(ws.pageSetup).toMatchObject({
      orientation: 'landscape',
      scale: 80,
      blackAndWhite: true,
      verticalCentered: true,
    });
    expect(ws.headerFooter).toEqual({
      differentOddEven: true,
      oddFooter: '&LLeft&RRight',
      evenFooter: 'even',
    });
  });

  it('default page margins are always written', () => {
    const wb = new Workbook();
    wb.addWorksheet('A');
    expect(parts(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')).toContain(
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    );
  });
});

describe('dxfs', () => {
  it('intern, write after cellStyles, read back', async () => {
    const wb = new Workbook();
    const a = wb.styles.internDxf({
      font: { bold: true, color: { argb: 'ff9c0006' } },
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
    });
    const b = wb.styles.internDxf({ numFmt: '0.000%', border: { top: { style: 'thin' } } });
    expect(
      wb.styles.internDxf({
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
        font: { color: { argb: 'FF9C0006' }, bold: true },
      }),
    ).toBe(a);
    expect([a, b]).toEqual([0, 1]);
    wb.addWorksheet('S').setCell('A1', 1);
    const buf = writeXlsx(wb);
    const styles = parts(buf).get('xl/styles.xml')!;
    expect(styles.indexOf('</cellStyles><dxfs count="2"><dxf><font><b/>')).toBeGreaterThan(0);
    expect(styles.indexOf('</dxfs><tableStyles')).toBeGreaterThan(0);
    expect(styles).toMatch(/<dxf><numFmt numFmtId="\d+" formatCode="0.000%"\/><border>/);
    const back = readXlsx(buf).styles.dxfs;
    expect(back).toEqual(wb.styles.dxfs);
    await oracle(buf); // still a valid styles part for exceljs
  });
});
