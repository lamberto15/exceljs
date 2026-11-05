/// <reference types="node" />
// Sheet views (one per Excel window), sheet properties (<sheetPr> / <sheetFormatPr>) and the
// language / revision document properties (#82), checked against exceljs.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { createXlsxWriter, readXlsx, Workbook, writeXlsx } from '../src/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
const part = (buf: Uint8Array, p: string): string => dec.decode(unzipSync(buf).get(p)!);
const SHEET1 = 'xl/worksheets/sheet1.xml';
const rezip = (m: Map<string, Uint8Array>): Uint8Array =>
  zipSync(
    [...m].map(([path, data]) => ({ path, data })),
    6,
  );

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

async function streamed(f: (w: ReturnType<typeof createXlsxWriter>) => void): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  f(w);
  await w.end();
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

/** Index of each tag's first occurrence; asserts they appear in this order. */
function expectOrder(xml: string, tags: string[]): void {
  const at = tags.map((t) => xml.search(new RegExp(`<${t}[\\s/>]`)));
  for (let i = 0; i < tags.length; i++) expect(at[i], tags[i]).toBeGreaterThanOrEqual(0);
  expect([...at].sort((a, b) => a - b)).toEqual(at);
}

describe('several sheet views', () => {
  it('writes one <sheetView> per view, each pointing at its window; windows are added', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.views = [{ state: 'frozen', ySplit: 1 }, { zoomScale: 50 }];
    wb.addWorksheet('T').getCell('A1').value = 2;
    const buf = writeXlsx(wb);
    const s = part(buf, SHEET1);
    const views = s.match(/<sheetView\b[^>]*>/g)!;
    expect(views).toHaveLength(2);
    expect(views[0]).toContain('workbookViewId="0"');
    expect(views[0]).toContain('tabSelected="1"');
    expect(views[1]).toContain('workbookViewId="1"');
    expect(views[1]).toContain('zoomScale="50"');
    expect(s.match(/<sheetViews>/g)).toHaveLength(1);
    expectOrder(s, ['sheetViews', 'sheetFormatPr', 'sheetData']);
    // every workbookViewId has a <workbookView>
    const book = part(buf, 'xl/workbook.xml');
    expect(book.match(/<workbookView\b/g)).toHaveLength(2);
    // sheet T has one view: window 0 only
    expect(part(buf, 'xl/worksheets/sheet2.xml').match(/<sheetView\b/g)).toHaveLength(1);

    const back = readXlsx(buf);
    expect(back.getWorksheet('S')!.views).toEqual([
      { state: 'frozen', ySplit: 1, topLeftCell: 'A2', tabSelected: true },
      { state: 'normal', zoomScale: 50, tabSelected: true },
    ]);
    expect(back.views).toHaveLength(2);
    const x = (await oracle(buf)).getWorksheet('S')!;
    expect(x.views).toHaveLength(2);
    expect(x.views[1]).toMatchObject({ workbookViewId: 1, zoomScale: 50 });
  });

  it('explicit workbookViewId is honoured and read back; wb.views entries are all written', () => {
    const wb = new Workbook();
    wb.views = [
      { x: 0, y: 0, width: 10000, height: 20000 },
      { x: 1, y: 2, width: 10001, height: 20002, visibility: 'hidden', activeTab: 1 },
    ];
    wb.addWorksheet('one').views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
    wb.addWorksheet('two').views = [{ workbookViewId: 1, state: 'split', xSplit: 2000 }];
    const buf = writeXlsx(wb);
    expect(part(buf, 'xl/worksheets/sheet2.xml')).toContain('<sheetView workbookViewId="1"');
    const book = part(buf, 'xl/workbook.xml');
    expect(book.match(/<workbookView\b[^>]*>/g)![1]).toBe(
      '<workbookView visibility="hidden" xWindow="1" yWindow="2" windowWidth="10001" windowHeight="20002" firstSheet="0" activeTab="1"/>',
    );
    const back = readXlsx(buf);
    expect(back.views).toEqual([
      { x: 0, y: 0, width: 10000, height: 20000 },
      {
        x: 1,
        y: 2,
        width: 10001,
        height: 20002,
        visibility: 'hidden',
        activeTab: 1,
        firstSheet: 0,
      },
    ]);
    expect(back.getWorksheet('two')!.views).toEqual([
      { workbookViewId: 1, state: 'split', xSplit: 2000 },
    ]);
  });

  // #102
  it('style / showRuler / zoomScaleNormal / activePane round-trip and match exceljs', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.views = [
      {
        style: 'pageLayout',
        showRuler: false,
        zoomScale: 80,
        zoomScaleNormal: 120,
        state: 'frozen',
        xSplit: 1,
        ySplit: 2,
        activePane: 'topRight',
      },
    ];
    const buf = writeXlsx(wb);
    const s = part(buf, SHEET1);
    expect(s).toContain('view="pageLayout"');
    expect(s).toContain('showRuler="0"');
    expect(s).toContain('zoomScaleNormal="120"');
    expect(s).toContain('activePane="topRight"');
    expect(s).toContain('<selection pane="topRight"');
    expect(readXlsx(buf).worksheets[0]!.views[0]).toEqual({
      state: 'frozen',
      xSplit: 1,
      ySplit: 2,
      topLeftCell: 'B3',
      activePane: 'topRight',
      style: 'pageLayout',
      showRuler: false,
      zoomScale: 80,
      zoomScaleNormal: 120,
      tabSelected: true,
    });
    const x = (await oracle(buf)).worksheets[0]!.views[0]!;
    expect(x).toMatchObject({
      style: 'pageLayout',
      showRuler: false,
      zoomScale: 80,
      zoomScaleNormal: 120,
    }); // (exceljs does not read activePane back for frozen panes)
  });

  it('reads exceljs-written view fields; a derived activePane is not kept', async () => {
    const x = new ExcelJS.Workbook();
    x
      .addWorksheet('S', {
        views: [
          { style: 'pageBreakPreview', showRuler: false, zoomScaleNormal: 90 },
          // biome-ignore lint/suspicious/noExplicitAny: exceljs typings lack activePane on frozen views
          { state: 'frozen', ySplit: 1, activePane: 'bottomLeft' } as any,
        ],
      })
      .getCell('A1').value = 1;
    const buf = new Uint8Array(await x.xlsx.writeBuffer());
    const [a, b] = readXlsx(buf).worksheets[0]!.views;
    expect(a).toMatchObject({ style: 'pageBreakPreview', showRuler: false, zoomScaleNormal: 90 });
    expect(b!.activePane).toBeUndefined();
  });

  it('padded windows copy the first wb.views entry', () => {
    const wb = new Workbook();
    wb.views = [{ width: 20000, height: 10000 }];
    wb.addWorksheet('S').views = [{}, {}, {}];
    const book = part(writeXlsx(wb), 'xl/workbook.xml');
    const v = book.match(/<workbookView\b[^>]*>/g)!;
    expect(v).toHaveLength(3);
    for (const e of v) expect(e).toContain('windowWidth="20000" windowHeight="10000"');
  });

  it('streaming writer: views from addWorksheet options', async () => {
    const buf = await streamed((w) => {
      const s = w.addWorksheet('S', { views: [{ showGridLines: false }, { zoomScale: 75 }] });
      s.addRow([1]);
    });
    const s = part(buf, SHEET1);
    expect(s.match(/<sheetView\b/g)).toHaveLength(2);
    expect(part(buf, 'xl/workbook.xml').match(/<workbookView\b/g)).toHaveLength(2);
    expect(readXlsx(buf).worksheets[0]!.views[1]).toMatchObject({ zoomScale: 75 });
  });

  it('a file with several windows round-trips its views', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const src = unzipSync(writeXlsx(wb));
    const s = dec.decode(src.get(SHEET1)!);
    src.set(
      SHEET1,
      enc.encode(
        s.replace(
          /<sheetViews>.*<\/sheetViews>/,
          '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/><sheetView zoomScale="80" workbookViewId="1"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>',
        ),
      ),
    );
    const b = dec.decode(src.get('xl/workbook.xml')!);
    src.set(
      'xl/workbook.xml',
      enc.encode(b.replace('</bookViews>', '<workbookView xWindow="5" yWindow="5"/></bookViews>')),
    );
    const read = readXlsx(rezip(src));
    expect(read.worksheets[0]!.views).toEqual([
      { state: 'normal', tabSelected: true },
      { state: 'normal', zoomScale: 80, activeCell: 'B2' },
    ]);
    const out = writeXlsx(read);
    expect(part(out, SHEET1).match(/<sheetView\b/g)).toHaveLength(2);
    expect(part(out, 'xl/workbook.xml').match(/<workbookView\b/g)).toHaveLength(2);
  });
});

describe('sheet properties', () => {
  it('addWorksheet({ properties }) writes sheetPr / sheetFormatPr in schema order', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('P', {
      properties: {
        tabColor: { argb: 'FF00B050' },
        defaultRowHeight: 20,
        defaultColWidth: 14,
        outlineLevelRow: 2,
        outlineLevelCol: 1,
        dyDescent: 0.25,
      },
    });
    ws.getCell('A1').value = 'x';
    expect(ws.tabColor).toEqual({ argb: 'FF00B050' });
    const buf = writeXlsx(wb);
    const s = part(buf, SHEET1);
    expect(s).toMatch(
      /^<\?xml[^>]*>\n<worksheet xmlns="[^"]+" xmlns:r="[^"]+" xmlns:mc="http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006" mc:Ignorable="x14ac" xmlns:x14ac="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2009\/9\/ac">/,
    );
    expect(s).toContain('<sheetPr><tabColor rgb="FF00B050"/></sheetPr>');
    expect(s.match(/<sheetFormatPr[^>]*\/>/)![0]).toBe(
      '<sheetFormatPr defaultColWidth="14" defaultRowHeight="20" customHeight="1" outlineLevelRow="2" outlineLevelCol="1" x14ac:dyDescent="0.25"/>',
    );
    expectOrder(s, ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'sheetData']);

    const back = readXlsx(buf).getWorksheet('P')!;
    // outline levels are derived from rows / columns on write, not kept on read
    expect(back.properties).toEqual({
      tabColor: { argb: 'FF00B050' },
      defaultRowHeight: 20,
      defaultColWidth: 14,
      dyDescent: 0.25,
    });
    const x = (await oracle(buf)).getWorksheet('P')!;
    expect(x.properties).toMatchObject({
      tabColor: { argb: 'FF00B050' },
      defaultRowHeight: 20,
      defaultColWidth: 14,
      outlineLevelRow: 2,
      outlineLevelCol: 1,
      dyDescent: 0.25,
    });
  });

  it('ws.tabColor and ws.properties.tabColor are the same value', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.properties.tabColor = { argb: 'FFFF0000' };
    expect(ws.tabColor).toEqual({ argb: 'FFFF0000' });
    ws.tabColor = { theme: 4 };
    expect(ws.properties.tabColor).toEqual({ theme: 4 });
    const s = part(writeXlsx(ws.workbook), SHEET1);
    expect(s.match(/<tabColor/g)).toHaveLength(1);
  });

  it('outline levels: the higher of the property and the data', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setRowProps(2, { outlineLevel: 3 });
    ws.getColumn(2).outlineLevel = 1;
    ws.properties.outlineLevelRow = 1;
    ws.properties.outlineLevelCol = 4;
    const el = part(writeXlsx(wb), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el).toContain('outlineLevelRow="3"');
    expect(el).toContain('outlineLevelCol="4"');
  });

  it('defaultRowHeight equal to the default font height is not custom; no dyDescent → no mc', () => {
    const wb = new Workbook();
    wb.addWorksheet('S', { properties: { defaultRowHeight: 15 } });
    const s = part(writeXlsx(wb), SHEET1);
    expect(s.match(/<sheetFormatPr[^>]*\/>/)![0]).toBe('<sheetFormatPr defaultRowHeight="15"/>');
    expect(s).not.toContain('mc:Ignorable');
  });

  it('streaming writer: properties are written in the sheet head (outline levels too)', async () => {
    const buf = await streamed((w) => {
      const s = w.addWorksheet('S', {
        properties: { outlineLevelRow: 1, defaultRowHeight: 18, dyDescent: 0.3 },
        tabColor: { argb: 'FF0070C0' },
      });
      s.addRow(['a'], undefined, { outlineLevel: 1 });
    });
    const s = part(buf, SHEET1);
    expect(s.match(/<sheetFormatPr[^>]*\/>/)![0]).toBe(
      '<sheetFormatPr defaultRowHeight="18" customHeight="1" outlineLevelRow="1" x14ac:dyDescent="0.3"/>',
    );
    expect(s).toContain('<tabColor rgb="FF0070C0"/>');
    expect(s).toContain('mc:Ignorable="x14ac"');
    const x = (await oracle(buf)).getWorksheet('S')!;
    expect(x.properties.outlineLevelRow).toBe(1);
  });

  it('defaultColWidth is the width of <col> records without their own width', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S', { properties: { defaultColWidth: 20 } });
    ws.getColumn(2).hidden = true;
    expect(part(writeXlsx(wb), SHEET1)).toContain('<col min="2" max="2" width="20" hidden="1"/>');
  });

  it('a read file: modelled attributes and kept ones are written once', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const src = unzipSync(writeXlsx(wb));
    const s = dec.decode(src.get(SHEET1)!);
    src.set(
      SHEET1,
      enc.encode(
        s.replace(
          /<sheetFormatPr[^>]*\/>/,
          '<sheetFormatPr baseColWidth="10" defaultRowHeight="16" customHeight="1" zeroHeight="1" x14ac:dyDescent="0.2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"/>',
        ),
      ),
    );
    const read = readXlsx(rezip(src));
    expect(read.worksheets[0]!.properties).toEqual({ defaultRowHeight: 16, dyDescent: 0.2 });
    read.worksheets[0]!.properties.defaultRowHeight = 22;
    const el = part(writeXlsx(read), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el).toBe(
      '<sheetFormatPr defaultRowHeight="22" customHeight="1" x14ac:dyDescent="0.2" baseColWidth="10" zeroHeight="1"/>',
    );
  });
});

describe('document properties: language, revision', () => {
  it('written to core.xml and read back; exceljs agrees', async () => {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    wb.language = 'en-US';
    wb.revision = 3;
    const buf = writeXlsx(wb);
    const core = part(buf, 'docProps/core.xml');
    expect(core).toContain('<dc:language>en-US</dc:language>');
    expect(core).toContain('<cp:revision>3</cp:revision>');
    const back = readXlsx(buf);
    expect(back.language).toBe('en-US');
    expect(back.revision).toBe(3);
    // exceljs's typings lack language / revision
    const x = (await oracle(buf)) as unknown as { language: string; revision: number };
    expect(x.language).toBe('en-US');
    expect(x.revision).toBe(3);
  });

  it('streaming writer writes them too', async () => {
    const buf = await streamed((w) => {
      w.workbook.language = 'de-DE';
      w.workbook.revision = 7;
      w.addWorksheet('S').addRow([1]);
    });
    const back = readXlsx(buf);
    expect(back.language).toBe('de-DE');
    expect(back.revision).toBe(7);
  });
});
