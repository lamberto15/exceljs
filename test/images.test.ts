/// <reference types="node" />
// Pictures (features/images): workbook media + per-sheet drawings, checked against exceljs.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { addImage, addNote, type ImagePlacement, placeImage } from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
const b64 = (s: string): Uint8Array => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
// 1x1 PNG / GIF
const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const PNG = b64(PNG_B64);
const GIF = b64('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7');

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

const SHEET1: ImagePlacement[] = [
  { imageId: 0, anchor: { from: { col: 1, row: 1 }, to: { col: 4, row: 6 } } },
  {
    imageId: 0,
    anchor: { from: { col: 5, row: 1, colOff: 3, rowOff: 2 }, ext: { width: 120, height: 80 } },
    hyperlink: 'https://example.com/?a=1&b=2',
    tooltip: 'Go',
    description: 'alt <text>',
  },
  {
    imageId: 1,
    anchor: { from: { col: 0, row: 8 }, to: { col: 2, row: 12, colOff: 5 } },
    editAs: 'oneCell',
    name: 'Logo',
  },
  { imageId: 1, anchor: { pos: { x: 10, y: 20 }, ext: { width: 30, height: 40 } } },
];
const SHEET2: ImagePlacement[] = [
  { imageId: 0, anchor: { from: { col: 2, row: 2 }, to: { col: 3, row: 3 } } },
];

function build(wb: Workbook): void {
  const png = addImage(wb, { data: PNG, extension: 'png' });
  const gif = addImage(wb, {
    base64: `data:image/gif;base64,${btoa(String.fromCharCode(...GIF))}`,
    extension: 'gif',
  });
  expect([png, gif]).toEqual([0, 1]);
  expect(addImage(wb, { base64: PNG_B64, extension: 'png' })).toBe(0); // deduped
  const a = wb.addWorksheet('A');
  a.setCell('A1', 1);
  placeImage(a, png, 'B2:D6');
  placeImage(
    a,
    png,
    { tl: { col: 5, row: 1, colOff: 3, rowOff: 2 }, ext: { width: 120, height: 80 } },
    {
      hyperlink: 'https://example.com/?a=1&b=2',
      tooltip: 'Go',
      description: 'alt <text>',
    },
  );
  placeImage(
    a,
    gif,
    { from: { col: 0, row: 8 }, to: { col: 2, row: 12, colOff: 5 } },
    {
      editAs: 'oneCell',
      name: 'Logo',
    },
  );
  placeImage(a, gif, { pos: { x: 10, y: 20 }, ext: { width: 30, height: 40 } });
  const b = wb.addWorksheet('B');
  b.setCell('A1', 2);
  placeImage(b, png, 'C3');
}

async function allWriters(): Promise<[string, Uint8Array][]> {
  const wb = new Workbook();
  build(wb);
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  w.workbook.images = wb.images;
  for (const src of wb.worksheets) {
    const s = w.addWorksheet(src.name);
    s.addRow([1]);
    s.worksheet.images = src.images;
    await s.commit();
  }
  await w.end();
  return [
    ['writeXlsx', writeXlsx(wb)],
    ['xlsxStream', new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())],
    ['createXlsxWriter', new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())],
  ];
}

describe('images', () => {
  it('writes media once, a drawing per sheet, rels and content types', () => {
    const wb = new Workbook();
    build(wb);
    const files = unzipSync(writeXlsx(wb));
    const media = [...files.keys()].filter((k) => k.startsWith('xl/media/'));
    expect(media.sort()).toEqual(['xl/media/image1.png', 'xl/media/image2.gif']);
    expect(files.get('xl/media/image1.png')).toEqual(PNG);
    const p = parts(writeXlsx(wb));
    const d1 = p.get('xl/drawings/drawing1.xml')!;
    expect(d1).toContain(
      '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col>',
    );
    expect(d1).toContain('<xdr:ext cx="1143000" cy="762000"/>');
    expect(d1).toContain('<xdr:twoCellAnchor editAs="oneCell">');
    expect(d1).toContain('<xdr:absoluteAnchor><xdr:pos x="95250" y="190500"/>');
    expect(d1).toContain('descr="alt &lt;text&gt;"><a:hlinkClick r:id="rId2" tooltip="Go"/>');
    const r1 = p.get('xl/drawings/_rels/drawing1.xml.rels')!;
    expect(r1.match(/relationships\/image"/g)!.length).toBe(2); // one rel per distinct image
    expect(r1).toContain('Target="../media/image1.png"');
    expect(r1).toContain('Target="https://example.com/?a=1&amp;b=2" TargetMode="External"');
    expect(p.get('xl/drawings/_rels/drawing2.xml.rels')).toContain('Target="../media/image1.png"');
    expect(p.get('xl/worksheets/sheet1.xml')).toMatch(
      /<pageMargins[^>]*\/><drawing r:id="rId1"\/>/,
    );
    expect(p.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain(
      'relationships/drawing" Target="../drawings/drawing1.xml"',
    );
    const ct = p.get('[Content_Types].xml')!;
    expect(ct).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(ct).toContain('<Default Extension="gif" ContentType="image/gif"/>');
    expect(ct).toContain(
      '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
    );
  });

  it('round-trips through ours in every writer (shared image read once)', async () => {
    for (const [name, buf] of await allWriters()) {
      const wb = readXlsx(buf);
      expect(wb.images, name).toEqual([
        { id: 0, data: PNG, extension: 'png' },
        { id: 1, data: GIF, extension: 'gif' },
      ]);
      expect(wb.getWorksheet('A')!.images, name).toEqual(SHEET1);
      expect(wb.getWorksheet('B')!.images, name).toEqual(SHEET2);
      const again = readXlsx(writeXlsx(wb));
      expect(again.getWorksheet('A')!.images, name).toEqual(SHEET1);
      expect(again.images!.length, name).toBe(2);
    }
  });

  it('exceljs reads our pictures', async () => {
    for (const [name, buf] of await allWriters()) {
      const x = await oracle(buf);
      const imgs = x.getWorksheet('A')!.getImages();
      // exceljs only knows two/one-cell anchors (the absolute one is ignored)
      expect(imgs.length, name).toBe(3);
      const [i0, i1, i2] = imgs as unknown as {
        imageId: string;
        range: {
          tl: { nativeCol: number; nativeRow: number; nativeColOff: number };
          br?: { nativeCol: number; nativeRow: number };
          ext?: { width: number; height: number };
          editAs?: string;
          hyperlinks?: { hyperlink: string; tooltip: string };
        };
      }[];
      expect([i0!.range.tl.nativeCol, i0!.range.tl.nativeRow], name).toEqual([1, 1]);
      expect([i0!.range.br!.nativeCol, i0!.range.br!.nativeRow], name).toEqual([4, 6]);
      expect(i1!.range.ext, name).toEqual({ width: 120, height: 80 });
      expect(i1!.range.tl.nativeColOff, name).toBe(3 * 9525);
      expect(i1!.range.hyperlinks, name).toEqual({
        hyperlink: 'https://example.com/?a=1&b=2',
        tooltip: 'Go',
      });
      expect(i2!.range.editAs, name).toBe('oneCell');
      const m0 = x.getImage(Number(i0!.imageId));
      expect(m0.extension, name).toBe('png');
      expect(new Uint8Array(m0.buffer as unknown as Uint8Array), name).toEqual(PNG);
      expect(x.getImage(Number(i2!.imageId)).extension, name).toBe('gif');
      const b = x.getWorksheet('B')!.getImages();
      expect(b.length, name).toBe(1);
      expect(b[0]!.imageId, name).toBe(i0!.imageId); // same medium shared by both sheets
    }
  });

  it('reads exceljs pictures and dedupes identical media', async () => {
    const x = new ExcelJS.Workbook();
    // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
    const id = x.addImage({ buffer: Buffer.from(PNG) as any, extension: 'png' });
    // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
    const dup = x.addImage({ buffer: Buffer.from(PNG) as any, extension: 'png' });
    const s1 = x.addWorksheet('S1');
    s1.getCell('A1').value = 1;
    s1.addImage(id, 'B2:D6');
    s1.addImage(id, { tl: { col: 1.5, row: 7 }, ext: { width: 50, height: 60 } });
    s1.addImage(dup, {
      tl: { col: 0, row: 0 },
      ext: { width: 10, height: 10 },
      hyperlinks: { hyperlink: 'https://x.test', tooltip: 'tip' },
    } as unknown as ExcelJS.ImageRange);
    const s2 = x.addWorksheet('S2');
    s2.addImage(dup, 'A1:A1');
    s2.addBackgroundImage(id);
    const wb = readXlsx(new Uint8Array(await x.xlsx.writeBuffer()));
    expect(wb.images).toEqual([{ id: 0, data: PNG, extension: 'png' }]);
    const a = wb.getWorksheet('S1')!.images!;
    expect(a[0]).toEqual({
      imageId: 0,
      anchor: { from: { col: 1, row: 1 }, to: { col: 4, row: 6 } },
      editAs: 'oneCell',
    });
    expect(a[1]!.anchor).toEqual({
      from: { col: 1, row: 7, colOff: 320000 / 9525 }, // exceljs: half of a 640000 EMU column
      ext: { width: 50, height: 60 },
    });
    expect(a[2]).toMatchObject({ imageId: 0, hyperlink: 'https://x.test', tooltip: 'tip' });
    const b = wb.getWorksheet('S2')!;
    expect(b.images!.map((p) => p.imageId)).toEqual([0]);
    expect(b.backgroundImage).toBe(0);
  });

  it('media that is neither placed nor a background is not written', async () => {
    const wb = new Workbook();
    const unused = addImage(wb, { data: PNG, extension: 'png' });
    const placed = addImage(wb, { data: GIF, extension: 'gif' });
    const bg = addImage(wb, { data: new Uint8Array([1, 2, 3]), extension: 'bmp' });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 1);
    placeImage(ws, placed, 'B2:C3');
    ws.backgroundImage = bg;
    expect(unused).toBe(0);
    const bufs = [writeXlsx(wb), new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())];
    for (const buf of bufs) {
      const p = parts(buf);
      // part names keep the position in wb.images (stable across sheets): image1 is skipped
      expect([...p.keys()].filter((k) => k.startsWith('xl/media/')).sort()).toEqual([
        'xl/media/image2.gif',
        'xl/media/image3.bmp',
      ]);
      expect(p.get('[Content_Types].xml')).not.toContain('Extension="png"');
      expect(p.get('[Content_Types].xml')).toContain('Extension="gif"');
      const back = readXlsx(buf);
      expect(back.images!.map((i) => i.extension).sort()).toEqual(['bmp', 'gif']);
      expect(back.getWorksheet('S')!.images!.length).toBe(1);
      expect(back.getWorksheet('S')!.backgroundImage).toBeDefined();
      await oracle(buf);
    }
  });

  it('background image round trip and exceljs', async () => {
    const wb = new Workbook();
    const id = addImage(wb, { data: GIF, extension: 'gif' });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 1);
    ws.backgroundImage = id;
    const buf = writeXlsx(wb);
    expect(parts(buf).get('xl/worksheets/sheet1.xml')).toMatch(/<picture r:id="rId1"\/>/);
    const back = readXlsx(buf);
    expect(back.getWorksheet('S')!.backgroundImage).toBe(0);
    expect(back.images).toEqual([{ id: 0, data: GIF, extension: 'gif' }]);
    const x = (await oracle(buf)).getWorksheet('S')!;
    expect(x.getBackgroundImageId()).toBeDefined();
  });

  it('notes and pictures on the same sheet do not collide', async () => {
    const wb = new Workbook();
    const id = addImage(wb, { data: PNG, extension: 'png' });
    for (const n of ['A', 'B']) {
      const ws = wb.addWorksheet(n);
      ws.setCell('B2', 1);
      addNote(ws, 'B2', `note ${n}`);
      placeImage(ws, id, 'D2:E4');
    }
    for (const buf of [writeXlsx(wb)]) {
      const p = parts(buf);
      expect(p.has('xl/drawings/drawing1.xml') && p.has('xl/drawings/drawing2.xml')).toBe(true);
      expect(p.has('xl/drawings/vmlDrawing1.vml') && p.has('xl/drawings/vmlDrawing2.vml')).toBe(
        true,
      );
      const sheet = p.get('xl/worksheets/sheet2.xml')!;
      expect(sheet).toMatch(/<drawing r:id="rId1"\/><legacyDrawing r:id="rId3"\/>/);
      const rels = p.get('xl/worksheets/_rels/sheet2.xml.rels')!;
      expect(rels).toContain(
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"',
      );
      expect(rels).toContain(
        'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments2.xml"',
      );
      expect(rels).toContain(
        'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing2.vml"',
      );
      const back = readXlsx(buf);
      for (const n of ['A', 'B']) {
        const ws = back.getWorksheet(n)!;
        expect(ws.comments).toEqual([{ ref: 'B2', text: `note ${n}` }]);
        expect(ws.images!.length).toBe(1);
      }
      const x = await oracle(buf);
      for (const n of ['A', 'B']) {
        const xs = x.getWorksheet(n)!;
        expect(xs.getCell('B2').note).toBe(`note ${n}`);
        expect(xs.getImages().length).toBe(1);
      }
    }
  });

  it('skips charts / shapes in a drawing but keeps its pictures', () => {
    const wb = new Workbook();
    const id = addImage(wb, { data: PNG, extension: 'png' });
    placeImage(wb.addWorksheet('S'), id, 'B2:C3');
    const files = unzipSync(writeXlsx(wb));
    const path = 'xl/drawings/drawing1.xml';
    const chart =
      '<xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="9" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId99"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
      '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="10" cy="10"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="10" name="Shape"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>' +
      '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>';
    files.set(
      path,
      enc.encode(
        dec.decode(files.get(path)!).replace('<xdr:twoCellAnchor>', `${chart}<xdr:twoCellAnchor>`),
      ),
    );
    const ws = readXlsx(
      zipSync(
        [...files].map(([p, data]) => ({ path: p, data })),
        6,
      ),
    ).getWorksheet('S')!;
    expect(ws.images).toEqual([
      { imageId: 0, anchor: { from: { col: 1, row: 1 }, to: { col: 3, row: 3 } } },
    ]);
  });

  it('rejects unknown image ids and missing data', () => {
    const wb = new Workbook();
    placeImage(wb.addWorksheet('S'), 7, 'A1');
    expect(() => writeXlsx(wb)).toThrow(/Image 7/);
    expect(() => addImage(wb, { extension: 'png' })).toThrow(TypeError);
  });
});
