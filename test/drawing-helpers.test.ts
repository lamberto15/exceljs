/// <reference types="node" />
// Fractional / corner image anchors, fitImage + imageSize, in-cell checkboxes.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addChart,
  addImage,
  fitImage,
  imageSize,
  placeImage,
  readXlsx,
  setCheckbox,
  Workbook,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

/** Minimal PNG header bytes of the given size (enough for imageSize; not a full image). */
function pngHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(24);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, w);
  new DataView(b.buffer).setUint32(20, h);
  return b;
}

// 8x8 red PNG (valid)
const PNG8 = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFWEXHbQSACj/P8Fu7N9hAAAAAElFTkSuQmCC',
  ),
  (c) => c.charCodeAt(0),
);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('anchors', () => {
  it('fractional positions use column widths / row heights; tl + br corners', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getColumn(2).width = 20; // 140 px
    ws.setRowProps(3, { height: 30 }); // 40 px
    const id = addImage(wb, { data: PNG8, extension: 'png' });
    const a = placeImage(ws, id, { tl: { col: 1.5, row: 2.25 }, ext: { width: 10, height: 10 } });
    expect(a.anchor).toEqual({
      from: { col: 1, row: 2, colOff: 70, rowOff: 10 },
      ext: { width: 10, height: 10 },
    });
    const b = placeImage(ws, id, { tl: { col: 0.5, row: 0 }, br: 'C4' });
    expect(b.anchor).toEqual({ from: { col: 0, row: 0, colOff: 32 }, to: { col: 2, row: 3 } });
    addChart(ws, {
      type: 'line',
      series: [{ values: 'A1:A3' }],
      anchor: { tl: { col: 3.5, row: 1 }, br: { col: 8, row: 15 } },
    });
    const xml = dec.decode(unzipSync(writeXlsx(wb)).get('xl/drawings/drawing1.xml')!);
    expect(xml).toContain(
      '<xdr:col>1</xdr:col><xdr:colOff>666750</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>95250</xdr:rowOff>',
    );
    expect(xml).toContain('<xdr:col>3</xdr:col><xdr:colOff>304800</xdr:colOff>'); // 32 px of a default column
  });
});

describe('imageSize / fitImage', () => {
  it('reads PNG, GIF and JPEG sizes', () => {
    expect(imageSize(PNG8)).toEqual({ width: 8, height: 8 });
    expect(imageSize(pngHeader(640, 480))).toEqual({ width: 640, height: 480 });
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x2c, 0x01, 0xc8, 0x00]);
    expect(imageSize(gif)).toEqual({ width: 300, height: 200 });
    // SOI, APP0 (length 4), SOF0: length 17, precision 8, height 0x0100, width 0x0200
    const jpg = new Uint8Array([
      0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 17, 8, 1, 0, 2, 0, 3, 0, 0,
    ]);
    expect(imageSize(jpg)).toEqual({ width: 512, height: 256 });
    expect(imageSize(new Uint8Array([1, 2, 3]))).toBeUndefined();
  });

  it('scales down to fit, keeps the aspect ratio and centers in the range', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const big = addImage(wb, { data: pngHeader(640, 480), extension: 'png' });
    // B2:C4 = 2 default columns (128 px) x 3 default rows (60 px)
    const p = fitImage(ws, big, 'B2:C4');
    // box minus 2 px padding: 124 x 56 → scale 56/480 → 75 x 56; x = (128-75)/2 = 26.5
    expect(p.anchor).toEqual({
      from: { col: 1, row: 1, colOff: 27, rowOff: 2 },
      ext: { width: 75, height: 56 },
    });
    const small = addImage(wb, { data: PNG8, extension: 'png' });
    const q = fitImage(ws, small, 'E5'); // not upscaled: 8x8 centered in 64x20
    expect(q.anchor).toEqual({
      from: { col: 4, row: 4, colOff: 28, rowOff: 6 },
      ext: { width: 8, height: 8 },
    });
    const u = fitImage(ws, small, 'E5', { upscale: true, padding: 0 });
    expect(u.anchor).toMatchObject({ ext: { width: 20, height: 20 } });
    const wide = addImage(wb, { data: pngHeader(1000, 10), extension: 'png' });
    const w = fitImage(ws, wide, 'A10:D10'); // centered vertically, starts in column A
    expect(w.anchor).toMatchObject({ from: { col: 0, row: 9, colOff: 2 } });
    const svg = addImage(wb, { data: new TextEncoder().encode('<svg/>'), extension: 'svg' });
    expect(() => fitImage(ws, svg, 'A1')).toThrow(/unknown pixel size/);
  });
});

describe('in-cell checkboxes', () => {
  it('setCheckbox writes the xf extension and the property bag, and reads back', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A2', true);
    setCheckbox(ws, 'A1:A3');
    setCheckbox(ws, 'B1', true);
    expect([ws.getValue('A1'), ws.getValue('A2'), ws.getValue('B1')]).toEqual([false, true, true]);
    expect(ws.getStyle('A1').checkbox).toBe(true);
    const buf = writeXlsx(wb);
    const files = unzipSync(buf);
    const styles = dec.decode(files.get('xl/styles.xml')!);
    expect(styles).toContain('<xfpb:xfComplement i="0"/>');
    expect(dec.decode(files.get('xl/featurePropertyBag/featurePropertyBag.xml')!)).toContain(
      '<bag type="Checkbox"/>',
    );
    expect(dec.decode(files.get('xl/_rels/workbook.xml.rels')!)).toContain(
      'relationships/FeaturePropertyBag',
    );
    expect(dec.decode(files.get('[Content_Types].xml')!)).toContain(
      'application/vnd.ms-excel.featurepropertybag+xml',
    );
    const back = readXlsx(buf);
    expect(back.getWorksheet('S')!.getStyle('B1').checkbox).toBe(true);
    // round trip: the bag is regenerated once, not kept twice
    const again = unzipSync(writeXlsx(back));
    const rels = dec.decode(again.get('xl/_rels/workbook.xml.rels')!);
    expect(rels.match(/FeaturePropertyBag/g)!.length).toBe(1);
    expect([...again.keys()].filter((k) => k.includes('featurePropertyBag')).length).toBe(1);
    await oracle(buf);
  });

  it('no checkbox: no property bag part', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').setCell('A1', true);
    expect([...unzipSync(writeXlsx(wb)).keys()].some((k) => k.includes('featurePropertyBag'))).toBe(
      false,
    );
  });
});
