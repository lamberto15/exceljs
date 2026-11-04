// Pictures: spec/integration/workbook/images.spec.js, workbook-xlsx-writer.spec.js (background
// image) and the images.xlsx anchors of workbook-xlsx-reader.spec.js. pr 2244 is in prs.test.ts.

import { describe, expect, it } from 'vitest';
import {
  addImage,
  type ImagePlacement,
  placeImage,
  readXlsx,
  type Workbook as WB,
  Workbook,
  writeXlsx,
} from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';

const png = fixture('image.png');

function imageData(wb: WB, id: number | undefined): Uint8Array | undefined {
  return wb.images?.find((i) => i.id === id)?.data;
}

function only(wb: WB, sheet: string): ImagePlacement {
  const list = wb.getWorksheet(sheet)!.images ?? [];
  expect(list).toHaveLength(1);
  return list[0]!;
}

describe('exceljs parity: images', () => {
  // exceljs: spec/integration/workbook/images.spec.js — "stores background image"
  it('stores background image', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    // ➖ no `filename` source: pass the bytes (the exceljs test also labels this PNG 'jpeg')
    const id = addImage(wb, { data: png, extension: 'jpeg' });
    ws.getCell('A1').value = 'Hello, World!';
    ws.backgroundImage = id;
    const bytes = writeXlsx(wb);
    const wb2 = readXlsx(bytes);
    const ws2 = wb2.getWorksheet('blort')!;
    expect(imageData(wb2, ws2.backgroundImage)).toEqual(png);
    const x = (await exceljsLoad(bytes)).getWorksheet('blort')!;
    expect(x.getBackgroundImageId()).toBeDefined();
  });

  // exceljs: spec/integration/workbook/images.spec.js — "stores embedded image and hyperlink"
  it('stores an embedded image over a range', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const id = addImage(wb, { data: png, extension: 'jpeg' });
    ws.getCell('A2').value = { hyperlink: 'http://www.somewhere.com', text: 'www.somewhere.com' };
    placeImage(ws, id, 'C3:E6');
    const wb2 = roundtrip(wb);
    const p = only(wb2, 'blort');
    // exceljs: tl { col: 2, row: 2 }, br { col: 5, row: 6 } (0-based, br exclusive)
    expect(p.anchor).toMatchObject({ from: { col: 2, row: 2 }, to: { col: 5, row: 6 } });
    expect(imageData(wb2, p.imageId)).toEqual(png);
    expect(wb2.getWorksheet('blort')!.getValue('A2')).toEqual({
      hyperlink: 'http://www.somewhere.com',
      text: 'www.somewhere.com',
    });
  });

  // exceljs: spec/integration/workbook/images.spec.js — "stores embedded image with oneCell"
  it('stores an embedded image with editAs oneCell (two-cell anchor with offsets)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const id = addImage(wb, { data: png, extension: 'jpeg' });
    // ➖ exceljs fractional cells ({ col: 0.1125, row: 0.4 }); turbosheet: cell + pixel offset
    placeImage(
      ws,
      id,
      {
        from: { col: 0, row: 0, colOff: 7, rowOff: 8 },
        to: { col: 2, row: 3, colOff: 6, rowOff: 8 },
      },
      { editAs: 'oneCell' },
    );
    const p = only(roundtrip(wb), 'blort');
    expect(p.editAs).toBe('oneCell');
    expect(p.anchor).toEqual({
      from: { col: 0, row: 0, colOff: 7, rowOff: 8 },
      to: { col: 2, row: 3, colOff: 6, rowOff: 8 },
    });
  });

  // exceljs: spec/integration/workbook/images.spec.js — "stores embedded image with one-cell-anchor"
  it('stores an embedded image with a one-cell anchor (fixed size)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const id = addImage(wb, { data: png, extension: 'jpeg' });
    placeImage(ws, id, {
      tl: { col: 0, row: 0, colOff: 7, rowOff: 8 },
      ext: { width: 100, height: 100 },
    });
    const wb2 = roundtrip(wb);
    const p = only(wb2, 'blort');
    expect(p.anchor).toMatchObject({ ext: { width: 100, height: 100 } });
    expect(imageData(wb2, p.imageId)).toEqual(png);
  });

  // exceljs: spec/integration/workbook/images.spec.js — "stores embedded image with hyperlinks"
  it('stores an embedded image with a hyperlink and tooltip', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const id = addImage(wb, { data: png, extension: 'jpeg' });
    placeImage(
      ws,
      id,
      { tl: { col: 0, row: 0, colOff: 7, rowOff: 8 }, ext: { width: 100, height: 100 } },
      { hyperlink: 'http://www.somewhere.com', tooltip: 'www.somewhere.com' },
    );
    const bytes = writeXlsx(wb);
    const p = only(readXlsx(bytes), 'blort');
    expect(p.anchor).toMatchObject({ ext: { width: 100, height: 100 } });
    expect(p.hyperlink).toBe('http://www.somewhere.com');
    expect(p.tooltip).toBe('www.somewhere.com');
    const xi = (await exceljsLoad(bytes)).getWorksheet('blort')!.getImages();
    expect((xi[0]!.range as unknown as { hyperlinks: unknown }).hyperlinks).toEqual({
      hyperlink: 'http://www.somewhere.com',
      tooltip: 'www.somewhere.com',
    });
  });

  // exceljs: spec/integration/workbook/images.spec.js — "image extensions should not be case sensitive"
  it('image extensions are not case sensitive', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    const id1 = addImage(wb, { data: png, extension: 'PNG' });
    const id2 = addImage(wb, { data: png, extension: 'JPEG' });
    expect(wb.images!.map((i) => i.extension)).toEqual(['png', 'jpeg']);
    placeImage(ws, id1, { tl: { col: 0, row: 0 }, ext: { width: 100, height: 100 } });
    placeImage(
      ws,
      id2,
      { from: { col: 0, row: 0 }, to: { col: 2, row: 3 } },
      { editAs: 'oneCell' },
    );
    const wb2 = roundtrip(wb);
    const imgs = wb2.getWorksheet('blort')!.images!;
    expect(imgs).toHaveLength(2);
    expect(imgs[0]!.anchor).toMatchObject({ ext: { width: 100, height: 100 } });
    expect(imgs[1]!.editAs).toBe('oneCell');
    expect(imageData(wb2, imgs[0]!.imageId)).toEqual(png);
    expect(imageData(wb2, imgs[1]!.imageId)).toEqual(png);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "with background image"
  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "with background image where worksheet is commited in advance"
  it.each([false, true])(
    'streaming writer: background image (commit first: %s)',
    async (commit) => {
      const bytes = await streamWrite(async (w) => {
        const s = w.addWorksheet('Hello');
        const id = addImage(w.workbook, { data: png, extension: 'jpeg' });
        s.worksheet.backgroundImage = id;
        s.addRow(['Hello, World!']);
        if (commit) await s.commit();
      });
      const wb2 = readXlsx(bytes);
      expect(imageData(wb2, wb2.getWorksheet('Hello')!.backgroundImage)).toEqual(png);
    },
  );

  describe('with a spreadsheet that contains images (images.xlsx)', () => {
    const imgs = () => readXlsx(fixture('images.xlsx')).worksheets[0]!.images ?? [];
    const corners = (p: ImagePlacement) => {
      const a = p.anchor as { from?: object; to?: object };
      return [a.from, a.to].filter(Boolean) as { col: number; row: number }[];
    };

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should integer part of col equals nativeCol" (tl)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should integer part of col equals nativeCol" (br)
    it('anchor columns are whole cell indexes', () => {
      expect(imgs().length).toBeGreaterThan(0);
      for (const p of imgs())
        for (const c of corners(p)) expect(Number.isInteger(c.col)).toBe(true);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should integer part of row equals nativeRow" (tl)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should integer part of row equals nativeRow" (br)
    it('anchor rows are whole cell indexes', () => {
      for (const p of imgs())
        for (const c of corners(p)) expect(Number.isInteger(c.row)).toBe(true);
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should anchor width equals to column width when custom" (tl, br)
    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "Should anchor height equals to row height" (tl, br)
    // exceljs: spec/unit/doc/anchor.spec.js (13 tests of the Anchor class)
    it.skip('Anchor colWidth / rowHeight / nativeColOff internals — not applicable (no Anchor class)', () => {});

    it('pictures round-trip', () => {
      const wb = readXlsx(fixture('images.xlsx'));
      const back = roundtrip(wb).worksheets[0]!.images ?? [];
      expect(back.map((p) => p.anchor)).toEqual(imgs().map((p) => p.anchor));
    });
  });
});
