/// <reference types="node" />
// Gradient fills, extra font attributes, workbook views.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { type RichText, readXlsx, Workbook, writeXlsx } from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const part = (buf: Uint8Array, p: string): string => dec.decode(unzipSync(buf).get(p)!);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('gradient fills', () => {
  it('angle and path gradients write, dedupe, read back; exceljs agrees', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const linear = {
      type: 'gradient' as const,
      gradient: 'angle' as const,
      degree: 90,
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    };
    ws.setCell('A1', 'linear', { fill: linear });
    ws.setCell('A2', 'same', { fill: { ...linear } });
    ws.setCell('B1', 'path', {
      fill: {
        type: 'gradient',
        gradient: 'path',
        center: { left: 0.5, top: 0.5 },
        stops: [
          { position: 0, color: { argb: 'FFFFFFFF' } },
          { position: 1, color: { theme: 4 } },
        ],
      },
    });
    expect(ws.getStyleIndex('A1')).toBe(ws.getStyleIndex('A2'));
    const buf = writeXlsx(wb);
    const styles = part(buf, 'xl/styles.xml');
    expect(styles).toContain(
      '<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>',
    );
    expect(styles).toContain(
      '<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">',
    );
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.getStyle('A1').fill).toEqual(linear);
    expect(back.getStyle('B1').fill).toMatchObject({
      gradient: 'path',
      center: { left: 0.5, top: 0.5 },
    });
    const ex = await oracle(buf);
    expect(ex.getWorksheet('S')!.getCell('A1').fill).toMatchObject({
      type: 'gradient',
      gradient: 'angle',
      degree: 90,
    });
  });
});

describe('font attributes', () => {
  it('charset, outline, shadow, condense, extend round-trip and dedupe', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const font = {
      name: 'MS Gothic',
      charset: 128,
      outline: true,
      shadow: true,
      condense: true,
      extend: true,
    };
    ws.setCell('A1', 'x', { font });
    ws.setCell('A2', 'plain', { font: { name: 'MS Gothic' } });
    expect(ws.getStyleIndex('A1')).not.toBe(ws.getStyleIndex('A2'));
    const buf = writeXlsx(wb);
    expect(part(buf, 'xl/styles.xml')).toMatch(
      /<condense\/><extend\/><outline\/><shadow\/>.*<charset val="128"\/>/,
    );
    expect(readXlsx(buf).getWorksheet('S')!.getStyle('A1').font).toMatchObject(font);
    // in rich text runs too
    ws.setCell('B1', { richText: [{ text: 'o', font: { outline: true, charset: 0 } }] });
    const rt = readXlsx(writeXlsx(wb)).getWorksheet('S')!.getValue('B1') as RichText;
    expect(rt.richText[0]!.font).toMatchObject({ outline: true, charset: 0 });
  });
});

describe('workbook views', () => {
  it('window position / size, first sheet, active tab, visibility', async () => {
    const wb = new Workbook();
    for (const n of ['A', 'B', 'C']) wb.addWorksheet(n).setCell('A1', n);
    wb.views = [{ x: 100, y: 200, width: 20000, height: 10000, firstSheet: 1, activeTab: 2 }];
    const buf = writeXlsx(wb);
    expect(part(buf, 'xl/workbook.xml')).toContain(
      '<workbookView xWindow="100" yWindow="200" windowWidth="20000" windowHeight="10000" firstSheet="1" activeTab="2"/>',
    );
    const back = readXlsx(buf);
    expect(back.views).toEqual([{ x: 100, y: 200, width: 20000, height: 10000, firstSheet: 1 }]);
    const ex = await oracle(buf);
    expect(ex.views[0]).toMatchObject({ activeTab: 2, firstSheet: 1 });
    // defaults unchanged without views
    const plain = new Workbook();
    plain.addWorksheet('S');
    expect(part(writeXlsx(plain), 'xl/workbook.xml')).toContain(
      '<workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/>',
    );
  });
});
