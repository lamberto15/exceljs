// Styles: spec/integration/workbook/styles.spec.js, the style tests of worksheet.spec.js /
// worksheet-xlsx-writer.spec.js / workbook-xlsx-writer.spec.js / workbook-xlsx-reader.spec.js,
// and regressions for the style bugs found while porting.

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx, type Style, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad, fixture, fixtureText, partText, roundtrip, streamWrite } from './helpers.js';
import { alignment, font, fontBack, numFmt1, numFmt2, styles } from './test-book.js';

const comic = font(styles.fonts.comicSansUdB16);
const comicBack = fontBack(styles.fonts.comicSansUdB16);
const broadway = font(styles.fonts.broadwayRedOutline20);
const broadwayBack = fontBack(styles.fonts.broadwayRedOutline20);
const arial = font(styles.fonts.arialBlackUI14);
const centre = alignment(styles.alignments[1].alignment);
const middleCentre = alignment(styles.namedAlignments.middleCentre);
const bottomRight = alignment(styles.namedAlignments.bottomRight);

/** The row/column style the ExcelJS "sets row styles" / "sets col styles" tests apply. */
const lineStyle: Style = {
  numFmt: numFmt2,
  font: comic,
  alignment: middleCentre,
  border: styles.borders.thin,
  fill: styles.fills.redGreenDarkTrellis,
};

function expectLineStyle(s: Style): void {
  expect(s.numFmt).toBe(numFmt2);
  expect(s.font).toEqual(comicBack);
  expect(s.alignment).toEqual(middleCentre);
  expect(s.border).toEqual(styles.borders.thin);
  expect(s.fill).toEqual(styles.fills.redGreenDarkTrellis);
}

describe('exceljs parity: styles', () => {
  // exceljs: spec/integration/workbook/styles.spec.js — "row styles and columns properly"
  it('row styles and column styles round-trip and are inherited by new cells', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    ws.columns = [
      { header: 'A1', width: 10 },
      { header: 'B1', width: 20, style: { font: comic, alignment: centre } },
      { header: 'C1', width: 30 },
    ];
    // ➖ exceljs `row.font = ...`; turbosheet sets the row style through row props
    ws.setRowProps(2, { style: { font: broadway } });
    for (const ref of ['A2', 'B2', 'C2', 'A3', 'B3', 'C3']) ws.getCell(ref).value = ref;

    const ws2 = roundtrip(wb).getWorksheet('blort')!;
    for (const ref of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']) {
      expect(ws2.getValue(ref)).toBe(ref);
    }
    expect(ws2.getStyle('B1').font).toEqual(comicBack);
    expect(ws2.getStyle('B1').alignment).toEqual(centre);
    expect(ws2.getStyle('A2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('B2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('C2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('B3').font).toEqual(comicBack);
    expect(ws2.getStyle('B3').alignment).toEqual(centre);
    expect(ws2.getColumn(2).style.font).toEqual(comicBack);
    expect(ws2.getColumn(2).style.alignment).toEqual(centre);
    expect(ws2.getRowProps(2).style?.font).toEqual(broadwayBack);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "serializes row styles and columns properly"
  it('streaming writer: row styles and column styles', async () => {
    const colStyle: Style = { font: comic, alignment: middleCentre };
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('blort', {
        columns: [
          { header: 'A1', width: 10 },
          { header: 'B1', style: colStyle },
          { header: 'C1', width: 30 },
          { header: 'D1' },
        ],
      });
      // ➖ a row style is given to addRow (cells) and as row props (empty cells)
      const rowStyle: Style = { font: broadway };
      s.addRow(['A2', 'B2', 'C2'], rowStyle, { style: rowStyle });
      s.addRow(['A3', 'B3', 'C3']);
    });
    const ws2 = readXlsx(bytes).getWorksheet('blort')!;
    for (const ref of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']) {
      expect(ws2.getValue(ref)).toBe(ref);
    }
    expect(ws2.getStyle('B1').font).toEqual(comicBack);
    expect(ws2.getStyle('B1').alignment).toEqual(middleCentre);
    expect(ws2.getStyle('A2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('B2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('C2').font).toEqual(broadwayBack);
    expect(ws2.getStyle('B3').font).toEqual(comicBack);
    expect(ws2.getStyle('B3').alignment).toEqual(middleCentre);
    expect(ws2.getColumn(2).style.font).toEqual(comicBack);
    expect(ws2.getColumn(2).style.alignment).toEqual(middleCentre);
    // a styled column without a width gets Excel's default width (exceljs reports 9)
    expect(ws2.getColumn(2).width).toBeCloseTo(9.14, 1);
    expect(ws2.getColumn(4).width).toBeUndefined();
    expect(ws2.getRowProps(2).style?.font).toEqual(broadwayBack);
  });

  // exceljs: spec/integration/workbook/styles.spec.js — "in-cell formats properly in xlsx file"
  it('reads in-cell (rich text) formats', () => {
    const src = fixtureText('rich-text-sample.js');
    const b64 = /'([A-Za-z0-9+/=]+)'/.exec(src)![1]!;
    const expected = JSON.parse(fixtureText('rich-text-sample-a1.json')) as {
      richText: { text: string; font: Record<string, unknown> }[];
    };
    const ws = readXlsx(Uint8Array.from(Buffer.from(b64, 'base64'))).worksheets[0]!;
    const a1 = ws.getValue('A1') as { richText: { text: string; font?: object }[] };
    // 🚫 font family / scheme are not modelled; everything else matches
    expect(a1).toEqual({
      richText: expected.richText.map((r) => {
        const { family: _f, scheme: _s, ...f } = r.font;
        return { text: r.text, font: f };
      }),
    });
    // ➖ no `cell.text`: the plain text is the concatenated runs
    expect(a1.richText.map((r) => r.text).join('')).toBe(ws.getValue('A2'));
  });

  // exceljs: spec/integration/workbook/styles.spec.js — "null cells retain style"
  it('null cells retain style', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('B2').value = 'hello';
    ws.getCell('B4').fill = styles.fills.redDarkVertical;
    ws.getCell('B4').font = broadway;
    const ws2 = roundtrip(wb).getWorksheet('blort')!;
    expect(ws2.getStyle('B4').fill).toEqual(styles.fills.redDarkVertical);
    expect(ws2.getStyle('B4').font).toEqual(broadwayBack);
    expect(ws2.getKind('B4')).toBe(CellKind.Empty);
  });

  // exceljs: spec/integration/workbook/styles.spec.js — "sets row styles"
  // exceljs: spec/integration/worksheet.spec.js — "sets row styles"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "sets row styles"
  // exceljs: spec/unit/doc/worksheet.styles.spec.js — "sets row styles"
  it('sets row styles (new and empty cells inherit; existing cells keep theirs)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('basket');
    ws.getCell('A1').value = 5;
    ws.getCell('A1').numFmt = numFmt1;
    ws.getCell('A1').font = arial;
    ws.getCell('C1').value = 'Hello, World!';
    ws.getCell('C1').alignment = bottomRight;
    ws.getCell('C1').border = styles.borders.doubleRed;
    ws.getCell('C1').fill = styles.fills.redDarkVertical;
    ws.setRowProps(1, { style: lineStyle });

    // ➖ exceljs restyles A1 and C1 too; turbosheet (like Excel's file model) keeps cell styles
    expect(ws.getStyle('A1').numFmt).toBe(numFmt1);
    expect(ws.getStyle('C1').fill).toEqual(styles.fills.redDarkVertical);
    expect(ws.getKind('B1')).toBe(CellKind.Empty); // findCell('B1') === undefined
    expectLineStyle(ws.getStyle('B1'));
    ws.getCell('B1').value = 'new';
    expectLineStyle(ws.getStyle('B1'));
    // to restyle the existing cells, set them explicitly
    for (const ref of ['A1', 'C1']) ws.setStyle(ref, lineStyle);
    const ws2 = roundtrip(wb).getWorksheet('basket')!;
    for (const ref of ['A1', 'B1', 'C1', 'D1']) expectLineStyle(ws2.getStyle(ref));
  });

  // exceljs: spec/integration/workbook/styles.spec.js — "sets col styles"
  // exceljs: spec/integration/worksheet.spec.js — "sets col styles"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "sets col styles"
  // exceljs: spec/unit/doc/worksheet.styles.spec.js — "sets col styles"
  it('sets column styles (new and empty cells inherit; existing cells keep theirs)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('basket');
    ws.getCell('A1').value = 5;
    ws.getCell('A1').numFmt = numFmt1;
    ws.getCell('A1').font = arial;
    ws.getCell('A3').value = 'Hello, World!';
    ws.getCell('A3').alignment = bottomRight;
    ws.getCell('A3').border = styles.borders.doubleRed;
    ws.getCell('A3').fill = styles.fills.redDarkVertical;
    ws.getColumn('A').style = lineStyle;

    expect(ws.getStyle('A1').numFmt).toBe(numFmt1); // ➖ not restyled
    expect(ws.rowData[1]).toBeUndefined(); // findRow(2) === undefined
    expectLineStyle(ws.getStyle('A2'));
    ws.getCell('A2').value = 'new';
    expectLineStyle(ws.getStyle('A2'));
    const ws2 = roundtrip(wb).getWorksheet('basket')!;
    expectLineStyle(ws2.getStyle('A2'));
    expectLineStyle(ws2.getStyle('A9'));
    expect(ws2.getColumn('A').style.font).toEqual(comicBack);
  });

  // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "edit styles of single row instead of all"
  it('editing the style of some rows does not leak into rows sharing the style', () => {
    const ws = readXlsx(fixture('test-row-styles.xlsx')).worksheets[0]!;
    const color = (r: number) => ws.getStyle(r, 1).font?.color?.argb;
    const before = color(3);
    ws.eachRow((row, rowNo) => {
      if (rowNo % 5 === 0) {
        row.eachCell((_v, c) => {
          const s = ws.getStyle(rowNo, c);
          ws.setStyle(rowNo, c, { ...s, font: { ...s.font, color: { argb: '00ff00' } } });
        });
      }
    });
    expect(color(3)).toBe(before);
    expect(color(6)).toBe(color(3));
    expect(color(9)).toBe(color(6));
    expect(color(12)).toBe(color(9));
    expect(color(12)).not.toBe(color(15));
    expect(color(15)).not.toBe(color(18));
    expect(color(15)).toBe(color(10));
    expect(color(10)).toBe(color(5));
    expect(color(5)).toBe('FF00FF00');
  });

  // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file" (gradient fills of row 8)
  it.skip('gradient fills (angle / path) — not supported: Fill is pattern-only', () => {});

  // exceljs: spec/integration/workbook/workbook.spec.js — "xlsx file" (font family / outline)
  it.skip('font family, scheme, charset, outline — not modelled (dropped on read/write)', () => {});

  describe('regressions found by the parity port', () => {
    // exceljs: spec/utils/data/styles.json — alignments "Rotate -15..-90" and badAlignments
    it('negative text rotation is written in Excel encoding (90 + |deg|), not as an invalid negative', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('s');
      ws.setCell('A1', 'x', { alignment: { textRotation: -45 } });
      ws.setCell('A2', 'x', { alignment: { textRotation: -90 } });
      ws.setCell('A3', 'x', { alignment: { textRotation: -91 } }); // invalid → dropped
      ws.setCell('A4', 'x', { alignment: { textRotation: 181 } }); // invalid → dropped
      ws.setCell('A5', 'x', { alignment: { textRotation: 255 } }); // vertical text
      expect(ws.getStyle('A1').alignment).toEqual({ textRotation: 135 });
      const bytes = writeXlsx(wb);
      const xml = partText(bytes, 'xl/styles.xml')!;
      expect(xml).not.toMatch(/textRotation="-/);
      expect(xml).not.toMatch(/textRotation="181"/);
      const ws2 = readXlsx(bytes).getWorksheet('s')!;
      expect(ws2.getStyle('A1').alignment).toEqual({ textRotation: 135 });
      expect(ws2.getStyle('A2').alignment).toEqual({ textRotation: 180 });
      expect(ws2.getStyle('A3').alignment).toBeUndefined();
      expect(ws2.getStyle('A4').alignment).toBeUndefined();
      expect(ws2.getStyle('A5').alignment).toEqual({ textRotation: 255 });
      const x = (await exceljsLoad(bytes)).getWorksheet('s')!;
      expect(x.getCell('A1').alignment.textRotation).toBe(-45);
      expect(x.getCell('A5').alignment.textRotation).toBe('vertical');
    });

    // exceljs: spec/utils/data/styles.json — badAlignments "Indent -1"
    it('negative indent is dropped instead of being written as an invalid attribute', () => {
      const wb = new Workbook();
      wb.addWorksheet('s').setCell('A1', 'x', { alignment: { indent: -1 } });
      const bytes = writeXlsx(wb);
      expect(partText(bytes, 'xl/styles.xml')).not.toMatch(/indent="-/);
      expect(readXlsx(bytes).getWorksheet('s')!.getStyle('A1').alignment).toBeUndefined();
    });

    // exceljs: spec/utils/data/styles.json — borders.thickRainbow right color "FF00FF"
    it('6-digit RGB colors are written as 8-digit ARGB (opaque) and read back the same', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('s', { tabColor: { argb: '00FF00' } });
      ws.setCell('A1', 'x', { border: { right: { style: 'double', color: { argb: 'FF00FF' } } } });
      ws.setCell('A2', { richText: [{ text: 'red', font: { color: { argb: 'ff0000' } } }] });
      expect(ws.getStyle('A1').border?.right?.color).toEqual({ argb: 'FFFF00FF' });
      const bytes = writeXlsx(wb);
      const styleXml = partText(bytes, 'xl/styles.xml')!;
      expect(styleXml).toContain('rgb="FFFF00FF"');
      expect(styleXml).not.toMatch(/rgb="[0-9A-Fa-f]{6}"/);
      const sheetXml = partText(bytes, 'xl/worksheets/sheet1.xml')!;
      expect(sheetXml).toContain('<tabColor rgb="FF00FF00"/>');
      // rich text is written inline in the sheet
      expect(sheetXml).toContain('rgb="FFff0000"');
      const ws2 = readXlsx(bytes).getWorksheet('s')!;
      expect(ws2.getStyle('A1').border?.right?.color).toEqual({ argb: 'FFFF00FF' });
      expect(ws2.tabColor).toEqual({ argb: 'FF00FF00' });
    });
  });
});
