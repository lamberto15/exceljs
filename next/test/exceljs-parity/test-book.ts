// Port of spec/utils/test-values-sheet.js (`createTestBook` / `checkTestBook` with the 'values'
// sheet), the workbook most ExcelJS serialisation specs write and read back.
//
// Substitutions (see docs/exceljs-parity.md):
// - JSON objects as cell values: exceljs stringifies them; turbosheet rejects them, the caller
//   stores JSON.stringify(...) (➖).
// - Fonts: `family` / `outline` are not modelled (dropped); `underline: true` reads back as
//   'single' (🚫 / ➖).
// - Alignment: exceljs `vertical: 'middle'` is OOXML 'center'; negative rotations are stored in
//   Excel's encoding (-45 → 135); 'vertical' text is 255 (➖).
// - Gradient fills are not modelled (🚫): row 8 only checks the pattern fills.
// - `{ sharedFormula }` values are not written (➖): C9:E9 are plain formulas.
// - Merged cells: slave cells read null (exceljs returns the master value) (➖).
// - Workbook views, `collapsed`, sheet `properties` (defaultRowHeight, dyDescent,
//   outlineLevelCol/Row) are not modelled (🚫); outline levels are per row/column.

import { expect } from 'vitest';
import {
  type Alignment,
  type Border,
  CellKind,
  type CellValue,
  type Font,
  getPrintArea,
  getPrintTitles,
  type PageSetup,
  type Style,
  setPrintArea,
  setPrintTitles,
  type Workbook,
  type Worksheet,
  type XlsxWriter,
} from '../../src/index.js';
import { utilData } from './helpers.js';

// biome-ignore-start lint/suspicious/noExplicitAny: untyped ExcelJS JSON fixtures
export const testValues: any = utilData('sheet-values');
export const styles: any = utilData('styles');
export const properties: any = utilData('sheet-properties');
export const pageSetupJson: any = utilData('page-setup');
export const headerFooter: any = utilData('header-footer');
// biome-ignore-end lint/suspicious/noExplicitAny: untyped ExcelJS JSON fixtures

/** exceljs font → turbosheet font (family/outline are not modelled). */
export function font(f: Record<string, unknown>): Font {
  const { family: _family, outline: _outline, ...rest } = f;
  return rest as Font;
}

/** The font turbosheet reads back (underline true → 'single', name/size completed). */
export function fontBack(f: Record<string, unknown>): Font {
  const o = { ...font(f) } as Record<string, unknown>;
  if (o.underline === true) o.underline = 'single';
  return o as Font;
}

/** exceljs alignment → turbosheet alignment. */
export function alignment(a: Record<string, unknown>): Alignment {
  const o = { ...a } as Record<string, unknown>;
  if (o.vertical === 'middle') o.vertical = 'center';
  if (o.textRotation === 'vertical') o.textRotation = 255;
  return o as Alignment;
}

/** The alignment turbosheet reads back (Excel's rotation encoding, defaults dropped). */
export function alignmentBack(a: Record<string, unknown>): Alignment | undefined {
  const o = alignment(a) as Record<string, unknown>;
  const r = o.textRotation as number | undefined;
  if (r !== undefined && r < 0) o.textRotation = 90 - r;
  if (o.vertical === 'bottom') delete o.vertical;
  return Object.keys(o).length ? (o as Alignment) : undefined;
}

/** argb as turbosheet normalizes it (6 hex digits get an opaque alpha). */
function argb(b: Border): Border {
  return JSON.parse(
    JSON.stringify(b, (k, v) => (k === 'argb' && /^[0-9a-f]{6}$/i.test(v) ? `FF${v}` : v)),
  );
}

export const numFmt1: string = testValues.numFmt1;
export const numFmt2: string = testValues.numFmt2;

/** exceljs page-setup.json minus the fields that live elsewhere in turbosheet. */
export function pageSetup(): PageSetup {
  const {
    printArea: _a,
    printTitlesRow: _r,
    printTitlesColumn: _c,
    useFirstPageNumber: _u,
    ...rest
  } = pageSetupJson;
  return rest as PageSetup;
}

interface CellSpec {
  ref: string;
  value?: CellValue;
  style?: Style;
}

/** Cells of the 'values' sheet in row order. */
function cells(checkBadAlignments: boolean): CellSpec[] {
  const tv = testValues;
  const out: CellSpec[] = [
    { ref: 'A1', value: 7 },
    { ref: 'B1', value: tv.str },
    { ref: 'C1', value: tv.date },
    { ref: 'D1', value: tv.formulas[0] },
    { ref: 'E1', value: tv.formulas[1] },
    { ref: 'F1', value: tv.hyperlink },
    { ref: 'G1', value: tv.str2 },
    { ref: 'H1', value: JSON.stringify(tv.json.raw) },
    { ref: 'I1', value: true },
    { ref: 'J1', value: false },
    { ref: 'K1', value: tv.Errors.NotApplicable },
    { ref: 'L1', value: tv.Errors.Value },
    { ref: 'A2', value: 5 },
    { ref: 'A4', value: 1.5, style: { numFmt: numFmt1, border: styles.borders.thin } },
    { ref: 'C4', value: 1.5, style: { numFmt: numFmt2, border: styles.borders.doubleRed } },
    { ref: 'E4', value: 1.5, style: { border: styles.borders.thickRainbow } },
    { ref: 'A5', value: tv.str, style: { font: font(styles.fonts.arialBlackUI14) } },
    { ref: 'B5', value: tv.str, style: { font: font(styles.fonts.broadwayRedOutline20) } },
    { ref: 'C5', value: tv.str, style: { font: font(styles.fonts.comicSansUdB16) } },
    {
      ref: 'D5',
      value: 1.6,
      style: { numFmt: numFmt1, font: font(styles.fonts.arialBlackUI14) },
    },
    {
      ref: 'E5',
      value: 1.6,
      style: { numFmt: numFmt2, font: font(styles.fonts.broadwayRedOutline20) },
    },
    {
      ref: 'F5',
      value: tv.date,
      style: { numFmt: tv.numFmtDate, font: font(styles.fonts.comicSansUdB16) },
    },
  ];
  styles.alignments.forEach((a: { text: string; alignment: object }, i: number) => {
    out.push({
      ref: `${String.fromCharCode(65 + i)}6`,
      value: a.text,
      style: { alignment: alignment(a.alignment as Record<string, unknown>) },
    });
  });
  if (checkBadAlignments) {
    styles.badAlignments.forEach((a: { text: string; alignment: object }, i: number) => {
      out.push({
        ref: `${String.fromCharCode(65 + i)}7`,
        value: a.text,
        style: { alignment: a.alignment as Alignment },
      });
    });
  }
  out.push(
    { ref: 'A8', value: 'Blue White Horizontal Gradient' },
    { ref: 'B8', value: 'Red Dark Vertical', style: { fill: styles.fills.redDarkVertical } },
    {
      ref: 'C8',
      value: 'Red Green Dark Trellis',
      style: { fill: styles.fills.redGreenDarkTrellis },
    },
    { ref: 'D8', value: 'RGB Path Gradient' },
    { ref: 'A9', value: 1 },
    { ref: 'B9', value: { formula: 'A9+1', result: 2 } },
    { ref: 'C9', value: { formula: 'B9+1', result: 3 } },
    { ref: 'D9', value: { formula: 'C9+1', result: 4 } },
    { ref: 'E9', value: { formula: 'D9+1', result: 5 } },
    { ref: 'J10', value: 1 },
    { ref: 'A11', value: { formula: 'A9', result: 1, array: 'A11:E11' } },
    { ref: 'B11', value: 1 },
    { ref: 'C11', value: 1 },
    { ref: 'D11', value: 1 },
    { ref: 'E11', value: 1 },
  );
  return out;
}

export interface BookOptions {
  checkStyles?: boolean;
  checkBadAlignments?: boolean;
  checkSheetProperties?: boolean;
}

function sheetSetup(ws: Worksheet): void {
  ws.tabColor = properties.tabColor;
  ws.pageSetup = pageSetup();
  ws.headerFooter = headerFooter;
  setPrintArea(ws, pageSetupJson.printArea);
  setPrintTitles(ws, { rows: pageSetupJson.printTitlesRow, cols: pageSetupJson.printTitlesColumn });
  ws.getColumn(10).outlineLevel = 1;
}

/** createTestBook(new Workbook(), 'xlsx') */
export function createTestBook(wb: Workbook, o: BookOptions = {}): Workbook {
  const ws = wb.addWorksheet('sheet1');
  sheetSetup(ws);
  for (const c of cells(o.checkBadAlignments !== false)) ws.setCell(c.ref, c.value, c.style);
  ws.getRow(10).outlineLevel = 1;
  ws.mergeCells('A2:B3');
  ws.mergeCells('C2:D3');
  ws.getRow(6).height = 42;
  ws.getRow(8).height = 40;
  return wb;
}

/**
 * The same sheet through the streaming writer. Rows are written whole, so cell styles are not
 * applied (exceljs' "Without styles" variant); row heights / outline levels are row props.
 */
export function createTestBookStream(w: XlsxWriter): void {
  const s = w.addWorksheet('sheet1');
  sheetSetup(s.worksheet);
  s.mergeCells('A2:B3');
  s.mergeCells('C2:D3');
  const rows: CellValue[][] = [];
  for (const c of cells(true)) {
    const m = /^([A-Z]+)(\d+)$/.exec(c.ref)!;
    const r = Number(m[2]) - 1;
    (rows[r] ??= [])[m[1]!.charCodeAt(0) - 65] = c.value;
  }
  for (let r = 0; r < rows.length; r++) {
    const props =
      r === 5
        ? { height: 42 }
        : r === 7
          ? { height: 40 }
          : r === 9
            ? { outlineLevel: 1 }
            : undefined;
    s.addRow(rows[r] ?? [], undefined, props);
  }
}

const near = (a: Date, b: Date, ms: number) =>
  expect(Math.abs(a.getTime() - b.getTime())).toBeLessThan(ms);

/** checkTestBook(wb, 'xlsx') */
export function checkTestBook(wb: Workbook, o: BookOptions = {}): void {
  const checkStyles = o.checkStyles !== false;
  const ws = wb.getWorksheet('sheet1')!;
  expect(ws).toBeDefined();
  const tv = testValues;

  if (o.checkSheetProperties !== false) {
    expect(ws.getColumn(10).outlineLevel).toBe(1);
    expect(ws.getRow(10).outlineLevel).toBe(1);
    expect(ws.tabColor).toEqual(properties.tabColor);
    expect(ws.pageSetup).toEqual(pageSetup());
    expect(ws.headerFooter).toEqual(headerFooter);
    expect(getPrintArea(ws)).toEqual([pageSetupJson.printArea]);
    expect(getPrintTitles(ws)).toEqual({
      rows: pageSetupJson.printTitlesRow,
      cols: pageSetupJson.printTitlesColumn,
    });
  }

  expect(ws.getValue('A1')).toBe(7);
  expect(ws.getKind('A1')).toBe(CellKind.Number);
  expect(ws.getValue('B1')).toBe(tv.str);
  expect(ws.getKind('B1')).toBe(CellKind.String);
  near(ws.getValue('C1') as Date, tv.date, 3);
  expect(ws.getKind('C1')).toBe(CellKind.Date);
  expect(ws.getValue('D1')).toEqual(tv.formulas[0]);
  expect(ws.getKind('D1')).toBe(CellKind.Formula);
  expect((ws.getValue('E1') as { formula: string }).formula).toBe(tv.formulas[1].formula);
  expect((ws.getValue('E1') as { result?: unknown }).result).toBeUndefined();
  expect(ws.getKind('E1')).toBe(CellKind.Formula);
  expect(ws.getValue('F1')).toEqual(tv.hyperlink);
  expect(ws.getKind('F1')).toBe(CellKind.Hyperlink);
  expect(ws.getValue('G1')).toBe(tv.str2);
  expect(ws.getValue('H1')).toBe(tv.json.string);
  expect(ws.getKind('H1')).toBe(CellKind.String);
  expect(ws.getValue('I1')).toBe(true);
  expect(ws.getKind('I1')).toBe(CellKind.Boolean);
  expect(ws.getValue('J1')).toBe(false);
  expect(ws.getValue('K1')).toEqual(tv.Errors.NotApplicable);
  expect(ws.getKind('K1')).toBe(CellKind.Error);
  expect(ws.getValue('L1')).toEqual(tv.Errors.Value);
  expect(ws.getKind('L1')).toBe(CellKind.Error);

  // merges (➖ slave cells hold no value of their own)
  expect(ws.getValue('A2')).toBe(5);
  expect(ws.merges).toEqual(expect.arrayContaining(['A2:B3', 'C2:D3']));
  for (const ref of ['A3', 'B2', 'B3', 'C2', 'D2', 'C3', 'D3']) expect(ws.getValue(ref)).toBeNull();

  expect(ws.getRow(5).height).toBeUndefined();
  expect(ws.getRow(6).height).toBe(42);
  expect(ws.getRow(8).height).toBe(40);

  if (checkStyles) {
    expect(ws.getStyle('A4').numFmt).toBe(numFmt1);
    expect(ws.getKind('A4')).toBe(CellKind.Number);
    expect(ws.getStyle('A4').border).toEqual(styles.borders.thin);
    expect(ws.getStyle('C4').numFmt).toBe(numFmt2);
    expect(ws.getStyle('C4').border).toEqual(styles.borders.doubleRed);
    expect(ws.getStyle('E4').border).toEqual(argb(styles.borders.thickRainbow));

    expect(ws.getValue('A5')).toBe(tv.str);
    const fontOf = (ref: string) => ws.getStyle(ref).font;
    expect(fontOf('A5')).toEqual(fontBack(styles.fonts.arialBlackUI14));
    expect(fontOf('B5')).toEqual(fontBack(styles.fonts.broadwayRedOutline20));
    expect(fontOf('C5')).toEqual(fontBack(styles.fonts.comicSansUdB16));
    expect(Math.abs((ws.getValue('D5') as number) - 1.6)).toBeLessThan(1e-8);
    expect(ws.getStyle('D5').numFmt).toBe(numFmt1);
    expect(fontOf('D5')).toEqual(fontBack(styles.fonts.arialBlackUI14));
    expect(ws.getStyle('E5').numFmt).toBe(numFmt2);
    expect(fontOf('E5')).toEqual(fontBack(styles.fonts.broadwayRedOutline20));
    near(ws.getValue('F5') as Date, tv.date, 3);
    expect(ws.getKind('F5')).toBe(CellKind.Date);
    expect(ws.getStyle('F5').numFmt).toBe(tv.numFmtDate);
    expect(fontOf('F5')).toEqual(fontBack(styles.fonts.comicSansUdB16));

    styles.alignments.forEach((a: { text: string; alignment: object }, i: number) => {
      const ref = `${String.fromCharCode(65 + i)}6`;
      expect(ws.getValue(ref)).toBe(a.text);
      expect(ws.getStyle(ref).alignment, ref).toEqual(
        alignmentBack(a.alignment as Record<string, unknown>),
      );
    });
    if (o.checkBadAlignments !== false) {
      styles.badAlignments.forEach((a: { text: string; alignment: object }, i: number) => {
        const ref = `${String.fromCharCode(65 + i)}7`;
        expect(ws.getValue(ref)).toBe(a.text);
        const al = a.alignment as Record<string, unknown>;
        // ➖ 91 is a valid Excel rotation (= -1°), so it is kept; the others are invalid
        if (al.textRotation === 91) expect(ws.getStyle(ref).alignment).toEqual(al);
        else expect(ws.getStyle(ref).alignment, ref).toBeUndefined();
      });
    }
    expect(ws.getStyle('B8').fill).toEqual(styles.fills.redDarkVertical);
    expect(ws.getStyle('C8').fill).toEqual(styles.fills.redGreenDarkTrellis);
  }

  expect(ws.getValue('A9')).toBe(1);
  expect(ws.getValue('B9')).toEqual({ formula: 'A9+1', result: 2 });
  expect(ws.getValue('C9')).toEqual({ formula: 'B9+1', result: 3 });
  expect(ws.getValue('E9')).toEqual({ formula: 'D9+1', result: 5 });
  expect(ws.getValue('A11')).toEqual({ formula: 'A9', result: 1, array: 'A11:E11' });
  expect(ws.getValue('E11')).toBe(1);
}
