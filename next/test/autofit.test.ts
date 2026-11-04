// Column auto-fit. References come from Excel's documented column-width formula
//   width = Truncate((textPixels + 5) / MDW * 256) / 256      (MDW = 7 px for Calibri 11)
// with textPixels summed from Calibri 11 hinted glyph advances at 96 dpi (digits 7, 'i'/'l' 3,
// space 3, 'e'/'o'/'d' 7, 'H' 9, 'W' 12, 'r' 5, '.' 3.6, ',' 3.6, '-' 4.2, '%' 9.9). Our model is
// not hinted, so the check is +-10%. `padding: 5 / 7` reproduces Excel's bare 5 px padding.

import { describe, expect, it } from 'vitest';
import {
  autoFitColumns,
  type CellValue,
  estimateColumnWidths,
  Workbook,
} from '../src/core/index.js';
import * as lite from '../src/lite.js';

const PAD = 5 / 7;
const excel = (px: number): number => Math.trunc(((px + 5) / 7) * 256) / 256;

function fitCol(values: CellValue[], style?: object, opts: object = { padding: PAD }): number {
  const ws = new Workbook().addWorksheet('S');
  values.forEach((v, i) => {
    ws.setCell(i + 1, 1, v, style as never);
  });
  return autoFitColumns(ws, opts)[0]!;
}

function near(actual: number, expected: number, tol = 0.1): void {
  expect(Math.abs(actual - expected) / expected).toBeLessThanOrEqual(tol);
}

describe('autoFitColumns: text', () => {
  it.each([
    ['1234567890', 70],
    ['iiiiiiiiii', 30],
    ['WWWWWWWWWW', 120],
    ['Hello World', 9 + 7 + 3 + 3 + 7 + 3 + 12 + 7 + 5 + 3 + 7],
  ])('%s is within 10%% of the Excel formula', (text, px) => {
    near(fitCol([text]), excel(px));
  });

  it('digits are exact', () => {
    expect(fitCol(['1234567890'])).toBeCloseTo(excel(70), 1);
  });

  it('default padding is 1.5 characters', () => {
    expect(fitCol(['1234567890'], undefined, {})).toBeCloseTo(11.5, 5);
  });

  it('takes the widest cell', () => {
    near(
      fitCol(['a', 'a much longer text value', 'bb']),
      fitCol(['a much longer text value']),
      0.001,
    );
  });

  it('CJK and emoji count as two characters', () => {
    near(fitCol(['日本語']), excel(6 * 7));
    near(fitCol(['😀']), excel(14));
  });

  it('uses the widest line of multi-line text, wrapped or not', () => {
    const single = fitCol(['short', 'the longest line']);
    expect(fitCol(['short\nthe longest line\nmid'])).toBeCloseTo(single, 5);
    expect(fitCol(['short\nthe longest line'], { alignment: { wrapText: true } })).toBeCloseTo(
      single,
      5,
    );
  });
});

describe('autoFitColumns: fonts', () => {
  it('bold is ~1.1x wider, font size scales linearly', () => {
    const base = fitCol(['Quarterly report'], undefined, { padding: 0 });
    near(fitCol(['Quarterly report'], { font: { bold: true } }, { padding: 0 }), base * 1.1, 0.001);
    near(fitCol(['Quarterly report'], { font: { size: 22 } }, { padding: 0 }), base * 2, 0.001);
    near(
      fitCol(['Quarterly report'], { font: { size: 8 } }, { padding: 0 }),
      (base * 8) / 11,
      0.001,
    );
  });

  it('rich text uses the widest run font', () => {
    const plain = fitCol(['abcdef ghijkl'], undefined, { padding: 0 });
    const rich = fitCol(
      [{ richText: [{ text: 'abcdef ' }, { text: 'ghijkl', font: { size: 22 } }] }],
      undefined,
      { padding: 0 },
    );
    near(rich, plain * 2, 0.001);
  });

  it('a larger default font rescales the unit', () => {
    const wb = new Workbook();
    wb.styles.setDefaultFont({ name: 'Calibri', size: 22 });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 'abc');
    const [w] = autoFitColumns(ws, { padding: 0 });
    // text in the default font measures the same number of default-font characters
    expect(w).toBeCloseTo(fitCol(['abc'], undefined, { padding: 0 }), 5);
  });
});

describe('autoFitColumns: displayed values', () => {
  const num = (v: number, numFmt?: string): number => fitCol([v], numFmt ? { numFmt } : undefined);

  it('General: integers, decimals capped at 11 characters, exponent', () => {
    near(num(12345), excel(35), 0.001);
    near(num(-12345), excel(35 + 4.2));
    near(num(1 / 3), excel(70 + 3.6)); // 0.333333333
    near(num(Math.PI), excel(70 + 3.6)); // 3.141592654
    near(num(1.5), excel(7 + 3.6 + 7));
    near(num(1.23e20), excel(5 * 7 + 3.6 + 6.7 + 6.9)); // 1.23E+20
  });

  it('fixed decimals and thousands separators', () => {
    near(num(1234.5, '0.00'), excel(6 * 7 + 3.6));
    near(num(12345678, '#,##0'), excel(8 * 7 + 2 * 3.6));
    near(num(1234567.891, '#,##0.00'), excel(9 * 7 + 3.6 * 3));
    near(num(0.5, '0.00'), excel(3 * 7 + 3.6));
  });

  it('percent', () => {
    near(num(0.256, '0.0%'), excel(3 * 7 + 3.6 + 9.9));
    near(num(12, '0%'), excel(4 * 7 + 9.9));
  });

  it('currency literals, parentheses and negative sections', () => {
    near(num(1234.5, '"$"#,##0.00'), excel(6 * 7 + 7 + 3.6 * 2));
    near(num(-1234.5, '#,##0.00;(#,##0.00)'), excel(6 * 7 + 3.6 * 2 + 2 * 4.2));
    near(num(-5, '0.00'), excel(3 * 7 + 3.6 + 4.2));
  });

  it('scientific', () => {
    near(num(12345, '0.00E+00'), excel(3 * 7 + 3.6 + 7 + 7 * 2 + 7));
  });

  it('dates use the width of the format', () => {
    const d = new Date(Date.UTC(2024, 0, 5));
    near(fitCol([d]), excel(8 * 7 + 2 * 4.2)); // default yyyy-mm-dd
    near(fitCol([new Date(Date.UTC(2024, 0, 5, 13, 30))]), excel(14 * 7 + 4.2 * 2 + 3.7 * 2 + 3));
    near(fitCol([d], { numFmt: 'dd/mm/yy' }), excel(6 * 7 + 2 * 4));
    near(fitCol([44000], { numFmt: 'm/d/yy' }), excel(6 * 7 + 2 * 4)); // serial in a date style
    // long names use the widest month/day: an upper bound, never narrower than the real thing
    expect(fitCol([d], { numFmt: 'dddd, mmmm d, yyyy' })).toBeGreaterThan(22);
  });

  it('booleans, errors, formula results, hyperlinks', () => {
    near(fitCol([true]), excel(29.8));
    near(fitCol([false]), excel(33.2));
    near(fitCol([{ error: '#DIV/0!' }]), excel(42.7)); // # D I V / 0 !
    near(fitCol([{ formula: 'A1*2', result: 123456 }]), excel(42));
    near(fitCol([{ formula: 'A1&"x"', result: 'hello world' }]), fitCol(['hello world']), 0.001);
    expect(fitCol([{ formula: 'NOW()' }])).toBeUndefined(); // no cached result -> nothing measured
    near(
      fitCol([{ text: 'turbosheet', hyperlink: 'https://example.com/very/long' }]),
      fitCol(['turbosheet']),
      0.001,
    );
  });

  it('rotated text gets a small fixed width', () => {
    const label = 'a very long rotated label';
    expect(fitCol([label], { alignment: { textRotation: 90 } })).toBeLessThan(4);
    expect(fitCol([label], { alignment: { textRotation: 255 } })).toBeLessThan(4);
  });

  it('indent adds width', () => {
    expect(fitCol(['abc'], { alignment: { indent: 2 } })).toBeGreaterThan(fitCol(['abc']) + 2);
  });
});

describe('autoFitColumns: options', () => {
  function sheet() {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('A1', 'Header that is rather long');
    ws.setCell('A2', 'x');
    ws.setCell('B1', 'h');
    ws.setCell('B2', 'a-fairly-long-value-here');
    ws.setCell('C3', 'third');
    return ws;
  }

  it('applies widths to the columns and returns them', () => {
    const ws = sheet();
    const w = autoFitColumns(ws);
    expect(w[0]).toBe(ws.getColumn(1).width);
    expect(w[1]).toBe(ws.getColumn(2).width);
    expect(w[3]).toBeUndefined();
    expect(w.length).toBe(3);
  });

  it('apply: false only returns', () => {
    const ws = sheet();
    autoFitColumns(ws, { apply: false });
    expect(ws.getColumn(1).width).toBeUndefined();
  });

  it('columns by number, letter and key; min fills empty requested columns', () => {
    const ws = sheet();
    ws.getColumn(2).key = 'second';
    const w = autoFitColumns(ws, { columns: [1, 'second', 'E'], min: 8 });
    expect(w[0]).toBeGreaterThan(8);
    expect(w[1]).toBeGreaterThan(8);
    expect(w[2]).toBeUndefined();
    expect(w[4]).toBe(8);
    expect(ws.getColumn(3).width).toBeUndefined();
  });

  it('empty columns are untouched by default', () => {
    const w = autoFitColumns(sheet(), { columns: ['E'] });
    expect(w[4]).toBeUndefined();
  });

  it('min and max clamp', () => {
    const ws = sheet();
    const [a, b] = autoFitColumns(ws, { min: 30, max: 40 });
    expect(a).toBe(30);
    expect(b).toBe(30);
    expect(autoFitColumns(ws, { max: 10 })[1]).toBe(10);
    ws.setCell('D1', 'x'.repeat(500));
    expect(autoFitColumns(ws)[3]).toBe(100);
  });

  it('includeHeader: false skips the first row', () => {
    const ws = sheet();
    const [withH] = autoFitColumns(ws, { apply: false });
    const [noH] = autoFitColumns(ws, { includeHeader: false, apply: false });
    expect(noH).toBeLessThan(withH!);
    near(noH!, fitCol(['x'], undefined, {}), 0.001);
  });

  it('rows range and sample', () => {
    const ws = new Workbook().addWorksheet('S');
    for (let r = 1; r <= 100; r++) ws.setCell(r, 1, r === 100 ? 'z'.repeat(40) : 'ab');
    const small = fitCol(['ab'], undefined, {});
    const wide = autoFitColumns(ws, { apply: false })[0]!;
    expect(wide).toBeGreaterThan(30);
    expect(autoFitColumns(ws, { sample: 10, apply: false })[0]).toBeCloseTo(small, 5);
    expect(autoFitColumns(ws, { rows: { from: 2, to: 99 }, apply: false })[0]).toBeCloseTo(
      small,
      5,
    );
    expect(autoFitColumns(ws, { rows: { from: 100 }, apply: false })[0]).toBeCloseTo(wide, 5);
  });

  it('skips cells of multi-column merges but not of single-column merges', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.setCell('A1', 'a title that spans several columns');
    ws.mergeCells('A1:D1');
    ws.setCell('A2', 'ab');
    ws.setCell('B3', 'long vertical merge text');
    ws.mergeCells('B3:B5');
    const w = autoFitColumns(ws, { apply: false });
    expect(w[0]).toBeCloseTo(fitCol(['ab'], undefined, {}), 5);
    expect(w[1]).toBeGreaterThan(15);
  });

  it('measures headers written via ws.columns', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.columns = [
      { header: 'Customer name', key: 'n' },
      { header: 'Qty', key: 'q' },
    ];
    ws.getRow(2).setValue(1, 'Al');
    const w = autoFitColumns(ws);
    expect(w[0]).toBeGreaterThan(w[1]!);
  });

  it('is exported from lite as well', () => {
    expect(lite.autoFitColumns).toBe(autoFitColumns);
    expect(lite.estimateColumnWidths).toBe(estimateColumnWidths);
  });
});

describe('estimateColumnWidths', () => {
  const rows = [
    ['Name', 'Amount', 'When', 'OK'],
    ['Alexander Hamilton', 1234567.891, new Date(Date.UTC(2024, 0, 5)), true],
    ['Bo', 5, null, false],
  ];

  it('matches autoFitColumns on the same data', () => {
    const ws = new Workbook().addWorksheet('S');
    rows.forEach((r, i) => {
      r.forEach((v, j) => {
        ws.setCell(i + 1, j + 1, v);
      });
    });
    const real = autoFitColumns(ws, { apply: false });
    const est = estimateColumnWidths(rows);
    for (let c = 0; c < 4; c++) expect(est[c]).toBeCloseTo(real[c]!, 5);
  });

  it('numFmts, headerBold, includeHeader, sample, min/max', () => {
    const plain = estimateColumnWidths(rows)[1]!;
    expect(estimateColumnWidths(rows, { numFmts: [undefined, '0'] })[1]!).toBeLessThan(plain);
    const two = [['Quarterly totals'], ['x']];
    expect(estimateColumnWidths(two, { headerBold: true })[0]!).toBeGreaterThan(
      estimateColumnWidths(two)[0]!,
    );
    expect(estimateColumnWidths(rows, { includeHeader: false })[0]!).toBeGreaterThan(10);
    expect(estimateColumnWidths(rows, { sample: 1 })[0]!).toBeLessThan(8);
    expect(estimateColumnWidths(rows, { min: 50 })[0]).toBe(50);
    expect(estimateColumnWidths(rows, { max: 5 })[0]).toBe(5);
  });

  it('accepts any iterable and leaves empty columns undefined', () => {
    function* gen(): Generator<(string | null)[]> {
      yield ['a', null, 'ccc'];
    }
    const w = estimateColumnWidths(gen());
    expect(w[1]).toBeUndefined();
    expect(w[2]).toBeGreaterThan(w[0]!);
  });
});

describe('autoFitColumns: performance', () => {
  it('1M cells (200k rows x 5 columns, mixed content)', () => {
    const ws = new Workbook().addWorksheet('S');
    const date = new Date(Date.UTC(2024, 5, 1));
    const money = { numFmt: '#,##0.00' };
    for (let r = 1; r <= 200_000; r++) {
      ws.setCell(r, 1, `Customer ${r % 977} with a name`);
      ws.setCell(r, 2, r);
      ws.setCell(r, 3, r * 1.37, money);
      ws.setCell(r, 4, date);
      ws.setCell(r, 5, r % 2 === 0);
    }
    const t0 = performance.now();
    const w = autoFitColumns(ws, { apply: false });
    const ms = performance.now() - t0;
    console.log(`autoFitColumns: 1,000,000 cells in ${ms.toFixed(0)} ms`);
    expect(w.length).toBe(5);
    expect(ms).toBeLessThan(1000);
  });
});
