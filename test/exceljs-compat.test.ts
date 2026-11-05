// ExcelJS migration compatibility (#80): `[options]` / includeEmpty on the iteration methods,
// maxRows / maxCols read limits and `{ sharedFormula }` cell values. exceljs 4.4.0 is the oracle
// for the iteration semantics.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  type CellValue,
  eachColumnCell,
  readXlsx,
  SheetLimitError,
  Workbook,
  writeXlsx,
  xlsxRows,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';
import { streamWrite } from './exceljs-parity/helpers.js';

/** Same cells in a turbosheet and an exceljs sheet: [address, value]. */
function both(cells: [string, number][]): {
  ws: ReturnType<Workbook['addWorksheet']>;
  x: ExcelJS.Worksheet;
} {
  const ws = new Workbook().addWorksheet('s');
  const x = new ExcelJS.Workbook().addWorksheet('s');
  for (const [a, v] of cells) {
    ws.setCell(a, v);
    x.getCell(a).value = v;
  }
  return { ws, x };
}

const CELLS: [string, number][] = [
  ['B2', 1],
  ['D2', 2],
  ['A5', 3],
  ['F7', 4],
];

describe('includeEmpty', () => {
  it('eachRow matches exceljs, with and without includeEmpty', () => {
    const { ws, x } = both(CELLS);
    for (const opts of [undefined, { includeEmpty: false }, { includeEmpty: true }]) {
      const ours: number[] = [];
      const theirs: number[] = [];
      if (opts) {
        ws.eachRow(opts, (_r, n) => ours.push(n));
        x.eachRow(opts, (_r, n) => theirs.push(n));
      } else {
        ws.eachRow((_r, n) => ours.push(n));
        x.eachRow((_r, n) => theirs.push(n));
      }
      expect(ours).toEqual(theirs);
    }
  });

  it("row.eachCell includeEmpty goes to the row's last cell (not the sheet width), like exceljs", () => {
    const { ws, x } = both(CELLS);
    for (const r of [2, 5, 7]) {
      const ours: [number, CellValue][] = [];
      const theirs: [number, CellValue][] = [];
      ws.getRow(r).eachCell({ includeEmpty: true }, (v, c) => ours.push([c, v]));
      x.getRow(r).eachCell({ includeEmpty: true }, (cell, c) =>
        theirs.push([c, cell.value as CellValue]),
      );
      expect(ours).toEqual(theirs);
    }
    // an empty row visits nothing
    const seen: number[] = [];
    ws.getRow(3).eachCell({ includeEmpty: true }, (_v, c) => seen.push(c));
    expect(seen).toEqual([]);
  });

  it('row.eachCell includeEmpty counts styled empty cells and passes effective styles', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.setCell('B1', 5);
    ws.setStyle('D1', { font: { bold: true } });
    const bold = ws.getStyleIndex('D1');
    ws.getColumn('A').style = { font: { italic: true } };
    const italic = ws.getColumn('A').styleIndex;
    const seen: [number, CellValue, number][] = [];
    ws.getRow(1).eachCell({ includeEmpty: true }, (v, c, xf) => seen.push([c, v, xf]));
    expect(seen).toEqual([
      [1, null, italic],
      [2, 5, ws.getStyleIndex('B1')],
      [3, null, 0],
      [4, null, bold],
    ]);
    // the fast path still skips empty (styled) cells
    const fast: number[] = [];
    ws.getRow(1).eachCell((_v, c) => fast.push(c));
    expect(fast).toEqual([2]);
  });

  it('column eachCell / eachColumnCell match exceljs for includeEmpty', () => {
    const ws = new Workbook().addWorksheet('s');
    const x = new ExcelJS.Workbook().addWorksheet('s');
    for (const [a, v] of [
      ['A1', 1],
      ['A3', 3],
      ['C6', 6],
    ] as const) {
      ws.setCell(a, v);
      x.getCell(a).value = v;
    }
    const theirs: [number, CellValue][] = [];
    x.getColumn('A').eachCell({ includeEmpty: true }, (cell, n) =>
      theirs.push([n, cell.value as CellValue]),
    );
    const ours: [number, CellValue][] = [];
    ws.getColumn('A').eachCell({ includeEmpty: true }, (v, n) => ours.push([n, v]));
    const fn: [number, CellValue][] = [];
    eachColumnCell(ws, 'A', { includeEmpty: true }, (v, n) => fn.push([n, v]));
    expect(ours).toEqual(theirs);
    expect(fn).toEqual(theirs);
    const plain: number[] = [];
    eachColumnCell(ws, 1, (_v, n) => plain.push(n));
    expect(plain).toEqual([1, 3]);
    // eachColumnCell does not create a column record
    expect(ws.columnData[2]).toBeUndefined();
  });

  it('a missing callback throws a TypeError naming the signature', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.setCell('A1', 1);
    expect(() => ws.eachRow({ includeEmpty: true } as never)).toThrow(
      'eachRow([options], callback): callback must be a function',
    );
    expect(() => ws.getRow(1).eachCell({} as never)).toThrow(TypeError);
    expect(() => ws.getColumn(1).eachCell({} as never)).toThrow(TypeError);
    expect(() => eachColumnCell(ws, 1, {} as never)).toThrow(TypeError);
  });
});

describe('maxRows / maxCols', () => {
  const book = (rows: number, cols: number): Uint8Array => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    for (let r = 1; r <= rows; r++) for (let c = 1; c <= cols; c++) ws.setCell(r, c, r * c);
    return writeXlsx(wb);
  };

  it('limits equal to the used range pass; one less throws', () => {
    const bytes = book(5, 4);
    expect(readXlsx(bytes, { maxRows: 5, maxCols: 4 }).worksheets[0]!.rowCount).toBe(5);
    expect(() => readXlsx(bytes, { maxRows: 4 })).toThrow(
      'Max row count (4) exceeded in sheet "Data" (row 5)',
    );
    expect(() => readXlsx(bytes, { maxCols: 3 })).toThrow(
      'Max column count (3) exceeded in sheet "Data" (column 4)',
    );
  });

  it('0 or unset means no limit', () => {
    const bytes = book(3, 3);
    expect(readXlsx(bytes, { maxRows: 0, maxCols: 0 }).worksheets[0]!.rowCount).toBe(3);
  });

  it('SheetLimitError carries the sheet, option and value', () => {
    try {
      readXlsx(book(3, 3), { maxRows: 2 });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(SheetLimitError);
      const err = e as SheetLimitError;
      expect(err.name).toBe('SheetLimitError');
      expect([err.sheet, err.limit, err.max]).toEqual(['Data', 'maxRows', 2]);
    }
  });

  it('only sheets that are read are checked (opts.sheets)', () => {
    const wb = new Workbook();
    wb.addWorksheet('Small').setCell('A1', 1);
    wb.addWorksheet('Big').setCell('A100', 1);
    const bytes = writeXlsx(wb);
    expect(readXlsx(bytes, { maxRows: 10, sheets: ['Small'] }).worksheets).toHaveLength(1);
    expect(() => readXlsx(bytes, { maxRows: 10 })).toThrow(SheetLimitError);
  });

  it('xlsxRows yields the rows before the limit, then throws', async () => {
    const rows: number[] = [];
    await expect(async () => {
      for await (const r of xlsxRows(book(50, 2), { maxRows: 10 })) rows.push(r.row);
    }).rejects.toThrow(SheetLimitError);
    expect(Math.max(...rows)).toBeLessThanOrEqual(10);
    await expect(async () => {
      for await (const _ of xlsxRows(book(2, 30), { maxCols: 20 })) {
        // drain
      }
    }).rejects.toThrow('Max column count (20) exceeded');
  });
});

describe('{ sharedFormula } values', () => {
  it('translates relative references, keeps absolute ones', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.setCell('C2', { formula: 'A2*$B$1+SUM(A$1:A2)' });
    ws.setCell('D4', { sharedFormula: 'C2' });
    expect(ws.getValue('D4')).toEqual({ formula: 'B4*$B$1+SUM(B$1:B4)' });
    ws.setCell(5, 3, { sharedFormula: '$C$2', result: 7 });
    expect(ws.getValue('C5')).toEqual({ formula: 'A5*$B$1+SUM(A$1:A5)', result: 7 });
  });

  it('works through addRow (array and keyed objects), setValue and chains', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.columns = [{ key: 'a' }, { key: 'b' }];
    ws.addRow([1, { formula: 'A1*2' }]);
    ws.addRow([2, { sharedFormula: 'B1' }]);
    ws.addRow({ a: 3, b: { sharedFormula: 'B2', result: 6 } });
    ws.getRow(4).setValue(2, { sharedFormula: 'B3' });
    expect(ws.getValue('B2')).toEqual({ formula: 'A2*2' });
    expect(ws.getValue('B3')).toEqual({ formula: 'A3*2', result: 6 });
    expect(ws.getValue('B4')).toEqual({ formula: 'A4*2' });
  });

  it('is written as an Excel shared formula (read back expanded)', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    ws.setCell('A1', 2);
    ws.setCell('B1', { formula: 'A1+1', result: 3 });
    ws.setCell('C1', { sharedFormula: 'B1', result: 4 });
    const bytes = writeXlsx(wb);
    expect(readXlsx(bytes).worksheets[0]!.getValue('C1')).toEqual({ formula: 'B1+1', result: 4 });
    const sheet = new TextDecoder().decode(unzipSync(bytes).get('xl/worksheets/sheet1.xml')!);
    expect(sheet).toContain('<f t="shared" ref="B1:C1" si="0">A1+1</f>');
    expect(sheet).toContain('<c r="C1"><f t="shared" si="0"/><v>4</v></c>');
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(bytes as unknown as ArrayBuffer);
    expect(x.worksheets[0]!.getCell('C1').value).toEqual({ sharedFormula: 'B1', result: 4 });
    expect(x.worksheets[0]!.getCell('C1').formula).toBe('B1+1');
  });

  it('a master without a formula throws, and addRow stays atomic', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.setCell('A1', 1);
    expect(() => ws.setCell('B1', { sharedFormula: 'A1' })).toThrow(
      'Shared formula master A1 holds no formula',
    );
    expect(() => ws.setCell('B1', { sharedFormula: 'Z9' })).toThrow(TypeError);
    expect(() => ws.setCell('B1', { sharedFormula: 'nope' })).toThrow(RangeError);
    expect(ws.getValue('B1')).toBeNull();
    expect(ws.rowData[0]!.n).toBe(1);
    expect(() => ws.addRow([5, { sharedFormula: 'A1' }])).toThrow(TypeError);
    expect(ws.rowCount).toBe(1);
  });

  it('streaming writer: the master must be in a row not yet flushed', async () => {
    const bytes = await streamWrite(
      (w) => {
        const s = w.addWorksheet('s');
        s.addRow([{ formula: 'ROW()' }]);
        s.addRow([{ sharedFormula: 'A1' }]); // same batch: fine (then flushed)
        expect(() => s.addRow([{ sharedFormula: 'A1' }])).toThrow(TypeError);
        s.addRow([{ formula: 'B1*2' }]);
        s.addRow([5, { sharedFormula: 'A3' }]); // absolute addresses in a later batch
      },
      { flushRows: 2 },
    );
    const ws = readXlsx(bytes).worksheets[0]!;
    expect(ws.getValue('A2')).toEqual({ formula: 'ROW()' });
    expect(ws.getValue('A3')).toEqual({ formula: 'B1*2' });
    expect(ws.getValue('B4')).toEqual({ formula: 'C2*2' });
  });
});
