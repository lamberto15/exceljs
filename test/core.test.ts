import { describe, expect, it } from 'vitest';
import { SharedStrings } from '../src/core/sharedstrings.js';
import { StyleRegistry } from '../src/core/styles.js';
import {
  BUILTIN_NUM_FMTS,
  builtinFormatId,
  CellKind,
  dateToSerial,
  decodeAddr,
  decodeCol,
  decodeRange,
  encodeAddr,
  encodeCol,
  encodeRange,
  isDateFormat,
  serialToDate,
  Workbook,
} from '../src/index.js';

describe('address', () => {
  it('encodes/decodes edge columns', () => {
    const cases: [number, string][] = [
      [0, 'A'],
      [25, 'Z'],
      [26, 'AA'],
      [51, 'AZ'],
      [52, 'BA'],
      [701, 'ZZ'],
      [702, 'AAA'],
      [16383, 'XFD'],
    ];
    for (const [n, s] of cases) {
      expect(encodeCol(n)).toBe(s);
      expect(decodeCol(s)).toBe(n);
    }
    expect(decodeCol('xfd')).toBe(16383);
  });

  it('round-trips every column', () => {
    for (let c = 0; c < 16384; c++) expect(decodeCol(encodeCol(c))).toBe(c);
  });

  it('rejects out of bounds', () => {
    expect(() => encodeCol(16384)).toThrow();
    expect(() => encodeCol(-1)).toThrow();
    expect(() => decodeCol('XFE')).toThrow();
    expect(() => decodeCol('A1')).toThrow();
    expect(() => decodeAddr('A0')).toThrow();
    expect(() => decodeAddr('A1048577')).toThrow();
    expect(() => decodeAddr('XFE1')).toThrow();
    expect(() => decodeAddr('1A')).toThrow();
    expect(() => decodeAddr('')).toThrow();
    expect(() => encodeAddr(1048576, 0)).toThrow();
  });

  it('addresses', () => {
    expect(decodeAddr('B3')).toEqual({ row: 2, col: 1 });
    expect(decodeAddr('$B$3')).toEqual({ row: 2, col: 1 });
    expect(decodeAddr("'My Sheet'!c10")).toEqual({ row: 9, col: 2 });
    expect(decodeAddr('XFD1048576')).toEqual({ row: 1048575, col: 16383 });
    expect(encodeAddr(2, 1)).toBe('B3');
    expect(encodeAddr(1048575, 16383)).toBe('XFD1048576');
  });

  it('ranges', () => {
    expect(decodeRange('A1:C3')).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
    expect(decodeRange('C3:A1')).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
    expect(decodeRange('B2')).toEqual({ top: 1, left: 1, bottom: 1, right: 1 });
    expect(decodeRange('$A$1:$B$2')).toEqual({ top: 0, left: 0, bottom: 1, right: 1 });
    expect(decodeRange('B:D')).toEqual({ top: 0, left: 1, bottom: 1048575, right: 3 });
    expect(decodeRange('2:4')).toEqual({ top: 1, left: 0, bottom: 3, right: 16383 });
    expect(() => decodeRange('A1:B')).toThrow();
    expect(encodeRange({ top: 0, left: 0, bottom: 2, right: 2 })).toBe('A1:C3');
    expect(encodeRange({ top: 4, left: 1, bottom: 4, right: 1 })).toBe('B5');
  });
});

describe('dates', () => {
  const utc = (y: number, m: number, d: number, h = 0, mi = 0, s = 0) =>
    new Date(Date.UTC(y, m - 1, d, h, mi, s));

  it('1900 system incl. leap-year bug', () => {
    expect(dateToSerial(utc(1900, 1, 1))).toBe(1);
    expect(dateToSerial(utc(1900, 2, 28))).toBe(59);
    expect(dateToSerial(utc(1900, 3, 1))).toBe(61);
    expect(dateToSerial(utc(1970, 1, 1))).toBe(25569);
    expect(dateToSerial(utc(2024, 1, 1))).toBe(45292);
    expect(serialToDate(1).toISOString()).toBe('1900-01-01T00:00:00.000Z');
    expect(serialToDate(59).toISOString()).toBe('1900-02-28T00:00:00.000Z');
    expect(serialToDate(60).toISOString()).toBe('1900-03-01T00:00:00.000Z'); // fictitious 02-29
    expect(serialToDate(61).toISOString()).toBe('1900-03-01T00:00:00.000Z');
    expect(serialToDate(45292).toISOString()).toBe('2024-01-01T00:00:00.000Z');
  });

  it('times', () => {
    expect(dateToSerial(utc(2024, 1, 1, 12))).toBe(45292.5);
    expect(dateToSerial(utc(2024, 1, 1, 6))).toBe(45292.25);
    expect(serialToDate(45292.75).toISOString()).toBe('2024-01-01T18:00:00.000Z');
    const d = utc(2023, 7, 15, 13, 45, 30);
    expect(serialToDate(dateToSerial(d)).getTime()).toBe(d.getTime());
    expect(serialToDate(0.5).toISOString()).toBe('1899-12-31T12:00:00.000Z');
  });

  it('1904 system', () => {
    expect(dateToSerial(utc(1904, 1, 1), true)).toBe(0);
    expect(dateToSerial(utc(1904, 1, 2), true)).toBe(1);
    expect(dateToSerial(utc(2024, 1, 1), true)).toBe(45292 - 1462);
    expect(serialToDate(0, true).toISOString()).toBe('1904-01-01T00:00:00.000Z');
    expect(serialToDate(43830.5, true).toISOString()).toBe('2024-01-01T12:00:00.000Z');
  });
});

describe('numfmt', () => {
  it('detects date formats', () => {
    const dates = [
      'mm-dd-yy',
      'd-mmm-yy',
      'yyyy-mm-dd',
      'yyyy-mm-dd hh:mm:ss',
      '[$-409]mmmm d, yyyy',
      '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy',
      'h:mm AM/PM',
      'h:mm:ss',
      '[h]:mm:ss',
      'mm:ss',
      '[mm]',
      'm/d/yy h:mm',
      'dd/mm/yyyy;@',
      '[Red]yyyy-mm-dd',
      'yyyy"年"m"月"d"日"',
      'm',
      'D',
    ];
    for (const f of dates) expect(isDateFormat(f), f).toBe(true);
  });

  it('rejects non-date formats', () => {
    const nums = [
      'General',
      '0',
      '0.00',
      '0.00%',
      '#,##0',
      '"Date:" 0',
      '[Red]#,##0',
      '#,##0 ;[Red](#,##0)',
      '0.00E+00',
      '# ?/?',
      '@',
      '_("$"* #,##0.00_)',
      '"days" 0',
      '\\d0',
      '[>=100][Red]0;[Blue]0',
      '[$€-407] #,##0.00',
      '',
    ];
    for (const f of nums) expect(isDateFormat(f), f).toBe(false);
  });

  it('builtin ids', () => {
    expect(builtinFormatId('General')).toBe(0);
    expect(builtinFormatId('0.00%')).toBe(10);
    expect(builtinFormatId('mm-dd-yy')).toBe(14);
    expect(builtinFormatId('@')).toBe(49);
    expect(builtinFormatId('yyyy-mm-dd')).toBeUndefined();
    for (const k of Object.keys(BUILTIN_NUM_FMTS)) {
      expect(builtinFormatId(BUILTIN_NUM_FMTS[+k]!)).toBe(+k);
    }
  });
});

describe('StyleRegistry', () => {
  it('reserves defaults', () => {
    const s = new StyleRegistry();
    expect(s.fills[0]).toEqual({ type: 'pattern', pattern: 'none' });
    expect(s.fills[1]).toEqual({ type: 'pattern', pattern: 'gray125' });
    expect(s.xfs.length).toBe(1);
    expect(s.intern(undefined)).toBe(0);
    expect(s.intern({})).toBe(0);
    expect(s.intern({ fill: { type: 'pattern', pattern: 'none' } })).toBe(0);
    expect(s.intern({ numFmt: 'General' })).toBe(0);
  });

  it('dedupes regardless of key order', () => {
    const s = new StyleRegistry();
    const a = s.intern({
      font: { bold: true, name: 'Arial', color: { argb: 'FFFF0000' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } },
      alignment: { horizontal: 'center', wrapText: true },
      numFmt: '0.00%',
    });
    const b = s.intern({
      numFmt: '0.00%',
      alignment: { wrapText: true, horizontal: 'center' },
      fill: { fgColor: { argb: 'ff00ff00' }, pattern: 'solid', type: 'pattern' },
      font: { color: { argb: 'FFFF0000' }, name: 'Arial', bold: true, italic: false },
    });
    expect(a).toBeGreaterThan(0);
    expect(b).toBe(a);
    expect(s.fonts.length).toBe(2);
    expect(s.fills.length).toBe(3);
    expect(s.numFmts.length).toBe(0); // 0.00% is built-in 10
    expect(s.xfs[a]!.numFmtId).toBe(10);
    expect(s.intern({ font: { bold: true } })).not.toBe(a);
  });

  it('custom numFmts start at 164 and round-trip via get()', () => {
    const s = new StyleRegistry();
    const x = s.intern({ numFmt: 'yyyy-mm-dd', font: { italic: true } });
    const y = s.intern({ numFmt: '0.000' });
    expect(s.numFmts).toEqual([
      { id: 164, formatCode: 'yyyy-mm-dd' },
      { id: 165, formatCode: '0.000' },
    ]);
    const st = s.get(x);
    expect(st.numFmt).toBe('yyyy-mm-dd');
    expect(st.font).toEqual({
      name: 'Calibri',
      size: 11,
      italic: true,
      family: 2,
      scheme: 'minor',
    });
    expect(s.intern({ ...st })).toBe(x);
    expect(s.isDateXf(x)).toBe(true);
    expect(s.isDateXf(y)).toBe(false);
  });

  it('reader hooks', () => {
    const s = new StyleRegistry();
    expect(s.registerNumFmt(170, 'dd/mm/yyyy')).toBe(170);
    const f = s.addFont({ name: 'Arial', size: 10 });
    const xf = s.internXf({ numFmtId: 170, fontId: f, fillId: 0, borderId: 0 });
    expect(s.get(xf).numFmt).toBe('dd/mm/yyyy');
    expect(s.isDateXf(xf)).toBe(true);
    expect(s.numFmtId('0.0')).toBe(171);
    expect(s.internXf({ numFmtId: 14, fontId: 0, fillId: 0, borderId: 0 })).toBeGreaterThan(0);
  });
});

describe('SharedStrings', () => {
  it('interns', () => {
    const ss = new SharedStrings();
    expect(ss.intern('a')).toBe(0);
    expect(ss.intern('b')).toBe(1);
    expect(ss.intern('a')).toBe(0);
    expect(ss.get(1)).toBe('b');
    expect(ss.count).toBe(3);
    expect(ss.uniqueCount).toBe(2);
  });
});

describe('Workbook / Worksheet', () => {
  it('manages sheets', () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('Data', { tabColor: { argb: 'FFFF0000' }, state: 'hidden' });
    const b = wb.addWorksheet();
    expect(b.name).toBe('Sheet2');
    expect(a.state).toBe('hidden');
    expect(wb.getWorksheet('data')).toBe(a);
    expect(wb.getWorksheet(2)).toBe(b);
    expect(() => wb.addWorksheet('DATA')).toThrow();
    expect(() => wb.addWorksheet('a/b')).toThrow();
    expect(() => wb.addWorksheet('x'.repeat(32))).toThrow();
    wb.removeWorksheet('Data');
    expect(wb.worksheets).toEqual([b]);
    expect(wb.getWorksheet()).toBe(b);
  });

  it('stores and reconstructs every value kind', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const d = new Date(Date.UTC(2024, 0, 2));
    ws.addRow([
      1.5,
      'text',
      true,
      d,
      { error: '#N/A' },
      { formula: '=A1*2', result: 3 },
      { richText: [{ text: 'a' }, { text: 'b', font: { bold: true } }] },
      { hyperlink: 'https://x.y', text: 'link' },
      Number.NaN,
      Number.POSITIVE_INFINITY,
      null,
      false,
    ]);
    expect(ws.getValue('A1')).toBe(1.5);
    expect(ws.getValue('B1')).toBe('text');
    expect(ws.getValue('C1')).toBe(true);
    expect(ws.getValue('D1')).toEqual(d);
    expect(ws.getKind('D1')).toBe(CellKind.Date);
    expect(ws.getStyle('D1').numFmt).toBe('yyyy-mm-dd');
    expect(ws.getValue('E1')).toEqual({ error: '#N/A' });
    expect(ws.getValue('F1')).toEqual({ formula: 'A1*2', result: 3 });
    expect(ws.getKind('G1')).toBe(CellKind.RichText);
    expect(ws.getValue('H1')).toEqual({ hyperlink: 'https://x.y', text: 'link' });
    expect(ws.getValue('I1')).toEqual({ error: '#NUM!' });
    expect(ws.getValue('J1')).toEqual({ error: '#NUM!' });
    expect(ws.getValue('K1')).toBeNull();
    expect(ws.getKind('K1')).toBe(CellKind.Empty);
    expect(ws.getValue('L1')).toBe(false);
    expect(ws.rowData[0]!.n).toBe(11);
    // datetime + explicit style without numFmt keeps the style and gets a datetime format
    ws.setCell('A2', new Date(Date.UTC(2024, 0, 2, 10, 30)), { font: { bold: true } });
    expect(ws.getStyle('A2').numFmt).toBe('yyyy-mm-dd hh:mm:ss');
    expect(ws.getStyle('A2').font?.bold).toBe(true);
    ws.setCell('B2', d, { numFmt: 'dd/mm/yyyy' });
    expect(ws.getStyle('B2').numFmt).toBe('dd/mm/yyyy');
  });

  it('respects date1904', () => {
    const wb = new Workbook({ date1904: true });
    const ws = wb.addWorksheet();
    const d = new Date(Date.UTC(2024, 0, 1));
    ws.setCell(1, 1, d);
    expect(ws.rowData[0]!.num[0]).toBe(45292 - 1462);
    expect(ws.getValue(1, 1)).toEqual(d);
  });

  it('inserts out of order and keeps columns sorted', () => {
    const ws = new Workbook().addWorksheet();
    ws.setCell('E1', 5);
    ws.setCell('A1', 'a');
    ws.setCell('C1', 3);
    ws.setCell('B1', 'b');
    ws.setCell('D1', 4);
    ws.setCell('C1', 'c'); // overwrite numeric with string
    const row = ws.rowData[0]!;
    expect(Array.from(row.col.subarray(0, row.n))).toEqual([0, 1, 2, 3, 4]);
    expect(ws.getRow(1).values).toEqual([undefined, 'a', 'b', 'c', 4, 5]);
    ws.setCell('B1', 2); // string → number clears obj slot
    expect(row.obj![1]).toBeUndefined();
    const seen: [number, number, unknown][] = [];
    ws.rowData[0]!.forEachCell((c, k, n, o) => {
      seen.push([c, k, k === CellKind.Number ? n : o]);
    });
    expect(seen).toEqual([
      [0, CellKind.String, 'a'],
      [1, CellKind.Number, 2],
      [2, CellKind.String, 'c'],
      [3, CellKind.Number, 4],
      [4, CellKind.Number, 5],
    ]);
    // many random inserts
    const ws2 = new Workbook().addWorksheet();
    const cols = Array.from({ length: 500 }, (_, i) => i + 1).sort(() => Math.random() - 0.5);
    for (const c of cols) ws2.setCell(7, c, c % 3 ? c : `s${c}`);
    for (let c = 1; c <= 500; c++) expect(ws2.getValue(7, c)).toBe(c % 3 ? c : `s${c}`);
  });

  it('clears cells', () => {
    const ws = new Workbook().addWorksheet();
    ws.addRow([1, 'x', 3]);
    ws.setCell('B1', null);
    expect(ws.getValue('B1')).toBeNull();
    expect(ws.rowData[0]!.n).toBe(2);
    expect(ws.getValue('C1')).toBe(3);
    // styled cell keeps formatting when its value is cleared
    ws.setCell('A2', 'y', { font: { bold: true } });
    ws.setCell('A2', undefined);
    expect(ws.getValue('A2')).toBeNull();
    expect(ws.getStyle('A2').font?.bold).toBe(true);
    expect(ws.rowData[1]!.n).toBe(1);
    ws.clearCell('A2');
    expect(ws.rowData[1]!.n).toBe(0);
    expect(ws.getStyle('A2')).toEqual({});
    ws.setStyle('D5', { numFmt: '0.0' });
    expect(ws.getKind('D5')).toBe(CellKind.Empty);
    ws.setStyle('D5', undefined);
    expect(ws.rowData[4]!.n).toBe(0);
  });

  it('merges reject overlaps', () => {
    const ws = new Workbook().addWorksheet();
    ws.mergeCells('A1:C2');
    expect(() => ws.mergeCells('C2:D4')).toThrow(/overlaps/);
    expect(() => ws.mergeCells('B1')).toThrow();
    ws.mergeCells(3, 1, 4, 2); // A3:B4
    ws.mergeCells('D1:E2');
    expect(ws.merges).toEqual(['A1:C2', 'A3:B4', 'D1:E2']);
    ws.unmergeCells('A1');
    expect(ws.merges).toEqual(['A3:B4', 'D1:E2']);
    ws.mergeCells('A1:C2');
  });

  it('merge overlap check scales (indexed) and stays correct after unmerge', () => {
    const ws = new Workbook().addWorksheet();
    const t = performance.now();
    for (let r = 1; r <= 40_000; r += 2) ws.mergeCells(r, 1, r + 1, 3);
    expect(performance.now() - t).toBeLessThan(2000);
    expect(ws.mergeRanges.length).toBe(20_000);
    expect(() => ws.mergeCells('C40000:D40001')).toThrow(/overlaps A39999:C40000/);
    expect(() => ws.mergeCells('B100:B300')).toThrow(/overlaps/);
    ws.mergeCells('D1:E40000'); // tall, no overlap
    expect(() => ws.mergeCells('E20000:F20001')).toThrow(/overlaps D1:E40000/);
    ws.unmergeCells('A99:A100');
    ws.mergeCells('A99:B100');
    expect(() => ws.mergeCells('B99:B100')).toThrow(/overlaps A99:B100/);
  });

  it('addMergeRaw skips validation; _ensureRow pre-sizes; styleIndex row props', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.addMergeRaw({ top: 0, left: 0, bottom: 1, right: 1 });
    ws.addMergeRaw({ top: 1, left: 1, bottom: 2, right: 2 }); // overlapping: accepted
    expect(ws.merges).toEqual(['A1:B2', 'B2:C3']);
    const row = ws._ensureRow(4, 32);
    expect(row.capacity).toBe(32);
    expect(ws._ensureRow(4)).toBe(row);
    expect(ws._ensureRow(4, 64).capacity).toBe(64);
    const xf = wb.styles.intern({ font: { bold: true } });
    ws.setRowProps(3, { styleIndex: xf, height: 20 });
    expect(ws.getRowProps(3).height).toBe(20);
    expect(ws.getRowProps(3).style?.font?.bold).toBe(true);
    expect(ws.rowData[2]!.xf).toBe(xf);
  });

  it('columns with headers and keys', () => {
    const ws = new Workbook().addWorksheet();
    ws.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'Price', key: 'price', style: { numFmt: '0.00' }, hidden: true },
    ];
    expect(ws.getRow(1).values).toEqual([undefined, 'Id', 'Name', 'Price']);
    const r = ws.addRow({ name: 'Widget', price: 9.5, id: 1, unknown: 'ignored' });
    expect(r).toBe(2);
    expect(ws.getRow(2).values).toEqual([undefined, 1, 'Widget', 9.5]);
    expect(ws.getStyle('C2').numFmt).toBe('0.00'); // inherited column style
    expect(ws.getStyle('A2')).toEqual({});
    expect(ws.getColumn('price').width).toBeUndefined();
    expect(ws.getColumn('price').hidden).toBe(true);
    expect(ws.getColumn('B').width).toBe(32);
    expect(ws.getColumn(1).key).toBe('id');
    expect(ws.getColumn('name').letter).toBe('B');
    expect(ws.addRow([2, 'Gadget'])).toBe(3);
    ws.addRows([[3], { id: 4 }]);
    expect(ws.rowCount).toBe(5);
    expect(ws.getValue(5, 'A')).toBe(4);
    expect(ws.columnCount).toBe(3);
    expect(ws.dimensions).toEqual({ top: 1, left: 1, bottom: 5, right: 3, ref: 'A1:C5' });
  });

  it('row props, styles, views, iteration', () => {
    const ws = new Workbook().addWorksheet();
    ws.setRowProps(3, { height: 30, hidden: true, style: { font: { italic: true } } });
    ws.setCell('B3', 'inherits');
    expect(ws.getStyle('B3').font?.italic).toBe(true);
    expect(ws.getRowProps(3)).toMatchObject({ height: 30, hidden: true });
    expect(ws.rowCount).toBe(3);
    ws.setCell(1, 2, 7);
    const visited: number[] = [];
    ws.eachRow((row, n) => {
      visited.push(n);
      expect(row.number).toBe(n);
    });
    expect(visited).toEqual([1, 3]);
    ws.freeze(1);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1, topLeftCell: 'A2' });
    expect(ws.views[0]!.xSplit).toBeUndefined();
    ws.autoFilter = 'A1:B1';
    const cell = ws.getCell('C4');
    cell.value = 12;
    cell.font = { bold: true };
    cell.numFmt = '0.0';
    expect(ws.getStyle('C4')).toMatchObject({ numFmt: '0.0', font: { bold: true } });
    expect(cell.value).toBe(12);
    expect(ws.getStyleIndex('C4')).toBe(ws.rowData[3]!.style[0]);
  });

  it('setCellRaw fast path', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.setCellRaw(0, 0, CellKind.String, 0, 'a', 0);
    ws.setCellRaw(0, 2, CellKind.Number, 3, undefined, 0);
    ws.setCellRaw(0, 1, CellKind.Boolean, 1, undefined, 0);
    expect(ws.getRow(1).values).toEqual([undefined, 'a', true, 3]);
  });

  it('memory: 1,000,000 numeric cells', async () => {
    // no @types/node in this package: load node builtins untyped
    type NodeProcess = { memoryUsage(): { heapUsed: number; arrayBuffers: number } };
    const v8: { setFlagsFromString(f: string): void } = await import('node:v8' as string);
    const vm: { runInNewContext(code: string): unknown } = await import('node:vm' as string);
    const proc = (globalThis as unknown as { process: NodeProcess }).process;
    v8.setFlagsFromString('--expose-gc');
    const gc = vm.runInNewContext('gc') as () => void;
    const measure = () => {
      gc();
      gc();
      const m = proc.memoryUsage();
      return m.heapUsed + m.arrayBuffers;
    };
    const ROWS = 50_000;
    const COLS = 20;
    const vals: number[] = new Array(COLS).fill(0);
    const before = measure();
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    const t0 = performance.now();
    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) vals[c] = r * COLS + c;
      ws.addRow(vals);
    }
    const ms = performance.now() - t0;
    const after = measure();
    const perCell = (after - before) / (ROWS * COLS);
    console.log(
      `[core] 1M numeric cells via addRow: ${ms.toFixed(0)} ms, ${perCell.toFixed(1)} bytes/cell (heap+arrayBuffers)`,
    );
    expect(ws.getValue(ROWS, COLS)).toBe(ROWS * COLS - 1);
    expect(perCell).toBeLessThan(60);

    // sequential setCell path (growth slack included)
    const before2 = measure();
    const ws2 = wb.addWorksheet();
    const t1 = performance.now();
    for (let r = 1; r <= ROWS; r++) for (let c = 1; c <= COLS; c++) ws2.setCell(r, c, r + c);
    const ms2 = performance.now() - t1;
    const per2 = (measure() - before2) / (ROWS * COLS);
    console.log(
      `[core] 1M numeric cells via setCell: ${ms2.toFixed(0)} ms, ${per2.toFixed(1)} bytes/cell`,
    );
    expect(ws2.getValue(ROWS, COLS)).toBe(ROWS + COLS);
    expect(per2).toBeLessThan(80);
  });
});
