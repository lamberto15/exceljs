// Every code sample in README.md and docs/*.md lives here (same code, imports from '../src').
// If you change a sample in the docs, change it here too: this file keeps the docs honest.

import {
  createReadStream,
  createWriteStream,
  mkdtempSync,
  openAsBlob,
  readFileSync,
  rmSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { finished } from 'node:stream/promises';
import { describe, expect, it } from 'vitest';
import { decryptXlsx, encryptXlsx, isEncrypted, WrongPasswordError } from '../src/crypto/index.js';
import { csvRows, csvStream, parseCsv, writeCsv } from '../src/csv/index.js';
import {
  createEvaluator,
  evaluate,
  evaluateCell,
  evaluateFormula,
  toNumber,
  toText,
} from '../src/formula/index.js';
import {
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addImage,
  addNote,
  addTable,
  createXlsxWriter,
  EncryptedFileError,
  listPreservedParts,
  placeImage,
  readXlsx,
  setPrintArea,
  Workbook,
  writeXlsx,
  XlsxParseError,
  xlsxRows,
  xlsxStream,
} from '../src/index.js';
import { objectsToSheet, sheetToNdjson, sheetToObjects } from '../src/json/index.js';
import * as lite from '../src/lite.js';
import { readXls } from '../src/xls/index.js';
import { readXlsb } from '../src/xlsb/index.js';

const tmp = (): string => mkdtempSync(join(tmpdir(), 'turbosheet-readme-'));

describe('README quick start', () => {
  it('write and read', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sales');
    ws.columns = [
      { header: 'Region', key: 'region', width: 14 },
      { header: 'Units', key: 'units', width: 10 },
      { header: 'Shipped', key: 'shipped', width: 12 },
    ];
    ws.addRow({ region: 'North', units: 120, shipped: new Date(Date.UTC(2025, 0, 15)) });
    ws.addRow(['South', 80, new Date(Date.UTC(2025, 1, 3))]);
    ws.setCell('B4', { formula: 'SUM(B2:B3)', result: 200 });
    ws.setStyle('A1', { font: { bold: true } });
    ws.getCell('C2').numFmt = 'yyyy-mm-dd';
    ws.freeze(1);

    const bytes: Uint8Array = writeXlsx(wb); // synchronous

    const back = readXlsx(bytes);
    const sheet = back.getWorksheet('Sales')!;
    expect(sheet.getValue('A2')).toBe('North');
    expect(sheet.getValue(3, 2)).toBe(80); // (row, col), both 1-based
    expect(sheet.getValue('C2')).toEqual(new Date(Date.UTC(2025, 0, 15)));
    const seen: number[] = [];
    sheet.eachRow((row, n) => {
      seen.push(n);
      void row.values;
    });
    expect(seen).toEqual([1, 2, 3, 4]);
  });

  it('stream write to a file, stream read rows', async () => {
    const dir = tmp();
    try {
      const path = join(dir, 'big.xlsx');
      const file = createWriteStream(path);
      const w = createXlsxWriter({
        sink: (chunk) => new Promise((resolve) => file.write(chunk, () => resolve())),
      });
      const sheet = w.addWorksheet('Data', {
        columns: [
          { header: 'Id', key: 'id', width: 10 },
          { header: 'Name', key: 'name', width: 24 },
        ],
        freeze: 1,
      });
      for (let i = 1; i <= 5000; i++) sheet.addRow({ id: i, name: `row ${i}` });
      await sheet.commit();
      await w.end();
      file.end();
      await finished(file);

      let n = 0;
      let last: unknown;
      for await (const row of xlsxRows(await openAsBlob(path))) {
        // row.sheet, row.row (1-based), row.values (0-based by column)
        n++;
        last = row.cells[1];
      }
      expect(n).toBe(5001);
      expect(last).toBe('row 5000');

      // streams work too (createXlsxWriter output is buffered per sheet when streamed)
      let m = 0;
      const ends: string[] = [];
      const src = createReadStream(path);
      for await (const row of xlsxRows(src, { onSheetEnd: ({ sheet }) => ends.push(sheet) })) {
        if (row.sheet === 'Data') m++;
      }
      expect(m).toBe(5001);
      expect(ends).toEqual(['Data']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('xlsxStream returns a web ReadableStream', async () => {
    const wb = new Workbook();
    wb.addWorksheet('A').addRow([1, 2, 3]);
    const res = new Response(xlsxStream(wb, { profile: 'max-speed' }));
    const bytes = new Uint8Array(await res.arrayBuffer());
    expect(readXlsx(bytes).getWorksheet('A')!.getValue('C1')).toBe(3);
  });

  it('csv', async () => {
    const text = writeCsv(
      [
        ['name', 'note'],
        ['Ann', 'said "hi", left'],
        ['Bob', '=HYPERLINK("http://evil")'],
      ],
      { escapeFormulas: true },
    );
    const rows = parseCsv(text);
    expect(rows[1]).toEqual(['Ann', 'said "hi", left']);
    expect(rows[2]![1]!.startsWith("'=")).toBe(true);
    expect(parseCsv('a,1\r\nb,2', { parseNumbers: true })).toEqual([
      ['a', 1],
      ['b', 2],
    ]);
    expect(
      parseCsv('TRUE,#N/A,04/11/2019', {
        parseBooleans: true,
        parseErrors: true,
        dateFormats: ['DD/MM/YYYY'],
      }),
    ).toEqual([[true, { error: '#N/A' }, new Date('2019-11-04T00:00:00.000Z')]]);

    const bytes = new Uint8Array(await new Response(csvStream([[1, 'x']])).arrayBuffer());
    const got: string[][] = [];
    for await (const r of csvRows([bytes])) got.push(r);
    expect(got).toEqual([['1', 'x']]);
    const win1252 = new Uint8Array([0x63, 0x61, 0x66, 0xe9]); // 'café' saved by Excel on Windows
    const auto: string[][] = [];
    for await (const r of csvRows([win1252], { encoding: 'auto' })) auto.push(r);
    expect(auto).toEqual([['café']]);
  });

  it('json helpers', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('People');
    objectsToSheet(ws, [
      { name: 'Ann', age: 31 },
      { name: 'Bob', age: 42 },
    ]);
    expect(sheetToObjects(ws)).toEqual([
      { name: 'Ann', age: 31 },
      { name: 'Bob', age: 42 },
    ]);
    const lines: string[] = [];
    for await (const line of sheetToNdjson(ws)) lines.push(line);
    expect(lines[0]).toBe('{"name":"Ann","age":31}\n');
  });

  it('encryption', async () => {
    const wb = new Workbook();
    wb.addWorksheet('Secret').addRow(['top', 'secret']);
    const encrypted = await encryptXlsx(writeXlsx(wb), 'hunter2');
    expect(isEncrypted(encrypted)).toBe(true);
    expect(() => readXlsx(encrypted)).toThrow(EncryptedFileError);
    await expect(decryptXlsx(encrypted, 'wrong')).rejects.toThrow(WrongPasswordError);
    const back = readXlsx(await decryptXlsx(encrypted, 'hunter2'));
    expect(back.getWorksheet('Secret')!.getValue('B1')).toBe('secret');
  }, 30_000);

  it('.xlsb and legacy .xls (README, docs/features.md)', () => {
    const file = (n: string) =>
      new Uint8Array(readFileSync(new URL(`./fixtures/binary/${n}`, import.meta.url)));
    const xlsbBytes = file('excel-saved.xlsb');
    const xlsBytes = file('excel-saved.xls');
    const fromXlsb = readXlsb(xlsbBytes); // same Workbook model as readXlsx
    const fromXls = readXls(xlsBytes); // Excel 97-2003 (BIFF8)
    const xlsx = writeXlsx(fromXls); // convert
    expect(fromXlsb.getWorksheet('Data')?.getValue('B1')).toBe(42);
    expect(readXlsx(xlsx).getWorksheet('Data')?.getValue('A5')).toEqual({
      formula: 'B1*2',
      result: 84,
    });
    for (const name of ['excel-saved.xls', 'excel-saved.xlsb']) {
      const bytes = file(name);
      const wb = name.endsWith('.xls') ? readXls(bytes) : readXlsb(bytes);
      expect(wb.getWorksheet('Data')?.getValue('B2')).toBe(false);
    }
  });

  it('round-trip preservation', () => {
    const wb = new Workbook();
    wb.addWorksheet('S').addRow([1]);
    const bytes = writeXlsx(wb);
    const loaded = readXlsx(bytes); // preserveUnknown defaults to true
    expect(listPreservedParts(loaded)).toBeInstanceOf(Array);
    expect(readXlsx(writeXlsx(loaded)).getWorksheet('S')!.getValue('A1')).toBe(1);
    expect(() => readXlsx(new Uint8Array([1, 2, 3, 4]))).toThrow();
    expect(XlsxParseError).toBeDefined();
  });
});

describe('README features', () => {
  it('tables, notes, validation, conditional formatting, images, names', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    addTable(ws, {
      name: 'Sales',
      columns: ['Region', { name: 'Units', totalsRowFunction: 'sum' }],
      rows: [
        ['North', 120],
        ['South', 80],
      ],
      totalsRow: true,
    });
    addNote(ws, 'A1', { text: 'Region name', author: 'me' });
    addDataValidation(ws, { sqref: 'A2:A3', type: 'list', values: ['North', 'South'] });
    addConditionalFormatting(ws, 'B2:B3', [
      { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } },
    ]);
    const png = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      ),
      (c) => c.charCodeAt(0),
    );
    const id = addImage(wb, { data: png, extension: 'png' });
    placeImage(ws, id, 'D2:F6');
    addDefinedName(wb, 'Rate', '0.2');
    setPrintArea(ws, 'A1:C10');
    ws.protection = { password: 's3cret', spinCount: 1000 };
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    const back = readXlsx(writeXlsx(wb));
    const d = back.getWorksheet('Data')!;
    expect(d.tables?.[0]?.name).toBe('Sales');
    expect(d.comments?.[0]?.ref).toBe('A1');
    expect(d.dataValidations?.[0]?.type).toBe('list');
    expect(d.conditionalFormatting).toHaveLength(1);
    expect(d.images).toHaveLength(1);
    expect(d.pageSetup?.orientation).toBe('landscape');
  });
});

describe('README lite', () => {
  it('opt-in features', () => {
    const wb = new lite.Workbook();
    lite.addTable(wb.addWorksheet('Data'), {
      name: 'T',
      columns: ['Region', 'Units'],
      rows: [['N', 1]],
    });
    const bytes = lite.writeXlsx(wb, { features: [lite.tablesWriter] });
    const back = lite.readXlsx(bytes, { features: [lite.tablesReader] });
    expect(back.getWorksheet('Data')!.tables).toHaveLength(1);

    const plain = new lite.Workbook();
    plain.addWorksheet('Plain').addRow([1, 'a']);
    expect(lite.readXlsx(lite.writeXlsx(plain)).getWorksheet('Plain')!.getValue('B1')).toBe('a');

    // Using a feature without enabling it throws instead of silently dropping data.
    expect(() => lite.writeXlsx(wb)).toThrow(/tables/);
  });
});

describe('README / docs/formula.md formulas', () => {
  it('evaluate, evaluateFormula (README)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['Units', 'Price']);
    ws.addRow([12, 2.5]);
    ws.setCell('C2', { formula: 'A2*B2' });
    const report = evaluate(wb); // fills Formula.result of every formula, in dependency order
    expect(report.circular).toEqual([]);
    expect(report.unsupported).toEqual([]);
    expect(ws.getValue('C2')).toEqual({ formula: 'A2*B2', result: 30 });
    expect(evaluateFormula(wb, 'SUMIF(A:A,">10")')).toBe(12); // ad-hoc
  });

  it('read, evaluate, write (docs/formula.md)', () => {
    const src = new Workbook();
    const s = src.addWorksheet('Sheet1');
    for (let r = 2; r <= 9; r++) s.setCell(r, 2, r);
    s.setCell('D10', { formula: 'SUM(B2:B9)*Rate' });
    addDefinedName(src, 'Rate', '0.5');
    const bytes = writeXlsx(src);

    const wb = readXlsx(bytes);
    const report = evaluate(wb); // every formula cell gets a fresh Formula.result
    expect(report).toEqual({ evaluated: 1, circular: [], unsupported: [], errors: [] });
    const out = writeXlsx(wb); // the results are written as cached values (<v>)
    expect(readXlsx(out).getWorksheet()!.getValue('D10')).toEqual({
      formula: 'SUM(B2:B9)*Rate',
      result: 22,
    });
    expect(evaluateCell(wb, 'Sheet1', 'D10')).toBe(22); // one cell and what it depends on
    expect(evaluateFormula(wb, 'SUM(B2:B9)*Rate', { sheet: 'Sheet1' })).toBe(22);
  });

  it('custom functions (docs/formula.md)', () => {
    const ev = createEvaluator({
      functions: { DOUBLE: (args) => toNumber(args[0]!) * 2 },
      now: () => new Date(2024, 0, 15),
    });
    ev.registerFunction('GREET', (args) => `Hello ${toText(args[0]!)}`);
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 21);
    ws.setCell('B1', { formula: 'DOUBLE(A1)&" "&GREET("you")&" "&TODAY()' });
    ev.evaluate(wb);
    expect((ws.getValue('B1') as { result: unknown }).result).toBe('42 Hello you 45306');
  });
});
