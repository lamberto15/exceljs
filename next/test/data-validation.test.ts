/// <reference types="node" />
// Data validation (features/data-validation): ours ↔ ours in every writer, ours → exceljs,
// exceljs → ours, the x14 extLst form, list limits and element order.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { addDataValidation, type DataValidation } from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const sheetXml = (buf: Uint8Array, n = 1): string =>
  dec.decode(unzipSync(buf).get(`xl/worksheets/sheet${n}.xml`)!);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

type XModel = Record<string, Record<string, unknown> | undefined>;
const xdv = (ws: ExcelJS.Worksheet): XModel =>
  (ws as unknown as { dataValidations: { model: XModel } }).dataValidations.model;

const DATE = new Date(Date.UTC(2024, 0, 15));

const VALIDATIONS: DataValidation[] = [
  {
    sqref: 'A1:A3',
    type: 'list',
    values: ['Yes', 'No', 'Say "maybe"'],
    allowBlank: true,
    showErrorMessage: true,
    errorStyle: 'warning',
    errorTitle: 'Oops',
    error: 'Pick <one> & only one',
  },
  { sqref: 'B1', type: 'list', formula1: 'Lists!$A$1:$A$5', showInputMessage: true, prompt: 'p' },
  {
    sqref: 'C1:C2 E1',
    type: 'whole',
    operator: 'between',
    formula1: 1,
    formula2: 10,
    promptTitle: 'Range',
  },
  { sqref: 'D1', type: 'decimal', operator: 'greaterThan', formula1: 2.5 },
  { sqref: 'F1', type: 'date', operator: 'lessThan', formula1: DATE },
  { sqref: 'G1', type: 'textLength', operator: 'lessThanOrEqual', formula1: '=5' },
  { sqref: 'H1', type: 'custom', formula1: 'ISNUMBER(H1)', hideDropDown: true },
  { sqref: 'I1', type: 'any', showInputMessage: true, promptTitle: 'Note', prompt: 'anything' },
];

/** What our reader returns for VALIDATIONS (formulas become raw strings, values → formula1). */
const READ_BACK: DataValidation[] = [
  {
    sqref: 'A1:A3',
    type: 'list',
    formula1: '"Yes,No,Say ""maybe"""',
    allowBlank: true,
    showErrorMessage: true,
    errorStyle: 'warning',
    errorTitle: 'Oops',
    error: 'Pick <one> & only one',
  },
  { sqref: 'B1', type: 'list', formula1: 'Lists!$A$1:$A$5', showInputMessage: true, prompt: 'p' },
  {
    sqref: 'C1:C2 E1',
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '10',
    promptTitle: 'Range',
  },
  { sqref: 'D1', type: 'decimal', operator: 'greaterThan', formula1: '2.5' },
  { sqref: 'F1', type: 'date', operator: 'lessThan', formula1: '45306' },
  { sqref: 'G1', type: 'textLength', operator: 'lessThanOrEqual', formula1: '5' },
  { sqref: 'H1', type: 'custom', formula1: 'ISNUMBER(H1)', hideDropDown: true },
  { sqref: 'I1', type: 'any', showInputMessage: true, promptTitle: 'Note', prompt: 'anything' },
];

function build(wb: Workbook): void {
  const ws = wb.addWorksheet('V');
  ws.setCell('A1', 'Yes');
  for (const dv of VALIDATIONS) addDataValidation(ws, dv);
  ws.setCell('J1', { text: 'link', hyperlink: 'https://example.com' });
  ws.pageSetup = { orientation: 'landscape' };
  wb.addWorksheet('Lists').setCell('A1', 'a');
}

async function allWriters(): Promise<[string, Uint8Array][]> {
  const wb = new Workbook();
  build(wb);
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  const s = w.addWorksheet('V');
  s.addRow(['Yes']);
  s.worksheet.dataValidations = VALIDATIONS;
  w.addWorksheet('Lists').addRow(['a']);
  await w.end();
  return [
    ['writeXlsx', writeXlsx(wb)],
    ['xlsxStream', new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())],
    ['createXlsxWriter', new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())],
  ];
}

describe('data validation', () => {
  it('round-trips through ours in every writer', async () => {
    for (const [name, buf] of await allWriters()) {
      expect(readXlsx(buf).getWorksheet('V')!.dataValidations, name).toEqual(READ_BACK);
      expect(readXlsx(buf).getWorksheet('Lists')!.dataValidations, name).toBeUndefined();
    }
  });

  it('writes dataValidations after conditionalFormatting slot and before hyperlinks', async () => {
    for (const [name, buf] of await allWriters()) {
      const xml = sheetXml(buf);
      const dv = xml.indexOf('<dataValidations count="8">');
      expect(dv, name).toBeGreaterThan(xml.indexOf('</sheetData>'));
      if (name !== 'createXlsxWriter') {
        expect(dv, name).toBeLessThan(xml.indexOf('<hyperlinks>'));
      }
      expect(dv, name).toBeLessThan(xml.indexOf('<pageMargins'));
      expect(xml, name).toContain(
        '<dataValidation type="list" errorStyle="warning" allowBlank="1" showErrorMessage="1" errorTitle="Oops" error="Pick &lt;one&gt; &amp; only one" sqref="A1:A3"><formula1>"Yes,No,Say ""maybe"""</formula1></dataValidation>',
      );
      expect(xml, name).toContain(
        '<dataValidation showInputMessage="1" promptTitle="Note" prompt="anything" sqref="I1"/>',
      );
    }
  });

  it('is read by exceljs with the same semantics', async () => {
    for (const [name, buf] of await allWriters()) {
      const m = xdv((await oracle(buf)).getWorksheet('V')!);
      expect(m.A2, name).toMatchObject({
        type: 'list',
        formulae: ['"Yes,No,Say ""maybe"""'],
        allowBlank: true,
        showErrorMessage: true,
        errorStyle: 'warning',
        errorTitle: 'Oops',
        error: 'Pick <one> & only one',
      });
      expect(m.B1, name).toMatchObject({
        type: 'list',
        formulae: ['Lists!$A$1:$A$5'],
        showInputMessage: true,
        prompt: 'p',
      });
      for (const a of ['C1', 'C2', 'E1']) {
        expect(m[a], `${name} ${a}`).toMatchObject({
          type: 'whole',
          operator: 'between',
          formulae: [1, 10],
          promptTitle: 'Range',
        });
      }
      expect(m.D1, name).toMatchObject({
        type: 'decimal',
        operator: 'greaterThan',
        formulae: [2.5],
      });
      expect(m.F1, name).toMatchObject({ type: 'date', operator: 'lessThan', formulae: [DATE] });
      expect(m.G1, name).toMatchObject({ type: 'textLength', formulae: [5] });
      expect(m.H1, name).toMatchObject({ type: 'custom', formulae: ['ISNUMBER(H1)'] });
      expect(m.I1, name).toMatchObject({ type: 'any', showInputMessage: true, prompt: 'anything' });
      expect(m.D2, name).toBeUndefined();
    }
  });

  it('reads validations written by exceljs', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('X');
    xs.getCell('A1').dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"One,Two"'],
      showErrorMessage: true,
      errorStyle: 'error',
      error: 'bad',
    };
    xs.getCell('A2').dataValidation = xs.getCell('A1').dataValidation;
    xs.getCell('B1').dataValidation = {
      type: 'decimal',
      operator: 'notBetween',
      formulae: [1.5, 7],
      showInputMessage: true,
      promptTitle: 't',
      prompt: 'p',
    };
    xs.getCell('C1').dataValidation = {
      type: 'date',
      operator: 'greaterThan',
      formulae: [DATE],
    };
    const ws = readXlsx(new Uint8Array(await x.xlsx.writeBuffer())).getWorksheet('X')!;
    const bySqref = new Map(ws.dataValidations!.map((d) => [d.sqref, d]));
    expect(bySqref.get('A1:A2')).toEqual({
      sqref: 'A1:A2',
      type: 'list',
      formula1: '"One,Two"',
      allowBlank: true,
      showErrorMessage: true,
      errorStyle: 'error',
      error: 'bad',
    } as unknown as DataValidation);
    expect(bySqref.get('B1')).toEqual({
      sqref: 'B1',
      type: 'decimal',
      operator: 'notBetween',
      formula1: '1.5',
      formula2: '7',
      showInputMessage: true,
      promptTitle: 't',
      prompt: 'p',
    });
    expect(bySqref.get('C1')).toMatchObject({ type: 'date', formula1: '45306' });
  });

  it('reads the x14 extLst form (cross-sheet list sources)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addDataValidation(ws, { sqref: 'A1', type: 'whole', operator: 'equal', formula1: 3 });
    const files = unzipSync(writeXlsx(wb));
    const xml = dec
      .decode(files.get('xl/worksheets/sheet1.xml')!)
      .replace(
        '</worksheet>',
        '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB8-1C0A8D0A6E5B}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="E"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>B1:B5 D1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>',
      );
    files.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(xml));
    const entries = [...files].map(([path, data]) => ({ path, data }));
    const back = readXlsx(zipSync(entries, 6)).getWorksheet('S')!;
    expect(back.dataValidations).toEqual([
      { sqref: 'A1', type: 'whole', operator: 'equal', formula1: '3' },
      {
        sqref: 'B1:B5 D1',
        type: 'list',
        formula1: 'Lists!$A$1:$A$3',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        errorTitle: 'E',
      },
    ]);
  });

  it('enforces Excel list limits with a clear error', () => {
    const ws = new Workbook().addWorksheet('S');
    const long = Array.from({ length: 60 }, (_, i) => `item${i}`);
    expect(() => addDataValidation(ws, { sqref: 'A1', type: 'list', values: long })).toThrow(
      /at most 255.*range formula/,
    );
    expect(() => addDataValidation(ws, { sqref: 'A1', type: 'list', values: ['a,b'] })).toThrow(
      /comma/,
    );
    expect(ws.dataValidations).toBeUndefined();
    // set directly: the writer still refuses
    ws.dataValidations = [{ sqref: 'A1', type: 'list', values: long }];
    expect(() => writeXlsx(ws.workbook)).toThrow(/at most 255/);
    // exactly 255 characters is fine
    const ok = ['x'.repeat(255)];
    ws.dataValidations = [{ sqref: 'A1', type: 'list', values: ok }];
    expect(readXlsx(writeXlsx(ws.workbook)).getWorksheet('S')!.dataValidations![0]!.formula1).toBe(
      `"${ok[0]}"`,
    );
  });

  it('time validations use the time-of-day fraction; 1904 dates shift the serial', () => {
    const wb = new Workbook({ date1904: true });
    const ws = wb.addWorksheet('S');
    addDataValidation(ws, {
      sqref: 'A1',
      type: 'time',
      operator: 'greaterThan',
      formula1: new Date(Date.UTC(2024, 0, 1, 18)),
    });
    addDataValidation(ws, { sqref: 'B1', type: 'date', formula1: DATE, formula2: DATE });
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!.dataValidations!;
    expect(back[0]!.formula1).toBe('0.75');
    expect(back[1]).toMatchObject({ formula1: '43844', formula2: '43844' });
  });
});
