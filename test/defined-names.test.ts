/// <reference types="node" />
// Defined-name helpers (print area / titles, named ranges) against exceljs.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { isValidName } from '../src/xlsx/features/defined-names.js';
import {
  addDefinedName,
  getDefinedName,
  getPrintArea,
  getPrintTitles,
  removeDefinedName,
  setPrintArea,
  setPrintTitles,
} from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

function build(): Workbook {
  const wb = new Workbook();
  const a = wb.addWorksheet('My Sheet');
  a.addRow([1, 2]);
  const b = wb.addWorksheet("Bob's");
  b.addRow([3]);
  setPrintArea(a, 'A1:B2');
  setPrintTitles(a, { rows: '1:2', cols: 'A:B' });
  setPrintArea(b, ['A1:B2', 'D4:E5']);
  addDefinedName(wb, 'Rate', '0.2', { comment: 'tax rate' });
  addDefinedName(wb, 'Total', "Bob's!A1");
  addDefinedName(wb, 'Local', 'B2:C3', { localSheet: a, hidden: true });
  return wb;
}

describe('defined-name helpers', () => {
  it('quote sheet names and make ranges absolute', () => {
    const wb = build();
    expect(getDefinedName(wb, 'Total')?.value).toBe("'Bob''s'!$A$1");
    expect(getDefinedName(wb, 'Local', 'My Sheet')).toMatchObject({
      value: "'My Sheet'!$B$2:$C$3",
      hidden: true,
      localSheetId: 0,
    });
    expect(getDefinedName(wb, 'Rate')).toMatchObject({ value: '0.2', comment: 'tax rate' });
    expect(wb.definedNames.find((d) => d.name === '_xlnm.Print_Titles')?.value).toBe(
      "'My Sheet'!$A:$B,'My Sheet'!$1:$2",
    );
    expect(getPrintArea(wb.worksheets[1]!)).toEqual(['A1:B2', 'D4:E5']);
    expect(getPrintTitles(wb.worksheets[0]!)).toEqual({ rows: '1:2', cols: 'A:B' });
  });

  it('replace, remove and clear', () => {
    const wb = build();
    addDefinedName(wb, 'rate', '0.3');
    expect(wb.definedNames.filter((d) => d.name.toLowerCase() === 'rate')).toHaveLength(1);
    expect(getDefinedName(wb, 'RATE')?.value).toBe('0.3');
    expect(removeDefinedName(wb, 'Rate')).toBe(true);
    expect(removeDefinedName(wb, 'Rate')).toBe(false);
    const ws = wb.worksheets[0]!;
    setPrintArea(ws, undefined);
    setPrintTitles(ws, {});
    expect(getPrintArea(ws)).toBeUndefined();
    expect(getPrintTitles(ws)).toBeUndefined();
    setPrintTitles(ws, { rows: '3' });
    expect(getPrintTitles(ws)).toEqual({ rows: '3:3' });
  });

  it('validate names and references', () => {
    for (const ok of [
      'Rate',
      '_x',
      '\\a',
      'a.b',
      'Tax_2020',
      'ABCD1',
      'XFE1',
      'Rx',
      'RC1x',
      'ñandú',
    ]) {
      expect(isValidName(ok), ok).toBe(true);
    }
    for (const bad of ['', 'A1', 'xfd1048576', 'R1C1', 'R', 'c', 'RC', 'R1', '1a', 'a b', 'a-b']) {
      expect(isValidName(bad), bad).toBe(false);
    }
    expect(isValidName('x'.repeat(256))).toBe(false);
    const wb = build();
    expect(() => addDefinedName(wb, 'A1', '0')).toThrow(/Invalid defined name/);
    expect(() => addDefinedName(wb, 'a b', '0')).toThrow(/Invalid/);
    expect(() => addDefinedName(wb, 'X', 'A1:B2')).toThrow(/no sheet/);
    expect(() => addDefinedName(wb, 'X', '0', { localSheet: 'nope' })).toThrow(/not a sheet/);
    expect(() => setPrintArea(wb.worksheets[0]!, 'Other!A1:B2')).toThrow(/Invalid print area/);
    expect(() => setPrintTitles(wb.worksheets[0]!, { rows: 'A' })).toThrow(/Invalid rows/);
    expect(() => setPrintTitles(wb.worksheets[0]!, { cols: '1' })).toThrow(/Invalid columns/);
  });

  it('round-trips through writeXlsx -> readXlsx', () => {
    const buf = writeXlsx(build());
    const back = readXlsx(buf);
    const a = back.worksheets[0]!;
    expect(getPrintArea(a)).toEqual(['A1:B2']);
    expect(getPrintTitles(a)).toEqual({ rows: '1:2', cols: 'A:B' });
    expect(getPrintArea(back.worksheets[1]!)).toEqual(['A1:B2', 'D4:E5']);
    expect(getDefinedName(back, 'Rate')).toMatchObject({ value: '0.2', comment: 'tax rate' });
    expect(getDefinedName(back, 'Local', a)?.hidden).toBe(true);
    expect(dec.decode(unzipSync(buf).get('xl/workbook.xml')!)).toContain('comment="tax rate"');
  });

  it('are written by xlsxStream', async () => {
    const buf = new Uint8Array(await new Response(xlsxStream(build())).arrayBuffer());
    const back = readXlsx(buf);
    expect(getPrintArea(back.worksheets[0]!)).toEqual(['A1:B2']);
    expect(getDefinedName(back, 'Total')?.value).toBe("'Bob''s'!$A$1");
  });

  it('are written by createXlsxWriter', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    s.addRow([1]);
    w.addWorksheet('T').addRow([1]);
    setPrintArea(s.worksheet, 'A1:A1');
    addDefinedName(w.workbook, 'N', 'T!A1');
    await w.end();
    const back = readXlsx(new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer()));
    expect(getPrintArea(back.worksheets[0]!)).toEqual(['A1:A1']);
    expect(getDefinedName(back, 'N')?.value).toBe("'T'!$A$1");
  });

  it('are understood by exceljs', async () => {
    const x = await oracle(writeXlsx(build()));
    const ws = x.getWorksheet('My Sheet')!;
    expect(ws.pageSetup.printArea).toBe('A1:B2');
    expect(ws.pageSetup.printTitlesRow).toBe('1:2');
    expect(ws.pageSetup.printTitlesColumn).toBe('A:B');
    // exceljs only reads the first range of a multi-range Print_Area
    expect(x.getWorksheet("Bob's")!.pageSetup.printArea).toBe('A1:B2');
    expect(x.definedNames.getRanges('Total').ranges).toEqual(["'Bob''s'!$A$1"]);
  });

  it('read exceljs files', async () => {
    const x = new ExcelJS.Workbook();
    const ws = x.addWorksheet('Data Sheet');
    ws.addRow([1, 2, 3]);
    ws.pageSetup.printArea = 'A1:C5';
    ws.pageSetup.printTitlesRow = '1:1';
    ws.pageSetup.printTitlesColumn = 'A:A';
    x.definedNames.add("'Data Sheet'!$A$1:$B$2", 'Mine');
    const back = readXlsx(new Uint8Array(await x.xlsx.writeBuffer()));
    const w = back.worksheets[0]!;
    expect(getPrintArea(w)).toEqual(['A1:C5']);
    expect(getPrintTitles(w)).toEqual({ rows: '1:1', cols: 'A:A' });
    expect(getDefinedName(back, 'Mine')?.value).toBe("'Data Sheet'!$A$1:$B$2");
  });
});
