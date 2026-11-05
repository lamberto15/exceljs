// readXlsb (#116) / readXls (#117) against Excel itself: every fixture in fixtures/binary was
// saved by Excel for Mac from one open workbook as .xlsx, .xlsb and .xls (and the .xls opened
// again in Excel and saved as .xlsx: excel-saved-from-xls.xlsx). The binary readers must give
// the model readXlsx gives for Excel's .xlsx: values, formulas, styles, merges, columns, rows,
// sheet states, defined names, hyperlinks.

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decryptXlsx, encryptXlsx } from '../src/crypto/index.js';
import { encodeAddr, readXlsx, type Workbook, writeXlsx } from '../src/index.js';
import { BiffParseError, CfbError, EncryptedXlsError, readXls } from '../src/xls/index.js';
import { readXlsb, ZipLimitError } from '../src/xlsb/index.js';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/binary/${name}`, import.meta.url)));

/** Everything the binary readers model, in a form `toEqual` can compare. */
function model(wb: Workbook): unknown {
  return {
    date1904: wb.date1904,
    names: [...wb.definedNames].sort((a, b) => a.name.localeCompare(b.name)),
    sheets: wb.worksheets.map((ws) => {
      const cells: Record<string, unknown> = {};
      const rows: Record<number, unknown> = {};
      ws.rowData.forEach((row, r) => {
        if (!row) return;
        if (row.height || row.hidden || row.xf)
          rows[r + 1] = {
            height: row.height,
            hidden: row.hidden,
            style: ws.getRowProps(r + 1).style,
          };
        for (let i = 0; i < row.n; i++) {
          const c = row.col[i]!;
          const v = ws.getValue(r + 1, c + 1);
          cells[encodeAddr(r, c)] = {
            value: v instanceof Date ? { date: v.toISOString() } : v,
            style: ws.getStyle(r + 1, c + 1),
          };
        }
      });
      const cols: Record<number, unknown> = {};
      ws.columnData.forEach((col, c) => {
        if (col) cols[c + 1] = { width: col.width, hidden: col.hidden, style: col.styleIndex };
      });
      return {
        name: ws.name,
        state: ws.state,
        merges: [...ws.merges].sort(),
        cols,
        rows,
        cells,
      };
    }),
  };
}

const SETS = ['excel-saved', 'excel-saved-formulas', 'excel-saved-shared', 'excel-saved-1904'];

describe('readXlsb matches readXlsx on the same workbook saved by Excel', () => {
  for (const set of SETS) {
    it(set, () => {
      const want = model(readXlsx(fixture(`${set}.xlsx`)));
      expect(model(readXlsb(fixture(`${set}.xlsb`)))).toEqual(want);
    });
  }
});

describe('readXls matches readXlsx on the same workbook saved by Excel', () => {
  // the .xls format has no theme colors in rich-text runs: Excel itself reads them back as
  // palette indices, so the reference is Excel's own .xls → .xlsx conversion
  it('excel-saved (vs. Excel converting the .xls to .xlsx)', () => {
    const want = model(readXlsx(fixture('excel-saved-from-xls.xlsx')));
    expect(model(readXls(fixture('excel-saved.xls')))).toEqual(want);
  });
  for (const set of SETS.slice(1)) {
    it(set, () => {
      const want = model(readXlsx(fixture(`${set}.xlsx`)));
      expect(model(readXls(fixture(`${set}.xls`)))).toEqual(want);
    });
  }
});

describe('details', () => {
  const xlsb = readXlsb(fixture('excel-saved.xlsb'));
  const xls = readXls(fixture('excel-saved.xls'));

  it('reads sheets in order with their visibility', () => {
    for (const wb of [xlsb, xls]) {
      expect(wb.worksheets.map((s) => [s.name, s.state])).toEqual([
        ['Data', 'visible'],
        ['Other', 'visible'],
        ['Hidden', 'hidden'],
      ]);
    }
  });

  it('reads values, dates (numbers with a date format) and formulas', () => {
    for (const wb of [xlsb, xls]) {
      const ws = wb.getWorksheet('Data')!;
      expect(ws.getValue('B1')).toBe(42); // RK integer
      expect(ws.getValue('F1')).toBe(0.01); // RK / 100
      expect(ws.getValue('D1')).toBe(-1234567.891);
      expect(ws.getValue('A2')).toBe(true);
      expect(ws.getValue('C2')).toEqual({ error: '#DIV/0!' });
      expect(ws.getValue('A3')).toEqual(new Date(Date.UTC(2024, 0, 15)));
      expect(ws.getCell('A3').style.numFmt).toBe('yyyy\\-mm\\-dd');
      expect(ws.getValue('A5')).toEqual({ formula: 'B1*2', result: 84 });
      expect(ws.getValue('F5')).toEqual({ formula: 'Other!A1', result: 10 });
      expect(ws.getValue('C8')).toEqual({ formula: 'B1 C1:D1', result: { error: '#NULL!' } });
      expect(ws.getValue('B8')).toEqual({ formula: 'IFERROR(1/0,"err")', result: 'err' });
      expect(ws.getValue('C9')).toEqual({ formula: 'SUM(B:B)', result: expect.any(Number) });
      expect(ws.getValue('F9')).toEqual({ formula: '{1,2;3,4}', result: 1 });
    }
  });

  it('reads long shared strings split across CONTINUE records (.xls)', () => {
    const ws = xls.getWorksheet('Data')!;
    expect(ws.getValue('B21')).toBe(`${'y'.repeat(5000)}日本`);
    expect(ws.getValue('B22')).toBe(`${'z'.repeat(4000)}é${'w'.repeat(4000)}`);
    expect(ws.getValue('A319')).toBe('string number 299');
    expect(xlsb.getWorksheet('Data')!.getValue('B21')).toBe(`${'y'.repeat(5000)}日本`);
  });

  it('reads rich text runs with their fonts', () => {
    expect(xlsb.getWorksheet('Data')!.getValue('A4')).toEqual({
      richText: [
        {
          text: 'Bold red',
          font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFF0000' }, family: 2 },
        },
        { text: ' and italic', font: { name: 'Calibri', size: 11, italic: true, family: 2 } },
      ],
    });
  });

  it('reads styles into the registry like readXlsx', () => {
    for (const wb of [xlsb, xls]) {
      const s = wb.getWorksheet('Data')!.getCell('A10').style;
      expect(s.font).toMatchObject({
        name: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        strike: true,
      });
      expect(s.fill).toEqual({
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00FF00' },
        bgColor: { indexed: 64 },
      });
      expect(s.alignment).toEqual({ horizontal: 'center', vertical: 'top', wrapText: true });
      expect(s.border?.top).toEqual({ style: 'thin', color: { argb: 'FFFF0000' } });
      expect(wb.getWorksheet('Data')!.getCell('G10').style.protection).toEqual({
        locked: false,
        hidden: true,
      });
      expect(wb.getWorksheet('Data')!.getCell('H11').style.fill).toEqual({
        type: 'pattern',
        pattern: 'solid',
        fgColor: { theme: 5, tint: 13106 / 32767 },
        bgColor: { indexed: 64 },
      });
    }
  });

  it('reads merges, column widths and row heights', () => {
    for (const wb of [xlsb, xls]) {
      const ws = wb.getWorksheet('Data')!;
      expect(ws.merges.sort()).toEqual(['A13:C14', 'E13:E15']);
      expect(ws.getColumn(1).width).toBe(20);
      expect(ws.getColumn(9).hidden).toBe(true);
      expect(ws.getRow(4).height).toBe(30);
      expect(ws.getRow(16).hidden).toBe(true);
    }
  });

  it('reads defined names (ranges, constants, sheet scope, built-ins)', () => {
    for (const wb of [xlsb, xls]) {
      expect(wb.definedNames).toContainEqual({ name: 'Rate', value: '0.2' });
      expect(wb.definedNames).toContainEqual({
        name: 'Local',
        value: 'Other!$A$1:$A$3',
        localSheetId: 1,
      });
    }
    for (const wb of [
      readXlsb(fixture('excel-saved-formulas.xlsb')),
      readXls(fixture('excel-saved-formulas.xls')),
    ]) {
      expect(wb.definedNames).toContainEqual({
        name: '_xlnm.Print_Area',
        value: 'Calc!$A$1:$E$10',
        localSheetId: 0,
      });
      expect(wb.definedNames).toContainEqual({ name: 'Spaced', value: "'My Sheet'!$A$1" });
    }
  });

  it('reads shared and array formulas', () => {
    for (const wb of [
      readXlsb(fixture('excel-saved-shared.xlsb')),
      readXls(fixture('excel-saved-shared.xls')),
    ]) {
      const ws = wb.worksheets[0]!;
      expect(ws.getValue('B6')).toEqual({ formula: 'A6*2+$A$1', result: 26 });
      expect(ws.getValue('D3')).toEqual({ formula: 'SUM($A$1:B3)-C3', result: expect.any(Number) });
      expect(ws.getValue('E1')).toEqual({ formula: 'INDEX({1,"a";TRUE,#N/A},2,1)', result: true });
    }
    for (const wb of [
      readXlsb(fixture('excel-saved-formulas.xlsb')),
      readXls(fixture('excel-saved-formulas.xls')),
    ]) {
      const ws = wb.worksheets[0]!;
      expect(ws.getValue('D1')).toEqual({ formula: 'A1:A2*10', result: 15, array: 'D1:D2' });
      expect(ws.getValue('D2')).toBe(30);
      expect(ws.getValue('F1')).toEqual({ text: 'link', hyperlink: 'https://example.com/' });
    }
  });

  it('reads the 1904 date system', () => {
    for (const wb of [
      readXlsb(fixture('excel-saved-1904.xlsb')),
      readXls(fixture('excel-saved-1904.xls')),
    ]) {
      expect(wb.date1904).toBe(true);
      expect(wb.worksheets[0]!.getValue('A1')).toEqual(new Date(Date.UTC(2024, 0, 15)));
    }
  });

  it('writes what it read as .xlsx', () => {
    for (const wb of [xlsb, xls]) {
      const back = readXlsx(writeXlsx(wb));
      expect(model(back)).toEqual(model(wb));
    }
  });

  it('honors the sheets option', () => {
    const wb = readXlsb(fixture('excel-saved.xlsb'), { sheets: ['other'] });
    expect(wb.worksheets.map((s) => s.name)).toEqual(['Other']);
    // names scoped to sheets that were not loaded are dropped
    expect(wb.definedNames.find((n) => n.name === 'Local')).toEqual({
      name: 'Local',
      value: 'Other!$A$1:$A$3',
      localSheetId: 0,
    });
    expect(
      readXls(fixture('excel-saved.xls'), { sheets: ['Hidden'] }).worksheets.map((s) => s.name),
    ).toEqual(['Hidden']);
  });
});

describe('errors', () => {
  it('rejects the wrong container with a clear message', () => {
    expect(() => readXls(fixture('excel-saved.xlsb'))).toThrow(/zip package/);
    expect(() => readXlsb(fixture('excel-saved.xls'))).toThrow(BiffParseError);
    expect(() => readXlsb(fixture('excel-saved.xlsx'))).toThrow(/no binary workbook part/);
    expect(() => readXls(new Uint8Array(600))).toThrow(/compound file/);
  });

  it('reads an encrypted .xlsb after decryptXlsx', async () => {
    const enc = await encryptXlsx(fixture('excel-saved-1904.xlsb'), 'pw', { spinCount: 10 });
    expect(() => readXlsb(enc)).toThrow(/decryptXlsx/);
    expect(() => readXls(enc)).toThrow(/decryptXlsx/);
    expect(readXlsb(await decryptXlsx(enc, 'pw')).date1904).toBe(true);
  });

  it('throws EncryptedXlsError for a FILEPASS record', () => {
    const bytes = fixture('excel-saved-1904.xls').slice();
    // turn the record after BOF in the Workbook stream into FILEPASS (0x002F)
    const i = findWorkbookStream(bytes);
    const bofLen = bytes[i + 2]! | (bytes[i + 3]! << 8);
    const next = i + 4 + bofLen;
    bytes[next] = 0x2f;
    bytes[next + 1] = 0;
    expect(() => readXls(bytes)).toThrow(EncryptedXlsError);
    expect(() => readXls(bytes)).toThrow(/password-protected/);
  });

  it('throws BiffParseError for truncated data, never a raw RangeError', () => {
    const xlsb = fixture('excel-saved-1904.xlsb');
    const xls = fixture('excel-saved.xls');
    for (const n of [0.3, 0.5, 0.8]) {
      expect(() => readXls(xls.slice(0, Math.floor(xls.length * n)))).toThrow(CfbError);
      expect(() => readXlsb(xlsb.slice(0, Math.floor(xlsb.length * n)))).toThrow(ZipLimitError);
    }
    // a corrupt record length inside a sheet
    const bad = fixture('excel-saved-1904.xls').slice();
    const i = findWorkbookStream(bad);
    bad[i + 3] = 0xff; // BOF length 0xFF10: past the end of the stream records
    let err: unknown;
    try {
      readXls(bad);
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(BiffParseError);
  });
});

/** Offset of the BOF record that starts the Workbook stream in a .xls fixture. */
function findWorkbookStream(b: Uint8Array): number {
  for (let i = 512; i + 8 < b.length; i += 512) {
    if (b[i] === 0x09 && b[i + 1] === 0x08 && b[i + 4] === 0 && b[i + 5] === 6 && b[i + 6] === 5)
      return i;
  }
  throw new Error('no BOF');
}
