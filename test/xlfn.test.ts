// Future-function prefixes (_xlfn. / _xlfn._xlws.): added on write, stripped on read.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { type Formula, Workbook } from '../src/core/index.js';
import { addXlfn, stripXlfn } from '../src/xlsx/common/xlfn.js';
import {
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addTable,
  getDefinedName,
} from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const part = (bytes: Uint8Array, path: string): string => dec.decode(unzipSync(bytes).get(path)!);

describe('addXlfn / stripXlfn', () => {
  it('prefixes future functions only', () => {
    expect(addXlfn('XLOOKUP(A1,B:B,C:C)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)');
    expect(addXlfn('SUM(A1:A3)+IFNA(1,2)')).toBe('SUM(A1:A3)+_xlfn.IFNA(1,2)');
    expect(addXlfn('SORT(FILTER(A1:A9,B1:B9>0))')).toBe(
      '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))',
    );
    expect(addXlfn('stdev.s(A:A)')).toBe('_xlfn.stdev.s(A:A)'); // case-insensitive
    expect(addXlfn('IF(XOR(TRUE,FALSE),CONCAT("a",TEXTJOIN(",",1,DAYS(B1,C1))),0)')).toBe(
      'IF(_xlfn.XOR(TRUE,FALSE),_xlfn.CONCAT("a",_xlfn.TEXTJOIN(",",1,_xlfn.DAYS(B1,C1))),0)',
    );
    expect(addXlfn('A1+B1')).toBe('A1+B1');
    expect(addXlfn('XLOOKUP')).toBe('XLOOKUP'); // a name, not a call
    expect(addXlfn('MYXOR(1)+A1XOR(1)+ERROR.TYPE(A1)')).toBe('MYXOR(1)+A1XOR(1)+ERROR.TYPE(A1)');
    expect(addXlfn('NETWORKDAYS.INTL(A1,B1)')).toBe('_xlfn.NETWORKDAYS.INTL(A1,B1)');
  });

  it('is idempotent and leaves strings, sheet names and structured refs alone', () => {
    const once = addXlfn('XLOOKUP(1,A:A,B:B)&SORT(C1:C3)');
    expect(addXlfn(once)).toBe(once);
    expect(addXlfn('_XLFN.XOR(1,0)')).toBe('_XLFN.XOR(1,0)');
    expect(addXlfn('"XOR(1)"&\'IFS(x\'!A1&"a""XOR(""b"')).toBe(
      '"XOR(1)"&\'IFS(x\'!A1&"a""XOR(""b"',
    );
    expect(addXlfn('Table1[XOR(]')).toBe('Table1[XOR(]');
  });

  it('strips canonical prefixes only', () => {
    expect(stripXlfn('_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe('XLOOKUP(A1,B:B,C:C)');
    expect(stripXlfn('_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))')).toBe(
      'SORT(FILTER(A1:A9,B1:B9>0))',
    );
    expect(stripXlfn('_XLFN.ifna(1,2)')).toBe('ifna(1,2)');
    // unknown or non-canonical prefixes are kept, so writing them back is exact
    expect(stripXlfn('_xlfn.FUTUREFN(1)')).toBe('_xlfn.FUTUREFN(1)');
    expect(stripXlfn('_xlfn.SORT(A1:A3)')).toBe('_xlfn.SORT(A1:A3)');
    expect(stripXlfn('_xludf.MYFN(1)')).toBe('_xludf.MYFN(1)');
    expect(stripXlfn('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)')).toBe('LET(_xlpm.x,1,_xlpm.x+1)');
    expect(stripXlfn('"_xlfn.XOR(1)"&A1')).toBe('"_xlfn.XOR(1)"&A1');
    for (const f of [
      '_xlfn.XLOOKUP(1,A:A,B:B)',
      '_xlfn._xlws.SORT(A1:A3)',
      '_xlfn.LET(_xlpm.x,1,_xlpm.x)',
    ]) {
      expect(addXlfn(stripXlfn(f))).toBe(f);
    }
  });
});

function build(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow([1, 2]);
  ws.setCell('C1', { formula: 'XLOOKUP(1,A1:A1,B1:B1)', result: 2 });
  ws.setCell('C2', { formula: 'SUM(SORT(A1:B1))', array: 'C2' });
  ws.setCell('C3', { formula: '"IFS("&TEXTJOIN(",",TRUE,A1:B1)' });
  addDefinedName(wb, 'Pick', 'IFNA(MATCH(1,S!$A$1:$A$9,0),0)');
  addDataValidation(ws, { sqref: 'D1', type: 'custom', formula1: '=XOR(A1>0,B1>0)' });
  addConditionalFormatting(ws, 'A1:B1', [
    {
      type: 'expression',
      formulae: ['MAXIFS($A$1:$B$1,$A$1:$B$1,">0")=A1'],
      style: { font: { bold: true } },
    },
  ]);
  addTable(wb.addWorksheet('T'), {
    name: 'Tbl',
    columns: [
      'a',
      { name: 'b', calculatedColumnFormula: 'CONCAT(Tbl[[#This Row],[a]],"x")' },
      { name: 'c', totalsRowFunction: 'custom', totalsRowFormula: 'STDEV.S(Tbl[a])' },
    ],
    rows: [[1, null, 3]],
    totalsRow: true,
  });
  return wb;
}

describe('writers add the prefix, readers strip it', () => {
  it('writeXlsx: cells, array formulas, defined names, validations, conditional formats, tables', () => {
    const bytes = writeXlsx(build());
    const sheet = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<f>_xlfn.XLOOKUP(1,A1:A1,B1:B1)</f>');
    expect(sheet).toContain('<f t="array" ref="C2">SUM(_xlfn._xlws.SORT(A1:B1))</f>');
    expect(sheet).toContain('<f>"IFS("&amp;_xlfn.TEXTJOIN(",",TRUE,A1:B1)</f>');
    expect(sheet).toContain('<formula1>_xlfn.XOR(A1&gt;0,B1&gt;0)</formula1>');
    expect(sheet).toContain('<formula>_xlfn.MAXIFS(');
    expect(part(bytes, 'xl/workbook.xml')).toContain(
      '>_xlfn.IFNA(MATCH(1,S!$A$1:$A$9,0),0)</definedName>',
    );
    const table = part(bytes, 'xl/tables/table1.xml');
    expect(table).toContain('<calculatedColumnFormula>_xlfn.CONCAT(');
    expect(table).toContain('<totalsRowFormula>_xlfn.STDEV.S(Tbl[a])</totalsRowFormula>');
    expect(part(bytes, 'xl/worksheets/sheet2.xml')).toContain('<f>_xlfn.CONCAT(');

    // the model read back is prefix-free
    const back = readXlsx(bytes);
    const ws = back.getWorksheet('S')!;
    expect(ws.getValue('C1')).toEqual({ formula: 'XLOOKUP(1,A1:A1,B1:B1)', result: 2 });
    expect((ws.getValue('C2') as Formula).formula).toBe('SUM(SORT(A1:B1))');
    expect(getDefinedName(back, 'Pick')?.value).toBe('IFNA(MATCH(1,S!$A$1:$A$9,0),0)');
    expect(ws.dataValidations?.[0]?.formula1).toBe('XOR(A1>0,B1>0)');
    expect(ws.conditionalFormatting?.[0]?.rules[0]).toMatchObject({
      formulae: ['MAXIFS($A$1:$B$1,$A$1:$B$1,">0")=A1'],
    });
    const cols = back.getWorksheet('T')!.tables![0]!.columns;
    expect(cols[1]).toMatchObject({ calculatedColumnFormula: 'CONCAT(Tbl[[#This Row],[a]],"x")' });
    expect(cols[2]).toMatchObject({ totalsRowFormula: 'STDEV.S(Tbl[a])' });

    // round trip ours → ours is byte-exact for the formula parts
    const again = writeXlsx(back);
    for (const p of ['xl/worksheets/sheet1.xml', 'xl/tables/table1.xml']) {
      expect(part(again, p)).toBe(part(bytes, p));
    }
    expect(part(again, 'xl/workbook.xml')).toContain('>_xlfn.IFNA(');
  });

  it('streaming writers add it too', async () => {
    const s = part(
      new Uint8Array(await new Response(xlsxStream(build())).arrayBuffer()),
      'xl/worksheets/sheet1.xml',
    );
    expect(s).toContain('<f>_xlfn.XLOOKUP(');
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    w.addWorksheet('S').addRow([{ formula: 'IFS(TRUE,1)' }]);
    await w.end();
    const bytes = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<f>_xlfn.IFS(TRUE,1)</f>');
  });

  it('readXlsx strips prefixes in files written by Excel (incl. shared formulas)', () => {
    const files = unzipSync(writeXlsx(build()));
    let xml = dec.decode(files.get('xl/worksheets/sheet1.xml')!);
    xml = xml.replace(
      /<c r="C1"[\s\S]*?<\/c>/,
      '<c r="C1"><f t="shared" ref="C1:C2" si="0">_xlfn._xlws.SORT(A1:B1)+_xlfn.UNKNOWNFN(1)</f><v>1</v></c>',
    );
    xml = xml.replace(/<c r="C2"[\s\S]*?<\/c>/, '<c r="C2"><f t="shared" si="0"/><v>1</v></c>');
    files.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(xml));
    const back = readXlsx(
      zipSync(
        [...files].map(([path, data]) => ({ path, data })),
        6,
      ),
    );
    const ws = back.getWorksheet('S')!;
    expect((ws.getValue('C1') as Formula).formula).toBe('SORT(A1:B1)+_xlfn.UNKNOWNFN(1)');
    expect((ws.getValue('C2') as Formula).formula).toBe('SORT(A2:B2)+_xlfn.UNKNOWNFN(1)');
    // SORT spills: written as a dynamic-array formula (cm + t="array"), prefix kept
    expect(part(writeXlsx(back), 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="C1" cm="1"><f t="array" ref="C1">_xlfn._xlws.SORT(A1:B1)+_xlfn.UNKNOWNFN(1)</f>',
    );
  });

  it('exceljs loads the files', async () => {
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(writeXlsx(build()) as never);
    const c = x.getWorksheet('S')!.getCell('C1').value as { formula: string; result: number };
    expect(c.formula).toBe('_xlfn.XLOOKUP(1,A1:A1,B1:B1)');
    expect(c.result).toBe(2);
  });
});

describe('partial prefixes written by the caller', () => {
  it('addXlfn completes them and keeps unknown names as written', () => {
    expect(addXlfn('_xlfn.SORT(A1:A3)')).toBe('_xlfn._xlws.SORT(A1:A3)');
    expect(addXlfn('_xlws.FILTER(A1:A3,B1:B3)')).toBe('_xlfn._xlws.FILTER(A1:A3,B1:B3)');
    expect(addXlfn('_XLFN.sort(A1:A3)')).toBe('_xlfn._xlws.sort(A1:A3)');
    expect(addXlfn('_xlfn._xlws.SORT(A1:A3)')).toBe('_xlfn._xlws.SORT(A1:A3)');
    expect(addXlfn('_xlfn.XLOOKUP(1,A:A,B:B)')).toBe('_xlfn.XLOOKUP(1,A:A,B:B)');
    expect(addXlfn('_xlfn.MYFUNC(1)')).toBe('_xlfn.MYFUNC(1)');
    expect(addXlfn('"_xlfn.SORT(" & _xlfn.SORT(A1:A3)')).toBe(
      '"_xlfn.SORT(" & _xlfn._xlws.SORT(A1:A3)',
    );
  });

  it('_xlfn.SORT / _xlws.SORT are written as dynamic-array formulas with the full prefix', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('s');
    ws.setCell('A1', { formula: '_xlfn.SORT(C1:C3)' });
    ws.setCell('B1', { formula: '_xlws.SORT(C1:C3)' });
    const xml = new TextDecoder().decode(unzipSync(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!);
    for (const ref of ['A1', 'B1']) {
      expect(xml).toMatch(
        new RegExp(
          `<c r="${ref}"[^>]*cm="1"><f t="array" ref="${ref}">_xlfn\\._xlws\\.SORT\\(C1:C3\\)</f>`,
        ),
      );
    }
  });
});

describe('structured reference @ shorthand on write', () => {
  it.each([
    ['Sales[@Units]*Sales[@Price]', 'Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]'],
    ['Sales[@[Unit Price]]', 'Sales[[#This Row],[Unit Price]]'],
    ['[@Units]+1', '[[#This Row],[Units]]+1'],
    ['SUM(T[@[A]:[B]])', 'SUM(T[[#This Row],[A]:[B]])'],
    ['T[@]', 'T[#This Row]'],
    ["T[@[a']b]]", "T[[#This Row],[a']b]]"],
    ['"x[@y]"&T[@y]', '"x[@y]"&T[[#This Row],[y]]'],
    ["'Sheet [@1]'!A1", "'Sheet [@1]'!A1"],
    ['@A1:A3', '@A1:A3'],
  ])('%s → %s', (f, want) => {
    expect(addXlfn(f)).toBe(want);
  });
});

describe('LET names get _xlpm. on write', () => {
  it.each([
    ['LET(x,2,y,x*3,x+y)', '_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x*3,_xlpm.x+_xlpm.y)'],
    ['_xlfn.LET(_xlpm.x,1,_xlpm.x+1)', '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)'],
    [
      'LET(q, Sales[Qty], SUM(q)/ROWS(q))',
      '_xlfn.LET(_xlpm.q, Sales[Qty], SUM(_xlpm.q)/ROWS(_xlpm.q))',
    ],
    ['LET(x,"x",x&Sheet1!x&x(1))', '_xlfn.LET(_xlpm.x,"x",_xlpm.x&Sheet1!x&x(1))'],
    ['LET(a,1,LET(b,a,a+b))', '_xlfn.LET(_xlpm.a,1,_xlfn.LET(_xlpm.b,_xlpm.a,_xlpm.a+_xlpm.b))'],
    ['x+LET(x,1,x)', 'x+_xlfn.LET(_xlpm.x,1,_xlpm.x)'],
  ])('%s → %s', (f, want) => {
    expect(addXlfn(f)).toBe(want);
  });
});
