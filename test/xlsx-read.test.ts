import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import type { CellValue } from '../src/core/types.js';
import { Workbook } from '../src/core/workbook.js';
import { CellKind } from '../src/core/worksheet.js';
import {
  readXlsx,
  translateFormula,
  XlsxParseError,
  type XlsxRow,
  xlsxRows,
} from '../src/xlsx/read/index.js';
import { decodeXString } from '../src/xlsx/read/xml-util.js';
import { writeXlsx } from '../src/xlsx/write/index.js';
import { unzipSync, ZipLimitError, ZipPathError, zipSync } from '../src/zip/index.js';

// ---------------------------------------------------------------- helpers

async function exceljsBuffer(wb: ExcelJS.Workbook): Promise<Uint8Array> {
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

async function collectRows(
  src: Uint8Array | AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
  opts?: Parameters<typeof xlsxRows>[1],
): Promise<XlsxRow[]> {
  const out: XlsxRow[] = [];
  for await (const r of xlsxRows(src, opts)) out.push(r);
  return out;
}

async function* chunked(data: Uint8Array, size: number): AsyncGenerator<Uint8Array> {
  for (let i = 0; i < data.length; i += size) yield data.slice(i, i + size);
}

const NS =
  'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

interface MiniOpts {
  sheets?: { name: string; xml: string; path?: string; target?: string; rels?: string }[];
  sst?: string;
  styles?: string;
  workbook?: string;
  wbPath?: string;
  order?: 'meta-first' | 'sheets-first';
}

/** Build a minimal XLSX package by hand. */
function mini(o: MiniOpts): Uint8Array {
  const wbPath = o.wbPath ?? 'xl/workbook.xml';
  const wbDir = wbPath.slice(0, wbPath.lastIndexOf('/') + 1);
  const sheets = o.sheets ?? [];
  const meta: { path: string; data: string }[] = [];
  const parts: { path: string; data: string }[] = [];
  meta.push({
    path: '[Content_Types].xml',
    data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`,
  });
  meta.push({
    path: '_rels/.rels',
    data: `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="${wbPath}"/></Relationships>`,
  });
  let rels = '';
  let sheetEls = '';
  sheets.forEach((s, i) => {
    const path = s.path ?? `${wbDir}worksheets/sheet${i + 1}.xml`;
    const target = s.target ?? path.slice(wbDir.length);
    rels += `<Relationship Id="rIdS${i}" Type="${REL}/worksheet" Target="${target}"/>`;
    sheetEls += `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rIdS${i}"/>`;
    parts.push({ path, data: s.xml });
    if (s.rels) {
      const dir = path.slice(0, path.lastIndexOf('/') + 1);
      parts.push({ path: `${dir}_rels/${path.slice(dir.length)}.rels`, data: s.rels });
    }
  });
  if (o.styles !== undefined) {
    rels += `<Relationship Id="rIdSt" Type="${REL}/styles" Target="styles.xml"/>`;
    meta.push({ path: `${wbDir}styles.xml`, data: o.styles });
  }
  if (o.sst !== undefined) {
    rels += `<Relationship Id="rIdSs" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`;
    meta.push({ path: `${wbDir}sharedStrings.xml`, data: o.sst });
  }
  meta.push({
    path: `${wbDir}_rels/${wbPath.slice(wbDir.length)}.rels`,
    data: `<Relationships xmlns="${PKG_REL}">${rels}</Relationships>`,
  });
  meta.push({
    path: wbPath,
    data: o.workbook ?? `<workbook ${NS}><sheets>${sheetEls}</sheets></workbook>`,
  });
  const all = o.order === 'sheets-first' ? [...parts, ...meta.reverse()] : [...meta, ...parts];
  return zipSync(all, 6);
}

function sheetXml(data: string, extra = ''): string {
  return `<?xml version="1.0" encoding="UTF-8"?><worksheet ${NS}><sheetData>${data}</sheetData>${extra}</worksheet>`;
}

function values(wb: Workbook, sheet = 0): Record<string, CellValue> {
  const ws = wb.worksheets[sheet]!;
  const out: Record<string, CellValue> = {};
  ws.eachRow((row) => {
    row.eachCell((v, c) => {
      if (v !== null && v !== undefined) out[`${String.fromCharCode(64 + c)}${row.number}`] = v;
    });
  });
  return out;
}

// ---------------------------------------------------------------- exceljs round-trip

async function buildRichWorkbook(date1904: boolean): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.properties.date1904 = date1904;
  wb.creator = 'tester';
  const ws = wb.addWorksheet('Data', { views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }] });
  ws.getCell('A1').value = 42;
  ws.getCell('B1').value = -3.25;
  ws.getCell('C1').value = 'hello';
  ws.getCell('D1').value = true;
  ws.getCell('E1').value = false;
  ws.getCell('F1').value = new Date(Date.UTC(2024, 1, 29, 13, 45, 30));
  ws.getCell('G1').value = { error: '#N/A' } as ExcelJS.CellErrorValue;
  ws.getCell('H1').value = 'line1\r\nline2 & <tag> "q"';
  ws.getCell('A2').value = { formula: 'A1*2', result: 84 };
  ws.getCell('B2').value = { formula: 'C1&"!"', result: 'hello!' };
  ws.getCell('C2').value = { formula: 'D1', result: true };
  ws.getCell('D2').value = { formula: '1/0', result: { error: '#DIV/0!' } };
  ws.getCell('E2').value = { formula: 'NOW()' } as ExcelJS.CellFormulaValue;
  // shared formulas
  ws.getCell('A3').value = 1;
  ws.getCell('A4').value = 2;
  ws.getCell('A5').value = 3;
  ws.getCell('B3').value = { formula: 'A3*10+$A$3', result: 11 };
  ws.getCell('B4').value = { sharedFormula: 'B3', result: 21 } as ExcelJS.CellSharedFormulaValue;
  ws.getCell('B5').value = { sharedFormula: 'B3', result: 31 } as ExcelJS.CellSharedFormulaValue;
  // array formula
  ws.getCell('C3').value = {
    formula: 'A3:A5*2',
    result: 2,
    shareType: 'array',
    ref: 'C3:C5',
  } as unknown as ExcelJS.CellFormulaValue;
  ws.getCell('D3').value = {
    richText: [
      { text: 'bold', font: { bold: true, size: 12, name: 'Arial', color: { argb: 'FFFF0000' } } },
      { text: ' plain' },
      { text: ' it', font: { italic: true, underline: true } },
    ],
  };
  ws.getCell('E3').value = { text: 'site', hyperlink: 'https://example.com/a?b=1&c=2' };
  ws.getCell('F3').value = { text: 'internal', hyperlink: '#Other!B2' };
  ws.getCell('G3').value = 'merged';
  ws.mergeCells('G3:H4');
  // styles
  const s = ws.getCell('A6');
  s.value = 1234.5678;
  s.numFmt = '#,##0.00';
  s.font = {
    name: 'Courier New',
    size: 9,
    bold: true,
    italic: true,
    color: { theme: 4, tint: 0.4 } as Partial<ExcelJS.Color>,
  };
  s.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF00FF00' },
    bgColor: { indexed: 64 } as Partial<ExcelJS.Color>,
  };
  s.border = {
    top: { style: 'thin', color: { argb: 'FF0000FF' } },
    bottom: { style: 'double' },
    left: { style: 'dashed' },
    right: { style: 'thick', color: { theme: 2 } },
  };
  s.alignment = { horizontal: 'center', vertical: 'top', wrapText: true, indent: 1 };
  const p = ws.getCell('B6');
  p.value = 0.5;
  p.numFmt = '0.00%';
  p.protection = { locked: false };
  ws.getCell('C6').value = new Date(Date.UTC(1999, 11, 31));
  ws.getCell('C6').numFmt = 'yyyy-mm-dd';
  ws.getCell('D6').value = 45000.5;
  ws.getCell('D6').numFmt = 'dd/mm/yyyy hh:mm';
  // rows / cols
  ws.getRow(7).height = 33;
  ws.getRow(7).getCell(1).value = 'tall';
  ws.getRow(8).hidden = true;
  ws.getRow(8).getCell(1).value = 'hidden row';
  ws.getRow(9).outlineLevel = 2;
  ws.getRow(9).getCell(1).value = 'outlined';
  ws.getColumn(1).width = 25;
  ws.getColumn(3).hidden = true;
  ws.getColumn(4).outlineLevel = 1;
  ws.getColumn(4).width = 12.5;
  ws.autoFilter = 'A1:H1';

  const other = wb.addWorksheet('Other', { state: 'hidden' });
  other.getCell('B2').value = 'target';
  wb.addWorksheet('VeryHidden', { state: 'veryHidden' }).getCell('A1').value = 1;
  wb.addWorksheet('Empty');
  return exceljsBuffer(wb);
}

describe('readXlsx: exceljs round-trip', () => {
  for (const date1904 of [false, true]) {
    it(`reads values, formulas, styles and layout (date1904=${date1904})`, async () => {
      const buf = await buildRichWorkbook(date1904);
      const wb = readXlsx(buf);
      expect(wb.date1904).toBe(date1904);
      expect(wb.creator).toBe('tester');
      expect(wb.worksheets.map((w) => [w.name, w.state])).toEqual([
        ['Data', 'visible'],
        ['Other', 'hidden'],
        ['VeryHidden', 'veryHidden'],
        ['Empty', 'visible'],
      ]);
      const ws = wb.worksheets[0]!;
      const g = (a: string) => ws.getValue(a);
      expect(g('A1')).toBe(42);
      expect(g('B1')).toBe(-3.25);
      expect(g('C1')).toBe('hello');
      expect(g('D1')).toBe(true);
      expect(g('E1')).toBe(false);
      expect(g('F1')).toEqual(new Date(Date.UTC(2024, 1, 29, 13, 45, 30)));
      expect(ws.getKind('F1')).toBe(CellKind.Date);
      expect(g('G1')).toEqual({ error: '#N/A' });
      // exceljs writes a raw CR LF; XML end-of-line normalization turns it into LF (Excel itself
      // would have written _x000D_, which we decode)
      expect(g('H1')).toBe('line1\nline2 & <tag> "q"');
      expect(g('A2')).toEqual({ formula: 'A1*2', result: 84 });
      expect(g('B2')).toEqual({ formula: 'C1&"!"', result: 'hello!' });
      expect(g('C2')).toEqual({ formula: 'D1', result: true });
      expect(g('D2')).toEqual({ formula: '1/0', result: { error: '#DIV/0!' } });
      expect(g('E2')).toEqual({ formula: 'NOW()' });
      expect(g('B3')).toEqual({ formula: 'A3*10+$A$3', result: 11 });
      expect(g('B4')).toEqual({ formula: 'A4*10+$A$3', result: 21 });
      expect(g('B5')).toEqual({ formula: 'A5*10+$A$3', result: 31 });
      expect(g('C3')).toEqual({ formula: 'A3:A5*2', result: 2, array: 'C3:C5' });
      expect(g('D3')).toEqual({
        richText: [
          {
            text: 'bold',
            font: { bold: true, size: 12, name: 'Arial', color: { argb: 'FFFF0000' } },
          },
          { text: ' plain' },
          { text: ' it', font: { italic: true, underline: 'single' } },
        ],
      });
      expect(g('E3')).toEqual({ text: 'site', hyperlink: 'https://example.com/a?b=1&c=2' });
      // exceljs writes internal links as an external relationship with target "#Other!B2"
      expect(g('F3')).toEqual({ text: 'internal', hyperlink: '#Other!B2' });
      expect(ws.merges).toEqual(['G3:H4']);
      expect(ws.views[0]).toMatchObject({
        state: 'frozen',
        xSplit: 1,
        ySplit: 2,
        topLeftCell: 'B3',
      });
      expect(ws.autoFilter).toBe('A1:H1');

      const st = ws.getStyle('A6');
      expect(st.numFmt).toBe('#,##0.00');
      expect(st.font).toEqual({
        name: 'Courier New',
        size: 9,
        bold: true,
        italic: true,
        color: { theme: 4, tint: 0.4 },
      });
      expect(st.fill).toEqual({
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00FF00' },
        bgColor: { indexed: 64 },
      });
      expect(st.border).toEqual({
        top: { style: 'thin', color: { argb: 'FF0000FF' } },
        bottom: { style: 'double' },
        left: { style: 'dashed' },
        right: { style: 'thick', color: { theme: 2 } },
      });
      expect(st.alignment).toEqual({
        horizontal: 'center',
        vertical: 'top',
        wrapText: true,
        indent: 1,
      });
      expect(ws.getStyle('B6')).toMatchObject({ numFmt: '0.00%', protection: { locked: false } });
      expect(g('C6')).toEqual(new Date(Date.UTC(1999, 11, 31)));
      expect(ws.getStyle('C6').numFmt).toBe('yyyy-mm-dd');
      expect(ws.getKind('D6')).toBe(CellKind.Date);

      expect(ws.getRowProps(7)).toMatchObject({ height: 33 });
      expect(ws.getRowProps(8)).toMatchObject({ hidden: true });
      expect(ws.getRowProps(9)).toMatchObject({ outlineLevel: 2 });
      expect(ws.getColumn(1).width).toBe(25);
      expect(ws.getColumn(3).hidden).toBe(true);
      expect(ws.getColumn(4).outlineLevel).toBe(1);
      expect(ws.getColumn(4).width).toBe(12.5);

      expect(wb.worksheets[1]!.getValue('B2')).toBe('target');
      expect(wb.worksheets[3]!.rowCount).toBe(0);
    });
  }

  it('xlsxRows yields the same values (0-based sparse arrays)', async () => {
    const buf = await buildRichWorkbook(false);
    const wb = readXlsx(buf);
    const rows = await collectRows(chunked(buf, 1000));
    expect(new Set(rows.map((r) => r.sheet))).toEqual(new Set(['Data', 'Other', 'VeryHidden']));
    for (const r of rows) {
      const ws = wb.getWorksheet(r.sheet)!;
      expect(r.sheetIndex).toBe(wb.worksheets.indexOf(ws));
      r.values.forEach((v, c) => {
        const dom = ws.getValue(r.row, c + 1);
        // hyperlinks are not applied in streaming mode
        if (dom && typeof dom === 'object' && 'hyperlink' in dom) expect(v).toBe(dom.text);
        else expect(v).toEqual(dom);
      });
    }
    const data = rows.filter((r) => r.sheet === 'Data');
    expect(data[0]!.row).toBe(1);
    expect(data[0]!.values[5]).toEqual(new Date(Date.UTC(2024, 1, 29, 13, 45, 30)));
    expect(data[0]!.values.length).toBe(8);
  });

  it('opts.sheets filters sheets in both modes', async () => {
    const buf = await buildRichWorkbook(false);
    const wb = readXlsx(buf, { sheets: ['other'] });
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Other']);
    const rows = await collectRows(buf, { sheets: ['Other'] });
    expect(rows).toEqual([
      { sheet: 'Other', sheetIndex: 1, row: 2, values: [undefined, 'target'] },
    ]);
  });

  it('accepts ArrayBuffer and ReadableStream sources', async () => {
    const buf = await buildRichWorkbook(false);
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
    expect(readXlsx(ab).worksheets.length).toBe(4);
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        for (let i = 0; i < buf.length; i += 777) c.enqueue(buf.slice(i, i + 777));
        c.close();
      },
    });
    const rows = await collectRows(stream);
    expect(rows.length).toBeGreaterThan(5);
  });

  it('round-trips many rows through both modes', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Big');
    for (let r = 1; r <= 3000; r++) ws.addRow([r, `s${r % 50}`, r % 2 === 0, r * 0.5]);
    const buf = await exceljsBuffer(wb);
    const ours = readXlsx(buf).worksheets[0]!;
    expect(ours.rowCount).toBe(3000);
    expect(ours.getValue(2999, 2)).toBe('s49');
    const rows = await collectRows(chunked(buf, 4096));
    expect(rows.length).toBe(3000);
    expect(rows[2999]).toEqual({
      sheet: 'Big',
      sheetIndex: 0,
      row: 3000,
      values: [3000, 's0', true, 1500],
    });
  });
});

// ---------------------------------------------------------------- real-world fixtures

const SPEC = fileURLToPath(new URL('./fixtures/exceljs/', import.meta.url));

function findXlsx(dir: string): string[] {
  if (!existsSync(dir)) return [];
  const out: string[] = [];
  for (const f of readdirSync(dir)) {
    const p = join(dir, f);
    if (statSync(p).isDirectory()) out.push(...findXlsx(p));
    else if (f.endsWith('.xlsx')) out.push(p);
  }
  return out.sort();
}

/** Fixtures exceljs itself rejects (its own spec expects the error). */
const EXPECTED_ERRORS: Record<string, RegExp> = {
  'invalid-xml.xlsx': /text data outside of root element/,
};

/**
 * Fixtures exceljs itself cannot load in reasonable time, so the exceljs comparison is skipped.
 * test-issue-1842: a whole-column data validation that exceljs expands cell by cell (its bug
 * #1842) — it throws after seconds on fast machines and runs for many minutes on CI runners.
 */
const EXCELJS_ORACLE_SKIP = new Set(['test-issue-1842.xlsx']);

function exceljsValue(v: ExcelJS.CellValue): unknown {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.getTime();
  if (typeof v !== 'object') return v;
  if ('richText' in v) return v.richText.map((r) => r.text).join('');
  if ('hyperlink' in v) {
    const t = v.text as unknown;
    if (typeof t === 'string') return t;
    if (t && typeof t === 'object' && 'richText' in t) {
      return (t as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join('');
    }
    return ''; // exceljs: hyperlink on an empty cell has no text
  }
  if ('error' in v) return v.error;
  if ('formula' in v || 'sharedFormula' in v) {
    const r = (v as ExcelJS.CellFormulaValue).result;
    return r instanceof Date
      ? r.getTime()
      : r && typeof r === 'object'
        ? (r as { error: string }).error
        : r;
  }
  return '<other>';
}

function oursValue(v: CellValue): unknown {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.getTime();
  if (typeof v !== 'object') return v;
  if ('richText' in v) return v.richText.map((r) => r.text).join('');
  if ('hyperlink' in v) return v.text;
  if ('error' in v) return v.error;
  if ('formula' in v) {
    const r = v.result;
    return r && typeof r === 'object' ? r.error : r;
  }
  return '<other>';
}

describe('real-world fixtures', () => {
  const files = findXlsx(SPEC);
  it('finds fixtures', () => {
    expect(files.length).toBeGreaterThan(20);
  });
  for (const file of files) {
    const base = basename(file);
    const rel = relative(SPEC, file).replaceAll('\\', '/');
    it(`parses ${rel}`, { timeout: 60_000 }, async () => {
      const buf = new Uint8Array(readFileSync(file));
      const expected = EXPECTED_ERRORS[base];
      if (expected) {
        expect(() => readXlsx(buf)).toThrow(expected);
        expect(() => readXlsx(buf)).toThrow(XlsxParseError);
        await expect(collectRows(buf)).rejects.toThrow(expected);
        return;
      }
      const wb = readXlsx(buf);
      const rows = await collectRows(chunked(buf, 64 * 1024));

      // streamed values equal DOM values (hyperlinks excepted)
      for (const r of rows) {
        const ws = wb.getWorksheet(r.sheet)!;
        r.values.forEach((v, c) => {
          const dom = ws.getValue(r.row, c + 1);
          if (dom && typeof dom === 'object' && 'hyperlink' in dom) expect(v).toBe(dom.text);
          else expect(v).toEqual(dom);
        });
      }

      // compare with exceljs' own reading
      if (EXCELJS_ORACLE_SKIP.has(base)) return;
      const ex = new ExcelJS.Workbook();
      try {
        // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
        await ex.xlsx.load(Buffer.from(buf) as any);
      } catch {
        return; // exceljs cannot read it; ours did not throw — good enough
      }
      const exSheets = ex.worksheets.map((w) => w.name);
      const ourSheets = wb.worksheets.map((w) => w.name);
      // missing-bits.xlsx has no workbook.xml: exceljs yields no sheets, we synthesize "Sheet1"
      if (exSheets.length > 0) expect(ourSheets).toEqual(exSheets);
      for (const ews of ex.worksheets) {
        const ows = wb.getWorksheet(ews.name)!;
        expect(ows, ews.name).toBeDefined();
        let exCount = 0;
        let ourCount = 0;
        let checked = 0;
        ews.eachRow({ includeEmpty: false }, (row, rn) => {
          row.eachCell({ includeEmpty: false }, (cell, cn) => {
            const ev = exceljsValue(cell.value);
            if (ev === null) return;
            if (cell.isMerged && cell.master !== cell) return; // exceljs mirrors merged values
            exCount++;
            if (checked++ < 2000) {
              const ov = oursValue(ows.getValue(rn, cn));
              // exceljs drops a cached formula result of 0 on some cells (test-issue-880 J5)
              if (ev === undefined && ov === 0) return;
              expect(ov, `${ews.name}!${cell.address}`).toEqual(ev);
            }
          });
        });
        ows.eachRow((row) => {
          row.eachCell((v) => {
            if (oursValue(v) !== null) ourCount++;
          });
        });
        expect(ourCount, `${ews.name} cell count`).toBe(exCount);
      }
    });
  }
});

// ---------------------------------------------------------------- hand-crafted packages

describe('hand-crafted XML edge cases', () => {
  it('infers missing r attributes on cells and rows', async () => {
    const buf = mini({
      sheets: [
        {
          name: 'S',
          xml: sheetXml(
            '<row><c><v>1</v></c><c><v>2</v></c><c r="E1"><v>5</v></c><c><v>6</v></c></row>' +
              '<row><c t="inlineStr"><is><t>x</t></is></c></row>' +
              '<row r="5"><c><v>9</v></c></row><row><c><v>10</v></c></row>',
          ),
        },
      ],
    });
    const wb = readXlsx(buf);
    expect(values(wb)).toEqual({ A1: 1, B1: 2, E1: 5, F1: 6, A2: 'x', A5: 9, A6: 10 });
    const rows = await collectRows(buf);
    expect(rows.map((r) => [r.row, r.values])).toEqual([
      [1, [1, 2, undefined, undefined, 5, 6]],
      [2, ['x']],
      [5, [9]],
      [6, [10]],
    ]);
  });

  it('handles namespace-prefixed tags', () => {
    const x = 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const buf = mini({
      sheets: [
        {
          name: 'P',
          xml: `<x:worksheet ${x}><x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:f>1+1</x:f><x:v>2</x:v></x:c></x:row></x:sheetData><x:mergeCells><x:mergeCell ref="C1:D1"/></x:mergeCells></x:worksheet>`,
        },
      ],
      sst: `<x:sst ${x}><x:si><x:t>pref</x:t></x:si></x:sst>`,
      workbook: `<x:workbook ${x} xmlns:r="${REL}"><x:workbookPr date1904="1"/><x:sheets><x:sheet name="P" sheetId="1" r:id="rIdS0"/></x:sheets></x:workbook>`,
    });
    const wb = readXlsx(buf);
    expect(wb.date1904).toBe(true);
    expect(values(wb)).toEqual({ A1: 'pref', B1: { formula: '1+1', result: 2 } });
    expect(wb.worksheets[0]!.merges).toEqual(['C1:D1']);
  });

  it('resolves absolute and unusual relationship targets', async () => {
    const buf = mini({
      wbPath: 'custom/book.xml',
      sheets: [
        {
          name: 'Abs',
          xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>'),
          path: 'custom/ws/a.xml',
          target: '/custom/ws/a.xml',
        },
        {
          name: 'Up',
          xml: sheetXml('<row r="1"><c r="A1"><v>2</v></c></row>'),
          path: 'other/Sheet1.xml',
          target: '../other/Sheet1.xml',
        },
        {
          name: 'Dot',
          xml: sheetXml('<row r="1"><c r="A1"><v>3</v></c></row>'),
          path: 'custom/worksheets/Sheet3.xml',
          target: './worksheets/Sheet3.xml',
        },
      ],
    });
    const wb = readXlsx(buf);
    expect(wb.worksheets.map((w) => [w.name, w.getValue('A1')])).toEqual([
      ['Abs', 1],
      ['Up', 2],
      ['Dot', 3],
    ]);
    const rows = await collectRows(buf);
    expect(rows.map((r) => [r.sheet, r.values[0]])).toEqual([
      ['Abs', 1],
      ['Up', 2],
      ['Dot', 3],
    ]);
  });

  it('reads t="d" ISO dates, booleans, errors, str and inline rich text', () => {
    const buf = mini({
      sheets: [
        {
          name: 'T',
          xml: sheetXml(
            '<row r="1"><c r="A1" t="d"><v>2024-03-01T12:30:00</v></c><c r="B1" t="d"><v>2024-03-01</v></c>' +
              '<c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#REF!</v></c><c r="E1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>' +
              '<c r="F1" t="inlineStr"><is><r><rPr><b/><sz val="14"/></rPr><t>B</t></r><r><t xml:space="preserve"> n</t></r></is></c>' +
              '<c r="G1" t="inlineStr"><is><r><t>plain</t></r><r><t>runs</t></r></is></c>' +
              '<c r="H1" t="d"><v>not a date</v></c></row>',
          ),
        },
      ],
    });
    const wb = readXlsx(buf);
    const ws = wb.worksheets[0]!;
    expect(ws.getValue('A1')).toEqual(new Date(Date.UTC(2024, 2, 1, 12, 30)));
    expect(ws.getValue('B1')).toEqual(new Date(Date.UTC(2024, 2, 1)));
    expect(ws.getValue('C1')).toBe(true);
    expect(ws.getValue('D1')).toEqual({ error: '#REF!' });
    expect(ws.getValue('E1')).toEqual({ formula: '"a"&"b"', result: 'ab' });
    expect(ws.getValue('F1')).toEqual({
      richText: [{ text: 'B', font: { bold: true, size: 14 } }, { text: ' n' }],
    });
    expect(ws.getValue('G1')).toBe('plainruns');
    expect(ws.getValue('H1')).toBe('not a date');
  });

  it('expands shared formulas (incl. ranges, absolute refs, sheet refs, strings)', () => {
    const buf = mini({
      sheets: [
        {
          name: 'F',
          xml: sheetXml(
            '<row r="2"><c r="B2"><f t="shared" ref="B2:C4" si="0">SUM(A1:A2)+$A$1+A$1+$A1+Sheet2!A1+\'My Sheet\'!B1&amp;"A1"</f><v>1</v></c><c r="C2"><f t="shared" si="0"/><v>2</v></c></row>' +
              '<row r="3"><c r="B3"><f t="shared" si="0"/><v>3</v></c></row>' +
              '<row r="4"><c r="C4"><f t="shared" si="1" ref="C4:C5">A:A+1:1</f></c></row><row r="5"><c r="C5"><f t="shared" si="1"/></c></row>',
          ),
        },
      ],
    });
    const ws = readXlsx(buf).worksheets[0]!;
    expect(ws.getValue('C2')).toEqual({
      formula: 'SUM(B1:B2)+$A$1+B$1+$A1+Sheet2!B1+\'My Sheet\'!C1&"A1"',
      result: 2,
    });
    expect(ws.getValue('B3')).toEqual({
      formula: 'SUM(A2:A3)+$A$1+A$1+$A2+Sheet2!A2+\'My Sheet\'!B2&"A1"',
      result: 3,
    });
    expect(ws.getValue('C5')).toEqual({ formula: 'A:A+2:2' });
  });

  it('translateFormula edge cases', () => {
    expect(translateFormula('A1+B$2+$C3+$D$4', 1, 1)).toBe('B2+C$2+$C4+$D$4');
    expect(translateFormula('LOG10(A1)+ATAN2(1,2)', 1, 0)).toBe('LOG10(A2)+ATAN2(1,2)');
    expect(translateFormula('AB1!A1+Table1[Col A1]', 1, 0)).toBe('AB1!A2+Table1[Col A1]');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
    expect(translateFormula('SUM($A:$A,B:C,3:4)', 1, 1)).toBe('SUM($A:$A,C:D,4:5)');
    expect(translateFormula('"x""A1"&A1', 0, 2)).toBe('"x""A1"&C1');
    expect(translateFormula('XFD1+A1', 0, 1)).toBe('#REF!+B1');
    expect(translateFormula('TRUE+R1C1', 3, 3)).toBe('TRUE+R1C1');
  });

  it('decodes _xHHHH_ escapes in shared and inline strings', () => {
    expect(decodeXString('a_x000D__x000A_b')).toBe('a\r\nb');
    expect(decodeXString('_x005F_x000D_')).toBe('_x000D_');
    expect(decodeXString('no_xescape_x12')).toBe('no_xescape_x12');
    const buf = mini({
      sheets: [
        {
          name: 'E',
          xml: sheetXml(
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>x_x0009_y</t></is></c><c r="C1" t="s"><v>1</v></c></row>',
          ),
        },
      ],
      sst: `<sst ${NS}><si><t>line_x000D_two</t></si><si><t>ph</t><rPh sb="0" eb="1"><t>PHONETIC</t></rPh><phoneticPr fontId="1"/></si></sst>`,
    });
    expect(values(readXlsx(buf))).toEqual({ A1: 'line\rtwo', B1: 'x\ty', C1: 'ph' });
  });

  it('empty sheet and sheet with only styles', async () => {
    const styles = `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs></styleSheet>`;
    const buf = mini({
      styles,
      sheets: [
        { name: 'Empty', xml: `<worksheet ${NS}><sheetData/></worksheet>` },
        {
          name: 'Styled',
          xml: sheetXml('<row r="3" s="1" customFormat="1"><c r="B3" s="1"/></row>', '').replace(
            '<sheetData>',
            '<cols><col min="2" max="3" width="20" style="1" customWidth="1"/></cols><sheetData>',
          ),
        },
      ],
    });
    const wb = readXlsx(buf);
    expect(wb.styles.defaultFont).toEqual({ name: 'Arial', size: 10 });
    expect(wb.worksheets[0]!.rowCount).toBe(0);
    const ws = wb.worksheets[1]!;
    expect(ws.getValue('B3')).toBe(null);
    expect(ws.getStyle('B3')).toEqual({
      numFmt: '0.000',
      font: { name: 'Arial', size: 10, bold: true },
    });
    expect(ws.getRowProps(3).style?.numFmt).toBe('0.000');
    expect(ws.getColumn(2).width).toBe(20);
    expect(ws.getColumn(3).style.numFmt).toBe('0.000');
    expect(await collectRows(buf)).toEqual([]);
  });

  it('hyperlinks: external rel, location-only, tooltip, range', () => {
    const buf = mini({
      sheets: [
        {
          name: 'H',
          xml: sheetXml(
            '<row r="1"><c r="A1" t="inlineStr"><is><t>ext</t></is></c><c r="B1" t="inlineStr"><is><t>loc</t></is></c><c r="C1"><v>5</v></c><c r="D1"><v>6</v></c></row>',
            '<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="tip"/><hyperlink ref="B1" location="H!A1"/><hyperlink ref="C1:D1" r:id="rId1" location="frag"/></hyperlinks>',
          ),
          rels: `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://x.test/" TargetMode="External"/></Relationships>`,
        },
      ],
    });
    expect(values(readXlsx(buf))).toEqual({
      A1: { text: 'ext', hyperlink: 'https://x.test/', tooltip: 'tip' },
      B1: { text: 'loc', hyperlink: '#H!A1' },
      C1: { text: '5', hyperlink: 'https://x.test/#frag' },
      D1: { text: '6', hyperlink: 'https://x.test/#frag' },
    });
  });

  it('autoFilter from _xlnm._FilterDatabase, other defined names on wb.definedNames', () => {
    const buf = mini({
      sheets: [{ name: 'A', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }],
      workbook: `<workbook ${NS}><sheets><sheet name="A" sheetId="1" r:id="rIdS0"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">A!$A$1:$C$9</definedName><definedName name="Rate">0.2</definedName></definedNames></workbook>`,
    });
    const wb = readXlsx(buf);
    expect(wb.worksheets[0]!.autoFilter).toBe('A1:C9');
    expect(wb.definedNames).toEqual([{ name: 'Rate', value: '0.2' }]);
  });

  it('defined names round-trip; sheet scopes follow loaded sheets', async () => {
    const wb = new Workbook();
    wb.addWorksheet('A').autoFilter = 'A1:B2';
    wb.addWorksheet('B').setCell('A1', 1);
    wb.addWorksheet('C').autoFilter = 'A1:C1';
    wb.definedNames.push(
      { name: 'Rate', value: '0.2' },
      { name: '_xlnm.Print_Area', localSheetId: 2, value: "'C'!$A$1:$C$9" },
      { name: 'Local', localSheetId: 1, value: "'B'!$A$1", hidden: true },
    );
    const buf = writeXlsx(wb);
    const back = readXlsx(buf);
    expect(back.definedNames).toEqual(wb.definedNames);
    expect(back.worksheets[2]!.autoFilter).toBe('A1:C1');
    const only = readXlsx(buf, { sheets: ['C'] });
    expect(only.definedNames).toEqual([
      { name: 'Rate', value: '0.2' },
      { name: '_xlnm.Print_Area', localSheetId: 0, value: "'C'!$A$1:$C$9" },
    ]);
    const x = new ExcelJS.Workbook();
    // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
    await x.xlsx.load(buf as any);
    expect(x.definedNames.getRanges('Local').ranges).toEqual(['B!$A$1']);
    expect(x.getWorksheet('C')!.pageSetup.printArea).toBe('A1:C9');
  });

  it('date-formatted numbers become Dates (built-in and custom formats)', async () => {
    const styles = `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="170" formatCode="[$-409]d\\-mmm\\-yyyy;@"/></numFmts><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="3"><xf/><xf numFmtId="14"/><xf numFmtId="170"/></cellXfs></styleSheet>`;
    const buf = mini({
      styles,
      sheets: [
        {
          name: 'D',
          xml: sheetXml(
            '<row r="1"><c r="A1" s="1"><v>45292</v></c><c r="B1" s="2"><v>45292.5</v></c><c r="C1"><v>45292</v></c></row>',
          ),
        },
      ],
    });
    const ws = readXlsx(buf).worksheets[0]!;
    expect(ws.getValue('A1')).toEqual(new Date(Date.UTC(2024, 0, 1)));
    expect(ws.getValue('B1')).toEqual(new Date(Date.UTC(2024, 0, 1, 12)));
    expect(ws.getValue('C1')).toBe(45292);
    const rows = await collectRows(buf);
    expect(rows[0]!.values).toEqual([
      new Date(Date.UTC(2024, 0, 1)),
      new Date(Date.UTC(2024, 0, 1, 12)),
      45292,
    ]);
  });

  it('streams correctly when sheets precede workbook/styles/sharedStrings (buffering)', async () => {
    const styles = `<styleSheet ${NS}><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf/><xf numFmtId="14"/></cellXfs></styleSheet>`;
    const opts: MiniOpts = {
      styles,
      sst: `<sst ${NS}><si><t>s0</t></si></sst>`,
      sheets: [
        {
          name: 'One',
          xml: sheetXml(
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c></row>',
          ),
        },
        { name: 'Two', xml: sheetXml('<row r="2"><c r="A2"><v>7</v></c></row>') },
      ],
    };
    const first = await collectRows(mini({ ...opts, order: 'meta-first' }));
    const last = await collectRows(chunked(mini({ ...opts, order: 'sheets-first' }), 50));
    const expected = [
      { sheet: 'One', sheetIndex: 0, row: 1, values: ['s0', new Date(Date.UTC(1900, 0, 1))] },
      { sheet: 'Two', sheetIndex: 1, row: 2, values: [7] },
    ];
    expect(first).toEqual(expected);
    expect(last).toEqual(expected);
    const lowMem = await collectRows(mini({ ...opts, order: 'sheets-first' }), {
      profile: 'low-memory',
    });
    expect(lowMem).toEqual(expected);
    expect(values(readXlsx(mini({ ...opts, order: 'sheets-first' })))).toEqual({
      A1: 's0',
      B1: new Date(Date.UTC(1900, 0, 1)),
    });
  });

  it('package without _rels/.rels, workbook.xml or sharedStrings still loads', async () => {
    const buf = zipSync(
      [
        {
          path: 'xl/worksheets/sheet1.xml',
          data: sheetXml('<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>'),
        },
        {
          path: 'xl/worksheets/sheet2.xml',
          data: sheetXml('<row r="1"><c r="A1"><v>2</v></c></row>'),
        },
      ],
      6,
    );
    const wb = readXlsx(buf);
    expect(wb.worksheets.map((w) => [w.name, w.getValue('A1')])).toEqual([
      ['Sheet1', 1],
      ['Sheet2', 2],
    ]);
    const rows = await collectRows(buf);
    expect(rows.map((r) => [r.sheet, r.values])).toEqual([
      ['Sheet1', [1]],
      ['Sheet2', [2]],
    ]);
  });

  it('skips chartsheets and keeps worksheet order', () => {
    const buf = mini({
      sheets: [{ name: 'W', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }],
      workbook: `<workbook ${NS}><sheets><sheet name="Chart" sheetId="2" r:id="rIdC"/><sheet name="W" sheetId="1" r:id="rIdS0"/></sheets></workbook>`,
    });
    expect(readXlsx(buf).worksheets.map((w) => w.name)).toEqual(['W']);
  });
});

// ---------------------------------------------------------------- hostile input

describe('hostile input', () => {
  it('zip bomb surfaces ZipLimitError (sync and streaming)', async () => {
    const bomb = zipSync(
      [
        { path: '[Content_Types].xml', data: '<Types/>' },
        { path: 'xl/worksheets/sheet1.xml', data: new Uint8Array(50_000_000) },
      ],
      9,
    );
    expect(() => readXlsx(bomb)).toThrow(ZipLimitError);
    await expect(collectRows(bomb)).rejects.toBeInstanceOf(ZipLimitError);
    const small = mini({
      sheets: [{ name: 'S', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }],
    });
    expect(() => readXlsx(small, { limits: { maxEntries: 2 } })).toThrow(ZipLimitError);
  });

  it('path traversal surfaces ZipPathError', async () => {
    const bad = zipSync([{ path: '../evil.xml', data: '<x/>' }], 6);
    expect(() => readXlsx(bad)).toThrow(ZipPathError);
    await expect(collectRows(bad)).rejects.toBeInstanceOf(ZipPathError);
  });

  it('malformed XML throws XlsxParseError quickly', async () => {
    for (const xml of [
      `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row>`, // truncated
      `<worksheet ${NS}><sheetData><row></sheetData></worksheet>`, // mismatched
      `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v`, // inside markup
      `junk<worksheet ${NS}/>`,
      '',
    ]) {
      const buf = mini({ sheets: [{ name: 'S', xml }] });
      const t = performance.now();
      expect(() => readXlsx(buf)).toThrow(XlsxParseError);
      await expect(collectRows(buf)).rejects.toThrow(XlsxParseError);
      expect(performance.now() - t).toBeLessThan(1000);
    }
  });

  it('garbage bytes are rejected', () => {
    expect(() => readXlsx(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toThrow();
  });

  it('out-of-range cell references throw a RangeError, not a hang', () => {
    const buf = mini({
      sheets: [{ name: 'S', xml: sheetXml('<row r="1"><c r="XFE1"><v>1</v></c></row>') }],
    });
    expect(() => readXlsx(buf)).toThrow(RangeError);
  });
});

// ---------------------------------------------------------------- perf (RUN_BENCH=1)

describe.runIf(process.env.RUN_BENCH === '1')('perf: 100k x 20', () => {
  it('ours vs exceljs', { timeout: 600_000 }, async () => {
    const ROWS = 100_000;
    const COLS = 20;
    const src = new ExcelJS.Workbook();
    const ws = src.addWorksheet('Bench');
    for (let r = 0; r < ROWS; r++) {
      const row: (number | string | Date | boolean)[] = [];
      for (let c = 0; c < COLS; c++) {
        const k = c % 4;
        row.push(
          k === 0 ? r * c : k === 1 ? `str ${r % 1000}-${c}` : k === 2 ? r % 3 === 0 : r + c / 10,
        );
      }
      ws.addRow(row);
    }
    const buf = await exceljsBuffer(src);
    const gc = (globalThis as { gc?: () => void }).gc;

    async function measure(name: string, fn: () => Promise<number>): Promise<void> {
      for (let i = 0; i < 3; i++) {
        gc?.();
        await new Promise((r) => setTimeout(r, 50)); // let ArrayBuffer backing stores be freed
      }
      // JS heap + ArrayBuffers (decompressed bytes live outside the V8 heap)
      const used = () => {
        const m = process.memoryUsage();
        return m.heapUsed + m.arrayBuffers;
      };
      const base = used();
      let peak = base;
      const timer = setInterval(() => {
        peak = Math.max(peak, used());
      }, 5);
      const t = performance.now();
      const n = await fn();
      const ms = performance.now() - t;
      clearInterval(timer);
      peak = Math.max(peak, used());
      console.log(
        `${name.padEnd(34)} ${ms.toFixed(0).padStart(6)} ms  peak heap+buffers +${((peak - base) / 1048576).toFixed(0)} MB  cells=${n}`,
      );
    }

    console.log(
      `file size ${(buf.length / 1048576).toFixed(1)} MB, gc exposed: ${gc !== undefined}`,
    );
    await measure('ours readXlsx', async () => {
      const wb = readXlsx(buf);
      let n = 0;
      for (const r of wb.worksheets[0]!.rowData) if (r) n += r.n;
      return n;
    });
    await measure('ours xlsxRows', async () => {
      let n = 0;
      for await (const r of xlsxRows(chunked(buf, 64 * 1024))) n += r.values.length;
      return n;
    });
    await measure('ours xlsxRows low-memory', async () => {
      let n = 0;
      for await (const r of xlsxRows(chunked(buf, 64 * 1024), { profile: 'low-memory' })) {
        n += r.values.length;
      }
      return n;
    });
    // same data with metadata parts before the sheet (Excel's order): no buffering at all
    const reordered = (() => {
      const files = unzipSync(buf);
      const names = [...files.keys()].sort(
        (a, b) => Number(a.includes('worksheets/')) - Number(b.includes('worksheets/')),
      );
      return zipSync(
        names.map((n) => ({ path: n, data: files.get(n)! })),
        6,
      );
    })();
    await measure('ours xlsxRows (metadata first)', async () => {
      let n = 0;
      for await (const r of xlsxRows(chunked(reordered, 64 * 1024))) n += r.values.length;
      return n;
    });
    await measure('exceljs xlsx.load', async () => {
      const wb = new ExcelJS.Workbook();
      // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
      await wb.xlsx.load(Buffer.from(buf) as any);
      let n = 0;
      wb.worksheets[0]!.eachRow((row) => {
        n += row.cellCount;
      });
      return n;
    });
    await measure('exceljs stream WorkbookReader', async () => {
      const { Readable } = await import('node:stream');
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(
        Readable.from(chunked(buf, 64 * 1024)),
        {
          sharedStrings: 'cache',
          styles: 'cache',
          worksheets: 'emit',
        },
      );
      let n = 0;
      for await (const sheet of reader) {
        for await (const row of sheet as AsyncIterable<ExcelJS.Row>) n += row.cellCount;
      }
      return n;
    });
  });
});
