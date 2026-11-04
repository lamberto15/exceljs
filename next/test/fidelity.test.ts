// Round-trip fidelity of details the model does not (fully) represent: read -> write -> read
// keeps them, exceljs still loads the output, and edits next to them behave sensibly.

import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readXlsx, Workbook, writeXlsx } from '../src/index.js';

const DATA = new URL('../../spec/integration/data/', import.meta.url);
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(new URL(name, DATA)));

type Zip = Record<string, Uint8Array>;
const text = (z: Zip, p: string): string => strFromU8(z[p]!);

/** Patch a package: replace / add parts (strings or bytes); undefined deletes. */
function patch(
  src: Uint8Array,
  edits: Record<string, string | Uint8Array | undefined>,
): Uint8Array {
  const z: Zip = unzipSync(src);
  for (const [k, v] of Object.entries(edits)) {
    if (v === undefined) delete z[k];
    else z[k] = typeof v === 'string' ? strToU8(v) : v;
  }
  return zipSync(z);
}

/** Patch one part with a string transform. */
function edit(src: Uint8Array, part: string, f: (s: string) => string): Uint8Array {
  return patch(src, { [part]: f(text(unzipSync(src), part)) });
}

async function exceljsLoads(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(buf as unknown as ArrayBuffer);
  return x;
}

/** Every r:id in sheet / workbook XML resolves to a relationship of that part; overrides exist. */
function checkPackage(z: Zip): void {
  const ct = text(z, '[Content_Types].xml');
  for (const m of ct.matchAll(/PartName="\/([^"]+)"/g)) {
    expect(z[m[1]!], `override for missing part ${m[1]}`).toBeDefined();
  }
  for (const p of Object.keys(z)) {
    if (!/^xl\/(workbook|worksheets\/sheet\d+|chartsheets\/[^/]+)\.xml$/.test(p)) continue;
    const i = p.lastIndexOf('/');
    const rp = `${p.slice(0, i + 1)}_rels/${p.slice(i + 1)}.rels`;
    const ids = new Set(z[rp] ? [...text(z, rp).matchAll(/Id="([^"]+)"/g)].map((m) => m[1]) : []);
    for (const m of text(z, p).matchAll(/\sr:id="([^"]*)"/g)) {
      expect(ids.has(m[1]!), `${p}: dangling ${m[1]}`).toBe(true);
    }
  }
}

/** A small workbook of ours: Data (3x3 values) + Other. */
function base(): Uint8Array {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['h1', 'h2', 'h3']);
  for (let i = 1; i <= 5; i++) ws.addRow([i, i * 2, `t${i}`]);
  wb.addWorksheet('Other').addRow(['x']);
  return writeXlsx(wb);
}

const SHEET1 = 'xl/worksheets/sheet1.xml';
const trip = (src: Uint8Array): Uint8Array => writeXlsx(readXlsx(src));

// ---------------------------------------------------------------- sheetFormatPr

describe('fidelity: sheetFormatPr', () => {
  const FMT =
    '<sheetFormatPr baseColWidth="10" defaultColWidth="12.5" defaultRowHeight="18" customHeight="1" zeroHeight="1" thickBottom="1" x14ac:dyDescent="0.2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"/>';
  const src = (): Uint8Array =>
    edit(base(), SHEET1, (s) => s.replace(/<sheetFormatPr[^>]*\/>/, FMT));

  it('keeps every attribute; outline levels follow the model', async () => {
    const out = trip(src());
    checkPackage(unzipSync(out));
    const s = text(unzipSync(out), SHEET1);
    const el = s.match(/<sheetFormatPr[^>]*\/>/)![0];
    for (const a of [
      'baseColWidth="10"',
      'defaultColWidth="12.5"',
      'defaultRowHeight="18"',
      'customHeight="1"',
      'zeroHeight="1"',
      'thickBottom="1"',
      'x14ac:dyDescent="0.2"',
    ]) {
      expect(el).toContain(a);
    }
    expect(s.match(/<sheetFormatPr/g)!.length).toBe(1);
    // stable on a second trip
    expect(text(unzipSync(trip(out)), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0]).toBe(el);
    await exceljsLoads(out);
  });

  it('user edits nearby: outline levels are recomputed, kept attributes stay', () => {
    const wb = readXlsx(src());
    const ws = wb.worksheets[0]!;
    ws.setRowProps(3, { outlineLevel: 2 });
    ws.getColumn(2).outlineLevel = 1;
    const el = text(unzipSync(writeXlsx(wb)), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el).toContain('outlineLevelRow="2"');
    expect(el).toContain('outlineLevelCol="1"');
    expect(el).toContain('defaultColWidth="12.5"');
    ws.setRowProps(3, { outlineLevel: 0 });
    ws.getColumn(2).outlineLevel = 0;
    const el2 = text(unzipSync(writeXlsx(wb)), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el2).not.toContain('outlineLevel');
  });

  it('a stale outline level in the file is not kept', () => {
    const wb = readXlsx(
      edit(base(), SHEET1, (s) =>
        s.replace(
          /<sheetFormatPr[^>]*\/>/,
          '<sheetFormatPr defaultRowHeight="15" defaultColWidth="20" outlineLevelRow="3"/>',
        ),
      ),
    );
    const el = text(unzipSync(writeXlsx(wb)), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el).toBe('<sheetFormatPr defaultRowHeight="15" defaultColWidth="20"/>');
  });

  it('writer defaults are not captured (no preserved data for plain sheets)', () => {
    const wb = readXlsx(base());
    expect(wb.worksheets[0]!.preserved).toBeUndefined();
  });

  it('formulas.xlsx (real file): baseColWidth / defaultRowHeight kept', () => {
    const s = text(unzipSync(trip(fixture('formulas.xlsx'))), SHEET1);
    expect(s).toMatch(/<sheetFormatPr[^>]*baseColWidth="10"/);
    expect(s).toMatch(/<sheetFormatPr[^>]*defaultRowHeight="16"/);
  });

  it('preserveUnknown: false writes the defaults', () => {
    const wb = readXlsx(src(), { preserveUnknown: false });
    const el = text(unzipSync(writeXlsx(wb)), SHEET1).match(/<sheetFormatPr[^>]*\/>/)![0];
    expect(el).toBe('<sheetFormatPr defaultRowHeight="15"/>');
  });
});

// ---------------------------------------------------------------- autoFilter criteria

describe('fidelity: autoFilter criteria and sortState', () => {
  const CRITERIA =
    '<filterColumn colId="0"><filters blank="1"><filter val="1"/><filter val="3"/></filters></filterColumn>' +
    '<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="2"/><customFilter operator="lessThan" val="9"/></customFilters></filterColumn>' +
    '<filterColumn colId="2"><top10 percent="1" val="10"/></filterColumn>';
  const SORT = '<sortState ref="A2:C6"><sortCondition descending="1" ref="B2:B6"/></sortState>';
  const AF = `<autoFilter ref="A1:C6">${CRITERIA}${SORT}</autoFilter>`;
  const SHEET_SORT = '<sortState ref="A2:C6"><sortCondition ref="C2:C6"/></sortState>';
  const src = (): Uint8Array =>
    edit(base(), SHEET1, (s) => s.replace('</sheetData>', `</sheetData>${AF}${SHEET_SORT}`));

  it('filterColumn criteria + sortState (inside autoFilter and sheet level) round-trip', async () => {
    const wb = readXlsx(src());
    expect(wb.worksheets[0]!.autoFilter).toBe('A1:C6');
    const out = writeXlsx(wb);
    const z = unzipSync(out);
    checkPackage(z);
    const s = text(z, SHEET1);
    expect(s).toContain(AF);
    expect(s).toContain(`</autoFilter>${SHEET_SORT}`);
    expect(s.match(/<autoFilter/g)!.length).toBe(1);
    expect(text(z, 'xl/workbook.xml')).toContain('_xlnm._FilterDatabase');
    expect(text(unzipSync(trip(out)), SHEET1)).toContain(AF);
    const x = await exceljsLoads(out);
    expect(x.worksheets[0]!.autoFilter).toBeDefined();
  });

  it('dynamic / color / icon filters (with a dxfId) are kept', () => {
    const af =
      '<autoFilter ref="A1:C6"><filterColumn colId="0"><dynamicFilter type="aboveAverage"/></filterColumn><filterColumn colId="1"><colorFilter dxfId="0"/></filterColumn><filterColumn colId="2"><iconFilter iconSet="3Arrows" iconId="1"/></filterColumn></autoFilter>';
    const s = text(
      unzipSync(trip(edit(base(), SHEET1, (x) => x.replace('</sheetData>', `</sheetData>${af}`)))),
      SHEET1,
    );
    expect(s).toContain('<dynamicFilter type="aboveAverage"/>');
    expect(s).toContain('<colorFilter dxfId="0"/>');
    expect(s).toContain('<iconFilter iconSet="3Arrows" iconId="1"/>');
  });

  it('range grown to the right / down: criteria kept, sortState dropped', () => {
    const wb = readXlsx(src());
    wb.worksheets[0]!.autoFilter = 'A1:D20';
    const s = text(unzipSync(writeXlsx(wb)), SHEET1);
    expect(s).toContain(`<autoFilter ref="A1:D20">${CRITERIA}</autoFilter>`);
    // the sheet-level sortState is independent of the filter
    expect(s).toContain(SHEET_SORT);
  });

  it('range narrowed: columns outside the new range are dropped', () => {
    const wb = readXlsx(src());
    wb.worksheets[0]!.autoFilter = 'A1:B6';
    const s = text(unzipSync(writeXlsx(wb)), SHEET1);
    const af = s.match(/<autoFilter[\s\S]*?<\/autoFilter>/)![0];
    expect(af).toContain('colId="0"');
    expect(af).toContain('colId="1"');
    expect(af).not.toContain('colId="2"');
  });

  it('range moved to another first column: criteria dropped; removed: no autoFilter', () => {
    const wb = readXlsx(src());
    wb.worksheets[0]!.autoFilter = 'B1:C6';
    const s = text(unzipSync(writeXlsx(wb)), SHEET1);
    expect(s).toContain('<autoFilter ref="B1:C6"/>');
    expect(s).not.toContain('filterColumn');
    wb.worksheets[0]!.autoFilter = undefined;
    const s2 = text(unzipSync(writeXlsx(wb)), SHEET1);
    expect(s2).not.toContain('<autoFilter');
    expect(s2).not.toContain('filterColumn');
  });

  it('user edits nearby (cells inside the filtered range) keep the criteria', () => {
    const wb = readXlsx(src());
    wb.worksheets[0]!.setCell('B3', 42);
    expect(text(unzipSync(writeXlsx(wb)), SHEET1)).toContain(AF);
  });

  it('preserveUnknown: false keeps only the range', () => {
    const s = text(unzipSync(writeXlsx(readXlsx(src(), { preserveUnknown: false }))), SHEET1);
    expect(s).toContain('<autoFilter ref="A1:C6"/>');
    expect(s).not.toContain('sortState');
  });
});
