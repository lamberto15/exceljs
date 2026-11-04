// Round-trip fidelity of details the model does not (fully) represent: read -> write -> read
// keeps them, exceljs still loads the output, and edits next to them behave sensibly.

import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { addNote, readXlsx, Workbook, writeXlsx } from '../src/index.js';

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

// ---------------------------------------------------------------- named cell styles

describe('fidelity: named cell styles (cellStyleXfs / cellStyles)', () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const H1_FONT = '<font><b/><sz val="15"/><color theme="3"/><name val="Calibri"/></font>';
  // font 3 duplicates font 1: the registry merges them, so kept ids must be remapped
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}"><fonts count="4"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>${H1_FONT}<font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>${H1_FONT}</fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border></borders><cellStyleXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="2" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="10" fontId="3" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="1"/><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="2"/><xf numFmtId="10" fontId="3" fillId="0" borderId="0" xfId="3" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/></cellXfs><cellStyles count="4"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Heading 1" xfId="1" builtinId="16"/><cellStyle name="Good" xfId="2" builtinId="26"/><cellStyle name="My Percent" xfId="3"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>`;
  const src = (): Uint8Array => {
    let b = edit(base(), SHEET1, (s) =>
      s
        .replace('<c r="A1"', '<c r="A1" s="1"')
        .replace('<c r="B1"', '<c r="B1" s="2"')
        .replace('<c r="C1"', '<c r="C1" s="3"')
        .replace('<c r="A2"', '<c r="A2" s="4"'),
    );
    b = patch(b, { 'xl/styles.xml': STYLES });
    return b;
  };

  /** Attributes of the `i`-th `<xf>` of a styles.xml list. */
  function xfOf(styles: string, list: string, i: number): Record<string, string> {
    const body = styles.match(new RegExp(`<${list}[^>]*>([\\s\\S]*?)</${list}>`))![1]!;
    const xf = [...body.matchAll(/<xf ([^>]*?)\/?>/g)][i]![1]!;
    return Object.fromEntries([...xf.matchAll(/(\w+)="([^"]*)"/g)].map((m) => [m[1]!, m[2]!]));
  }
  const cellS = (sheet: string, ref: string): number =>
    +(new RegExp(`<c r="${ref}" s="(\\d+)"`).exec(sheet)?.[1] ?? 0);

  it('names, style xfs and cell links survive; ids are remapped to the written tables', async () => {
    const out = trip(src());
    const z = unzipSync(out);
    checkPackage(z);
    const st = text(z, 'xl/styles.xml');
    const sheet = text(z, SHEET1);
    for (const n of ['Normal', 'Heading 1', 'Good', 'My Percent']) {
      expect(st).toContain(`name="${n}"`);
    }
    expect(st).toContain('<cellStyleXfs count="4">');
    const fonts = [
      ...st.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)![1]!.matchAll(/<font>.*?<\/font>/g),
    ];
    for (const [ref, xfId] of [
      ['A1', '1'],
      ['B1', '2'],
      ['C1', '3'],
      ['A2', '0'],
    ] as const) {
      const cx = xfOf(st, 'cellXfs', cellS(sheet, ref));
      expect(cx.xfId, ref).toBe(xfId);
      // the named style's ids point at the same written font / fill / border as the cell's
      const sx = xfOf(st, 'cellStyleXfs', +xfId);
      if (xfId !== '0') {
        expect(sx.fontId, ref).toBe(cx.fontId);
        expect(sx.fillId, ref).toBe(cx.fillId);
        expect(sx.borderId, ref).toBe(cx.borderId);
      }
    }
    // the duplicate font was merged: Heading 1 and My Percent share the bold 15pt font
    expect(xfOf(st, 'cellStyleXfs', 3).fontId).toBe(xfOf(st, 'cellStyleXfs', 1).fontId);
    expect(fonts[+xfOf(st, 'cellStyleXfs', 1).fontId!]![0]).toMatch(/<b\/>.*<sz val="15"\/>/);
    expect(xfOf(st, 'cellStyleXfs', 2).fillId).not.toBe('0');
    // A1 (Heading 1) and A2 (same formatting, no named style) stay distinct xfs
    expect(cellS(sheet, 'A1')).not.toBe(cellS(sheet, 'A2'));
    // stable
    expect(text(unzipSync(trip(out)), 'xl/styles.xml')).toBe(st);
    const x = await exceljsLoads(out);
    expect(x.worksheets[0]!.getCell('A1').font?.bold).toBe(true);
    expect(x.worksheets[0]!.getCell('C1').numFmt).toBe('0.00%');
  });

  it('user edits nearby: new styles get xfId 0, reused style objects keep the link, dedupe works', () => {
    const wb = readXlsx(src());
    const ws = wb.worksheets[0]!;
    const n = wb.styles.xfs.length;
    ws.setStyle('D1', { font: { bold: true } });
    ws.setStyle('D2', ws.getStyle('A1')); // same style object: same xf (Heading 1 link kept)
    ws.setStyle('D3', { ...ws.getStyle('A1') }); // equal formatting, new object: dedupes to the unnamed xf
    expect(ws.getStyleIndex('D2')).toBe(ws.getStyleIndex('A1'));
    expect(ws.getStyleIndex('D3')).toBe(ws.getStyleIndex('A2'));
    expect(wb.styles.xfs.length).toBe(n + 1);
    const z = unzipSync(writeXlsx(wb));
    const st = text(z, 'xl/styles.xml');
    const sheet = text(z, SHEET1);
    expect(xfOf(st, 'cellXfs', cellS(sheet, 'D1')).xfId).toBe('0');
    expect(xfOf(st, 'cellXfs', cellS(sheet, 'D2')).xfId).toBe('1');
    expect(st).toContain('name="Heading 1"');
  });

  it('gold.xlsx (real file): the Hyperlink style is kept and still linked', () => {
    const st = text(unzipSync(trip(fixture('gold.xlsx'))), 'xl/styles.xml');
    expect(st).toContain('<cellStyle name="Hyperlink" xfId="1" builtinId="8"/>');
    expect(st).toContain('<cellStyleXfs count="2">');
    expect(st).toMatch(/<cellXfs[^>]*>.*xfId="1".*<\/cellXfs>/);
  });

  it('preserveUnknown: false / deleting wb.preserved: only Normal, every xfId 0', () => {
    for (const wb of [
      readXlsx(src(), { preserveUnknown: false }),
      (() => {
        const w = readXlsx(src());
        delete w.preserved;
        return w;
      })(),
    ]) {
      const st = text(unzipSync(writeXlsx(wb)), 'xl/styles.xml');
      expect(st).not.toContain('Heading 1');
      expect(st).toContain('<cellStyleXfs count="1">');
      expect(st).not.toMatch(/xfId="[1-9]/);
    }
  });

  it('files without named styles are written as before', () => {
    const st = text(unzipSync(trip(base())), 'xl/styles.xml');
    expect(st).toContain(
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    );
  });
});

// ---------------------------------------------------------------- form controls + notes

describe('fidelity: form controls on a sheet with notes (legacy VML merge)', () => {
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const CHECKBOX =
    '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>' +
    '<v:shape id="_x0000_s1026" type="#_x0000_t201" style="position:absolute;margin-left:200pt;margin-top:20pt;width:80pt;height:18pt;z-index:2" filled="f" stroked="f" o:insetmode="auto"><v:textbox style="mso-direction-alt:auto" o:singleclick="f"><div style="text-align:left"><font face="Segoe UI" size="160" color="auto">Done</font></div></v:textbox><x:ClientData ObjectType="Checkbox"><x:SizeWithCells/><x:Anchor>4, 0, 1, 0, 5, 40, 2, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine><x:TextVAlign>Center</x:TextVAlign><x:Checked>1</x:Checked><x:NoThreeD/></x:ClientData></v:shape>';
  const CONTROLS = `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14"><controls><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14"><control shapeId="1026" r:id="rIdCtl" name="Check Box 2"><controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1"><from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>5</xdr:col><xdr:colOff>40</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>4</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent>`;
  const CTRL_PROP =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked" lockText="1" noThreeD="1"/>';

  /** Data has a note (A1) + a checkbox in the same VML; Other has a note. */
  function src(): Uint8Array {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['a', 'b']);
    addNote(ws, 'A1', 'first note');
    addNote(wb.addWorksheet('Other'), 'B2', 'other note');
    const b = writeXlsx(wb);
    const z = unzipSync(b);
    const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    const vmlRel = /Target="\.\.\/drawings\/(vmlDrawing\d+\.vml)"/.exec(text(z, relsPath))![1]!;
    const vmlPath = `xl/drawings/${vmlRel}`;
    return patch(b, {
      [vmlPath]: text(z, vmlPath).replace('</xml>', `${CHECKBOX}</xml>`),
      [SHEET1]: text(z, SHEET1)
        .replace(
          '<worksheet ',
          '<worksheet xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
        )
        .replace('</worksheet>', `${CONTROLS}</worksheet>`),
      [relsPath]: text(z, relsPath).replace(
        '</Relationships>',
        `<Relationship Id="rIdCtl" Type="${REL}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>`,
      ),
      'xl/ctrlProps/ctrlProp1.xml': CTRL_PROP,
      '[Content_Types].xml': text(z, '[Content_Types].xml').replace(
        '</Types>',
        '<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>',
      ),
    });
  }

  /** The VML part linked from a sheet. */
  function vmlOf(z: Zip, sheet: number): string {
    const rels = text(z, `xl/worksheets/_rels/sheet${sheet}.xml.rels`);
    return text(z, `xl/drawings/${/Target="\.\.\/drawings\/([^"]+\.vml)"/.exec(rels)![1]}`);
  }

  const shapeIds = (vml: string): string[] =>
    [...vml.matchAll(/<v:shape id="_x0000_s(\d+)"/g)].map((m) => m[1]!);

  it('notes are regenerated and the checkbox shape kept in the same VML; controls stay linked', async () => {
    const wb = readXlsx(src());
    expect(wb.worksheets[0]!.comments!.length).toBe(1);
    const out = writeXlsx(wb);
    const z = unzipSync(out);
    checkPackage(z);
    const vml = vmlOf(z, 1);
    expect(vml).toContain('ObjectType="Note"');
    expect(vml).toContain('ObjectType="Checkbox"');
    expect(vml).toContain('<x:Checked>1</x:Checked>');
    expect(vml).toContain('id="_x0000_t201"');
    expect(vml.match(/id="_x0000_t202"/g)!.length).toBe(1);
    // the checkbox keeps its id (the control refers to it); notes use another block
    const ids = shapeIds(vml);
    expect(ids).toContain('1026');
    expect(new Set(ids).size).toBe(ids.length);
    const blocks = /<o:idmap v:ext="edit" data="([\d,]+)"/.exec(vml)![1]!.split(',');
    expect(blocks).toContain('1');
    for (const id of ids) expect(blocks).toContain(String(Math.floor(+id / 1024)));
    // no other VML in the package uses the kept block
    const other = vmlOf(z, 2);
    expect(/data="([\d,]+)"/.exec(other)![1]!.split(',')).not.toContain('1');
    const s = text(z, SHEET1);
    expect(s.match(/<legacyDrawing /g)!.length).toBe(1);
    expect(s).toContain('<control shapeId="1026" r:id="rId');
    expect(s).toMatch(/<mc:AlternateContent xmlns:mc="[^"]+"[^>]*><mc:Choice Requires="x14"/);
    expect(s.indexOf('<legacyDrawing')).toBeLessThan(s.indexOf('<mc:AlternateContent'));
    expect(text(z, 'xl/ctrlProps/ctrlProp1.xml')).toBe(CTRL_PROP);
    // stable on a second trip
    const z2 = unzipSync(trip(out));
    checkPackage(z2);
    expect(shapeIds(vmlOf(z2, 1)).sort()).toEqual(ids.sort());
    expect(text(z2, SHEET1)).toContain('<control shapeId="1026"');
    const x = await exceljsLoads(out);
    expect(x.worksheets[0]!.getCell('A1').note).toBeDefined();
  });

  it('user edits nearby: notes added / removed keep the checkbox', () => {
    const wb = readXlsx(src());
    const ws = wb.worksheets[0]!;
    addNote(ws, 'C3', 'new');
    let vml = vmlOf(unzipSync(writeXlsx(wb)), 1);
    expect(vml.match(/ObjectType="Note"/g)!.length).toBe(2);
    expect(vml).toContain('ObjectType="Checkbox"');
    ws.comments = [];
    const z = unzipSync(writeXlsx(wb));
    checkPackage(z);
    vml = vmlOf(z, 1);
    expect(vml).not.toContain('ObjectType="Note"');
    expect(vml).toContain('_x0000_s1026');
    expect(text(z, 'xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('/comments"');
  });

  it('renaming the sheet drops the controls and their shapes together', () => {
    const wb = readXlsx(src());
    wb.worksheets[0]!.name = 'Renamed';
    const z = unzipSync(writeXlsx(wb));
    checkPackage(z);
    expect(vmlOf(z, 1)).not.toContain('Checkbox');
    expect(text(z, SHEET1)).not.toContain('<control ');
  });

  it('controls on a sheet without notes: VML and controls kept as they are', () => {
    const b = edit(src(), 'xl/worksheets/sheet1.xml', (s) => s);
    const wb = readXlsx(b, { features: [] });
    const z = unzipSync(writeXlsx(wb));
    checkPackage(z);
    expect(vmlOf(z, 1)).toContain('Checkbox');
    expect(text(z, SHEET1)).toContain('<control shapeId="1026"');
  });
});
