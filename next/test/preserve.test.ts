import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addNote,
  listPreservedParts,
  readXlsx,
  Workbook,
  writeXlsx,
  xlsxStream,
} from '../src/index.js';

const DATA = new URL('../../spec/integration/data/', import.meta.url);
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(new URL(name, DATA)));

type Zip = Record<string, Uint8Array>;
const text = (z: Zip, p: string): string => strFromU8(z[p]!);

function resolve(base: string, target: string): string {
  const out: string[] = [];
  const full = target.startsWith('/') ? target.slice(1) : base + target;
  for (const seg of full.split('/')) {
    if (seg === '..') out.pop();
    else if (seg && seg !== '.') out.push(seg);
  }
  return out.join('/');
}

interface Rel {
  id: string;
  type: string;
  target: string;
  path: string;
  external: boolean;
}

function relsOf(z: Zip, part: string): Rel[] {
  const i = part.lastIndexOf('/');
  const relsPath = `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
  if (!z[relsPath]) return [];
  const base = part.slice(0, i + 1);
  const out: Rel[] = [];
  for (const m of text(z, relsPath).matchAll(/<Relationship ([^>]*?)\/?>/g)) {
    const a = Object.fromEntries([...m[1]!.matchAll(/(\w+)="([^"]*)"/g)].map((x) => [x[1], x[2]]));
    const external = a.TargetMode === 'External';
    out.push({
      id: a.Id!,
      type: a.Type!,
      target: a.Target!,
      path: external ? a.Target! : resolve(base, a.Target!.replace(/&amp;/g, '&')),
      external,
    });
  }
  return out;
}

/** OPC sanity: content types cover every part, rels resolve, no orphan rels, ids unique. */
function checkPackage(z: Zip): void {
  const ct = text(z, '[Content_Types].xml');
  const defaults = new Set(
    [...ct.matchAll(/Extension="([^"]+)"/g)].map((m) => m[1]!.toLowerCase()),
  );
  const overrides = [...ct.matchAll(/PartName="\/([^"]+)"/g)].map((m) => m[1]!);
  expect(new Set(overrides).size, 'duplicate overrides').toBe(overrides.length);
  for (const o of overrides) expect(z[o], `override for missing part ${o}`).toBeDefined();
  for (const p of Object.keys(z)) {
    if (p.endsWith('/')) continue;
    const ext = p.slice(p.lastIndexOf('.') + 1).toLowerCase();
    expect(overrides.includes(p) || defaults.has(ext), `content type for ${p}`).toBe(true);
    if (p.endsWith('.rels')) {
      const src = p.replace(/_rels\/([^/]*)\.rels$/, '$1');
      if (src !== '') expect(z[src], `rels without source ${p}`).toBeDefined();
      const rels = relsOf(z, src);
      expect(new Set(rels.map((r) => r.id)).size, `duplicate ids in ${p}`).toBe(rels.length);
      for (const r of rels) {
        if (!r.external) expect(z[r.path], `${p}: ${r.id} → ${r.path}`).toBeDefined();
      }
    }
  }
  // every r:id in sheet / workbook / drawing XML resolves to a relationship of that part
  for (const p of Object.keys(z)) {
    if (
      !/^xl\/(workbook|worksheets\/sheet\d+|drawings\/drawing\d+|chartsheets\/[^/]+)\.xml$/.test(p)
    ) {
      continue;
    }
    const ids = new Set(relsOf(z, p).map((r) => r.id));
    for (const m of text(z, p).matchAll(/\sr:(?:id|embed|link)="([^"]*)"/g)) {
      expect(ids.has(m[1]!), `${p}: dangling ${m[1]}`).toBe(true);
    }
  }
}

function roundTrip(src: Uint8Array): { wb: Workbook; out: Uint8Array; z: Zip } {
  const wb = readXlsx(src);
  const out = writeXlsx(wb);
  return { wb, out, z: unzipSync(out) };
}

async function exceljsLoads(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(buf as unknown as ArrayBuffer);
  return x;
}

const byteEq = (a: Uint8Array | undefined, b: Uint8Array | undefined): boolean =>
  a !== undefined && b !== undefined && a.length === b.length && a.every((v, i) => v === b[i]);

// ---------------------------------------------------------------- real fixtures

describe('preserve: real fixtures', () => {
  it('chart-sheet: chartsheet + drawing + chart + style/colors + theme kept', async () => {
    const src = fixture('chart-sheet.xlsx');
    const zs = unzipSync(src);
    const { wb, out, z } = roundTrip(src);
    checkPackage(z);
    for (const p of [
      'xl/charts/chart1.xml',
      'xl/charts/style1.xml',
      'xl/charts/colors1.xml',
      'xl/drawings/drawing1.xml',
      'xl/theme/theme1.xml',
    ]) {
      expect(byteEq(z[p], zs[p]), p).toBe(true);
    }
    expect(z['xl/chartsheets/sheet1.xml']).toBeDefined();
    expect(text(z, 'xl/chartsheets/sheet1.xml')).not.toContain('tabSelected');
    expect(text(z, 'xl/workbook.xml')).toMatch(
      /<sheet name="Sheet1" sheetId="1" r:id="rId1"\/><sheet name="Chart1" sheetId="9" r:id="rId\d+"\/>/,
    );
    // a new worksheet gets an id no kept sheet uses
    const wb2 = readXlsx(src);
    wb2.addWorksheet('New');
    const ids = [
      ...text(unzipSync(writeXlsx(wb2)), 'xl/workbook.xml').matchAll(/sheetId="(\d+)"/g),
    ];
    expect(new Set(ids.map((m) => m[1])).size).toBe(3);
    expect(text(z, 'xl/workbook.xml')).toContain('codeName="ThisWorkbook"');
    // the kept theme replaces the default one: one theme part, one relationship
    expect(Object.keys(z).filter((p) => p.startsWith('xl/theme/'))).toEqual([
      'xl/theme/theme1.xml',
    ]);
    expect(relsOf(z, 'xl/workbook.xml').filter((r) => r.type.endsWith('/theme')).length).toBe(1);
    expect(text(z, 'xl/workbook.xml')).toContain('<fileVersion appName="xl"');
    expect(listPreservedParts(wb).map((p) => p.path)).toContain('xl/charts/chart1.xml');
    const x = await exceljsLoads(out);
    expect(x.worksheets.map((w) => w.name)).toContain('Sheet1');
    const again = readXlsx(out);
    expect(again.worksheets.map((w) => w.name)).toEqual(['Sheet1']);
    expect(listPreservedParts(again).length).toBe(listPreservedParts(wb).length);
  });

  it('huge: connections + queryTables kept, linked from sheets', { timeout: 60000 }, async () => {
    const src = fixture('huge.xlsx');
    const { out, z } = roundTrip(src);
    checkPackage(z);
    expect(z['xl/connections.xml']).toBeDefined();
    const qt = Object.keys(z).filter((p) => p.startsWith('xl/queryTables/'));
    expect(qt.length).toBe(2);
    const rels = [1, 2].flatMap((i) => relsOf(z, `xl/worksheets/sheet${i}.xml`));
    expect(rels.filter((r) => r.type.endsWith('/queryTable')).length).toBe(2);
    await exceljsLoads(out);
  });

  it('calcChain is dropped', () => {
    const { z } = roundTrip(fixture('fibonacci.xlsx'));
    checkPackage(z);
    expect(z['xl/calcChain.xml']).toBeUndefined();
    expect(text(z, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
  });

  it('custom properties and thumbnails (root relationships) are kept', () => {
    for (const [f, p] of [
      ['test-issue-771.xlsx', 'docProps/custom.xml'],
      ['test-issue-257.xlsx', 'docProps/thumbnail.jpeg'],
    ] as const) {
      const zs = unzipSync(fixture(f));
      const { z } = roundTrip(fixture(f));
      checkPackage(z);
      expect(byteEq(z[p], zs[p]), p).toBe(true);
      expect(text(z, '_rels/.rels')).toContain(`Target="${p}"`);
    }
  });

  for (const f of [
    '1519293514-KRISHNAPATNAM_LINE_UP.xlsx',
    'images.xlsx',
    'test-issue-877.xlsx',
    'test-issue-880.xlsx',
    'test-issue-1669.xlsx',
    'test-pr-1204.xlsx',
    'formulas.xlsx',
    'gold.xlsx',
  ]) {
    it(`${f}: valid package, loads in exceljs and ours, stable on second trip`, async () => {
      const src = fixture(f);
      const zs = unzipSync(src);
      const { wb, out, z } = roundTrip(src);
      checkPackage(z);
      // untouched kept parts: same path, same bytes
      for (const p of listPreservedParts(wb))
        expect(byteEq(z[p.path], zs[p.path]), p.path).toBe(true);
      await exceljsLoads(out);
      const again = readXlsx(out);
      expect(again.worksheets.map((w) => w.name)).toEqual(wb.worksheets.map((w) => w.name));
      const z2 = unzipSync(writeXlsx(again));
      checkPackage(z2);
      expect(Object.keys(z2).sort()).toEqual(Object.keys(z).sort());
    });
  }

  it('test-issue-877: header/footer VML images (legacyDrawingHF) kept with their media', () => {
    const src = fixture('test-issue-877.xlsx');
    const zs = unzipSync(src);
    const { z } = roundTrip(src);
    const vmls = Object.keys(zs).filter((p) => p.endsWith('.vml'));
    for (const p of vmls) expect(z[p] ?? null, p).not.toBeNull();
  });

  it('preserveUnknown: false drops everything kept', () => {
    const wb = readXlsx(fixture('chart-sheet.xlsx'), { preserveUnknown: false });
    expect(wb.preserved).toBeUndefined();
    expect(listPreservedParts(wb)).toEqual([]);
    const z = unzipSync(writeXlsx(wb));
    checkPackage(z);
    expect(Object.keys(z).some((p) => p.includes('chart'))).toBe(false);
    // the default theme instead of the file's
    expect(text(z, 'xl/theme/theme1.xml')).toContain('name="Office Theme"');
  });

  it('xlsxStream writes the same preserved parts', async () => {
    const wb = readXlsx(fixture('chart-sheet.xlsx'));
    const parts: Uint8Array[] = [];
    for await (const c of xlsxStream(wb) as unknown as AsyncIterable<Uint8Array>) parts.push(c);
    const buf = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of parts) {
      buf.set(p, o);
      o += p.length;
    }
    const z = unzipSync(buf);
    checkPackage(z);
    expect(Object.keys(z).sort()).toEqual(Object.keys(unzipSync(writeXlsx(wb))).sort());
  });
});

// ---------------------------------------------------------------- crafted fixtures

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

/** A small base workbook written by ExcelJS (an independent producer). */
async function base(withImage = false): Promise<Uint8Array> {
  const x = new ExcelJS.Workbook();
  const ws = x.addWorksheet('Data');
  ws.addRow(['a', 'b', 'c']);
  ws.addRow([1, 2, 3]);
  ws.addRow([4, 5, 6]);
  x.addWorksheet('Other').addRow(['x']);
  if (withImage) {
    const id = x.addImage({
      buffer: readFileSync(new URL('image.png', DATA)) as never,
      extension: 'png',
    });
    ws.addImage(id, 'E2:F4');
  }
  return new Uint8Array((await x.xlsx.writeBuffer()) as ArrayBuffer);
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
// ---------------------------------------------------------------- sheet elements

const SHEET_TAIL = `<rowBreaks count="1" manualBreakCount="1"><brk id="2" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks><ignoredErrors><ignoredError sqref="A1:C3" numberStoredAsText="1"/></ignoredErrors>`;
const SHEET_EXT = `<extLst><ext uri="{FAKE-0000}" xmlns:my="urn:my"><my:thing a="1">t &amp; u</my:thing></ext></extLst>`;

async function withSheetXml(): Promise<Uint8Array> {
  const src = await base();
  const z = unzipSync(src);
  let s = text(z, 'xl/worksheets/sheet1.xml');
  // elements go to their CT_Worksheet positions: breaks/ignoredErrors before drawing etc.
  s = s.replace(/(<pageMargins[^>]*\/>)(<pageSetup[^>]*\/>)?/, `$1$2${SHEET_TAIL}`);
  s = s.replace(/<sheetPr([^>]*)>/, '<sheetPr codeName="Sheet1"$1>');
  s = s.replace(/<sheetPr\/>/, '<sheetPr codeName="Sheet1"/>');
  if (!s.includes('codeName'))
    s = s.replace(/(<worksheet[^>]*>)/, '$1<sheetPr codeName="Sheet1"/>');
  s = s.replace(/<\/worksheet>$/, `${SHEET_EXT}</worksheet>`);
  return patch(src, { 'xl/worksheets/sheet1.xml': s });
}

describe('preserve: unknown worksheet elements', () => {
  it('rowBreaks, colBreaks, ignoredErrors, sheetPr codeName and an unknown ext round-trip in order', async () => {
    const src = await withSheetXml();
    expect(text(unzipSync(src), 'xl/worksheets/sheet1.xml')).toContain('<rowBreaks');
    const wb = readXlsx(src);
    const out = writeXlsx(wb);
    const z = unzipSync(out);
    checkPackage(z);
    const s = text(z, 'xl/worksheets/sheet1.xml');
    expect(s).toContain('<sheetPr codeName="Sheet1"');
    expect(s).toContain(
      '<rowBreaks count="1" manualBreakCount="1"><brk id="2" max="16383" man="1"/></rowBreaks>',
    );
    expect(s).toContain(
      '<ignoredErrors><ignoredError sqref="A1:C3" numberStoredAsText="1"/></ignoredErrors>',
    );
    expect(s).toContain(
      '<ext uri="{FAKE-0000}" xmlns:my="urn:my"><my:thing a="1">t &amp; u</my:thing></ext>',
    );
    const order = [
      '<sheetData',
      '<pageMargins',
      '<rowBreaks',
      '<colBreaks',
      '<ignoredErrors',
      '<extLst',
    ];
    const pos = order.map((t) => s.indexOf(t));
    expect(pos.every((p) => p >= 0)).toBe(true);
    expect([...pos].sort((a, b) => a - b)).toEqual(pos);
    await exceljsLoads(out);
    // stable
    const s2 = text(unzipSync(writeXlsx(readXlsx(out))), 'xl/worksheets/sheet1.xml');
    expect(s2).toBe(s);
  });

  it('ours win over preserved elements in the same slot (no duplicates)', async () => {
    const wb = readXlsx(await withSheetXml());
    addNote(wb.worksheets[0]!, 'A1', 'hello');
    const s = text(unzipSync(writeXlsx(wb)), 'xl/worksheets/sheet1.xml');
    expect(s.match(/<legacyDrawing /g)!.length).toBe(1);
    expect(s.match(/<extLst>/g)!.length).toBe(1);
  });

  it('preserveUnknown: false drops them', async () => {
    const wb = readXlsx(await withSheetXml(), { preserveUnknown: false });
    const s = text(unzipSync(writeXlsx(wb)), 'xl/worksheets/sheet1.xml');
    expect(s).not.toContain('rowBreaks');
    expect(s).not.toContain('codeName');
    expect(s).not.toContain('FAKE');
  });
});

// ---------------------------------------------------------------- threaded comments

const VML_CTRL = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shape id="_x0000_s1025" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"><x:Anchor>1, 0, 1, 0, 2, 0, 2, 0</x:Anchor></x:ClientData></v:shape></xml>`;

/** A sheet whose legacy VML holds a form control (no notes), plus a note on another sheet. */
async function withControlVml(): Promise<Uint8Array> {
  const wb = new Workbook();
  wb.addWorksheet('Data').addRow([1]);
  addNote(wb.addWorksheet('Notes'), 'A1', 'n');
  const src = writeXlsx(wb);
  const z = unzipSync(src);
  return patch(src, {
    'xl/drawings/vmlDrawing9.vml': VML_CTRL,
    'xl/worksheets/sheet1.xml': text(z, 'xl/worksheets/sheet1.xml').replace(
      '</worksheet>',
      '<legacyDrawing r:id="rId1"/></worksheet>',
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing9.vml"/></Relationships>`,
  });
}

describe('preserve: legacy VML with form controls', () => {
  it('kept with its <legacyDrawing>; shape-id blocks of kept VML are not reused by notes', async () => {
    const wb = readXlsx(await withControlVml());
    expect(wb.worksheets[0]!.comments).toBeUndefined();
    const out = writeXlsx(wb);
    const z = unzipSync(out);
    checkPackage(z);
    expect(text(z, 'xl/drawings/vmlDrawing9.vml')).toBe(VML_CTRL);
    const s = text(z, 'xl/worksheets/sheet1.xml');
    expect(s).toMatch(/<legacyDrawing r:id="rId\d+"\/>/);
    // the notes VML of sheet 2 must not use idmap block 1 (taken by the kept VML)
    const notesVml = Object.keys(z).find((p) => p.endsWith('.vml') && !p.endsWith('9.vml'))!;
    expect(text(z, notesVml)).not.toContain('data="1"');
    await exceljsLoads(out);
  });

  it('a notes-only VML is regenerated (not kept twice)', async () => {
    const wb = readXlsx(await withControlVml());
    const z = unzipSync(writeXlsx(wb));
    expect(Object.keys(z).filter((p) => p.endsWith('.vml')).length).toBe(2);
    expect(text(z, 'xl/worksheets/sheet2.xml').match(/<legacyDrawing /g)!.length).toBe(1);
  });
});

// ---------------------------------------------------------------- workbook / styles / VBA

async function withWorkbookBits(): Promise<Uint8Array> {
  const src = await base();
  const z = unzipSync(src);
  let wbx = text(z, 'xl/workbook.xml');
  wbx = wbx.replace(
    '</workbook>',
    '<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>',
  );
  wbx = wbx.replace(/<calcPr[^>]*\/>/, '<calcPr calcId="171027" calcMode="manual"/>');
  let styles = text(z, 'xl/styles.xml');
  const colors =
    '<colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="FF123456"/></indexedColors></colors>';
  styles = styles.replace(/(<extLst>|<\/styleSheet>)/, `${colors}$1`);
  return patch(src, { 'xl/workbook.xml': wbx, 'xl/styles.xml': styles });
}

describe('preserve: workbook.xml and styles.xml pieces', () => {
  it('workbook extLst, calcPr attributes and indexedColors kept', async () => {
    const wb = readXlsx(await withWorkbookBits());
    const out = writeXlsx(wb);
    const z = unzipSync(out);
    checkPackage(z);
    const w = text(z, 'xl/workbook.xml');
    expect(w).toContain('<x15:workbookPr chartTrackingRefBase="1"/>');
    expect(w).toContain('calcMode="manual"');
    expect(w.match(/<calcPr/g)!.length).toBe(1);
    expect(text(z, 'xl/styles.xml')).toContain('<rgbColor rgb="FF123456"/>');
    expect(text(z, 'xl/styles.xml')).toMatch(/<\/dxfs><tableStyles[^>]*\/><colors>/);
    await exceljsLoads(out);
  });

  it('macro-enabled workbook: vbaProject.bin and the macro main content type kept', async () => {
    const src = await base();
    const z = unzipSync(src);
    const vba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4]);
    const MACRO = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
    const ct = text(z, '[Content_Types].xml')
      .replace(
        /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet\.main\+xml/,
        MACRO,
      )
      .replace(
        '<Default Extension="xml"',
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"',
      );
    const xlsm = patch(src, {
      'xl/vbaProject.bin': vba,
      '[Content_Types].xml': ct,
      'xl/_rels/workbook.xml.rels': text(z, 'xl/_rels/workbook.xml.rels').replace(
        '</Relationships>',
        '<Relationship Id="rId77" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
      ),
    });
    const wb = readXlsx(xlsm);
    const out = writeXlsx(wb); // save as .xlsm
    const zo = unzipSync(out);
    checkPackage(zo);
    expect(byteEq(zo['xl/vbaProject.bin'], vba)).toBe(true);
    const oct = text(zo, '[Content_Types].xml');
    expect(oct).toContain(`<Override PartName="/xl/workbook.xml" ContentType="${MACRO}"/>`);
    expect(oct).toContain(
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
    );
    expect(oct.match(/PartName="\/xl\/workbook.xml"/g)!.length).toBe(1);
    expect(text(zo, 'xl/_rels/workbook.xml.rels')).toContain('Target="vbaProject.bin"');
    // dropped → plain xlsx main content type
    const plain = unzipSync(writeXlsx(readXlsx(xlsm, { preserveUnknown: false })));
    expect(text(plain, '[Content_Types].xml')).not.toContain('macroEnabled');
    expect(plain['xl/vbaProject.bin']).toBeUndefined();
  });
});
