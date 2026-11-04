// `./lite` entry: opt-in features (write / stream / streaming writer / read), the missing-feature
// error, round-trip preservation of everything a disabled reader leaves raw, no double emission
// when raw content and model data meet, and byte-identity with the root for the same features.

import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import * as root from '../src/index.js';
import * as lite from '../src/lite.js';
import { unzipSync, zipSync } from '../src/zip/index.js';
import { allFeatures, PNG } from './fixtures/all-features.js';

const dec = new TextDecoder();
const enc = new TextEncoder();

const ALL_W = [
  lite.commentsWriter,
  lite.conditionalFormattingWriter,
  lite.dataValidationWriter,
  lite.imagesWriter,
  lite.printSetupWriter,
  lite.sheetProtectionWriter,
  lite.tablesWriter,
  lite.themeWriter,
];
const ALL_R = [
  lite.commentsReader,
  lite.conditionalFormattingReader,
  lite.dataValidationReader,
  lite.imagesReader,
  lite.printSetupReader,
  lite.sheetProtectionReader,
  lite.tablesReader,
];

function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

/** Re-zip a package, transforming its XML parts with `f` (other parts are copied). */
function rezip(buf: Uint8Array, f: (path: string, xml: string) => string): Uint8Array {
  return zipSync(
    [...unzipSync(buf)].map(([path, data]) => ({
      path,
      data: /\.(xml|rels)$/.test(path) ? enc.encode(f(path, dec.decode(data))) : data,
    })),
    6,
  );
}

function count(s: string | undefined, re: RegExp): number {
  return (s ?? '').match(new RegExp(re, 'g'))?.length ?? 0;
}

async function readAll(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(rs).arrayBuffer());
}

const fixture = (): lite.Workbook => allFeatures(lite.Workbook, lite);

function plain(): lite.Workbook {
  const wb = new lite.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['a', 1, true]);
  ws.addRow(['b', 2, false]);
  return wb;
}

/** Every built-in feature's model content, as read by the root (full) reader. */
function expectAllFeatures(wb: lite.Workbook): void {
  const ws = wb.getWorksheet('Data')!;
  expect(ws.getCell('A1').value).toBe('plain');
  expect(ws.tables?.map((t) => [t.name, t.ref])).toEqual([['Sales', 'B2:C5']]);
  expect(ws.tables![0]!.headerStyle?.font).toMatchObject({
    bold: true,
    color: { argb: 'FF0000FF' },
  });
  expect(ws.comments?.map((n) => [n.ref, n.text, n.author])).toEqual([['A1', 'a note', 'Ann']]);
  expect(ws.images).toHaveLength(1);
  expect(wb.images?.[ws.images![0]!.imageId]?.data).toEqual(PNG);
  const cf = ws.conditionalFormatting!;
  expect(cf.map((b) => b.ref)).toEqual(['C3:C4']);
  expect(cf[0]!.rules.map((r) => r.type)).toEqual(['cellIs', 'dataBar']);
  expect((cf[0]!.rules[0] as lite.CfCellIsRule).style?.font).toMatchObject({
    bold: true,
    color: { argb: 'FF9C0006' },
  });
  expect(ws.dataValidations?.map((d) => [d.sqref, d.type])).toEqual([['A2', 'list']]);
  expect(ws.pageSetup).toMatchObject({
    orientation: 'landscape',
    paperSize: 9,
    margins: { left: 1, top: 1.5, header: 0.5 },
  });
  expect(ws.headerFooter?.oddHeader).toBe('&CReport');
  expect(ws.protection?.formatCells).toBe(true);
  expect(wb.getWorksheet('Two')!.getCell('B2').value).toBe(42);
}

// fflate stamps zip entries with the current time: freeze it for byte comparisons
beforeAll(() => {
  vi.useFakeTimers({ now: new Date(Date.UTC(2024, 5, 1, 12)), toFake: ['Date'] });
});
afterAll(() => {
  vi.useRealTimers();
});

describe('lite: no features by default', () => {
  it('writes plain data without any feature part (and no theme)', () => {
    const buf = lite.writeXlsx(plain());
    const p = parts(buf);
    expect([...p.keys()].some((k) => k.includes('theme'))).toBe(false);
    expect(p.get('xl/worksheets/sheet1.xml')).not.toContain('pageMargins');
    const back = root.readXlsx(buf);
    expect(back.getWorksheet('Data')!.getCell('B2').value).toBe(2);
  });

  it('themeWriter opts the default theme back in; theme: false still removes it', () => {
    const has = (b: Uint8Array): boolean => parts(b).has('xl/theme/theme1.xml');
    expect(has(lite.writeXlsx(plain(), { features: [lite.themeWriter] }))).toBe(true);
    expect(has(lite.writeXlsx(plain(), { features: [lite.themeWriter], theme: false }))).toBe(
      false,
    );
    expect(has(root.writeXlsx(plain()))).toBe(true);
    expect(has(root.writeXlsx(plain(), { theme: false }))).toBe(false);
  });
});

describe('lite: opt-in', () => {
  it('writeXlsx / xlsxStream / createXlsxWriter with the enabled writers', async () => {
    const wb = new lite.Workbook();
    const ws = wb.addWorksheet('Data');
    lite.addTable(ws, { name: 'T', columns: ['a', 'b'], rows: [[1, 2]] });
    lite.addNote(ws, 'D1', 'n');
    const features = [lite.tablesWriter, lite.commentsWriter];
    const bufs = [
      lite.writeXlsx(wb, { features }),
      await readAll(lite.xlsxStream(wb, { features })),
    ];
    const chunks: Uint8Array[] = [];
    const w = lite.createXlsxWriter({ features, sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('Data');
    s.addRow(['a', 'b']);
    s.addRow([1, 2]);
    s.worksheet.tables = [{ name: 'T', ref: 'A1:B2', columns: [{ name: 'a' }, { name: 'b' }] }];
    s.worksheet.comments = [{ ref: 'D1', text: 'n' }];
    await s.commit();
    await w.end();
    bufs.push(new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer()));
    for (const buf of bufs) {
      const back = root.readXlsx(buf).getWorksheet('Data')!;
      expect(back.tables?.map((t) => t.name)).toEqual(['T']);
      expect(back.comments?.map((n) => n.text)).toEqual(['n']);
    }
  });

  it('readXlsx models only the enabled readers', () => {
    const buf = root.writeXlsx(fixture());
    const wb = lite.readXlsx(buf, { features: [lite.tablesReader] });
    const ws = wb.getWorksheet('Data')!;
    expect(ws.tables?.map((t) => t.name)).toEqual(['Sales']);
    expect(ws.comments).toBeUndefined();
    expect(ws.images).toBeUndefined();
    expect(ws.conditionalFormatting).toBeUndefined();
    expect(ws.dataValidations).toBeUndefined();
    expect(ws.protection).toBeUndefined();
    expect(ws.pageSetup).toBeUndefined();
    expect(lite.readXlsx(buf).getWorksheet('Data')!.tables).toBeUndefined();
  });

  it('the root accepts extra (custom) features and ignores duplicates of built-ins', () => {
    const custom: lite.FeatureWriter = {
      name: 'custom',
      book: { write: (c) => c.addPart('xl/custom.xml', '<c/>', 'application/x-c+xml') },
    };
    const buf = root.writeXlsx(fixture(), { features: [custom, lite.tablesWriter] });
    expect(parts(buf).get('xl/custom.xml')).toBe('<c/>');
    expect(count(parts(buf).get('xl/worksheets/sheet1.xml'), /<tableParts/)).toBe(1);
  });
});

describe('lite: missing feature', () => {
  const cases: [string, (ws: lite.Worksheet) => void, string, string][] = [
    [
      'tables',
      (ws) => lite.addTable(ws, { name: 'T', columns: ['a'], rows: [[1]] }),
      'tables',
      'tablesWriter',
    ],
    ['comments', (ws) => lite.addNote(ws, 'A1', 'x'), 'comments', 'commentsWriter'],
    [
      'images',
      (ws) =>
        lite.placeImage(ws, lite.addImage(ws.workbook, { data: PNG, extension: 'png' }), 'A1'),
      'images',
      'imagesWriter',
    ],
    [
      'backgroundImage',
      (ws) =>
        void (ws.backgroundImage = lite.addImage(ws.workbook, { data: PNG, extension: 'png' })),
      'images',
      'imagesWriter',
    ],
    [
      'conditionalFormatting',
      (ws) => lite.addConditionalFormatting(ws, 'A1', [{ type: 'top10', style: {} }]),
      'conditional-formatting',
      'conditionalFormattingWriter',
    ],
    [
      'dataValidations',
      (ws) =>
        lite.addDataValidation(ws, { sqref: 'A1', type: 'whole', operator: 'equal', formula1: 1 }),
      'data-validation',
      'dataValidationWriter',
    ],
    ['protection', (ws) => void (ws.protection = {}), 'sheet-protection', 'sheetProtectionWriter'],
    [
      'pageSetup',
      (ws) => void (ws.pageSetup = { orientation: 'landscape' }),
      'print-setup',
      'printSetupWriter',
    ],
    [
      'headerFooter',
      (ws) => void (ws.headerFooter = { oddHeader: 'h' }),
      'print-setup',
      'printSetupWriter',
    ],
  ];
  for (const [field, set, name, ex] of cases) {
    it(`ws.${field} without the '${name}' feature throws`, () => {
      const wb = plain();
      set(wb.getWorksheet('Data')!);
      const msg = `Worksheet 'Data': ws.${field} is set but the '${name}' feature is not enabled; pass features: [${ex}]`;
      expect(() => lite.writeXlsx(wb)).toThrow(msg);
      const feature = ALL_W.find((w) => w.name === name)!;
      expect(() => lite.writeXlsx(wb, { features: [feature] })).not.toThrow();
    });
  }

  it('empty lists are not data', () => {
    const wb = plain();
    const ws = wb.getWorksheet('Data')!;
    ws.tables = [];
    ws.comments = [];
    ws.images = [];
    expect(() => lite.writeXlsx(wb)).not.toThrow();
  });

  it('xlsxStream and the streaming writer throw too (fields set after rows: at commit)', async () => {
    const wb = plain();
    wb.getWorksheet('Data')!.protection = {};
    await expect(readAll(lite.xlsxStream(wb))).rejects.toThrow(/'sheet-protection' feature/);
    const w = lite.createXlsxWriter({ sink: () => {} });
    const s = w.addWorksheet('S');
    s.addRow([1]);
    s.worksheet.tables = [{ name: 'T', ref: 'A1:A2', columns: [{ name: 'a' }] }];
    await expect(s.commit()).rejects.toThrow(/ws\.tables is set/);
  });

  it('a workbook read with every feature cannot be written without them', () => {
    const wb = root.readXlsx(root.writeXlsx(fixture()));
    expect(() => lite.writeXlsx(wb)).toThrow(/feature is not enabled/);
    expect(() => lite.writeXlsx(wb, { features: ALL_W })).not.toThrow();
  });

  it('kept chart anchors merged by the images feature require it', () => {
    // a drawing with a picture (modelled) and a non-picture anchor (kept in ws.preserved.drawing)
    const shape =
      '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="99" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>';
    const buf = rezip(root.writeXlsx(fixture()), (k, v) =>
      k === 'xl/drawings/drawing1.xml' ? v.replace('</xdr:wsDr>', `${shape}</xdr:wsDr>`) : v,
    );
    const wb = root.readXlsx(buf);
    const ws = wb.getWorksheet('Data')!;
    expect(ws.preserved?.drawing).toBeDefined();
    ws.images = undefined;
    expect(() =>
      lite.writeXlsx(wb, { features: ALL_W.filter((f) => f.name !== 'images') }),
    ).toThrow(
      /ws\.preserved is set but the 'images' feature is not enabled; pass features: \[imagesWriter\]/,
    );
    const out = parts(lite.writeXlsx(wb, { features: [lite.imagesWriter, ...ALL_W] }));
    expect(out.get('xl/drawings/drawing1.xml')).toContain('name="S"');
  });
});

describe('lite: round-trip preservation without features', () => {
  const full = (): Uint8Array => root.writeXlsx(fixture());

  it('lite read -> lite write keeps every feature raw (full reader sees it all)', () => {
    const wb = lite.readXlsx(full());
    const ws = wb.getWorksheet('Data')!;
    expect(ws.tables ?? ws.comments ?? ws.images ?? ws.protection).toBeUndefined();
    const out = lite.writeXlsx(wb);
    expectAllFeatures(root.readXlsx(out));
    // the source theme is kept raw (lite writes no default one)
    expect(parts(out).has('xl/theme/theme1.xml')).toBe(true);
  });

  it('lite read -> root write: raw content is written once (features have no model data)', () => {
    const out = root.writeXlsx(lite.readXlsx(full()));
    expectAllFeatures(root.readXlsx(out));
    const p = parts(out);
    const sheet = p.get('xl/worksheets/sheet1.xml')!;
    for (const el of [
      /<tableParts/,
      /<legacyDrawing /,
      /<drawing /,
      /<conditionalFormatting /,
      /<dataValidations /,
      /<sheetProtection /,
      /<pageSetup /,
      /<headerFooter/,
      /78C0D931-6437-407d-A8EE-F0AAD7539E65/i,
    ])
      expect(count(sheet, el), String(el)).toBe(1);
    const rels = p.get('xl/worksheets/_rels/sheet1.xml.rels')!;
    expect(count(rels, /\/comments"/)).toBe(1);
    expect(count(rels, /\/table"/)).toBe(1);
    expect(count(p.get('xl/_rels/workbook.xml.rels'), /\/theme"/)).toBe(1);
    expect([...p.keys()].filter((k) => /theme/.test(k))).toHaveLength(1);
  });

  it('model data added for a feature replaces the raw element (no duplicates)', () => {
    const wb = lite.readXlsx(full());
    const ws = wb.getWorksheet('Data')!;
    lite.addNote(ws, 'H1', 'new note');
    lite.addConditionalFormatting(ws, 'A1', [
      { type: 'dataBar', color: { argb: 'FF00FF00' }, gradient: false },
    ]);
    ws.pageSetup = { orientation: 'portrait', fitToPage: true };
    const p = parts(root.writeXlsx(wb));
    const sheet = p.get('xl/worksheets/sheet1.xml')!;
    expect(count(sheet, /<legacyDrawing /)).toBe(1);
    expect(count(sheet, /<pageSetup /)).toBe(1);
    expect(count(sheet, /<pageSetUpPr/)).toBe(1);
    expect(count(sheet, /78C0D931-6437-407d-A8EE-F0AAD7539E65/i)).toBe(1); // x14 CF ext
    expect(count(p.get('xl/worksheets/_rels/sheet1.xml.rels'), /\/comments"/)).toBe(1);
    expect([...p.keys()].filter((k) => /comments\d*\.xml$/.test(k))).toHaveLength(1);
    const back = root.readXlsx(root.writeXlsx(wb)).getWorksheet('Data')!;
    expect(back.comments?.map((n) => n.text)).toEqual(['new note']);
    expect(back.tables?.map((t) => t.name)).toEqual(['Sales']); // still raw, still there
  });

  it('kept cfRule / table dxfIds follow the registry when the reader merges dxfs', () => {
    // prepend a duplicate of dxf 0 and shift every dxfId: file dxfs 0 and 1 are identical, so
    // the registry maps 0,1 -> 0 and 2.. -> 1..
    const shifted = rezip(full(), (k, x) => {
      if (k === 'xl/styles.xml') {
        const m = /<dxfs count="(\d+)">(<dxf>.*?<\/dxf>)/.exec(x)!;
        return x.replace(m[0], `<dxfs count="${+m[1]! + 1}">${m[2]}${m[2]}`);
      }
      if (/^xl\/(worksheets\/sheet\d|tables\/table\d)/.test(k))
        return x.replace(/([dD]xfId)="(\d+)"/g, (_, a: string, n: string) => `${a}="${+n + 1}"`);
      return x;
    });
    expectAllFeatures(root.readXlsx(shifted));
    expectAllFeatures(root.readXlsx(lite.writeXlsx(lite.readXlsx(shifted))));
  });

  it('preserveUnknown: false drops what no enabled reader models', () => {
    const back = root.readXlsx(lite.writeXlsx(lite.readXlsx(full(), { preserveUnknown: false })));
    const ws = back.getWorksheet('Data')!;
    expect(ws.getCell('A1').value).toBe('plain');
    expect(ws.tables ?? ws.comments ?? ws.conditionalFormatting).toBeUndefined();
  });
});

describe('lite vs root', () => {
  it('identical bytes for the same feature set (writeXlsx, xlsxStream, round trip)', async () => {
    expect(lite.writeXlsx(fixture(), { features: ALL_W })).toEqual(root.writeXlsx(fixture()));
    expect(
      lite.writeXlsx(fixture(), { features: ALL_W, profile: 'low-memory', theme: false }),
    ).toEqual(root.writeXlsx(fixture(), { profile: 'low-memory', theme: false }));
    expect(await readAll(lite.xlsxStream(fixture(), { features: ALL_W }))).toEqual(
      await readAll(root.xlsxStream(fixture())),
    );
    const buf = root.writeXlsx(fixture());
    expect(lite.writeXlsx(lite.readXlsx(buf, { features: ALL_R }), { features: ALL_W })).toEqual(
      root.writeXlsx(root.readXlsx(buf)),
    );
  });

  it('identical bytes from the streaming writers', async () => {
    const run = async (w: lite.XlsxWriter, chunks: Uint8Array[]): Promise<Uint8Array> => {
      const s = w.addWorksheet('S');
      s.worksheet.pageSetup = { orientation: 'landscape' };
      lite.addNote(s.worksheet, 'A1', 'n');
      s.addRow([1, 'a']);
      await s.commit();
      await w.end();
      return new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    };
    const a: Uint8Array[] = [];
    const b: Uint8Array[] = [];
    expect(
      await run(lite.createXlsxWriter({ features: ALL_W, sink: (c) => void a.push(c) }), a),
    ).toEqual(await run(root.createXlsxWriter({ sink: (c) => void b.push(c) }), b));
  });

  it('root exports the same names as lite', () => {
    expect(Object.keys(root).sort()).toEqual(Object.keys(lite).sort());
  });
});
