/// <reference types="node" />
// Conditional formatting (features/conditional-formatting): ours ↔ ours in every writer,
// ours → exceljs, exceljs → ours, x14 data bar / icon set extension, priorities, generated
// text formulas and element order.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import {
  addConditionalFormatting,
  addDataValidation,
  type CfRule,
  type ConditionalFormatting,
} from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const sheetXml = (buf: Uint8Array): string =>
  dec.decode(unzipSync(buf).get('xl/worksheets/sheet1.xml')!);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

type XRule = Record<string, unknown>;
type XCf = { ref: string; rules: XRule[] };
const xcf = (ws: ExcelJS.Worksheet): XCf[] =>
  (ws as unknown as { conditionalFormattings: XCf[] }).conditionalFormattings;

const RED = { font: { bold: true, color: { argb: 'FF9C0006' } } };
const FILL = {
  fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
} as const;

const BLOCKS: ConditionalFormatting[] = [
  {
    ref: 'A1:A10',
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: RED },
      { type: 'cellIs', operator: 'between', formulae: ['$B$1', '=$B$2'], style: FILL },
      { type: 'expression', formulae: ['MOD(ROW(),2)=0'], style: FILL, stopIfTrue: true },
      { type: 'top10', rank: 3, percent: true, bottom: true, style: RED },
      { type: 'aboveAverage', aboveAverage: false, equalAverage: true, style: RED },
      { type: 'duplicateValues', style: FILL },
      { type: 'uniqueValues', style: RED },
    ],
  },
  {
    ref: 'B2:C5 E1:E3',
    rules: [
      { type: 'containsText', text: 'a "b"', style: RED },
      { type: 'notContainsText', text: 'x', style: RED },
      { type: 'beginsWith', text: 'pre', style: FILL },
      { type: 'endsWith', text: 'post', style: FILL },
      { type: 'containsBlanks', style: RED },
      { type: 'notContainsBlanks', style: RED },
      { type: 'containsErrors', style: RED },
      { type: 'notContainsErrors', style: RED, priority: 50 },
      { type: 'timePeriod', timePeriod: 'last7Days', style: FILL },
      { type: 'containsText', text: 'z', formulae: ['ISNUMBER(SEARCH("z",B2))'], style: FILL },
    ],
  },
  {
    ref: 'D1:D10',
    rules: [
      { type: 'colorScale', colors: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }] },
      {
        type: 'colorScale',
        cfvo: [{ type: 'num', value: 0 }, { type: 'percentile', value: 50 }, { type: 'max' }],
        colors: [{ argb: 'FFF8696B' }, { theme: 4, tint: 0.5 }, { argb: 'FF63BE7B' }],
      },
      { type: 'dataBar', color: { argb: 'FF638EC6' }, showValue: false },
      {
        type: 'dataBar',
        cfvo: [{ type: 'autoMin' }, { type: 'num', value: 100 }],
        color: { argb: 'FF00B050' },
        minLength: 0,
        maxLength: 100,
        gradient: false,
        border: true,
        borderColor: { argb: 'FF008000' },
        negativeFillColor: { argb: 'FFFF0000' },
        axisColor: { argb: 'FF000000' },
        axisPosition: 'middle',
        direction: 'rightToLeft',
      },
      {
        type: 'iconSet',
        iconSet: '3Arrows',
        cfvo: [
          { type: 'percent', value: 0 },
          { type: 'percent', value: 33 },
          { type: 'percent', value: 67, gte: false },
        ],
        reverse: true,
        showValue: false,
      },
      { type: 'iconSet', iconSet: '5Rating' },
      { type: 'iconSet', iconSet: '3Stars', showValue: false },
    ],
  },
];

/** What our reader returns: auto priorities filled in, generated formulas dropped, defaults. */
function expected(): ConditionalFormatting[] {
  let p = 50;
  const out = BLOCKS.map((b) => ({
    ref: b.ref,
    rules: b.rules.map((r) => {
      const x = { ...r, priority: r.priority ?? ++p } as Record<string, unknown>;
      if (x.type === 'cellIs') {
        x.formulae = (x.formulae as (string | number)[]).map((f) => String(f).replace(/^=/, ''));
      }
      if (x.type === 'top10') x.rank ??= 10;
      if (x.type === 'dataBar' || x.type === 'colorScale')
        x.cfvo ??= [{ type: 'min' }, { type: 'max' }];
      if (x.type === 'iconSet' && !x.cfvo) {
        const n = +String(x.iconSet)[0]!;
        x.cfvo = Array.from({ length: n }, (_, i) => ({
          type: 'percent',
          value: Math.round((i * 100) / n),
        }));
      }
      return x;
    }),
  }));
  return out as unknown as ConditionalFormatting[];
}

function build(wb: Workbook): void {
  const ws = wb.addWorksheet('CF');
  ws.setCell('A1', 1);
  for (const b of BLOCKS) addConditionalFormatting(ws, b.ref, b.rules);
  addDataValidation(ws, { sqref: 'A1', type: 'whole', operator: 'greaterThan', formula1: 0 });
  ws.setCell('F1', { text: 'link', hyperlink: 'https://example.com' });
}

async function allWriters(): Promise<[string, Uint8Array][]> {
  const wb = new Workbook();
  build(wb);
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  const s = w.addWorksheet('CF');
  s.addRow([1]);
  s.worksheet.conditionalFormatting = BLOCKS;
  s.worksheet.dataValidations = wb.getWorksheet('CF')!.dataValidations;
  await w.end();
  return [
    ['writeXlsx', writeXlsx(wb)],
    ['xlsxStream', new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())],
    ['createXlsxWriter', new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())],
  ];
}

/** Our read-back rules with styles reduced to what we wrote (dxfs are normalized). */
function strip(cfs: ConditionalFormatting[] | undefined): unknown {
  return cfs?.map((b) => ({
    ref: b.ref,
    rules: b.rules.map((r) => {
      const x = { ...r } as Record<string, unknown>;
      if (x.style) x.style = (x.style as { font?: unknown }).font ? 'RED' : 'FILL';
      return x;
    }),
  }));
}

describe('conditional formatting', () => {
  it('round-trips through ours in every writer', async () => {
    const want = strip(expected());
    for (const [name, buf] of await allWriters()) {
      const ws = readXlsx(buf).getWorksheet('CF')!;
      expect(strip(ws.conditionalFormatting), name).toEqual(want);
      const st = ws.conditionalFormatting![0]!.rules[0]! as { style: unknown };
      expect(st.style, name).toMatchObject(RED);
      const fill = ws.conditionalFormatting![0]!.rules[1]! as { style: unknown };
      expect(fill.style, name).toMatchObject(FILL);
    }
  });

  it('writes Excel-shaped XML in CT_Worksheet order', async () => {
    for (const [name, buf] of await allWriters()) {
      const xml = sheetXml(buf);
      const cf = xml.indexOf('<conditionalFormatting ');
      const dv = xml.indexOf('<dataValidations');
      const ext = xml.indexOf('<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"');
      expect(cf, name).toBeGreaterThan(xml.indexOf('</sheetData>'));
      expect(cf, name).toBeLessThan(dv);
      if (name !== 'createXlsxWriter') expect(dv, name).toBeLessThan(xml.indexOf('<hyperlinks>'));
      expect(dv, name).toBeLessThan(xml.indexOf('<pageMargins'));
      expect(ext, name).toBeGreaterThan(xml.indexOf('<pageMargins'));
      // generated formulas: top-left of the first range, quotes doubled
      expect(xml, name).toContain(
        '<cfRule type="containsText" dxfId="0" priority="58" operator="containsText" text="a &quot;b&quot;"><formula>NOT(ISERROR(SEARCH("a ""b""",B2)))</formula></cfRule>',
      );
      expect(xml, name).toContain('<formula>LEFT(B2,LEN("pre"))="pre"</formula>');
      expect(xml, name).toContain('<formula>RIGHT(B2,LEN("post"))="post"</formula>');
      expect(xml, name).toContain('<formula>ISERROR(SEARCH("x",B2))</formula>');
      expect(xml, name).toContain('<formula>LEN(TRIM(B2))=0</formula>');
      expect(xml, name).toContain(
        '<formula>AND(TODAY()-FLOOR(B2,1)&lt;=6,FLOOR(B2,1)&lt;=TODAY())</formula>',
      );
      // x14 data bar linked by id, x14-only icon set only in the extLst
      const id = /<x14:id>(\{[0-9A-F-]+\})<\/x14:id>/.exec(xml)![1]!;
      expect(xml, name).toContain(`<x14:cfRule type="dataBar" id="${id}"><x14:dataBar`);
      expect(xml, name).toContain(
        '<x14:cfvo type="autoMin"/><x14:cfvo type="num"><xm:f>100</xm:f>',
      );
      expect(xml, name).toMatch(
        /<x14:cfRule type="iconSet" priority="\d+" id="[^"]+"><x14:iconSet iconSet="3Stars"/,
      );
      expect(xml.match(/iconSet="3Stars"/g)!.length, name).toBe(1);
      expect(xml, name).toContain('<xm:sqref>D1:D10</xm:sqref>');
      // priorities are unique
      const pr = [...xml.matchAll(/priority="(\d+)"/g)].map((m) => m[1]);
      expect(new Set(pr).size, name).toBe(pr.length);
    }
  });

  it('is read by exceljs with the same semantics', async () => {
    for (const [name, buf] of await allWriters()) {
      const cfs = xcf((await oracle(buf)).getWorksheet('CF')!);
      expect(
        cfs.map((c) => c.ref),
        name,
      ).toEqual(['A1:A10', 'B2:C5 E1:E3', 'D1:D10']);
      const [a, b, d] = cfs as [XCf, XCf, XCf];
      expect(a.rules[0], name).toMatchObject({
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['100'],
        priority: 51,
        style: RED,
      });
      expect(a.rules[1], name).toMatchObject({ formulae: ['$B$1', '$B$2'], style: FILL });
      expect(a.rules[2], name).toMatchObject({ type: 'expression', formulae: ['MOD(ROW(),2)=0'] });
      expect(a.rules[3], name).toMatchObject({
        type: 'top10',
        rank: 3,
        percent: true,
        bottom: true,
      });
      expect(a.rules[4], name).toMatchObject({ type: 'aboveAverage', aboveAverage: false });
      expect(a.rules[5], name).toMatchObject({ type: 'duplicateValues', style: FILL });
      expect(b.rules[0], name).toMatchObject({
        type: 'containsText',
        operator: 'containsText',
        formulae: ['NOT(ISERROR(SEARCH("a ""b""",B2)))'],
        style: RED,
      });
      expect(b.rules[4], name).toMatchObject({ type: 'containsText', operator: 'containsBlanks' });
      expect(b.rules[7], name).toMatchObject({ operator: 'notContainsErrors', priority: 50 });
      expect(b.rules[8], name).toMatchObject({ type: 'timePeriod', timePeriod: 'last7Days' });
      expect(d.rules[0], name).toMatchObject({
        type: 'colorScale',
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }],
      });
      expect(d.rules[1], name).toMatchObject({
        cfvo: [{ type: 'num', value: 0 }, { type: 'percentile', value: 50 }, { type: 'max' }],
        color: [{ argb: 'FFF8696B' }, { theme: 4, tint: 0.5 }, { argb: 'FF63BE7B' }],
      });
      expect(d.rules[2], name).toMatchObject({ type: 'dataBar', color: { argb: 'FF638EC6' } });
      // exceljs merges the x14 twin into the classic rule
      expect(d.rules[3], name).toMatchObject({
        type: 'dataBar',
        color: { argb: 'FF00B050' },
        gradient: false,
        border: true,
        axisPosition: 'middle',
        direction: 'rightToLeft',
        negativeFillColor: { argb: 'FFFF0000' },
        axisColor: { argb: 'FF000000' },
      });
      expect(d.rules[4], name).toMatchObject({
        type: 'iconSet',
        iconSet: '3Arrows',
        reverse: true,
        showValue: false,
        cfvo: [
          { type: 'percent', value: 0 },
          { type: 'percent', value: 33 },
          { type: 'percent', value: 67 },
        ],
      });
      expect(d.rules[5], name).toMatchObject({ type: 'iconSet', iconSet: '5Rating' });
    }
  });

  it('reads conditional formatting written by exceljs', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('X');
    xs.getCell('A1').value = 1;
    xs.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [
        {
          type: 'cellIs',
          operator: 'lessThan',
          formulae: [5],
          priority: 1,
          style: { font: { italic: true, color: { argb: 'FF0000FF' } } },
        },
        {
          type: 'expression',
          formulae: ['$A1>2'],
          priority: 2,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } } },
        },
        {
          type: 'top10',
          rank: 2,
          percent: false,
          bottom: false,
          priority: 3,
          style: { font: { bold: true } },
        },
        {
          type: 'containsText',
          operator: 'containsText',
          text: 'hi',
          priority: 4,
          style: { font: { bold: true } },
        } as ExcelJS.ConditionalFormattingRule,
        { type: 'timePeriod', timePeriod: 'today', priority: 5, style: { font: { bold: true } } },
      ],
    });
    xs.addConditionalFormatting({
      ref: 'B1:B5',
      rules: [
        {
          type: 'colorScale',
          priority: 6,
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: [{ argb: 'FFFF0000' }, { argb: 'FF00FF00' }],
        },
        {
          type: 'iconSet',
          priority: 7,
          iconSet: '4Arrows',
          reverse: true,
          cfvo: [
            { type: 'percent', value: 0 },
            { type: 'percent', value: 25 },
            { type: 'percent', value: 50 },
            { type: 'percent', value: 75 },
          ],
        },
        {
          type: 'dataBar',
          priority: 8,
          gradient: false,
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: { argb: 'FF638EC6' },
        } as ExcelJS.ConditionalFormattingRule,
      ],
    });
    const ws = readXlsx(new Uint8Array(await x.xlsx.writeBuffer())).getWorksheet('X')!;
    const cfs = ws.conditionalFormatting!;
    expect(cfs.map((c) => c.ref)).toEqual(['A1:A5', 'B1:B5']);
    const [a, b] = cfs as [ConditionalFormatting, ConditionalFormatting];
    expect(a.rules[0]).toMatchObject({
      type: 'cellIs',
      operator: 'lessThan',
      formulae: ['5'],
      priority: 1,
      style: { font: { italic: true, color: { argb: 'FF0000FF' } } },
    });
    expect(a.rules[1]).toMatchObject({
      type: 'expression',
      formulae: ['$A1>2'],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } } },
    });
    expect(a.rules[2]).toEqual({
      type: 'top10',
      rank: 2,
      priority: 3,
      style: expect.objectContaining({ font: expect.objectContaining({ bold: true }) }),
    });
    // exceljs omits the text attribute: the formula is kept
    expect(a.rules[3]).toMatchObject({
      type: 'containsText',
      text: '',
      formulae: ['NOT(ISERROR(SEARCH("hi",A1)))'],
    });
    // exceljs's time-period formula equals ours → dropped
    expect(a.rules[4]).toEqual({
      type: 'timePeriod',
      timePeriod: 'today',
      priority: 5,
      style: expect.anything(),
    });
    expect(b.rules[0]).toEqual({
      type: 'colorScale',
      priority: 6,
      cfvo: [{ type: 'min' }, { type: 'max' }],
      colors: [{ argb: 'FFFF0000' }, { argb: 'FF00FF00' }],
    });
    expect(b.rules[1]).toEqual({
      type: 'iconSet',
      priority: 7,
      iconSet: '4Arrows',
      reverse: true,
      cfvo: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 25 },
        { type: 'percent', value: 50 },
        { type: 'percent', value: 75 },
      ],
    });
    expect(b.rules[2]).toMatchObject({
      type: 'dataBar',
      priority: 8,
      color: { argb: 'FF638EC6' },
    });
  });

  it('reads exceljs x14 data bars and x14-only icon sets', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('X');
    xs.getCell('A1').value = 1;
    xs.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [
        {
          type: 'dataBar',
          priority: 1,
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: { argb: 'FF638EC6' },
          border: true,
          negativeFillColor: { argb: 'FFFF0000' },
          axisPosition: 'middle',
        } as unknown as ExcelJS.ConditionalFormattingRule,
        {
          type: 'iconSet',
          priority: 2,
          iconSet: '3Triangles',
          cfvo: [
            { type: 'percent', value: 0 },
            { type: 'percent', value: 33 },
            { type: 'percent', value: 67 },
          ],
        } as unknown as ExcelJS.ConditionalFormattingRule,
      ],
    });
    const ws = readXlsx(new Uint8Array(await x.xlsx.writeBuffer())).getWorksheet('X')!;
    const rules = ws.conditionalFormatting!.flatMap((c) => c.rules);
    expect(rules.find((r) => r.type === 'dataBar')).toMatchObject({
      color: { argb: 'FF638EC6' },
      border: true,
      negativeFillColor: { argb: 'FFFF0000' },
      axisPosition: 'middle',
    });
    expect(rules.find((r) => r.type === 'iconSet')).toMatchObject({
      iconSet: '3Triangles',
      priority: 2,
      cfvo: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
    });
  });

  it('assigns unique priorities without mutating the model', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const rules: CfRule[] = [
      { type: 'duplicateValues', style: RED },
      { type: 'uniqueValues', style: RED, priority: 3 },
    ];
    addConditionalFormatting(ws, 'A1:A3', rules);
    addConditionalFormatting(ws, 'B1:B3', [{ type: 'top10', style: RED }]);
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!.conditionalFormatting!;
    expect(back.flatMap((c) => c.rules.map((r) => r.priority))).toEqual([4, 3, 5]);
    expect(rules[0]!.priority).toBeUndefined();
    // styles are deduped into one dxf
    expect(wb.styles.dxfs.length).toBe(1);
  });

  it('writes nothing for an empty model', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.conditionalFormatting = [];
    const xml = sheetXml(writeXlsx(wb));
    expect(xml).not.toContain('conditionalFormatting');
    expect(xml).not.toContain('extLst');
    expect(readXlsx(writeXlsx(wb)).getWorksheet('S')!.conditionalFormatting).toBeUndefined();
  });
});
