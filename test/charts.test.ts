/// <reference types="node" />
// Charts (features/charts, write only): chart parts per type checked against the ECMA-376
// dml-chart schema sequences (calibrated on the chart Excel wrote in the chart-sheet fixture),
// the shared drawing with pictures and kept anchors, rels / content types, caches, the streaming
// writer, lite opt-in, exceljs compatibility and raw round-trip through our reader.

import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { createXlsxWriter, listPreservedParts, writeXlsx, xlsxStream } from '../src/index.js';
import * as lite from '../src/lite.js';
import {
  addChart,
  addImage,
  type Chart,
  type ChartType,
  placeImage,
} from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { localName, parseXml } from '../src/xml/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

const count = (s: string | undefined, re: RegExp): number =>
  (s ?? '').match(new RegExp(re, 'g'))?.length ?? 0;

// ---------------------------------------------------------------- schema order (dml-chart.xsd)

const SER_TAIL = ['dPt', 'dLbls', 'trendline', 'errBars'];
const AX_HEAD = [
  'axId',
  'scaling',
  'delete',
  'axPos',
  'majorGridlines',
  'minorGridlines',
  'title',
  'numFmt',
  'majorTickMark',
  'minorTickMark',
  'tickLblPos',
  'spPr',
  'txPr',
  'crossAx',
  'crosses|crossesAt',
];
/** Child sequence per element (`a|b` = same position; repeats allowed). Key: name or parent/name. */
const SEQ: Record<string, string[]> = {
  chartSpace: [
    'date1904',
    'lang',
    'roundedCorners',
    'AlternateContent|style',
    'clrMapOvr',
    'pivotSource',
    'protection',
    'chart',
    'spPr',
    'txPr',
    'externalData',
    'printSettings',
    'userShapes',
    'extLst',
  ],
  'chartSpace/chart': [
    'title',
    'autoTitleDeleted',
    'pivotFmts',
    'view3D',
    'floor',
    'sideWall',
    'backWall',
    'plotArea',
    'legend',
    'plotVisOnly',
    'dispBlanksAs',
    'showDLblsOverMax',
    'extLst',
  ],
  title: ['tx', 'layout', 'overlay', 'spPr', 'txPr', 'extLst'],
  plotArea: [
    'layout',
    'areaChart|area3DChart|lineChart|line3DChart|stockChart|radarChart|scatterChart|pieChart|pie3DChart|doughnutChart|barChart|bar3DChart|ofPieChart|surfaceChart|surface3DChart|bubbleChart',
    'valAx|catAx|dateAx|serAx',
    'dTable',
    'spPr',
    'extLst',
  ],
  barChart: [
    'barDir',
    'grouping',
    'varyColors',
    'ser',
    'dLbls',
    'gapWidth',
    'overlap',
    'serLines',
    'axId',
    'extLst',
  ],
  lineChart: [
    'grouping',
    'varyColors',
    'ser',
    'dLbls',
    'dropLines',
    'hiLowLines',
    'upDownBars',
    'marker',
    'smooth',
    'axId',
    'extLst',
  ],
  pieChart: ['varyColors', 'ser', 'dLbls', 'firstSliceAng', 'extLst'],
  doughnutChart: ['varyColors', 'ser', 'dLbls', 'firstSliceAng', 'holeSize', 'extLst'],
  scatterChart: ['scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId', 'extLst'],
  areaChart: ['grouping', 'varyColors', 'ser', 'dLbls', 'dropLines', 'axId', 'extLst'],
  'barChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'invertIfNegative',
    'pictureOptions',
    ...SER_TAIL,
    'cat',
    'val',
    'shape',
    'extLst',
  ],
  'lineChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'marker',
    ...SER_TAIL,
    'cat',
    'val',
    'smooth',
    'extLst',
  ],
  'pieChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'explosion',
    'dPt',
    'dLbls',
    'cat',
    'val',
    'extLst',
  ],
  'doughnutChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'explosion',
    'dPt',
    'dLbls',
    'cat',
    'val',
    'extLst',
  ],
  'scatterChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'marker',
    ...SER_TAIL,
    'xVal',
    'yVal',
    'smooth',
    'extLst',
  ],
  'areaChart/ser': [
    'idx',
    'order',
    'tx',
    'spPr',
    'pictureOptions',
    ...SER_TAIL,
    'cat',
    'val',
    'extLst',
  ],
  catAx: [
    ...AX_HEAD,
    'auto',
    'lblAlgn',
    'lblOffset',
    'tickLblSkip',
    'tickMarkSkip',
    'noMultiLvlLbl',
    'extLst',
  ],
  valAx: [...AX_HEAD, 'crossBetween', 'majorUnit', 'minorUnit', 'dispUnits', 'extLst'],
  scaling: ['logBase', 'orientation', 'max', 'min', 'extLst'],
  legend: ['legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr', 'extLst'],
  'ser/dLbls': [
    'dLbl',
    'delete',
    'numFmt',
    'spPr',
    'txPr',
    'dLblPos',
    'showLegendKey',
    'showVal',
    'showCatName',
    'showSerName',
    'showPercent',
    'showBubbleSize',
    'separator',
    'showLeaderLines',
    'leaderLines',
    'extLst',
  ],
  'ser/marker': ['symbol', 'size', 'spPr', 'extLst'],
  'ser/trendline': [
    'name',
    'spPr',
    'trendlineType',
    'order',
    'period',
    'forward',
    'backward',
    'intercept',
    'dispRSqr',
    'dispEq',
    'trendlineLbl',
    'extLst',
  ],
  strRef: ['f', 'strCache', 'extLst'],
  numRef: ['f', 'numCache', 'extLst'],
  numCache: ['formatCode', 'ptCount', 'pt', 'extLst'],
  strCache: ['ptCount', 'pt', 'extLst'],
  // drawing
  twoCellAnchor: ['from', 'to', 'sp|grpSp|graphicFrame|cxnSp|pic|contentPart', 'clientData'],
  oneCellAnchor: ['from', 'ext', 'sp|grpSp|graphicFrame|cxnSp|pic|contentPart', 'clientData'],
  absoluteAnchor: ['pos', 'ext', 'sp|grpSp|graphicFrame|cxnSp|pic|contentPart', 'clientData'],
  graphicFrame: ['nvGraphicFramePr', 'xfrm', 'graphic', 'extLst'],
};
const REQUIRED: Record<string, string[]> = {
  chartSpace: ['chart'],
  'chartSpace/chart': ['plotArea'],
  plotArea: ['layout'],
  barChart: ['barDir', 'ser', 'axId'],
  lineChart: ['grouping', 'ser', 'axId'],
  scatterChart: ['scatterStyle', 'ser', 'axId'],
  areaChart: ['ser', 'axId'],
  pieChart: ['ser'],
  doughnutChart: ['ser'],
  ser: ['idx', 'order'],
  catAx: ['axId', 'scaling', 'axPos', 'crossAx'],
  valAx: ['axId', 'scaling', 'axPos', 'crossAx'],
  'ser/dLbls': [
    'showLegendKey',
    'showVal',
    'showCatName',
    'showSerName',
    'showPercent',
    'showBubbleSize',
  ],
};

/** Check child order + required children of every element in SEQ / REQUIRED; returns problems. */
function schemaProblems(xml: string): string[] {
  const out: string[] = [];
  const stack: { name: string; key: string; pos: number; kids: Set<string> }[] = [];
  const keyOf = (name: string, parent: string | undefined): string => {
    const k = parent ? `${parent}/${name}` : name;
    if (SEQ[k] || REQUIRED[k]) return k;
    if ((name === 'dLbls' || name === 'marker') && parent?.endsWith('ser')) return `ser/${name}`;
    return name;
  };
  const close = (): void => {
    const e = stack.pop()!;
    for (const r of REQUIRED[e.key] ?? REQUIRED[e.name] ?? [])
      if (!e.kids.has(r)) out.push(`${e.key}: missing ${r}`);
  };
  parseXml(xml, {
    open(name) {
      const ln = localName(name);
      const parent = stack[stack.length - 1];
      if (parent) {
        parent.kids.add(ln);
        const seq = SEQ[parent.key];
        if (seq) {
          const i = seq.findIndex((s) => s.split('|').includes(ln));
          if (i < 0) out.push(`${parent.key}: unexpected ${ln}`);
          else if (i < parent.pos) out.push(`${parent.key}: ${ln} out of order`);
          else parent.pos = i;
        }
      }
      const key = keyOf(ln, parent?.name);
      stack.push({ name: ln, key, pos: 0, kids: new Set() });
    },
    close() {
      close();
    },
  });
  if (stack.length) out.push('unbalanced');
  return out;
}

// ---------------------------------------------------------------- fixtures

function dataBook(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('My Data');
  ws.addRow(['Month', 'North', 'South', 'X', 'Y']);
  for (const [i, m] of ['Jan', 'Feb', 'Mar'].entries())
    ws.addRow([m, 10 + i, 20 - i, i + 1, (i + 1) ** 2]);
  return wb;
}

const D = "'My Data'!";
const TWO = [
  { name: { ref: `${D}$B$1` }, categories: `${D}$A$2:$A$4`, values: `${D}$B$2:$B$4` },
  { name: 'South', categories: 'A2:A4', values: 'C2:C4' },
];
const TYPES: ChartType[] = ['bar', 'column', 'line', 'pie', 'doughnut', 'scatter', 'area'];

function chartOf(type: ChartType, extra: Partial<Chart> = {}): Chart {
  const series =
    type === 'scatter'
      ? [{ name: 'y', xValues: 'D2:D4', values: 'E2:E4' }]
      : type === 'pie' || type === 'doughnut'
        ? [TWO[0]!]
        : TWO;
  return { type, title: `${type} chart`, series, anchor: 'G2:N16', ...extra };
}

function oneChart(type: ChartType, extra: Partial<Chart> = {}): Map<string, string> {
  const wb = dataBook();
  addChart(wb.worksheets[0]!, chartOf(type, extra));
  return parts(writeXlsx(wb));
}

// ---------------------------------------------------------------- tests

describe('charts: schema order', () => {
  it('the checker accepts the chart Excel wrote (fixture calibration)', () => {
    const z = parts(
      new Uint8Array(
        readFileSync(new URL('./fixtures/exceljs/data/chart-sheet.xlsx', import.meta.url)),
      ),
    );
    const xml = z.get('xl/charts/chart1.xml')!;
    expect(xml).toContain('<c:barChart>');
    expect(schemaProblems(xml)).toEqual([]);
    expect(schemaProblems(z.get('xl/drawings/drawing1.xml')!)).toEqual([]);
    // and rejects a misordered one
    expect(
      schemaProblems(
        xml.replace(
          '<c:barDir val="col"/><c:grouping val="clustered"/>',
          '<c:grouping val="clustered"/><c:barDir val="col"/>',
        ),
      ),
    ).toEqual(['barChart: barDir out of order']);
  });

  for (const type of TYPES) {
    it(`${type}: valid chartSpace, right chart element, axes`, () => {
      const z = oneChart(type, {
        dataLabels: { showValue: true, showPercent: type === 'pie' },
        axes: { x: { title: 'X' }, y: { title: 'Y', min: 0, max: 50, numFmt: '0.0' } },
      });
      const xml = z.get('xl/charts/chart1.xml')!;
      expect(schemaProblems(xml)).toEqual([]);
      expect(schemaProblems(z.get('xl/drawings/drawing1.xml')!)).toEqual([]);
      const tag = {
        bar: 'barChart',
        column: 'barChart',
        line: 'lineChart',
        pie: 'pieChart',
        doughnut: 'doughnutChart',
        scatter: 'scatterChart',
        area: 'areaChart',
      }[type];
      expect(count(xml, new RegExp(`<c:${tag}>`))).toBe(1);
      expect(xml).toContain(`<a:t>${type} chart</a:t>`);
      expect(xml).toContain('<c:autoTitleDeleted val="0"/>');
      expect(count(xml, /<c:ser>/)).toBe(
        type === 'scatter' || type === 'pie' || type === 'doughnut' ? 1 : 2,
      );
      expect(count(xml, /<c:dLbls>/)).toBe(count(xml, /<c:ser>/));
      if (type === 'pie' || type === 'doughnut') {
        expect(xml).not.toContain('Ax>');
        expect(xml).toContain('<c:varyColors val="1"/>');
        expect(xml).toContain('<c:showPercent val="' + (type === 'pie' ? 1 : 0));
      } else if (type === 'scatter') {
        expect(count(xml, /<c:valAx>/)).toBe(2);
        expect(xml).not.toContain('<c:catAx>');
        expect(xml).toContain("<c:xVal><c:numRef><c:f>'My Data'!$D$2:$D$4</c:f>");
        expect(xml).toContain('<c:scatterStyle val="lineMarker"/>');
      } else {
        expect(count(xml, /<c:catAx>/)).toBe(1);
        expect(count(xml, /<c:valAx>/)).toBe(1);
        expect(xml).toContain('<c:max val="50"/><c:min val="0"/>');
        expect(xml).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
        expect(xml).toContain('<a:bodyPr rot="-5400000" vert="horz"/>');
      }
      if (type === 'bar' || type === 'column') {
        expect(xml).toContain(
          `<c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/>`,
        );
        // horizontal bars: category axis on the left
        expect(xml).toContain(
          `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'l' : 'b'}"/>`,
        );
      }
    });
  }

  it('stacked / percentStacked groupings (+ overlap for bars); ignored for pie', () => {
    for (const type of ['bar', 'column', 'line', 'area'] as const) {
      const xml = oneChart(type, { stacked: 'percentStacked' }).get('xl/charts/chart1.xml')!;
      expect(xml).toContain('<c:grouping val="percentStacked"/>');
      expect(xml.includes('<c:overlap val="100"/>')).toBe(type === 'bar' || type === 'column');
      expect(schemaProblems(xml)).toEqual([]);
    }
    expect(oneChart('pie', { stacked: 'stacked' }).get('xl/charts/chart1.xml')).not.toContain(
      'grouping',
    );
  });

  it('legend, style, varyColors, no title, series options', () => {
    const z = oneChart('line', {
      title: undefined,
      legend: { position: 'b' },
      style: 26,
      varyColors: true,
      series: [
        { values: 'B2:B4', color: { argb: 'FF00FF00' }, smooth: true },
        { values: 'C2:C4', marker: false, color: { theme: 5 } },
      ],
    });
    const xml = z.get('xl/charts/chart1.xml')!;
    expect(schemaProblems(xml)).toEqual([]);
    expect(xml).toContain(
      '<c:roundedCorners val="0"/><c:style val="26"/><c:chart><c:autoTitleDeleted val="1"/>',
    );
    expect(xml).toContain('<c:legendPos val="b"/>');
    expect(xml).toContain('<c:varyColors val="1"/>');
    expect(xml).toContain('<a:srgbClr val="00FF00"/>');
    expect(xml).toContain('<a:schemeClr val="accent2"/>');
    expect(xml).toContain('<c:smooth val="1"/>');
    expect(xml).toContain('<c:marker><c:symbol val="none"/></c:marker>');
    expect(xml).not.toContain('<c:tx>');
    expect(oneChart('pie', { legend: false }).get('xl/charts/chart1.xml')).not.toContain(
      '<c:legend>',
    );
    // scatter: markers only unless `line`; smooth → smoothMarker
    const sc = oneChart('scatter', {
      series: [
        { values: 'E2:E4', xValues: 'D2:D4' },
        { values: 'E2:E4', xValues: 'D2:D4', line: true, smooth: true },
      ],
    }).get('xl/charts/chart1.xml')!;
    expect(schemaProblems(sc)).toEqual([]);
    expect(count(sc, /<a:ln w="28575" cap="rnd"><a:noFill\/>/)).toBe(1);
    expect(sc).toContain('<c:scatterStyle val="smoothMarker"/>');
  });

  it('rejects unknown types and empty series', () => {
    const wb = dataBook();
    expect(() => addChart(wb.worksheets[0]!, { type: 'line', series: [], anchor: 'A1' })).toThrow(
      /series/,
    );
    wb.worksheets[0]!.charts = [{ type: 'radar' as ChartType, series: TWO, anchor: 'A1' }];
    expect(() => writeXlsx(wb)).toThrow(/Unsupported chart type: radar/);
  });
});

describe('charts: combos, secondary axes, trendlines, data labels (#114)', () => {
  it('column + line on a secondary axis: two groups, four axes, valid order', () => {
    const z = oneChart('column', {
      series: [TWO[0]!, { ...TWO[1]!, type: 'line', axis: 'secondary', marker: false }],
      axes: { y: { title: 'Units' }, y2: { title: 'Share', numFmt: '0%' } },
    });
    const xml = z.get('xl/charts/chart1.xml')!;
    expect(schemaProblems(xml)).toEqual([]);
    expect(count(xml, /<c:barChart>/)).toBe(1);
    expect(count(xml, /<c:lineChart>/)).toBe(1);
    expect(count(xml, /<c:valAx>/)).toBe(2);
    expect(count(xml, /<c:catAx>/)).toBe(2);
    expect(xml).toMatch(
      /<c:lineChart>.*<c:idx val="1"\/>.*<c:axId val="30"\/><c:axId val="40"\/><\/c:lineChart>/,
    );
    expect(xml).toContain(
      '<c:axId val="30"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>',
    );
    expect(xml).toMatch(
      /<c:axId val="40"\/>.*<c:axPos val="r"\/>.*<a:t>Share<\/a:t>.*<c:crosses val="max"\/>/,
    );
  });

  it('stacked applies to the chart type only, not to combo series', () => {
    const xml = oneChart('area', {
      stacked: 'stacked',
      series: [TWO[0]!, { ...TWO[1]!, type: 'line' }],
    }).get('xl/charts/chart1.xml')!;
    expect(xml).toMatch(/<c:areaChart><c:grouping val="stacked"\/>/);
    expect(xml).toMatch(/<c:lineChart><c:grouping val="standard"\/>/);
  });

  it('column + line on one axis share the primary axes', () => {
    const xml = oneChart('column', { series: [TWO[0]!, { ...TWO[1]!, type: 'line' }] }).get(
      'xl/charts/chart1.xml',
    )!;
    expect(schemaProblems(xml)).toEqual([]);
    expect(xml).toMatch(/<c:lineChart>.*<c:axId val="10"\/><c:axId val="20"\/><\/c:lineChart>/);
    expect(count(xml, /<c:valAx>/)).toBe(1);
  });

  it('invalid combos throw', () => {
    const ws = dataBook().worksheets[0]!;
    expect(() => addChart(ws, chartOf('pie', { series: [{ ...TWO[0]!, type: 'line' }] }))).toThrow(
      'cannot be combined',
    );
    expect(() =>
      addChart(ws, chartOf('bar', { series: [{ ...TWO[0]!, type: 'column' }] })),
    ).toThrow('cannot be combined');
  });

  it('trendlines and per-series data labels', () => {
    for (const type of ['scatter', 'line', 'column', 'area'] as const) {
      const base = chartOf(type);
      const z = oneChart(type, {
        dataLabels: { showValue: true },
        series: base.series.map((s, i) => ({
          ...s,
          trendline:
            i === 0
              ? [
                  { type: 'linear', displayEquation: true, displayRSquared: true, forward: 1 },
                  { type: 'polynomial', order: 3, name: 'Poly', color: { argb: 'FFFF0000' } },
                ]
              : { type: 'movingAverage', period: 2 },
          dataLabels:
            i === 1
              ? false
              : { showValue: true, position: type === 'area' ? undefined : 't', numFmt: '0.0' },
        })),
      });
      const xml = z.get('xl/charts/chart1.xml')!;
      expect(schemaProblems(xml), type).toEqual([]);
      expect(xml).toContain(
        '<c:trendlineType val="linear"/><c:forward val="1"/><c:dispRSqr val="1"/><c:dispEq val="1"/><c:trendlineLbl>',
      );
      expect(xml).toContain('<c:name>Poly</c:name>');
      expect(xml).toContain('<c:trendlineType val="poly"/><c:order val="3"/>');
      expect(count(xml, /<c:dLbls>/)).toBe(1); // second series: no labels
      expect(xml).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
    }
  });

  it('pie ignores trendlines and secondary axes', () => {
    const xml = oneChart('pie', {
      series: [{ ...TWO[0]!, axis: 'secondary', trendline: { type: 'linear' } }],
    }).get('xl/charts/chart1.xml')!;
    expect(schemaProblems(xml)).toEqual([]);
    expect(xml).not.toContain('trendline');
    expect(xml).not.toContain('Ax>');
  });
});

describe('charts: references and caches', () => {
  it('qualifies / quotes / absolutizes references and fills caches from cells', () => {
    const xml = oneChart('column').get('xl/charts/chart1.xml')!;
    // name by reference → strRef + cache; literal name → c:v
    expect(xml).toContain(
      `<c:tx><c:strRef><c:f>'My Data'!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>`,
    );
    expect(xml).toContain('<c:tx><c:v>South</c:v></c:tx>');
    // unqualified range = the chart's own sheet
    expect(
      count(
        xml,
        /<c:f>'My Data'!\$A\$2:\$A\$4<\/c:f><c:strCache><c:ptCount val="3"\/><c:pt idx="0"><c:v>Jan<\/c:v>/,
      ),
    ).toBe(2);
    expect(xml).toContain(
      `<c:val><c:numRef><c:f>'My Data'!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>19</c:v></c:pt><c:pt idx="2"><c:v>18</c:v></c:pt></c:numCache></c:numRef></c:val>`,
    );
  });

  it('other sheets, quoting rules, gaps, formulas, numeric categories, number formats', () => {
    const wb = new Workbook();
    const own = wb.addWorksheet('Report');
    const a = wb.addWorksheet("Bob's Data");
    const b = wb.addWorksheet('Plain_1');
    const c = wb.addWorksheet('A1');
    a.addRow([2021, 5]);
    a.addRow([2022, { formula: 'B1*2', result: 10 }]);
    a.addRow([2023, null]);
    a.addRow([2024, 7]);
    b.setCell('A1', 0.25, { numFmt: '0%' });
    b.setCell('A2', 0.5);
    c.setCell('A1', 1);
    addChart(own, {
      type: 'line',
      series: [
        {
          name: "='Bob''s Data'!B1",
          categories: "'Bob''s Data'!A1:A4",
          values: "='Bob''s Data'!B1:B4",
        },
        { values: 'Plain_1!a1:a2' },
        { values: "'A1'!A1" },
        { values: 'Missing!A1:A2' },
      ],
      anchor: { tl: 'B2', ext: { width: 480, height: 300 } },
    });
    const z = parts(writeXlsx(wb));
    const xml = z.get('xl/charts/chart1.xml')!;
    expect(schemaProblems(xml)).toEqual([]);
    expect(xml).toContain(
      `<c:cat><c:numRef><c:f>'Bob''s Data'!$A$1:$A$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>2021</c:v></c:pt>`,
    );
    // gap (empty B3) → no pt, ptCount keeps the length; formula → cached result
    expect(xml).toContain(
      `<c:f>'Bob''s Data'!$B$1:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt><c:pt idx="3"><c:v>7</c:v></c:pt></c:numCache>`,
    );
    // number name ref → strCache of its text
    expect(xml).toContain(
      `<c:tx><c:strRef><c:f>'Bob''s Data'!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v>`,
    );
    expect(xml).toContain(
      '<c:f>Plain_1!$A$1:$A$2</c:f><c:numCache><c:formatCode>0%</c:formatCode>',
    );
    expect(xml).toContain(`<c:f>'A1'!$A$1</c:f>`); // looks like a cell reference → quoted
    expect(xml).toContain('<c:numRef><c:f>Missing!$A$1:$A$2</c:f></c:numRef>'); // no data → no cache
    expect(z.get('xl/drawings/drawing1.xml')).toContain(
      '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col>',
    );
  });

  it('user-provided caches win (and are the only caches in the streaming writer)', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    s.addRow(['a', 1]);
    s.addRow(['b', 2]);
    s.worksheet.charts = [
      {
        type: 'column',
        series: [
          { name: { ref: 'B1' }, categories: 'A1:A2', values: 'B1:B2' },
          {
            name: { ref: 'B1' },
            categories: 'A1:A2',
            values: 'B1:B2',
            cache: { name: 'N', categories: ['a', 'b'], values: [1, null] },
          },
        ],
        anchor: 'D2:H10',
      },
    ];
    await s.commit();
    await w.end();
    const z = parts(new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer()));
    const xml = z.get('xl/charts/chart1.xml')!;
    expect(schemaProblems(xml)).toEqual([]);
    // first series: references only (no cells in the streaming writer)
    expect(xml).toContain('<c:tx><c:strRef><c:f>S!$B$1</c:f></c:strRef></c:tx>');
    expect(xml).toContain(
      '<c:cat><c:strRef><c:f>S!$A$1:$A$2</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$1:$B$2</c:f></c:numRef></c:val>',
    );
    // second: the given caches
    expect(xml).toContain(
      '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>N</c:v></c:pt></c:strCache>',
    );
    expect(xml).toContain(
      '<c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache>',
    );
    expect(xml).toContain('<c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache>');
    expect(z.get('[Content_Types].xml')).toContain(
      `<Override PartName="/xl/charts/chart1.xml" ContentType="${CT_CHART}"/>`,
    );
    expect(z.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain('../drawings/drawing1.xml');
    // our reader keeps it
    const back = readXlsx(new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer()));
    expect(listPreservedParts(back).map((p) => p.path)).toContain('xl/charts/chart1.xml');
  });
});

describe('charts: package', () => {
  it('rels, content types and numbering across charts and sheets', () => {
    const wb = dataBook();
    const ws = wb.worksheets[0]!;
    addChart(ws, chartOf('column'));
    addChart(ws, chartOf('pie', { anchor: 'G18:N30', name: 'My pie' }));
    const other = wb.addWorksheet('Other');
    other.addRow([1]);
    const plain = wb.addWorksheet('NoCharts');
    plain.addRow([1]);
    const last = wb.addWorksheet('Last');
    addChart(last, chartOf('line'));
    const z = parts(writeXlsx(wb));
    expect([...z.keys()].filter((k) => k.includes('chart')).sort()).toEqual([
      'xl/charts/chart1.xml',
      'xl/charts/chart2.xml',
      'xl/charts/chart3.xml',
    ]);
    const ct = z.get('[Content_Types].xml')!;
    for (const n of [1, 2, 3])
      expect(ct).toContain(
        `<Override PartName="/xl/charts/chart${n}.xml" ContentType="${CT_CHART}"/>`,
      );
    expect(count(ct, /drawing\+xml/)).toBe(2);
    expect(z.get('xl/drawings/_rels/drawing1.xml.rels')).toMatch(
      /Id="rId1" Type="[^"]*\/chart" Target="..\/charts\/chart1.xml".*Id="rId2" Type="[^"]*\/chart" Target="..\/charts\/chart2.xml"/,
    );
    expect(z.get('xl/drawings/_rels/drawing2.xml.rels')).toContain('Target="../charts/chart3.xml"');
    expect(z.get('xl/worksheets/_rels/sheet4.xml.rels')).toContain('../drawings/drawing2.xml');
    expect(z.get('xl/worksheets/sheet2.xml')).not.toContain('<drawing');
    const d1 = z.get('xl/drawings/drawing1.xml')!;
    expect(d1).toContain('<xdr:cNvPr id="2" name="Chart 1"/>');
    expect(d1).toContain('<xdr:cNvPr id="3" name="My pie"/>');
    expect(d1).toContain(
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>',
    );
    expect(z.get('xl/worksheets/sheet1.xml')).toMatch(/<drawing r:id="rId\d+"\/>/);
  });

  it('a chart and a picture on the same sheet share one drawing with unique ids', () => {
    const wb = dataBook();
    const ws = wb.worksheets[0]!;
    const img = addImage(wb, { data: PNG, extension: 'png' });
    placeImage(ws, img, 'A10:B12');
    addChart(ws, chartOf('column'));
    placeImage(ws, img, 'C10:D12');
    addChart(
      ws,
      chartOf('line', { anchor: { from: { col: 1, row: 20 }, to: { col: 8, row: 30 } } }),
    );
    const z = parts(writeXlsx(wb));
    expect([...z.keys()].filter((k) => /^xl\/drawings\/drawing\d+\.xml$/.test(k))).toEqual([
      'xl/drawings/drawing1.xml',
    ]);
    const d = z.get('xl/drawings/drawing1.xml')!;
    expect(schemaProblems(d)).toEqual([]);
    expect(count(d, /<xdr:pic>/)).toBe(2);
    expect(count(d, /<xdr:graphicFrame /)).toBe(2);
    const ids = [...d.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(4);
    // pictures first (order 0), then charts (order 1)
    expect(d.indexOf('<xdr:pic>')).toBeLessThan(d.indexOf('<xdr:graphicFrame'));
    const rels = z.get('xl/drawings/_rels/drawing1.xml.rels')!;
    expect(count(rels, /\/image"/)).toBe(1);
    expect(count(rels, /\/chart"/)).toBe(2);
    expect(count(z.get('xl/worksheets/sheet1.xml'), /<drawing /)).toBe(1);
  });

  it('a new chart lands next to the charts kept from a read file', () => {
    const wb0 = dataBook();
    addChart(wb0.worksheets[0]!, chartOf('column'));
    placeImage(wb0.worksheets[0]!, addImage(wb0, { data: PNG, extension: 'png' }), 'A10:B12');
    const first = writeXlsx(wb0);
    const keptXml = parts(first).get('xl/charts/chart1.xml');

    const wb = readXlsx(first);
    const ws = wb.worksheets[0]!;
    expect(ws.charts).toHaveLength(1); // read charts are modelled (#113)
    expect(ws.charts![0]).toMatchObject({ type: 'column', title: 'column chart' });
    expect(ws.images).toHaveLength(1);
    expect(ws.preserved?.drawing?.xml ?? '').not.toContain('<c:chart ');
    addChart(ws, chartOf('pie', { anchor: 'P2:V16' }));
    const z = parts(writeXlsx(wb));
    // the kept chart keeps its name; ours skips the taken number
    expect(z.get('xl/charts/chart1.xml')).toBe(keptXml);
    expect(z.get('xl/charts/chart2.xml')).toContain('<c:pieChart>');
    const d = z.get('xl/drawings/drawing1.xml')!;
    expect(schemaProblems(d)).toEqual([]);
    expect(count(d, /<xdr:graphicFrame /)).toBe(2);
    expect(count(d, /<xdr:pic>/)).toBe(1);
    const ids = [...d.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(3);
    const rels = z.get('xl/drawings/_rels/drawing1.xml.rels')!;
    expect(rels).toContain('Target="../charts/chart1.xml"');
    expect(rels).toContain('Target="../charts/chart2.xml"');
    const ct = z.get('[Content_Types].xml')!;
    expect(count(ct, /drawingml\.chart\+xml/)).toBe(2);
    // and again: both are kept now
    const again = readXlsx(writeXlsx(wb));
    expect(listPreservedParts(again).map((p) => p.path)).toEqual(
      expect.arrayContaining(['xl/charts/chart1.xml', 'xl/charts/chart2.xml']),
    );
  });

  it('our reader round-trips charts raw (byte-identical chart parts)', async () => {
    const wb = dataBook();
    for (const t of TYPES) addChart(wb.worksheets[0]!, chartOf(t));
    const buf = writeXlsx(wb);
    const src = parts(buf);
    for (const out of [
      writeXlsx(readXlsx(buf)),
      new Uint8Array(await new Response(xlsxStream(readXlsx(buf))).arrayBuffer()),
    ]) {
      const z = parts(out);
      for (let n = 1; n <= TYPES.length; n++)
        expect(z.get(`xl/charts/chart${n}.xml`)).toBe(src.get(`xl/charts/chart${n}.xml`));
      expect(count(z.get('xl/drawings/drawing1.xml'), /<xdr:graphicFrame /)).toBe(TYPES.length);
    }
  });

  it('exceljs loads files with our charts (and sees the data)', async () => {
    const wb = dataBook();
    for (const t of TYPES) addChart(wb.worksheets[0]!, chartOf(t));
    placeImage(wb.worksheets[0]!, addImage(wb, { data: PNG, extension: 'png' }), 'A10:B12');
    const x = new ExcelJS.Workbook();
    // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
    await x.xlsx.load(writeXlsx(wb) as any);
    const ws = x.getWorksheet('My Data')!;
    expect(ws.getCell('B2').value).toBe(10);
    expect(ws.getImages()).toHaveLength(1);
  });

  it('streaming writer and xlsxStream write the same parts', async () => {
    const wb = dataBook();
    addChart(wb.worksheets[0]!, chartOf('area'));
    const a = parts(writeXlsx(wb));
    const b = parts(new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer()));
    expect(b.get('xl/charts/chart1.xml')).toBe(a.get('xl/charts/chart1.xml'));
    expect(b.get('xl/drawings/drawing1.xml')).toBe(a.get('xl/drawings/drawing1.xml'));
  });
});

describe('charts: lite', () => {
  it('opt-in: ws.charts without chartsWriter throws; chartsWriter alone works (no images)', () => {
    const wb = new lite.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow([1]);
    lite.addChart(ws, { type: 'pie', series: [{ values: 'A1' }], anchor: 'C2:F8' });
    expect(() => lite.writeXlsx(wb)).toThrow(
      /ws\.charts is set but the 'charts' feature is not enabled; pass features: \[chartsWriter\]/,
    );
    const z = parts(lite.writeXlsx(wb, { features: [lite.chartsWriter] }));
    expect(z.get('xl/charts/chart1.xml')).toContain('<c:pieChart>');
    expect(z.get('xl/drawings/drawing1.xml')).toContain('<xdr:graphicFrame ');
    // the same bytes as the root for the same content
    const r = parts(writeXlsx(wb));
    expect(z.get('xl/charts/chart1.xml')).toBe(r.get('xl/charts/chart1.xml'));
  });

  it('kept drawing anchors are merged by chartsWriter too (without imagesWriter)', () => {
    const wb0 = dataBook();
    addChart(wb0.worksheets[0]!, chartOf('column'));
    const wb = readXlsx(writeXlsx(wb0));
    const ws = wb.worksheets[0]!;
    const noImages = [lite.printSetupWriter, lite.themeWriter];
    // the read chart is modelled now: writing it needs chartsWriter
    expect(() => lite.writeXlsx(wb, { features: noImages })).toThrow(/chartsWriter/);
    addChart(ws, chartOf('line'));
    const z = parts(lite.writeXlsx(wb, { features: [lite.chartsWriter, ...noImages] }));
    expect(count(z.get('xl/drawings/drawing1.xml'), /<xdr:graphicFrame /)).toBe(2);
  });

  it('a sheet read without the images reader: the model chart replaces the raw drawing', () => {
    const wb0 = dataBook();
    addChart(wb0.worksheets[0]!, chartOf('column'));
    const wb = lite.readXlsx(writeXlsx(wb0));
    expect(parts(lite.writeXlsx(wb)).get('xl/charts/chart1.xml')).toContain('<c:barChart>');
    addChart(wb.worksheets[0]!, chartOf('line'));
    const z = parts(lite.writeXlsx(wb, { features: [lite.chartsWriter] }));
    // one drawing (no orphaned raw drawing / chart parts)
    expect([...z.keys()].filter((k) => /drawings\/drawing\d+\.xml$/.test(k))).toHaveLength(1);
    expect(count(z.get('xl/worksheets/sheet1.xml'), /<drawing /)).toBe(1);
    expect([...z.keys()].filter((k) => k.startsWith('xl/charts/'))).toHaveLength(1);
  });
});
