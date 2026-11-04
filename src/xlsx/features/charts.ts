// Feature: charts (write only) — DrawingML charts (xl/charts/chartN.xml) placed on a sheet through
// its drawing part (shared with pictures and kept anchors: features/drawing.ts).
//
//   addChart(ws, {
//     type: 'column', title: 'Sales',
//     series: [{ name: 'Units', categories: 'Data!$A$2:$A$7', values: 'Data!$B$2:$B$7' }],
//     anchor: 'D2:K16',
//   });
//
// Series data are A1 references (`'My Data'!$B$2:$B$10`; unqualified = the chart's own sheet).
// Each reference is written with a cache of the referenced cells' current values (Excel shows
// them immediately; LibreOffice needs them); the streaming writer holds no cells, so it writes no
// caches (Excel recalculates on open) unless `series.cache` provides them.
// Reading: charts in read files are kept raw (round-trip preservation), never modelled.

import { decodeRange, encodeCol } from '../../core/address.js';
import type { Color } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import { type AnchorInput, anchorXml, drawingEmitter, toAnchor } from './drawing.js';
import type { DrawingWriteCtx, FeatureWriter } from './types.js';
import { escapeAttr, escapeText } from './util.js';

export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area';

/** One data series. References are A1 ranges, optionally sheet-qualified. */
export interface ChartSeries {
  /** Series name: literal text, or `{ ref: 'Data!$B$1' }` (also a string starting with "="). */
  name?: string | { ref: string } | undefined;
  /** Category labels (pie slices, x axis of bar/column/line/area). */
  categories?: string | undefined;
  /** Values (scatter: y values). */
  values: string;
  /** Scatter only: x values. */
  xValues?: string | undefined;
  /** Fill (bar, column, area, pie) or line color (line, scatter). `argb` or `theme` (0-9). */
  color?: Color | undefined;
  /** Line / scatter: false hides the markers (default: automatic markers). */
  marker?: boolean | undefined;
  /** Line / scatter: smoothed line. */
  smooth?: boolean | undefined;
  /** Scatter only: draw lines between the points (default false: markers only). */
  line?: boolean | undefined;
  /**
   * Cached values written instead of reading the referenced cells (needed by the streaming
   * writer for LibreOffice, which does not recalculate charts).
   */
  cache?:
    | {
        name?: string | undefined;
        categories?: readonly (string | number)[] | undefined;
        values?: readonly (number | null)[] | undefined;
        xValues?: readonly (number | null)[] | undefined;
      }
    | undefined;
}

export interface ChartAxis {
  title?: string | undefined;
  /** Tick label number format, e.g. "0%" (default: linked to the source cells). */
  numFmt?: string | undefined;
  /** Value axes: scale bounds. */
  min?: number | undefined;
  max?: number | undefined;
}

/** A chart placed on a sheet (`ws.charts`). */
export interface Chart {
  type: ChartType;
  title?: string | undefined;
  series: ChartSeries[];
  /** Where the chart sits: same forms as `placeImage` ("D2:K16", `{ tl, ext }`, from/to). */
  anchor: AnchorInput;
  /** Legend position (default 'r'); false hides it. */
  legend?: { position: 'r' | 'l' | 't' | 'b' | 'tr' } | false | undefined;
  /** x = category axis (scatter: x values; horizontal bar: the vertical axis), y = value axis. */
  axes?: { x?: ChartAxis | undefined; y?: ChartAxis | undefined } | undefined;
  /** bar, column, line, area: stack the series. */
  stacked?: 'stacked' | 'percentStacked' | undefined;
  dataLabels?:
    | {
        showValue?: boolean | undefined;
        showPercent?: boolean | undefined;
        showCategoryName?: boolean | undefined;
        showSeriesName?: boolean | undefined;
      }
    | undefined;
  /** Built-in chart style 1-48. */
  style?: number | undefined;
  /** One color per point (default: pie / doughnut only). */
  varyColors?: boolean | undefined;
  /** Shape name (default "Chart N"). */
  name?: string | undefined;
}

/** Add a chart to a sheet (`ws.charts`); returns it. */
export function addChart(ws: Worksheet, chart: Chart): Chart {
  if (!chart.series?.length) throw new TypeError('addChart: at least one series is required');
  ws.charts ??= [];
  ws.charts.push(chart);
  return chart;
}

// ---------------------------------------------------------------- references

/** Sheet name as a formula prefix: quoted when needed ("'My Data'"), quotes doubled. */
function sheetRef(name: string): string {
  return /^[A-Za-z_][\w.]*$/.test(name) && !/^([A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*)$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

interface Ref {
  /** Normalized formula: `'Sheet'!$A$1:$A$5`. */
  f: string;
  sheet: string;
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Parse an A1 reference (optional "=", optional quoted sheet; default: the chart's sheet). */
function parseRef(ref: string, own: string): Ref {
  let s = ref.trim().replace(/^=/, '');
  let sheet = own;
  const bang = s.lastIndexOf('!');
  if (bang >= 0) {
    sheet = s.slice(0, bang);
    if (sheet[0] === "'") sheet = sheet.slice(1, -1).replace(/''/g, "'");
    s = s.slice(bang + 1);
  }
  const r = decodeRange(s);
  const a = (row: number, col: number): string => `$${encodeCol(col)}$${row + 1}`;
  const tl = a(r.top, r.left);
  const br = a(r.bottom, r.right);
  return { f: `${sheetRef(sheet)}!${tl === br ? tl : `${tl}:${br}`}`, sheet, ...r };
}

type Pt = string | number | undefined;
interface Cells {
  pts: Pt[];
  /** Number format of the first numeric cell. */
  fmt?: string | undefined;
}

/** Current values of a reference's cells (row-major), from the in-memory workbook. */
function cells(c: DrawingWriteCtx, r: Ref): Cells | undefined {
  if (c.streaming) return undefined;
  const ws = c.workbook.getWorksheet(r.sheet);
  if (!ws) return undefined;
  const pts: Pt[] = [];
  let fmt: string | undefined;
  for (let y = r.top; y <= r.bottom; y++) {
    const row = ws.rowData[y];
    for (let x = r.left; x <= r.right; x++) {
      const i = row ? row.find(x) : -1;
      let v: unknown;
      if (i >= 0) {
        const k = row!.kind[i];
        const o = row!.obj?.[i] as {
          result?: unknown;
          text?: string;
          richText?: { text: string }[];
        };
        v =
          k === 1 || k === 3 || k === 4
            ? row!.num[i]
            : k === 2
              ? o
              : k === 6
                ? o.result
                : k === 7
                  ? o.richText!.map((t) => t.text).join('')
                  : k === 8
                    ? o.text
                    : undefined;
        if (typeof v === 'boolean') v = +v;
        if (typeof v === 'number' && fmt === undefined) fmt = ws.getStyle(y + 1, x + 1).numFmt;
      }
      pts.push(typeof v === 'number' || typeof v === 'string' ? v : undefined);
    }
  }
  return { pts, fmt };
}

/**
 * `<c:strRef>` / `<c:numRef>` with its cache. `str`: text data (series name; the default of
 * `auto` sources without a cache); `auto` (categories, x values): text when any point is text,
 * else numbers. Number sources (values) skip non-numeric points.
 */
function dataRef(
  c: DrawingWriteCtx,
  ref: string,
  given: readonly (string | number | null)[] | undefined,
  str: boolean,
  auto?: boolean,
): string {
  const r = parseRef(ref, c.ws.name);
  const data: Cells | undefined = given ? { pts: given.map((v) => v ?? undefined) } : cells(c, r);
  const f = `<c:f>${escapeText(r.f)}</c:f>`;
  if (!data) return str ? `<c:strRef>${f}</c:strRef>` : `<c:numRef>${f}</c:numRef>`;
  let pts = data.pts;
  if (auto) str = pts.some((v) => typeof v === 'string');
  else if (!str) pts = pts.map((v) => (typeof v === 'number' ? v : undefined));
  let s = `<c:ptCount val="${pts.length}"/>`;
  for (let i = 0; i < pts.length; i++) {
    const v = pts[i];
    if (v !== undefined) s += `<c:pt idx="${i}"><c:v>${escapeText(String(v))}</c:v></c:pt>`;
  }
  return str
    ? `<c:strRef>${f}<c:strCache>${s}</c:strCache></c:strRef>`
    : `<c:numRef>${f}<c:numCache><c:formatCode>${escapeText(data.fmt ?? 'General')}</c:formatCode>${s}</c:numCache></c:numRef>`;
}

// ---------------------------------------------------------------- chart XML

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const SCHEME = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];
const val = (tag: string, v: string | number | boolean): string =>
  `<c:${tag} val="${typeof v === 'boolean' ? +v : v}"/>`;

function fill(col: Color): string {
  const clr =
    col.theme !== undefined
      ? `<a:schemeClr val="${SCHEME[col.theme] ?? 'accent1'}"/>`
      : `<a:srgbClr val="${escapeAttr((col.argb ?? '000000').slice(-6))}"/>`;
  return `<a:solidFill>${clr}</a:solidFill>`;
}

/** Rich text title (`rot` rotates vertical axis titles). */
function title(text: string, rot = ''): string {
  return `<c:title><c:tx><c:rich><a:bodyPr${rot}/><a:lstStyle/><a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function series(c: DrawingWriteCtx, ch: Chart, s: ChartSeries, i: number): string {
  const t = ch.type;
  const lined = t === 'line' || t === 'scatter';
  let x = val('idx', i) + val('order', i);
  const n = s.name;
  if (n !== undefined) {
    const ref = typeof n === 'object' ? n.ref : n[0] === '=' ? n : undefined;
    x += `<c:tx>${ref ? dataRef(c, ref, s.cache?.name === undefined ? undefined : [s.cache.name], true) : `<c:v>${escapeText(n as string)}</c:v>`}</c:tx>`;
  }
  // spPr: series fill / line; scatter without `line`: no connecting line
  const noLine = t === 'scatter' && !s.line;
  if (s.color || noLine) {
    const f = s.color ? fill(s.color) : '';
    x += `<c:spPr>${lined ? `<a:ln w="28575" cap="rnd">${noLine ? '<a:noFill/>' : f}<a:round/></a:ln>` : f}</c:spPr>`;
  }
  if (t === 'bar' || t === 'column') x += val('invertIfNegative', 0);
  if (lined && (s.marker === false || s.color)) {
    x +=
      s.marker === false
        ? '<c:marker><c:symbol val="none"/></c:marker>'
        : `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(s.color!)}</c:spPr></c:marker>`;
  }
  const dl = ch.dataLabels;
  if (dl) {
    x += `<c:dLbls>${val('showLegendKey', 0)}${val('showVal', !!dl.showValue)}${val('showCatName', !!dl.showCategoryName)}${val('showSerName', !!dl.showSeriesName)}${val('showPercent', !!dl.showPercent)}${val('showBubbleSize', 0)}</c:dLbls>`;
  }
  const cache = s.cache;
  if (t === 'scatter') {
    if (s.xValues) x += `<c:xVal>${dataRef(c, s.xValues, cache?.xValues, false, true)}</c:xVal>`;
    x += `<c:yVal>${dataRef(c, s.values, cache?.values, false)}</c:yVal>`;
  } else {
    if (s.categories)
      x += `<c:cat>${dataRef(c, s.categories, cache?.categories, true, true)}</c:cat>`;
    x += `<c:val>${dataRef(c, s.values, cache?.values, false)}</c:val>`;
  }
  if (lined) x += val('smooth', !!s.smooth);
  return `<c:ser>${x}</c:ser>`;
}

function axis(
  tag: 'catAx' | 'valAx',
  id: number,
  cross: number,
  pos: string,
  a: ChartAxis | undefined,
  value: boolean,
  between: string,
): string {
  let scaling = val('orientation', 'minMax');
  if (a?.max !== undefined) scaling += val('max', a.max);
  if (a?.min !== undefined) scaling += val('min', a.min);
  const vertical = pos === 'l' || pos === 'r';
  let x = `${val('axId', id)}<c:scaling>${scaling}</c:scaling>${val('delete', 0)}${val('axPos', pos)}`;
  if (value && vertical) x += '<c:majorGridlines/>';
  if (a?.title) x += title(a.title, vertical ? ' rot="-5400000" vert="horz"' : '');
  x += `<c:numFmt formatCode="${escapeAttr(a?.numFmt ?? 'General')}" sourceLinked="${a?.numFmt ? 0 : 1}"/>${val('majorTickMark', 'out')}${val('minorTickMark', 'none')}${val('tickLblPos', 'nextTo')}${val('crossAx', cross)}${val('crosses', 'autoZero')}`;
  x +=
    tag === 'catAx'
      ? `${val('auto', 1)}${val('lblAlgn', 'ctr')}${val('lblOffset', 100)}${val('noMultiLvlLbl', 0)}`
      : val('crossBetween', between);
  return `<c:${tag}>${x}</c:${tag}>`;
}

const TAGS: Record<ChartType, string> = {
  bar: 'barChart',
  column: 'barChart',
  line: 'lineChart',
  pie: 'pieChart',
  doughnut: 'doughnutChart',
  scatter: 'scatterChart',
  area: 'areaChart',
};

/** The chart part (c:chartSpace). */
function chartXml(c: DrawingWriteCtx, ch: Chart): string {
  const t = ch.type;
  const tag = TAGS[t];
  if (!tag) throw new TypeError(`Unsupported chart type: ${t}`);
  const pie = t === 'pie' || t === 'doughnut';
  const bar = t === 'bar' || t === 'column';
  const st = ch.stacked && !pie && t !== 'scatter' ? ch.stacked : undefined;
  let g = '';
  if (bar) g += val('barDir', t === 'bar' ? 'bar' : 'col');
  if (t === 'scatter')
    g += val('scatterStyle', ch.series.some((s) => s.smooth) ? 'smoothMarker' : 'lineMarker');
  else if (!pie) g += val('grouping', st ?? (bar ? 'clustered' : 'standard'));
  g += val('varyColors', ch.varyColors ?? pie);
  for (let i = 0; i < ch.series.length; i++) g += series(c, ch, ch.series[i]!, i);
  if (bar) g += val('gapWidth', 150) + (st ? val('overlap', 100) : '');
  if (t === 'line') g += val('marker', 1);
  if (pie) g += val('firstSliceAng', 0) + (t === 'doughnut' ? val('holeSize', 50) : '');
  let axes = '';
  if (!pie) {
    g += val('axId', 10) + val('axId', 20);
    const ax = ch.axes;
    const h = t === 'bar'; // horizontal bars: category axis on the left
    const between = t === 'scatter' || t === 'area' ? 'midCat' : 'between';
    axes =
      axis(
        t === 'scatter' ? 'valAx' : 'catAx',
        10,
        20,
        h ? 'l' : 'b',
        ax?.x,
        t === 'scatter',
        between,
      ) + axis('valAx', 20, 10, h ? 'b' : 'l', ax?.y, true, between);
  }
  const lg = ch.legend;
  const legend =
    lg === false
      ? ''
      : `<c:legend>${val('legendPos', lg?.position ?? 'r')}<c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS_C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${val('date1904', !!c.workbook.date1904)}${val('roundedCorners', 0)}${ch.style ? val('style', ch.style) : ''}<c:chart>${ch.title ? title(ch.title) + val('autoTitleDeleted', 0) : val('autoTitleDeleted', 1)}<c:plotArea><c:layout/><c:${tag}>${g}</c:${tag}>${axes}</c:plotArea>${legend}${val('plotVisOnly', 1)}${val('dispBlanksAs', 'gap')}</c:chart></c:chartSpace>`;
}

/** The sheet's charts: chart parts + graphicFrame anchors in the shared drawing. */
function charts(c: DrawingWriteCtx): string {
  const list = c.ws.charts ?? [];
  let body = '';
  for (let i = 0; i < list.length; i++) {
    const ch = list[i]!;
    const n = c.nextIndex('chart');
    c.addPart(
      `xl/charts/chart${n}.xml`,
      chartXml(c, ch),
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    );
    const id = c.drawingRel('chart', `../charts/chart${n}.xml`);
    body += anchorXml(
      toAnchor(ch.anchor, c.ws),
      undefined,
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${c.shapeId()}" name="${escapeAttr(ch.name ?? `Chart ${i + 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame>`,
    );
  }
  return body;
}

/**
 * Write half of the charts feature (ws.charts). Always on in the package root; in `./lite` pass
 * it in `features` of writeXlsx / xlsxStream / createXlsxWriter. There is no read half: charts
 * of read files are kept raw and written back.
 */
export const chartsWriter: FeatureWriter = {
  name: 'charts',
  sheet: {
    anchors: [{ order: 1, has: (ws) => !!ws.charts?.length, emit: charts }],
    emitters: [drawingEmitter],
  },
};
