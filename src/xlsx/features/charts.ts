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
// Reading (chartsReader): charts of the supported types (2-D bar / column / line / pie /
// doughnut / scatter / area, combos of column, line and area, secondary axes) become `ws.charts`
// with their original part kept: while a read chart is unchanged it is written back byte for
// byte (formatting the model does not hold included); once edited it is regenerated from the
// model. Other charts (3-D, bubble, stock, radar, chart sheets, ...) stay raw.

import { decodeRange, encodeCol } from '../../core/address.js';
import type { Color } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, parseXml } from '../../xml/index.js';
import { type AnchorInput, anchorXml, drawingEmitter, toAnchor } from './drawing.js';
import type { CellPosition, ImageAnchor } from './images.js';
import type { DrawingWriteCtx, FeatureReader, FeatureWriter, SheetReadCtx } from './types.js';
import { addPRel, escapeAttr, escapeText, presPart } from './util.js';

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
   * Combo charts: this series' type when it differs from the chart's ('column', 'line' or
   * 'area' — combined in one plot area; horizontal bars, pie, doughnut and scatter do not mix).
   */
  type?: 'column' | 'line' | 'area' | undefined;
  /** Plot on the secondary value axis (right side; `axes.y2` configures it). */
  axis?: 'primary' | 'secondary' | undefined;
  /** Trendline(s) of the series (not for pie / doughnut). */
  trendline?: ChartTrendline | ChartTrendline[] | undefined;
  /** Data labels of this series (overrides the chart's `dataLabels`; false: none). */
  dataLabels?: ChartDataLabels | false | undefined;
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

/** Data label settings (chart-wide `Chart.dataLabels` or per series). */
export interface ChartDataLabels {
  showValue?: boolean | undefined;
  showPercent?: boolean | undefined;
  showCategoryName?: boolean | undefined;
  showSeriesName?: boolean | undefined;
  /** Label position (support depends on the chart type, as in Excel). */
  position?: 'ctr' | 'inEnd' | 'inBase' | 'outEnd' | 't' | 'b' | 'l' | 'r' | 'bestFit' | undefined;
  /** Number format of the labels, e.g. "0.0%". */
  numFmt?: string | undefined;
}

/** A series trendline. */
export interface ChartTrendline {
  type: 'linear' | 'exponential' | 'logarithmic' | 'polynomial' | 'power' | 'movingAverage';
  /** polynomial: degree 2-6 (default 2). */
  order?: number | undefined;
  /** movingAverage: number of points (default 2). */
  period?: number | undefined;
  /** Forecast forward / backward, in x-axis units. */
  forward?: number | undefined;
  backward?: number | undefined;
  /** Set the intercept (linear, exponential, polynomial). */
  intercept?: number | undefined;
  displayEquation?: boolean | undefined;
  displayRSquared?: boolean | undefined;
  /** Legend name (default: Excel's "Linear (Series)"). */
  name?: string | undefined;
  color?: Color | undefined;
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
  /**
   * x = category axis (scatter: x values; horizontal bar: the vertical axis), y = value axis,
   * y2 = the secondary value axis of series with `axis: 'secondary'`.
   */
  axes?:
    | { x?: ChartAxis | undefined; y?: ChartAxis | undefined; y2?: ChartAxis | undefined }
    | undefined;
  /** bar, column, line, area: stack the series (of the chart's own type, in combos). */
  stacked?: 'stacked' | 'percentStacked' | undefined;
  /** Data labels of every series (a series' own `dataLabels` wins). */
  dataLabels?: ChartDataLabels | undefined;
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
  groupsOf(chart); // validates combos
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

const TREND: Record<ChartTrendline['type'], string> = {
  linear: 'linear',
  exponential: 'exp',
  logarithmic: 'log',
  polynomial: 'poly',
  power: 'power',
  movingAverage: 'movingAvg',
};

function trendlineXml(t: ChartTrendline): string {
  let x = t.name ? `<c:name>${escapeText(t.name)}</c:name>` : '';
  if (t.color)
    x += `<c:spPr><a:ln w="19050" cap="rnd">${fill(t.color)}<a:prstDash val="sysDot"/></a:ln></c:spPr>`;
  x += val('trendlineType', TREND[t.type] ?? 'linear');
  if (t.type === 'polynomial') x += val('order', Math.min(6, Math.max(2, t.order ?? 2)));
  if (t.type === 'movingAverage') x += val('period', Math.max(2, t.period ?? 2));
  if (t.forward) x += val('forward', t.forward);
  if (t.backward) x += val('backward', t.backward);
  if (t.intercept !== undefined) x += val('intercept', t.intercept);
  x += val('dispRSqr', !!t.displayRSquared) + val('dispEq', !!t.displayEquation);
  if (t.displayRSquared || t.displayEquation) {
    x += '<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/></c:trendlineLbl>';
  }
  return `<c:trendline>${x}</c:trendline>`;
}

function dLblsXml(dl: ChartDataLabels): string {
  let x = dl.numFmt ? `<c:numFmt formatCode="${escapeAttr(dl.numFmt)}" sourceLinked="0"/>` : '';
  if (dl.position) x += val('dLblPos', dl.position);
  x += `${val('showLegendKey', 0)}${val('showVal', !!dl.showValue)}${val('showCatName', !!dl.showCategoryName)}${val('showSerName', !!dl.showSeriesName)}${val('showPercent', !!dl.showPercent)}${val('showBubbleSize', 0)}`;
  return `<c:dLbls>${x}</c:dLbls>`;
}

function series(c: DrawingWriteCtx, ch: Chart, s: ChartSeries, i: number, t: ChartType): string {
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
  const dl = s.dataLabels === false ? undefined : (s.dataLabels ?? ch.dataLabels);
  if (dl) x += dLblsXml(dl);
  if (t !== 'pie' && t !== 'doughnut' && s.trendline) {
    for (const tl of Array.isArray(s.trendline) ? s.trendline : [s.trendline])
      x += trendlineXml(tl);
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
  opts: { hidden?: boolean; crosses?: string } = {},
): string {
  let scaling = val('orientation', 'minMax');
  if (a?.max !== undefined) scaling += val('max', a.max);
  if (a?.min !== undefined) scaling += val('min', a.min);
  const vertical = pos === 'l' || pos === 'r';
  let x = `${val('axId', id)}<c:scaling>${scaling}</c:scaling>${val('delete', !!opts.hidden)}${val('axPos', pos)}`;
  if (value && vertical && !opts.crosses) x += '<c:majorGridlines/>';
  if (a?.title) x += title(a.title, vertical ? ' rot="-5400000" vert="horz"' : '');
  x += `<c:numFmt formatCode="${escapeAttr(a?.numFmt ?? 'General')}" sourceLinked="${a?.numFmt ? 0 : 1}"/>${val('majorTickMark', 'out')}${val('minorTickMark', 'none')}${val('tickLblPos', 'nextTo')}${val('crossAx', cross)}${val('crosses', opts.crosses ?? 'autoZero')}`;
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

/** One chart group of the plot area: its type, axis side and series (with their indices). */
interface Group {
  type: ChartType;
  secondary: boolean;
  series: [ChartSeries, number][];
}

/** Split a chart's series into chart groups (combo charts, secondary axes); validates. */
function groupsOf(ch: Chart): Group[] {
  const t = ch.type;
  if (!TAGS[t]) throw new TypeError(`Unsupported chart type: ${t}`);
  const out: Group[] = [];
  ch.series.forEach((s, i) => {
    const st: ChartType = s.type ?? t;
    if (st !== t && (!COMBO.has(st) || !COMBO.has(t))) {
      throw new TypeError(
        `Chart: a ${st} series cannot be combined with a ${t} chart (combos mix column, line and area)`,
      );
    }
    const secondary = s.axis === 'secondary' && t !== 'pie' && t !== 'doughnut';
    let g = out.find((x) => x.type === st && x.secondary === secondary);
    if (!g) {
      g = { type: st, secondary, series: [] };
      out.push(g);
    }
    g.series.push([s, i]);
  });
  return out;
}

const COMBO = new Set<ChartType>(['column', 'line', 'area']);

/** One `<c:barChart>`-like group element. */
function groupXml(c: DrawingWriteCtx, ch: Chart, g: Group, ax: [number, number]): string {
  const t = g.type;
  const tag = TAGS[t];
  const pie = t === 'pie' || t === 'doughnut';
  const bar = t === 'bar' || t === 'column';
  // stacking applies to the chart's own type (combo series keep their default grouping)
  const st = ch.stacked && !pie && t !== 'scatter' && t === ch.type ? ch.stacked : undefined;
  let x = '';
  if (bar) x += val('barDir', t === 'bar' ? 'bar' : 'col');
  if (t === 'scatter')
    x += val('scatterStyle', g.series.some(([s]) => s.smooth) ? 'smoothMarker' : 'lineMarker');
  else if (!pie) x += val('grouping', st ?? (bar ? 'clustered' : 'standard'));
  x += val('varyColors', ch.varyColors ?? pie);
  for (const [s, i] of g.series) x += series(c, ch, s, i, t);
  if (bar) x += val('gapWidth', 150) + (st ? val('overlap', 100) : '');
  if (t === 'line') x += val('marker', 1);
  if (pie) x += val('firstSliceAng', 0) + (t === 'doughnut' ? val('holeSize', 50) : '');
  if (!pie) x += val('axId', ax[0]) + val('axId', ax[1]);
  return `<c:${tag}>${x}</c:${tag}>`;
}

/** The chart part (c:chartSpace). */
function chartXml(c: DrawingWriteCtx, ch: Chart): string {
  const t = ch.type;
  const groups = groupsOf(ch);
  const pie = t === 'pie' || t === 'doughnut';
  const two = groups.some((g) => g.secondary);
  let g = '';
  for (const gr of groups) g += groupXml(c, ch, gr, gr.secondary ? [30, 40] : [10, 20]);
  let axes = '';
  if (!pie) {
    const ax = ch.axes;
    const h = t === 'bar'; // horizontal bars: category axis on the left
    const between = groups.every((x) => x.type === 'scatter' || x.type === 'area')
      ? 'midCat'
      : 'between';
    const xTag = t === 'scatter' ? 'valAx' : 'catAx';
    axes =
      axis(xTag, 10, 20, h ? 'l' : 'b', ax?.x, t === 'scatter', between) +
      axis('valAx', 20, 10, h ? 'b' : 'l', ax?.y, true, between);
    if (two) {
      // secondary: a hidden category axis and a value axis on the other side
      axes +=
        axis(xTag, 30, 40, h ? 'l' : 'b', undefined, t === 'scatter', between, { hidden: true }) +
        axis('valAx', 40, 30, h ? 't' : 'r', ax?.y2, true, between, { crosses: 'max' });
    }
  }
  const lg = ch.legend;
  const legend =
    lg === false
      ? ''
      : `<c:legend>${val('legendPos', lg?.position ?? 'r')}<c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS_C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${val('date1904', !!c.workbook.date1904)}${val('roundedCorners', 0)}${ch.style ? val('style', ch.style) : ''}<c:chart>${ch.title ? title(ch.title) + val('autoTitleDeleted', 0) : val('autoTitleDeleted', 1)}<c:plotArea><c:layout/>${g}${axes}</c:plotArea>${legend}${val('plotVisOnly', 1)}${val('dispBlanksAs', 'gap')}</c:chart></c:chartSpace>`;
}

/** The sheet's charts: chart parts + graphicFrame anchors in the shared drawing. */
function charts(c: DrawingWriteCtx): string {
  const list = c.ws.charts ?? [];
  let body = '';
  for (let i = 0; i < list.length; i++) {
    const ch = list[i]!;
    // a chart read from a file and not changed since: its original part
    const kept = presPart(c.ws, ch);
    let id: string;
    if (kept) id = addPRel(kept, c.drawingPath, c.drawingRel);
    else {
      const n = c.nextIndex('chart');
      c.addPart(
        `xl/charts/chart${n}.xml`,
        chartXml(c, ch),
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      );
      id = c.drawingRel('chart', `../charts/chart${n}.xml`);
    }
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

// ---------------------------------------------------------------- reading

const GROUPS: Record<string, ChartType | 'bar*'> = {
  barChart: 'bar*',
  lineChart: 'line',
  pieChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  areaChart: 'area',
};
const TREND_IN: Record<string, ChartTrendline['type']> = {
  linear: 'linear',
  exp: 'exponential',
  log: 'logarithmic',
  poly: 'polynomial',
  power: 'power',
  movingAvg: 'movingAverage',
};

interface Node {
  name: string;
  a: Record<string, string>;
  kids: Node[];
  text: string;
}

/** A small DOM of the chart part (local names). */
function tree(xml: string): Node {
  const root: Node = { name: '#root', a: {}, kids: [], text: '' };
  const stack = [root];
  parseXml(xml, {
    open(name, a) {
      const n: Node = { name: localName(name), a: { ...a }, kids: [], text: '' };
      stack[stack.length - 1]!.kids.push(n);
      stack.push(n);
    },
    text(t) {
      stack[stack.length - 1]!.text += t;
    },
    close() {
      stack.pop();
    },
  });
  return root;
}

const kid = (n: Node | undefined, name: string): Node | undefined =>
  n?.kids.find((k) => k.name === name);
const kids = (n: Node | undefined, name: string): Node[] =>
  n?.kids.filter((k) => k.name === name) ?? [];
const valOf = (n: Node | undefined, name: string): string | undefined => kid(n, name)?.a.val;
const on = (v: string | undefined): boolean => v === '1' || v === 'true';

/** All `a:t` texts under a node (rich text titles). */
function texts(n: Node | undefined): string {
  if (!n) return '';
  if (n.name === 't') return n.text;
  return n.kids.map(texts).join('');
}

function color(spPr: Node | undefined): Color | undefined {
  const fillOf = (n: Node | undefined): Node | undefined => kid(n, 'solidFill');
  const f = fillOf(spPr) ?? fillOf(kid(spPr, 'ln'));
  const rgb = kid(f, 'srgbClr')?.a.val;
  if (rgb) return { argb: `FF${rgb.toUpperCase()}` };
  const scheme = kid(f, 'schemeClr')?.a.val;
  const i = scheme ? SCHEME.indexOf(scheme) : -1;
  return i >= 0 ? { theme: i } : undefined;
}

const refOf = (n: Node | undefined): string | undefined =>
  kid(kid(n, 'numRef') ?? kid(n, 'strRef'), 'f')?.text;

function labels(d: Node | undefined): ChartDataLabels | false | undefined {
  if (!d) return undefined;
  if (on(valOf(d, 'delete'))) return false;
  const out: ChartDataLabels = {};
  if (on(valOf(d, 'showVal'))) out.showValue = true;
  if (on(valOf(d, 'showPercent'))) out.showPercent = true;
  if (on(valOf(d, 'showCatName'))) out.showCategoryName = true;
  if (on(valOf(d, 'showSerName'))) out.showSeriesName = true;
  const pos = valOf(d, 'dLblPos');
  if (pos) out.position = pos as ChartDataLabels['position'];
  const nf = kid(d, 'numFmt');
  if (nf && nf.a.sourceLinked !== '1' && nf.a.formatCode) out.numFmt = nf.a.formatCode;
  return Object.keys(out).length ? out : false;
}

function trend(t: Node): ChartTrendline {
  const out: ChartTrendline = { type: TREND_IN[valOf(t, 'trendlineType') ?? ''] ?? 'linear' };
  const num = (k: string): number | undefined => {
    const v = valOf(t, k);
    return v === undefined ? undefined : +v;
  };
  const name = kid(t, 'name')?.text;
  if (name) out.name = name;
  const c = color(kid(t, 'spPr'));
  if (c) out.color = c;
  if (out.type === 'polynomial') out.order = num('order') ?? 2;
  if (out.type === 'movingAverage') out.period = num('period') ?? 2;
  if (num('forward')) out.forward = num('forward');
  if (num('backward')) out.backward = num('backward');
  if (num('intercept') !== undefined) out.intercept = num('intercept');
  if (on(valOf(t, 'dispEq'))) out.displayEquation = true;
  if (on(valOf(t, 'dispRSqr'))) out.displayRSquared = true;
  return out;
}

function axisOf(ax: Node | undefined): ChartAxis | undefined {
  if (!ax) return undefined;
  const out: ChartAxis = {};
  const t = texts(kid(ax, 'title'));
  if (t) out.title = t;
  const nf = kid(ax, 'numFmt');
  if (nf && nf.a.sourceLinked !== '1' && nf.a.formatCode && nf.a.formatCode !== 'General') {
    out.numFmt = nf.a.formatCode;
  }
  const sc = kid(ax, 'scaling');
  const mn = valOf(sc, 'min');
  const mx = valOf(sc, 'max');
  if (mn !== undefined) out.min = +mn;
  if (mx !== undefined) out.max = +mx;
  return Object.keys(out).length ? out : undefined;
}

/** The model of a chart part, or undefined when the model cannot express it (kept raw). */
export function parseChart(xml: string, anchor: AnchorInput, name?: string): Chart | undefined {
  const root = tree(xml);
  const space = kid(root, 'chartSpace');
  const chart = kid(space, 'chart');
  const plot = kid(chart, 'plotArea');
  if (!plot) return undefined;
  const groups = plot.kids.filter((k) => k.name.endsWith('Chart'));
  if (groups.length === 0 || groups.some((g) => !(g.name in GROUPS))) return undefined;
  const typeOf = (g: Node): ChartType => {
    const t = GROUPS[g.name]!;
    return t === 'bar*' ? (valOf(g, 'barDir') === 'bar' ? 'bar' : 'column') : t;
  };
  const main = typeOf(groups[0]!);
  const axIds = (g: Node) => kids(g, 'axId').map((a) => a.a.val);
  const primary = axIds(groups[0]!);
  const out: Chart = { type: main, series: [], anchor };
  if (name) out.name = name;
  const title = kid(chart, 'title');
  if (title) {
    const t = texts(title);
    if (t) out.title = t;
  }
  const grouping = valOf(groups[0], 'grouping');
  if (grouping === 'stacked' || grouping === 'percentStacked') out.stacked = grouping;
  const vary = valOf(groups[0], 'varyColors');
  const pie = main === 'pie' || main === 'doughnut';
  if (vary !== undefined && on(vary) !== pie) out.varyColors = on(vary);
  const style = valOf(space, 'style');
  if (style) out.style = +style;
  const series: [number, ChartSeries][] = [];
  for (const g of groups) {
    const t = typeOf(g);
    if (t !== main && (!COMBO.has(t) || !COMBO.has(main))) return undefined;
    const secondary = !pie && axIds(g).join() !== primary.join();
    for (const se of kids(g, 'ser')) {
      const values = refOf(kid(se, t === 'scatter' ? 'yVal' : 'val'));
      if (!values) return undefined; // literal data: not modelled
      const s: ChartSeries = { values };
      const tx = kid(se, 'tx');
      const nref = refOf(tx);
      if (nref) s.name = { ref: nref };
      else if (kid(tx, 'v')) s.name = kid(tx, 'v')!.text;
      const cats = refOf(kid(se, t === 'scatter' ? 'xVal' : 'cat'));
      if (cats) {
        if (t === 'scatter') s.xValues = cats;
        else s.categories = cats;
      }
      const c = color(kid(se, 'spPr'));
      if (c) s.color = c;
      if (valOf(kid(se, 'marker'), 'symbol') === 'none') s.marker = false;
      if (on(valOf(se, 'smooth'))) s.smooth = true;
      if (t === 'scatter') {
        const ln = kid(kid(se, 'spPr'), 'ln');
        if (!kid(ln, 'noFill')) s.line = true;
      }
      if (t !== main) s.type = t as ChartSeries['type'];
      if (secondary) s.axis = 'secondary';
      const tl = kids(se, 'trendline').map(trend);
      if (tl.length) s.trendline = tl.length === 1 ? tl[0] : tl;
      const dl = labels(kid(se, 'dLbls')) ?? labels(kid(g, 'dLbls'));
      if (dl) s.dataLabels = dl;
      series.push([+(valOf(se, 'order') ?? series.length), s]);
    }
  }
  out.series = series.sort((a, b) => a[0] - b[0]).map((x) => x[1]);
  if (out.series.length === 0) return undefined;
  if (!pie) {
    const axes = plot.kids.filter(
      (k) => k.name === 'catAx' || k.name === 'valAx' || k.name === 'dateAx',
    );
    const byId = (id: string | undefined) => axes.find((a) => valOf(a, 'axId') === id);
    // the first pair of the first group: x (category / scatter x) and y (values)
    let xa = byId(primary[0]);
    let ya = byId(primary[1]);
    if (xa?.name === 'valAx' && ya?.name !== 'valAx') [xa, ya] = [ya, xa];
    if (main === 'scatter' && valOf(xa, 'axPos') === 'l') [xa, ya] = [ya, xa];
    const ax: NonNullable<Chart['axes']> = {};
    const x = axisOf(xa);
    const y = axisOf(ya);
    if (x) ax.x = x;
    if (y) ax.y = y;
    const sec = groups.find((g) => axIds(g).join() !== primary.join());
    if (sec) {
      const ids = axIds(sec);
      const y2 = axisOf(
        axes.find(
          (a) => ids.includes(valOf(a, 'axId')) && a.name === 'valAx' && valOf(a, 'delete') !== '1',
        ),
      );
      if (y2) ax.y2 = y2;
    }
    if (Object.keys(ax).length) out.axes = ax;
  }
  const lg = kid(chart, 'legend');
  if (!lg) out.legend = false;
  else {
    const pos = valOf(lg, 'legendPos') ?? 'r';
    if (pos !== 'r') out.legend = { position: pos as 'l' | 't' | 'b' | 'tr' };
  }
  return out;
}

/** Chart anchors of a drawing part: graphicFrame + c:chart r:id. */
function chartAnchors(
  xml: string,
): { anchor: ImageAnchor; rid: string; name?: string | undefined }[] {
  const root = tree(xml);
  const out: { anchor: ImageAnchor; rid: string; name?: string | undefined }[] = [];
  const dr = root.kids[0];
  const pos = (n: Node | undefined): CellPosition | undefined => {
    if (!n) return undefined;
    const p: CellPosition = { col: +(kid(n, 'col')?.text ?? 0), row: +(kid(n, 'row')?.text ?? 0) };
    const co = +(kid(n, 'colOff')?.text ?? 0);
    const ro = +(kid(n, 'rowOff')?.text ?? 0);
    if (co) p.colOff = co / 9525;
    if (ro) p.rowOff = ro / 9525;
    return p;
  };
  for (const a of dr?.kids ?? []) {
    const gf = kid(a, 'graphicFrame');
    const ch = kid(kid(kid(gf, 'graphic'), 'graphicData'), 'chart');
    const rid = ch && Object.entries(ch.a).find(([k]) => k.endsWith(':id'))?.[1];
    if (!rid) continue;
    let anchor: ImageAnchor | undefined;
    const from = pos(kid(a, 'from'));
    const px = (v: string | undefined) => (+(v ?? 0) || 0) / 9525;
    const ext = kid(a, 'ext');
    if (a.name === 'twoCellAnchor' && from) {
      const to = pos(kid(a, 'to'));
      if (to) anchor = { from, to };
    } else if (a.name === 'oneCellAnchor' && from && ext) {
      anchor = { from, ext: { width: px(ext.a.cx), height: px(ext.a.cy) } };
    } else if (a.name === 'absoluteAnchor' && ext) {
      const p = kid(a, 'pos');
      anchor = {
        pos: { x: px(p?.a.x), y: px(p?.a.y) },
        ext: { width: px(ext.a.cx), height: px(ext.a.cy) },
      };
    }
    if (!anchor) continue;
    const nv = kid(kid(gf, 'nvGraphicFramePr'), 'cNvPr');
    out.push({ anchor, rid, name: nv?.a.name });
  }
  return out;
}

/**
 * Read half of the charts feature: charts the model can express become `ws.charts` (their
 * original parts are written back while unchanged); others stay raw. Always on in the package
 * root; in `./lite` pass it in `features` of readXlsx.
 */
export const chartsReader: FeatureReader = {
  name: 'charts',
  sheet: {
    finish(ctx: SheetReadCtx) {
      const dr = [...(ctx.rels?.values() ?? [])].find(
        (r) => !r.external && r.type.endsWith('/drawing'),
      );
      if (!dr) return;
      const xml = ctx.readText(dr.path);
      if (!xml) return;
      const drels = ctx.relsOf(dr.path);
      const found: [string, Chart][] = [];
      try {
        for (const a of chartAnchors(xml)) {
          const r = drels?.get(a.rid);
          if (!r || r.external) continue;
          const cx = ctx.readText(r.path);
          const ch = cx ? parseChart(cx, a.anchor, a.name) : undefined;
          if (ch) found.push([r.path, ch]);
        }
      } catch {
        return; // a drawing / chart we cannot parse stays raw
      }
      if (found.length === 0) return;
      // claim only once every chart parsed: claimed anchors leave the kept drawing
      for (const [path, ch] of found) ctx.claim(path, ch);
      ctx.ws.charts = [...(ctx.ws.charts ?? []), ...found.map((f) => f[1])];
    },
  },
};
