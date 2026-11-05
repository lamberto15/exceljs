// Feature: pivot table authoring (write) — xl/pivotCache/pivotCacheDefinitionN.xml + its records
// (xl/pivotCache/pivotCacheRecordsN.xml, linked from workbook.xml <pivotCaches>) and
// xl/pivotTables/pivotTableN.xml (linked from the sheet).
//
//   addPivotTable(ws, {
//     name: 'Sales by region', source: 'Data!A1:E500', ref: 'A3',
//     rows: ['Region', 'Product'], columns: ['Quarter'], values: [{ field: 'Units', fn: 'sum' }],
//     filters: ['Year'],
//   });
//
// Design: the cache holds every source column (shared items for the fields on an axis, numbers
// inline otherwise) and its records, and is marked refreshOnLoad, so Excel recomputes the
// report when it opens the file. `addPivotTable` also computes the report and writes its cells
// (Excel's compact layout: row labels in one column, nested fields indented below their parent,
// subtotals on the parent row, grand totals), so readers that do not refresh pivots show the
// totals too. Supported: one or more row fields, at most one column field, one or more value
// fields (not combined with a column field), report filters (all items shown).
// Pivot tables of read files are kept raw (round-trip preservation); adding one to a workbook
// that was read with pivot tables is not supported yet (their caches share workbook.xml).

import { decodeAddr, decodeRange, encodeAddr, encodeRange } from '../../core/address.js';
import type { CellValue } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';
import { BookSlot, type FeatureWriter, SheetSlot, type SheetWriteCtx } from './types.js';

export type PivotFunction =
  | 'sum'
  | 'count'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'countNums'
  | 'stdDev'
  | 'stdDevp'
  | 'var'
  | 'varp';

export interface PivotValue {
  /** Source column (header text). */
  field: string;
  /** Summary function (default 'sum'). */
  fn?: PivotFunction | undefined;
  /** Caption (default Excel's: "Sum of Units"). */
  name?: string | undefined;
}

/** A pivot table on a sheet (`ws.pivotTables`). */
export interface PivotTable {
  /** Unique name, e.g. "PivotTable1". */
  name: string;
  /** Source range with a header row: "Data!A1:E500" (unqualified = this sheet). */
  source: string;
  /** Top-left cell of the report (report filters, when any, start here). */
  ref: string;
  /** Row fields (source column headers), outermost first. */
  rows: string[];
  /** Column field (at most one in this version). */
  columns?: string[] | undefined;
  /** Value fields. */
  values: PivotValue[];
  /** Report filter fields (shown above the report; all items selected). */
  filters?: string[] | undefined;
  /** Pivot style (default "PivotStyleLight16"). */
  style?: string | undefined;
}

const LABEL: Record<PivotFunction, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
  product: 'Product',
  countNums: 'Count',
  stdDev: 'StdDev',
  stdDevp: 'StdDevp',
  var: 'Var',
  varp: 'Varp',
};

// ---------------------------------------------------------------- source data

type V = string | number | boolean | null;

interface Source {
  sheet: Worksheet;
  range: ReturnType<typeof decodeRange>;
  headers: string[];
  /** Records (one array per data row, one value per column). */
  rows: V[][];
}

function scalar(v: CellValue): V {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if ('formula' in v) {
    const r = v.result;
    return r === undefined || (typeof r === 'object' && r !== null) ? null : r;
  }
  if ('richText' in v) return v.richText.map((r) => r.text).join('');
  if ('hyperlink' in v) return v.text;
  return null;
}

function readSource(ws: Worksheet, src: string): Source {
  let s = src.trim().replace(/^=/, '');
  let sheet = ws;
  const bang = s.lastIndexOf('!');
  if (bang >= 0) {
    let name = s.slice(0, bang);
    if (name[0] === "'") name = name.slice(1, -1).replace(/''/g, "'");
    const t = ws.workbook.getWorksheet(name);
    if (!t) throw new Error(`Pivot table source: no sheet "${name}"`);
    sheet = t;
    s = s.slice(bang + 1);
  }
  const range = decodeRange(s.replace(/\$/g, ''));
  if (range.bottom <= range.top) throw new Error('Pivot table source needs a header row and data');
  const headers: string[] = [];
  for (let c = range.left; c <= range.right; c++) {
    const h = scalar(sheet.getValue(range.top + 1, c + 1));
    if (h === null || h === '')
      throw new Error(`Pivot table source: empty header in column ${c + 1}`);
    headers.push(String(h));
  }
  const rows: V[][] = [];
  for (let r = range.top + 1; r <= range.bottom; r++) {
    const rec: V[] = [];
    for (let c = range.left; c <= range.right; c++) rec.push(scalar(sheet.getValue(r + 1, c + 1)));
    rows.push(rec);
  }
  return { sheet, range, headers, rows };
}

/** Excel's item order: numbers, then text (case-insensitive), then booleans; blank last. */
function cmp(a: V, b: V): number {
  const rank = (x: V) =>
    x === null ? 3 : typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return a === b ? 0 : a ? 1 : -1;
}

const key = (v: V): string =>
  v === null ? 'm' : `${typeof v}:${typeof v === 'string' ? v.toLowerCase() : v}`;

// ---------------------------------------------------------------- layout

interface Field {
  name: string;
  col: number;
  /** Unique values in first-seen order (the cache's shared items). */
  shared: V[];
  /** Display order: indices into `shared`, sorted. */
  order: number[];
}

/** Excel's 15 significant digits (sums of decimals: 2.05, not 2.0500000000000003). */
const r15 = (x: number | null): number | null => (x === null ? null : Number(x.toPrecision(15)));

function aggregate(fn: PivotFunction, xs: V[]): number | null {
  return r15(raw(fn, xs));
}

function raw(fn: PivotFunction, xs: V[]): number | null {
  const nums = xs.filter((x): x is number => typeof x === 'number');
  const n = nums.length;
  const sum = nums.reduce((p, x) => p + x, 0);
  const variance = (sample: boolean): number | null => {
    if (n < (sample ? 2 : 1)) return null;
    const m = sum / n;
    return nums.reduce((p, x) => p + (x - m) ** 2, 0) / (sample ? n - 1 : n);
  };
  switch (fn) {
    case 'sum':
      return sum;
    case 'count':
      return xs.filter((x) => x !== null).length;
    case 'countNums':
      return n;
    case 'average':
      return n ? sum / n : null;
    case 'max':
      return n ? Math.max(...nums) : null;
    case 'min':
      return n ? Math.min(...nums) : null;
    case 'product':
      return n ? nums.reduce((p, x) => p * x, 1) : null;
    case 'stdDev':
    case 'stdDevp': {
      const v = variance(fn === 'stdDev');
      return v === null ? null : Math.sqrt(v);
    }
    case 'var':
    case 'varp':
      return variance(fn === 'var');
  }
}

interface Layout {
  src: Source;
  fields: Field[];
  rowF: Field[];
  colF: Field | undefined;
  pageF: Field[];
  values: { f: Field; fn: PivotFunction; name: string }[];
  /** Row items: path of item indices (into each row field's `order`), or 'grand'. */
  rowItems: (number[] | 'grand')[];
  /** Column items (order indices of the column field), or 'grand'; [] without a column field. */
  colItems: (number | 'grand')[];
  /** Top-left of the table (below the filters), 0-based. */
  top: number;
  left: number;
  width: number;
  height: number;
  headerRows: number;
}

function layoutOf(ws: Worksheet, p: PivotTable): Layout {
  if (!p.rows?.length)
    throw new Error(`Pivot table "${p.name}": at least one row field is required`);
  if (!p.values?.length)
    throw new Error(`Pivot table "${p.name}": at least one value field is required`);
  const cols = p.columns ?? [];
  if (cols.length > 1)
    throw new Error(`Pivot table "${p.name}": one column field at most (this version)`);
  if (cols.length && p.values.length > 1) {
    throw new Error(
      `Pivot table "${p.name}": several value fields with a column field are not supported yet`,
    );
  }
  const src = readSource(ws, p.source);
  const field = (name: string): Field => {
    const col = src.headers.findIndex((h) => h.toLowerCase() === name.toLowerCase());
    if (col < 0) throw new Error(`Pivot table "${p.name}": no source column "${name}"`);
    return fields[col]!;
  };
  const fields: Field[] = src.headers.map((name, col) => {
    const seen = new Map<string, number>();
    const shared: V[] = [];
    for (const r of src.rows) {
      const k = key(r[col]!);
      if (!seen.has(k)) {
        seen.set(k, shared.length);
        shared.push(r[col]!);
      }
    }
    const order = shared.map((_, i) => i).sort((a, b) => cmp(shared[a]!, shared[b]!));
    return { name, col, shared, order };
  });
  const rowF = p.rows.map(field);
  const colF = cols.length ? field(cols[0]!) : undefined;
  const pageF = (p.filters ?? []).map(field);
  const used = [...rowF, ...(colF ? [colF] : []), ...pageF];
  if (new Set(used).size !== used.length) {
    throw new Error(`Pivot table "${p.name}": a field can be on one axis only`);
  }
  const values = p.values.map((v) => {
    const f = field(v.field);
    const fn = v.fn ?? 'sum';
    return { f, fn, name: v.name ?? `${LABEL[fn]} of ${f.name}` };
  });
  // row items: depth-first over the combinations present in the data
  const rowItems: (number[] | 'grand')[] = [];
  const pos = (f: Field, v: V): number =>
    f.order.indexOf(f.shared.findIndex((x) => key(x) === key(v)));
  const walk = (depth: number, recs: V[][], path: number[]): void => {
    const f = rowF[depth]!;
    const groups = new Map<number, V[][]>();
    for (const r of recs) {
      const i = pos(f, r[f.col]!);
      const g = groups.get(i);
      if (g) g.push(r);
      else groups.set(i, [r]);
    }
    for (const i of [...groups.keys()].sort((a, b) => a - b)) {
      rowItems.push([...path, i]);
      if (depth + 1 < rowF.length) walk(depth + 1, groups.get(i)!, [...path, i]);
    }
  };
  walk(0, src.rows, []);
  rowItems.push('grand');
  const colItems: (number | 'grand')[] = [];
  if (colF) {
    const present = new Set(src.rows.map((r) => pos(colF, r[colF.col]!)));
    colItems.push(...[...present].sort((a, b) => a - b), 'grand');
  }
  const at = decodeAddr(p.ref);
  const top = at.row + (pageF.length ? pageF.length + 1 : 0);
  // a column field or several value fields (the "Values" field on columns): two header rows
  const headerRows = colF || values.length > 1 ? 2 : 1;
  const width = 1 + (colF ? colItems.length : values.length);
  return {
    src,
    fields,
    rowF,
    colF,
    pageF,
    values,
    rowItems,
    colItems,
    top,
    left: at.col,
    width,
    height: headerRows + rowItems.length,
    headerRows,
  };
}

/** The report cells (row-major, `height` x `width`) and the filter rows above them. */
function cellsOf(L: Layout): { filters: V[][]; grid: V[][] } {
  const { rowF, colF, values, src } = L;
  const match = (r: V[], path: number[]): boolean =>
    path.every((i, d) => key(r[rowF[d]!.col]!) === key(rowF[d]!.shared[rowF[d]!.order[i]!]!));
  const colMatch = (r: V[], c: number | 'grand'): boolean =>
    c === 'grand' || key(r[colF!.col]!) === key(colF!.shared[colF!.order[c]!]!);
  const label = (v: V): V => (v === null ? '(blank)' : v);
  const grid: V[][] = [];
  if (colF) {
    grid.push([values[0]!.name, 'Column Labels', ...Array(L.colItems.length - 1).fill(null)]);
    grid.push([
      'Row Labels',
      ...L.colItems.map((c) =>
        c === 'grand' ? 'Grand Total' : label(colF.shared[colF.order[c]!]!),
      ),
    ]);
  } else if (values.length > 1) {
    grid.push([null, 'Values', ...Array(values.length - 1).fill(null)]);
    grid.push(['Row Labels', ...values.map((v) => v.name)]);
  } else {
    grid.push(['Row Labels', ...values.map((v) => v.name)]);
  }
  for (const it of L.rowItems) {
    const recs = it === 'grand' ? src.rows : src.rows.filter((r) => match(r, it));
    const head: V =
      it === 'grand'
        ? 'Grand Total'
        : label(rowF[it.length - 1]!.shared[rowF[it.length - 1]!.order[it[it.length - 1]!]!]!);
    const row: V[] = [head];
    if (colF) {
      const v = values[0]!;
      for (const c of L.colItems) {
        row.push(
          aggregate(
            v.fn,
            recs.filter((r) => colMatch(r, c)).map((r) => r[v.f.col]!),
          ),
        );
      }
    } else {
      for (const v of values)
        row.push(
          aggregate(
            v.fn,
            recs.map((r) => r[v.f.col]!),
          ),
        );
    }
    grid.push(row);
  }
  const filters = L.pageF.map((f) => [f.name, '(All)'] as V[]);
  return { filters, grid };
}

/** Add a pivot table: validates it, writes its report cells and appends it to ws.pivotTables. */
export function addPivotTable(ws: Worksheet, p: PivotTable): PivotTable {
  const L = layoutOf(ws, p);
  const all = ws.workbook.worksheets.flatMap((w) => w.pivotTables ?? []);
  if (all.some((x) => x.name.toLowerCase() === p.name.toLowerCase())) {
    throw new Error(`Pivot table name "${p.name}" is used more than once`);
  }
  const { filters, grid } = cellsOf(L);
  const at = decodeAddr(p.ref);
  const area = {
    top: at.row,
    left: L.left,
    bottom: L.top + L.height - 1,
    right: L.left + Math.max(L.width, filters.length ? 2 : 0) - 1,
  };
  for (let r = area.top; r <= area.bottom; r++) {
    for (let c = area.left; c <= area.right; c++) {
      const v = ws.getValue(r + 1, c + 1);
      if (v !== null && v !== undefined) {
        throw new Error(
          `Pivot table "${p.name}": ${encodeAddr(r, c)} is not empty (the report needs ${encodeRange(area)})`,
        );
      }
    }
  }
  filters.forEach((row, i) => {
    row.forEach((v, j) => {
      if (v !== null) ws.setCell(at.row + i + 1, L.left + j + 1, v);
    });
  });
  grid.forEach((row, i) => {
    row.forEach((v, j) => {
      if (v !== null) ws.setCell(L.top + i + 1, L.left + j + 1, v);
    });
  });
  ws.pivotTables ??= [];
  ws.pivotTables.push(p);
  return p;
}

// ---------------------------------------------------------------- writing

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function sharedItemsXml(f: Field, onAxis: boolean): string {
  const vs = f.shared;
  const nums = vs.filter((v): v is number => typeof v === 'number');
  const hasStr = vs.some((v) => typeof v === 'string');
  const hasBlank = vs.some((v) => v === null);
  const hasBool = vs.some((v) => typeof v === 'boolean');
  let a = '';
  if (!hasStr) a += ' containsSemiMixedTypes="0" containsString="0"';
  if (nums.length) {
    a += ' containsNumber="1"';
    if (nums.every(Number.isInteger)) a += ' containsInteger="1"';
    a += ` minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"`;
  }
  if (hasBlank) a += ' containsBlank="1"';
  if (hasBool) a += ' containsMixedTypes="1"';
  if (!onAxis && !hasStr && !hasBool) return `<sharedItems${a}/>`;
  const items = vs
    .map((v) =>
      v === null
        ? '<m/>'
        : typeof v === 'number'
          ? `<n v="${v}"/>`
          : typeof v === 'boolean'
            ? `<b v="${+v}"/>`
            : `<s v="${escapeAttr(v)}"/>`,
    )
    .join('');
  return `<sharedItems${a} count="${vs.length}">${items}</sharedItems>`;
}

/** Fields written as shared items (x) in the records: on an axis, or holding text / booleans. */
const asShared = (f: Field, axis: Set<Field>): boolean =>
  axis.has(f) || f.shared.some((v) => typeof v === 'string' || typeof v === 'boolean');

function cacheXml(L: Layout, recordsRel: string): string {
  const axis = new Set([...L.rowF, ...(L.colF ? [L.colF] : []), ...L.pageF]);
  const src = L.src;
  const ref = encodeRange(src.range);
  let fields = '';
  for (const f of L.fields) {
    fields += `<cacheField name="${escapeAttr(f.name)}" numFmtId="0">${sharedItemsXml(f, asShared(f, axis))}</cacheField>`;
  }
  return `${HEAD}<pivotCacheDefinition xmlns="${NS}" xmlns:r="${NS_R}" r:id="${recordsRel}" refreshOnLoad="1" refreshedBy="turbosheet" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${src.rows.length}"><cacheSource type="worksheet"><worksheetSource ref="${ref}" sheet="${escapeAttr(src.sheet.name)}"/></cacheSource><cacheFields count="${L.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`;
}

function recordsXml(L: Layout): string {
  const axis = new Set([...L.rowF, ...(L.colF ? [L.colF] : []), ...L.pageF]);
  const idx = L.fields.map((f) => new Map(f.shared.map((v, i) => [key(v), i])));
  const shared = L.fields.map((f) => asShared(f, axis));
  let s = '';
  for (const r of L.src.rows) {
    let x = '';
    r.forEach((v, c) => {
      if (shared[c]) x += `<x v="${idx[c]!.get(key(v))}"/>`;
      else x += v === null ? '<m/>' : `<n v="${v}"/>`;
    });
    s += `<r>${x}</r>`;
  }
  return `${HEAD}<pivotCacheRecords xmlns="${NS}" xmlns:r="${NS_R}" count="${L.src.rows.length}">${s}</pivotCacheRecords>`;
}

function tableXml(p: PivotTable, L: Layout, cacheId: number): string {
  const loc = encodeRange({
    top: L.top,
    left: L.left,
    bottom: L.top + L.height - 1,
    right: L.left + L.width - 1,
  });
  const multi = L.values.length > 1;
  const firstHeaderRow = 1;
  const firstDataRow = L.headerRows;
  const page = L.pageF.length ? ` rowPageCount="${L.pageF.length}" colPageCount="1"` : '';
  let pf = '';
  for (const f of L.fields) {
    const isData = L.values.some((v) => v.f === f);
    const axis = L.rowF.includes(f)
      ? 'axisRow'
      : L.colF === f
        ? 'axisCol'
        : L.pageF.includes(f)
          ? 'axisPage'
          : undefined;
    let a = axis ? ` axis="${axis}"` : '';
    if (isData) a += ' dataField="1"';
    a += ' showAll="0"';
    if (axis) {
      const items = f.order.map((i) => `<item x="${i}"/>`).join('');
      pf += `<pivotField${a}><items count="${f.order.length + 1}">${items}<item t="default"/></items></pivotField>`;
    } else pf += `<pivotField${a}/>`;
  }
  let x = `<location ref="${loc}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${firstDataRow}" firstDataCol="1"${page}/>`;
  x += `<pivotFields count="${L.fields.length}">${pf}</pivotFields>`;
  x += `<rowFields count="${L.rowF.length}">${L.rowF.map((f) => `<field x="${f.col}"/>`).join('')}</rowFields>`;
  let ri = '';
  let prev: number[] = [];
  for (const it of L.rowItems) {
    if (it === 'grand') {
      ri += '<i t="grand"><x/></i>';
      continue;
    }
    // r: leading fields repeated from the previous row item
    let r = 0;
    while (r < it.length - 1 && prev[r] === it[r]) r++;
    const xs = it
      .slice(r)
      .map((v) => (v ? `<x v="${v}"/>` : '<x/>'))
      .join('');
    ri += `<i${r ? ` r="${r}"` : ''}>${xs}</i>`;
    prev = it;
  }
  x += `<rowItems count="${L.rowItems.length}">${ri}</rowItems>`;
  if (L.colF) {
    x += `<colFields count="1"><field x="${L.colF.col}"/></colFields>`;
    const ci = L.colItems
      .map((c) =>
        c === 'grand' ? '<i t="grand"><x/></i>' : `<i>${c ? `<x v="${c}"/>` : '<x/>'}</i>`,
      )
      .join('');
    x += `<colItems count="${L.colItems.length}">${ci}</colItems>`;
  } else if (multi) {
    x += '<colFields count="1"><field x="-2"/></colFields>';
    const ci = L.values
      .map((_, i) => (i ? `<i i="${i}"><x v="${i}"/></i>` : '<i><x/></i>'))
      .join('');
    x += `<colItems count="${L.values.length}">${ci}</colItems>`;
  } else x += '<colItems count="1"><i/></colItems>';
  if (L.pageF.length) {
    x += `<pageFields count="${L.pageF.length}">${L.pageF.map((f) => `<pageField fld="${f.col}" hier="-1"/>`).join('')}</pageFields>`;
  }
  const df = L.values
    .map(
      (v) =>
        `<dataField name="${escapeAttr(v.name)}" fld="${v.f.col}"${v.fn === 'sum' ? '' : ` subtotal="${v.fn}"`} baseField="0" baseItem="0"/>`,
    )
    .join('');
  x += `<dataFields count="${L.values.length}">${df}</dataFields>`;
  x += `<pivotTableStyleInfo name="${escapeAttr(p.style ?? 'PivotStyleLight16')}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`;
  return `${HEAD}<pivotTableDefinition xmlns="${NS}" name="${escapeAttr(p.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">${x}</pivotTableDefinition>`;
}

/** Caches written during one write, per workbook: cacheId and the definition part. */
const CACHES = new WeakMap<Workbook, { id: number; path: string }[]>();

function emit(c: SheetWriteCtx): string {
  const list = c.ws.pivotTables;
  if (!list?.length) return '';
  const wb = c.workbook;
  if (wb.preserved?.frags.some((f) => f.slot === BookSlot.pivotCaches)) {
    throw new Error(
      'Adding pivot tables to a workbook read with pivot tables is not supported yet (their caches share workbook.xml)',
    );
  }
  let caches = CACHES.get(wb);
  if (!caches) {
    caches = [];
    CACHES.set(wb, caches);
  }
  for (const p of list) {
    const L = layoutOf(c.ws, p);
    const n = c.nextIndex('pivotCacheDefinition');
    const def = `xl/pivotCache/pivotCacheDefinition${n}.xml`;
    // the cache part and its records (relationship from the definition, written by hand)
    c.addPart(
      `xl/pivotCache/pivotCacheRecords${n}.xml`,
      recordsXml(L),
      `${CT}pivotCacheRecords+xml`,
    );
    c.addPart(def, cacheXml(L, 'rId1'), `${CT}pivotCacheDefinition+xml`);
    c.addPart(
      `xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`,
      `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}pivotCacheRecords" Target="pivotCacheRecords${n}.xml"/></Relationships>`,
    );
    const id = caches.length + 1;
    caches.push({ id, path: `pivotCache/pivotCacheDefinition${n}.xml` });
    const t = c.nextIndex('pivotTable');
    c.addPart(`xl/pivotTables/pivotTable${t}.xml`, tableXml(p, L, id), `${CT}pivotTable+xml`);
    c.addPart(
      `xl/pivotTables/_rels/pivotTable${t}.xml.rels`,
      `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`,
    );
    c.addRel('pivotTable', `../pivotTables/pivotTable${t}.xml`);
  }
  return '';
}

/**
 * Write half of the pivot tables feature (ws.pivotTables). Always on in the package root; in
 * `./lite` pass it in `features`. Pivot tables of read files are kept raw.
 */
export const pivotTablesWriter: FeatureWriter = {
  name: 'pivot-tables',
  sheet: { emitters: [{ slot: SheetSlot.extLst, order: 99, emit }] },
  book: {
    emitters: [
      {
        slot: BookSlot.pivotCaches,
        emit(c) {
          const caches = CACHES.get(c.workbook);
          if (!caches?.length) return '';
          CACHES.delete(c.workbook);
          const items = caches
            .map(
              (x) =>
                `<pivotCache cacheId="${x.id}" r:id="${c.addRel(`${REL}pivotCacheDefinition`, x.path)}"/>`,
            )
            .join('');
          return `<pivotCaches>${items}</pivotCaches>`;
        },
      },
    ],
  },
};
