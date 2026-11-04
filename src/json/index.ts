// JSON helpers: convert worksheets to/from plain objects, arrays and NDJSON.
// Pure functions only (no top-level side effects), so bundlers can drop unused exports.

import { type CellRange, type CellValue, decodeRange, type Worksheet } from '../core/index.js';

export interface ToObjectsOptions {
  /** Header row number (1-based, default 1) or explicit keys (column i maps to the i-th key). */
  header?: number | readonly string[] | undefined;
  /** Region to read, e.g. "A2:D100". Default: the used range of the sheet. */
  range?: string | undefined;
  /** Value for empty cells. Default: the key is omitted. */
  defval?: CellValue;
  /** true (default): keep Date/number/CellValue objects. false: plain JSON-safe values. */
  raw?: boolean | undefined;
}

const RESERVED_KEYS: ReadonlySet<string> = new Set(['__proto__', 'constructor', 'prototype']);

/** Convert any cell value to a plain JSON-safe value. */
export function toJsonValue(v: CellValue): unknown {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'boolean') return v;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v !== 'object') return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.toISOString();
  if ('error' in v) return v.error;
  if ('formula' in v) return v.result === undefined ? null : toJsonValue(v.result);
  if ('richText' in v) return v.richText.map((run) => run.text).join('');
  if ('hyperlink' in v) return v.text;
  return null;
}

/** Convert an arbitrary JS value into something the worksheet can store. */
/** How `objectsToSheet` / `arraysToSheet` store `''`: no cell (default) or an empty-string cell. */
export type EmptyAs = 'empty' | 'string';

function toCell(v: unknown, emptyAs?: EmptyAs): CellValue {
  if (v === null || v === undefined || (v === '' && emptyAs !== 'string')) return null;
  switch (typeof v) {
    case 'string':
    case 'boolean':
      return v;
    case 'number':
      return Number.isFinite(v) ? v : null;
    case 'bigint':
      return v.toString();
    case 'object':
      if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
      return JSON.stringify(v);
    default:
      return null;
  }
}

/** Text of a header cell. Empty cells become ''. */
function headerText(v: CellValue): string {
  const j = toJsonValue(v);
  return j === null ? '' : String(j);
}

/** Sanitize header texts into unique, prototype-safe keys. */
function headerKeys(texts: readonly string[], firstCol: number): string[] {
  const seen = new Set<string>();
  return texts.map((t, i) => {
    let k = t === '' ? `__EMPTY_${firstCol + i}` : RESERVED_KEYS.has(t) ? `_${t}` : t;
    if (seen.has(k)) {
      const base = k;
      let n = 2;
      while (seen.has(`${base}_${n}`)) n++;
      k = `${base}_${n}`;
    }
    seen.add(k);
    return k;
  });
}

/** 0-based inclusive bounds of the region to read, or undefined for an empty sheet. */
function bounds(ws: Worksheet, range: string | undefined): CellRange | undefined {
  if (range !== undefined) return decodeRange(range);
  const d = ws.dimensions;
  return d
    ? { top: d.top - 1, left: d.left - 1, bottom: d.bottom - 1, right: d.right - 1 }
    : undefined;
}

function checkHeaderRow(n: number): number {
  if (!Number.isInteger(n) || n < 1)
    throw new RangeError(`Header row must be an integer >= 1: ${n}`);
  return n;
}

/** Lazily yields one object per non-blank data row. */
function* objectRows(ws: Worksheet, o: ToObjectsOptions): Generator<Record<string, unknown>> {
  const hdr = o.header;
  // Header: either a row to read keys from, or explicit keys (no header row in the data).
  const headerRow =
    hdr === undefined ? 1 : typeof hdr === 'number' ? checkHeaderRow(hdr) : undefined;
  const b = bounds(ws, o.range);
  if (!b) return;
  const raw = o.raw !== false;
  const dv = o.defval === undefined ? undefined : raw ? o.defval : toJsonValue(o.defval);
  const firstCol = b.left + 1;

  let keys: string[];
  let cols: number[];
  if (headerRow !== undefined) {
    const hv = ws.getRow(headerRow).values;
    cols = [];
    for (let c = firstCol; c <= b.right + 1; c++) cols.push(c);
    keys = headerKeys(
      cols.map((c) => headerText(hv[c] ?? null)),
      firstCol,
    );
  } else {
    const explicit = hdr as readonly string[];
    keys = headerKeys(
      explicit.map((k) => String(k)),
      firstCol,
    );
    cols = keys.map((_, i) => firstCol + i);
  }
  if (keys.length === 0) return;

  for (let r = b.top + 1; r <= b.bottom + 1; r++) {
    // Data rows start below the header row; rows above it are not data.
    if (headerRow !== undefined && r <= headerRow) continue;
    const vals = ws.getRow(r).values;
    const obj: Record<string, unknown> = {};
    let any = false;
    for (let i = 0; i < keys.length; i++) {
      const v = vals[cols[i]!];
      const key = keys[i]!;
      if (v === null || v === undefined) {
        if (dv !== undefined) obj[key] = dv;
      } else {
        any = true;
        obj[key] = raw ? v : toJsonValue(v);
      }
    }
    if (any) yield obj;
  }
}

/** Read rows as objects keyed by header. Blank data rows are skipped. */
export function sheetToObjects(ws: Worksheet, o: ToObjectsOptions = {}): Record<string, unknown>[] {
  return Array.from(objectRows(ws, o));
}

/** Read a region as a 2D array (empty cells are null). Row/column order follows the sheet. */
export function sheetToArrays(
  ws: Worksheet,
  o: { range?: string | undefined; raw?: boolean | undefined } = {},
): unknown[][] {
  const b = bounds(ws, o.range);
  if (!b) return [];
  const raw = o.raw !== false;
  const out: unknown[][] = [];
  for (let r = b.top + 1; r <= b.bottom + 1; r++) {
    const vals = ws.getRow(r).values;
    const row: unknown[] = [];
    for (let c = b.left + 1; c <= b.right + 1; c++) {
      const v = vals[c] ?? null;
      row.push(raw ? v : toJsonValue(v));
    }
    out.push(row);
  }
  return out;
}

/**
 * Write objects as rows. Header = `o.header` or the union of keys in first-seen order.
 * Nested objects/arrays are written as JSON strings. Materializes `rows` (needed for the
 * union of keys). Empty strings leave the cell empty unless `emptyAs: 'string'`.
 */
export function objectsToSheet(
  ws: Worksheet,
  rows: Iterable<Record<string, unknown>>,
  o: {
    header?: readonly string[] | undefined;
    writeHeader?: boolean | undefined;
    emptyAs?: EmptyAs | undefined;
  } = {},
): Worksheet {
  const list = Array.from(rows);
  const keys = o.header ? [...o.header] : unionKeys(list);
  if (keys.length === 0) return ws;
  if (o.writeHeader !== false) ws.addRow(keys);
  for (const row of list) {
    ws.addRow(keys.map((k) => (Object.hasOwn(row, k) ? toCell(row[k], o.emptyAs) : null)));
  }
  return ws;
}

function unionKeys(list: readonly Record<string, unknown>[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const row of list) {
    for (const k of Object.keys(row)) {
      if (!seen.has(k)) {
        seen.add(k);
        out.push(k);
      }
    }
  }
  return out;
}

/**
 * Append arrays as rows (index 0 maps to column A). Empty strings (e.g. empty CSV fields from
 * `parseCsv`) leave the cell empty unless `emptyAs: 'string'`.
 */
export function arraysToSheet(
  ws: Worksheet,
  rows: Iterable<readonly unknown[]>,
  o: { emptyAs?: EmptyAs | undefined } = {},
): Worksheet {
  for (const row of rows) ws.addRow(row.map((v) => toCell(v, o.emptyAs)));
  return ws;
}

/** One JSON object per row, each yielded with a trailing newline (raw = false). */
export async function* sheetToNdjson(
  ws: Worksheet,
  o: ToObjectsOptions = {},
): AsyncGenerator<string> {
  for (const row of objectRows(ws, { ...o, raw: false })) {
    yield `${JSON.stringify(row)}\n`;
  }
}
