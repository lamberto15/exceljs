import type { CellValue } from '../core/types.js';
import { CsvError } from './parser.js';

export interface CsvWriteOptions {
  delimiter?: string;
  quote?: string;
  /** Row terminator, default "\r\n" (RFC 4180). */
  eol?: string;
  /** Prefix output with a UTF-8 BOM (default false). */
  bom?: boolean;
  /** Prefix "'" to string values starting with = + - @ TAB CR (CSV injection). Numbers are never altered. */
  escapeFormulas?: boolean;
  /** Date formatter, default toISOString(). */
  dateFormat?: (d: Date) => string;
}

interface Ctx {
  d: string;
  q: string;
  qq: string;
  re: RegExp;
}
let lastKey = '';
let lastCtx: Ctx | undefined;

function ctxFor(o: CsvWriteOptions | undefined): Ctx {
  const d = o?.delimiter ?? ',';
  const q = o?.quote ?? '"';
  const key = d + q;
  if (lastCtx !== undefined && key === lastKey) return lastCtx;
  if (
    d.length !== 1 ||
    q.length !== 1 ||
    d === q ||
    d === '\n' ||
    d === '\r' ||
    q === '\n' ||
    q === '\r'
  ) {
    throw new CsvError('BAD_OPTION', 'delimiter and quote must be distinct single characters');
  }
  const esc = (c: string): string => c.replace(/[\\^$.*+?()[\]{}|\-/]/g, '\\$&');
  const ctx = { d, q, qq: q + q, re: new RegExp(`[${esc(d)}${esc(q)}\\r\\n]|^ | $`) };
  lastKey = key;
  lastCtx = ctx;
  return ctx;
}

function text(v: CellValue, o: CsvWriteOptions | undefined): { s: string; str: boolean } {
  switch (typeof v) {
    case 'string':
      return { s: v, str: true };
    case 'number':
      return { s: String(v), str: false };
    case 'boolean':
      return { s: v ? 'TRUE' : 'FALSE', str: false };
    case 'undefined':
      return { s: '', str: false };
  }
  if (v === null) return { s: '', str: false };
  if (v instanceof Date) {
    if (Number.isNaN(v.getTime())) return { s: '', str: false };
    return { s: o?.dateFormat ? o.dateFormat(v) : v.toISOString(), str: false };
  }
  if ('error' in v) return { s: v.error, str: false };
  if ('formula' in v) return v.result === undefined ? { s: '', str: false } : text(v.result, o);
  if ('richText' in v) {
    let s = '';
    for (const r of v.richText) s += r.text;
    return { s, str: true };
  }
  if ('hyperlink' in v) return { s: v.text, str: true };
  return { s: '', str: false };
}

function field(v: CellValue, o: CsvWriteOptions | undefined, c: Ctx): string {
  let { s, str } = text(v, o);
  if (str && o?.escapeFormulas && s.length > 0) {
    const ch = s.charCodeAt(0);
    if (ch === 61 || ch === 43 || ch === 45 || ch === 64 || ch === 9 || ch === 13) s = "'" + s;
  }
  if (s.length === 0) return '';
  if (!c.re.test(s)) return s;
  return c.q + (s.indexOf(c.q) >= 0 ? s.split(c.q).join(c.qq) : s) + c.q;
}

/** Format one row (no trailing eol). */
export function formatCsvRow(row: readonly CellValue[], o?: CsvWriteOptions): string {
  const c = ctxFor(o);
  const n = row.length;
  if (n === 0) return '';
  let out = field(row[0], o, c);
  // a lone empty field would be an empty line; quote it so it survives a round trip
  if (n === 1 && out === '') return c.qq;
  for (let i = 1; i < n; i++) out += c.d + field(row[i], o, c);
  return out;
}

/** Format rows; each row is followed by eol. */
export function writeCsv(rows: Iterable<readonly CellValue[]>, o?: CsvWriteOptions): string {
  const eol = o?.eol ?? '\r\n';
  const parts: string[] = [];
  if (o?.bom) parts.push('﻿');
  for (const r of rows) parts.push(formatCsvRow(r, o), eol);
  return parts.join('');
}

/** Pull-based byte stream: rows are only consumed as the reader demands. */
export function csvStream(
  rows: AsyncIterable<readonly CellValue[]> | Iterable<readonly CellValue[]>,
  o?: CsvWriteOptions,
): ReadableStream<Uint8Array> {
  const eol = o?.eol ?? '\r\n';
  const enc = new TextEncoder();
  ctxFor(o); // validate eagerly
  const it: AsyncIterator<readonly CellValue[]> | Iterator<readonly CellValue[]> =
    Symbol.asyncIterator in rows
      ? (rows as AsyncIterable<readonly CellValue[]>)[Symbol.asyncIterator]()
      : (rows as Iterable<readonly CellValue[]>)[Symbol.iterator]();
  let first = true;
  const FLUSH = 1 << 16;
  return new ReadableStream<Uint8Array>({
    async pull(ctrl) {
      let buf = first && o?.bom ? '﻿' : '';
      first = false;
      for (;;) {
        const r = await it.next();
        if (r.done) {
          if (buf) ctrl.enqueue(enc.encode(buf));
          ctrl.close();
          return;
        }
        buf += formatCsvRow(r.value, o) + eol;
        if (buf.length >= FLUSH) {
          ctrl.enqueue(enc.encode(buf));
          return;
        }
      }
    },
    async cancel() {
      await it.return?.();
    },
  });
}
