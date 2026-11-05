// RFC 4180 push parser. Chunks may split anywhere. Only the unconsumed tail
// (from the start of the current field) is carried across chunks.

import type { CellError, ErrorCode } from '../core/types.js';

/** Error raised for invalid options or limit violations. */
export class CsvError extends Error {
  readonly code: 'FIELD_TOO_LARGE' | 'BAD_OPTION';
  constructor(code: 'FIELD_TOO_LARGE' | 'BAD_OPTION', message: string) {
    super(message);
    this.name = 'CsvError';
    this.code = code;
  }
}

export interface CsvReadOptions {
  /** Single-character field delimiter (default ","). Multi-character delimiters are not supported. */
  delimiter?: string;
  /** Single-character quote (default '"'). */
  quote?: string;
  /** Strip a leading BOM (U+FEFF) from the text (default true). */
  bom?: boolean;
  /** Convert unquoted numeric-looking fields to numbers (rows become CsvValue[]). */
  parseNumbers?: boolean;
  /** Convert unquoted ISO 8601 fields to Date (rows become CsvValue[]). */
  parseDates?: boolean;
  /** Convert unquoted `true` / `false` (any case) to booleans (rows become CsvValue[]). */
  parseBooleans?: boolean;
  /** Convert unquoted Excel error literals (`#N/A`, `#DIV/0!`, ...) to `CellError` (rows become CsvValue[]). */
  parseErrors?: boolean;
  /**
   * Extra date formats tried in order after ISO 8601 (implies `parseDates`). Tokens: `YYYY`, `MM`,
   * `DD`, `HH`, `mm`, `ss` (two digits; `M`, `D`, `H`, `m`, `s` accept one or two), `[text]` is
   * literal, anything else must match exactly. Each format needs `YYYY`, a month and a day.
   */
  dateFormats?: readonly string[];
  /** Dates without a UTC offset use the local wall clock instead of UTC (like `WorkbookOptions.localDates`). */
  localDates?: boolean;
  /** Max length of a single field in UTF-16 code units (~bytes for ASCII). Default 16 MiB. */
  maxFieldBytes?: number;
  /** Skip lines that are completely empty (default true). */
  skipEmptyLines?: boolean;
}

/** Options that keep rows as `string[]`. */
export type CsvStringOptions = CsvReadOptions & {
  parseNumbers?: false;
  parseDates?: false;
  parseBooleans?: false;
  parseErrors?: false;
  dateFormats?: undefined;
};

export type CsvValue = string | number | boolean | Date | CellError;

const NUM_RE = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const DATE_RE =
  /^(\d{4})-(\d\d)-(\d\d)(?:[T ](\d\d):(\d\d)(?::(\d\d)(\.\d{1,9})?)?(Z|[+-]\d\d:?\d\d)?)?$/;
const ERR_RE = /^#(?:N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)$/;

/** A compiled date format: regex + capture index of [Y, M, D, h, m, s] (0 = absent). */
type Fmt = [RegExp, number[]];

function compileFormat(f: string): Fmt {
  const at = [0, 0, 0, 0, 0, 0];
  let n = 0;
  let ok = true;
  const src = f.replace(/\[([^\]]*)\]|YYYY|MM?|DD?|HH?|mm?|ss?|./gs, (t, lit?: string) => {
    const k = lit === undefined ? 'YMDHms'.indexOf(t[0]!) : -1;
    if (k < 0 || (!k && t.length < 4)) return (lit ?? t).replace(/\W/g, '\\$&');
    if (at[k]) ok = false;
    at[k] = ++n;
    return k ? (t.length > 1 ? '(\\d\\d)' : '(\\d\\d?)') : '(\\d{4})';
  });
  if (!ok || !at[0] || !at[1] || !at[2]) {
    throw new CsvError('BAD_OPTION', `bad date format: ${f}`);
  }
  return [new RegExp(`^${src}$`), at];
}

/** Calendar-checked Date from wall-clock parts (UTC, or local with `local`); undefined if invalid. */
function mkDate(p: number[], ms: number, local: boolean): Date | undefined {
  const [y, mo, d, h, mi, s] = p as [number, number, number, number, number, number];
  const u = new Date(Date.UTC(y, mo - 1, d, h, mi, s, ms));
  if (u.getUTCFullYear() !== y || u.getUTCMonth() !== mo - 1 || u.getUTCDate() !== d) return;
  if (h > 23 || mi > 59 || s > 59) return;
  return local ? new Date(y, mo - 1, d, h, mi, s, ms) : u;
}

function convert(s: string, o: Conv): CsvValue {
  const n = s.length;
  if (n === 0 || n > 40) return s;
  const c = s.charCodeAt(0);
  if (o.n && (c === 45 || (c >= 48 && c <= 57)) && NUM_RE.test(s)) {
    // keep long digit strings (ids) lossless
    if (n <= 15 || s.indexOf('.') >= 0 || s.indexOf('e') >= 0 || s.indexOf('E') >= 0) {
      const v = Number(s);
      if (Number.isFinite(v)) return v;
    }
    return s;
  }
  if (o.d) {
    const m = DATE_RE.exec(s);
    if (m) {
      const d = m[8]
        ? new Date(s.replace(' ', 'T'))
        : mkDate(
            m.slice(1, 7).map((x) => +(x ?? 0)),
            +(m[7] ?? '').slice(1, 4).padEnd(3, '0'),
            o.l,
          );
      if (d && !Number.isNaN(d.getTime())) return d;
    }
    for (const [re, at] of o.f) {
      const r = re.exec(s);
      if (r) {
        const d = mkDate(
          at.map((k) => (k ? +r[k]! : 0)),
          0,
          o.l,
        );
        if (d) return d;
      }
    }
  }
  if (o.b && /^(?:true|false)$/i.test(s)) return n === 4;
  if (o.e && ERR_RE.test(s)) return { error: s as ErrorCode };
  return s;
}

interface Conv {
  n: boolean;
  d: boolean;
  b: boolean;
  e: boolean;
  l: boolean;
  f: Fmt[];
}

export class CsvParser<R extends CsvValue[] = string[]> {
  #onRow: (row: R) => void;
  #delim: number;
  #quote: number;
  #stripBom: boolean;
  #conv: Conv | undefined;
  #max: number;
  #skipEmpty: boolean;
  #row: CsvValue[] = [];
  #buf = '';
  #qoff = 0; // chars of #buf (after the opening quote) already known quote-free
  #mode = 0; // 0 fresh field, 1 inside quotes (continuing), 2 after closing quote
  #qparts = '';
  #started = false;
  #ended = false;

  constructor(onRow: (row: R) => void, o: CsvReadOptions = {}) {
    const d = o.delimiter ?? ',';
    const q = o.quote ?? '"';
    if (d.length !== 1) throw new CsvError('BAD_OPTION', 'delimiter must be a single character');
    if (q.length !== 1) throw new CsvError('BAD_OPTION', 'quote must be a single character');
    if (d === q || d === '\n' || d === '\r' || q === '\n' || q === '\r') {
      throw new CsvError('BAD_OPTION', 'invalid delimiter/quote combination');
    }
    this.#onRow = onRow;
    this.#delim = d.charCodeAt(0);
    this.#quote = q.charCodeAt(0);
    this.#stripBom = o.bom ?? true;
    const f = (o.dateFormats ?? []).map(compileFormat);
    const c: Conv = {
      n: o.parseNumbers === true,
      d: o.parseDates === true || f.length > 0,
      b: o.parseBooleans === true,
      e: o.parseErrors === true,
      l: o.localDates === true,
      f,
    };
    this.#conv = c.n || c.d || c.b || c.e ? c : undefined;
    this.#max = o.maxFieldBytes ?? 16 * 1024 * 1024;
    this.#skipEmpty = o.skipEmptyLines ?? true;
  }

  write(chunk: string): void {
    if (this.#ended) throw new Error('CsvParser: write after end');
    if (chunk.length === 0) return;
    if (!this.#started) {
      this.#started = true;
      if (this.#stripBom && chunk.charCodeAt(0) === 0xfeff) {
        chunk = chunk.slice(1);
        if (chunk.length === 0) return;
      }
    }
    this.#parse(this.#buf.length === 0 ? chunk : this.#buf + chunk, false);
  }

  end(): void {
    if (this.#ended) return;
    this.#ended = true;
    this.#parse(this.#buf, true);
    if (this.#row.length > 0) {
      this.#row.push('');
      this.#emit();
    }
    this.#buf = '';
  }

  #tooLarge(): never {
    throw new CsvError('FIELD_TOO_LARGE', `field exceeds maxFieldBytes (${this.#max})`);
  }

  #parse(s: string, final: boolean): void {
    const n = s.length;
    const delim = this.#delim;
    const quote = this.#quote;
    const conv = this.#conv;
    let i = 0;
    let qoff = this.#qoff;
    this.#qoff = 0;
    const quoteStr = String.fromCharCode(quote);

    while (i < n || this.#mode !== 0) {
      const start = i;
      let val: string;
      let wasQuoted = false;
      let c = i < n ? s.charCodeAt(i) : -1;
      let term: number; // 0 = EOF(final), else terminator char
      let after: number; // index after terminator
      let mode = this.#mode;

      if (mode === 1 || (mode === 0 && c === quote)) {
        wasQuoted = true;
        let parts: string;
        let segStart: number;
        let j: number;
        if (mode === 1) {
          parts = this.#qparts;
          segStart = i;
          j = i + qoff;
        } else {
          parts = '';
          segStart = i + 1;
          j = i + 1;
        }
        qoff = 0;
        let closed = false;
        for (;;) {
          const k = s.indexOf(quoteStr, j);
          if (k < 0) break;
          if (k + 1 >= n) {
            if (final) {
              parts += s.slice(segStart, k);
              j = n;
              closed = true;
            }
            break;
          }
          if (s.charCodeAt(k + 1) === quote) {
            parts += s.slice(segStart, k + 1);
            segStart = j = k + 2;
            continue;
          }
          parts += s.slice(segStart, k);
          j = k + 1;
          closed = true;
          break;
        }
        if (!closed) {
          if (final) {
            parts += s.slice(segStart);
            j = n;
          } else {
            this.#mode = 1;
            this.#qparts = parts;
            this.#buf = s.slice(segStart);
            this.#qoff = Math.max(0, n - segStart - 1); // last char may begin an escaped quote
            if (parts.length + this.#buf.length > this.#max + 1) this.#tooLarge();
            return;
          }
        }
        val = parts;
        i = j;
        mode = 2;
      } else if (mode === 2) {
        wasQuoted = true;
        val = this.#qparts;
      } else {
        let e = i;
        while (e < n) {
          c = s.charCodeAt(e);
          if (c === delim || c === 10 || c === 13) break;
          e++;
        }
        if (e >= n && !final) {
          this.#buf = s.slice(start);
          if (this.#buf.length > this.#max) this.#tooLarge();
          return;
        }
        val = e === start ? '' : s.slice(start, e);
        i = e;
      }
      if (mode === 2) {
        // lenient: stray text between the closing quote and the terminator is appended
        let e = i;
        while (e < n) {
          c = s.charCodeAt(e);
          if (c === delim || c === 10 || c === 13) break;
          e++;
        }
        if (e >= n && !final) {
          this.#mode = 2;
          this.#qparts = e > i ? val + s.slice(i, e) : val;
          this.#buf = '';
          return;
        }
        if (e > i) val += s.slice(i, e);
        i = e;
      }
      this.#mode = 0;
      if (val.length > this.#max) this.#tooLarge();

      // terminator
      if (i >= n) {
        term = 0;
        after = n;
      } else {
        term = s.charCodeAt(i);
        after = i + 1;
        if (term === 13) {
          if (after >= n && !final) {
            this.#buf = s.slice(start);
            if (wasQuoted) {
              this.#mode = 2;
              this.#qparts = val;
              this.#buf = s.slice(i);
            }
            return;
          }
          if (after < n && s.charCodeAt(after) === 10) after++;
        }
      }

      if (term === delim) {
        this.#row.push(conv && !wasQuoted ? convert(val, conv) : val);
        i = after;
        continue;
      }
      // end of record (LF, CR, CRLF or EOF)
      if (this.#row.length === 0 && val.length === 0 && !wasQuoted && this.#skipEmpty) {
        // empty line
      } else {
        this.#row.push(conv && !wasQuoted ? convert(val, conv) : val);
        this.#emit();
      }
      i = after;
    }
    this.#buf = '';
  }

  #emit(): void {
    const r = this.#row;
    this.#row = [];
    this.#onRow(r as R);
  }
}

/** Parse a whole string. Rows are string[] unless a typing option is set. */
export function parseCsv(text: string, o?: CsvStringOptions): string[][];
export function parseCsv(text: string, o: CsvReadOptions): CsvValue[][];
export function parseCsv(text: string, o: CsvReadOptions = {}): CsvValue[][] {
  const out: CsvValue[][] = [];
  const p = new CsvParser<CsvValue[]>((r) => out.push(r), o);
  p.write(text);
  p.end();
  return out;
}

/** Options for csvRows: read options plus the byte decoding. */
export interface CsvRowsOptions extends CsvReadOptions {
  /**
   * Encoding of byte chunks (a `TextDecoder` label, e.g. 'windows-1252', 'utf-16le', 'shift_jis').
   * Default 'utf-8'. A BOM (EF BB BF, FF FE, FE FF) always wins. 'auto': without a BOM, UTF-8,
   * or windows-1252 if the first chunk holding non-ASCII bytes is not valid UTF-8.
   */
  encoding?: string;
}

type ByteOrStringSource =
  | AsyncIterable<string | Uint8Array>
  | Iterable<string | Uint8Array>
  | ReadableStream<Uint8Array>;

async function* chunksOf(source: ByteOrStringSource): AsyncGenerator<string | Uint8Array> {
  if (typeof (source as ReadableStream<Uint8Array>).getReader === 'function') {
    const reader = (source as ReadableStream<Uint8Array>).getReader();
    try {
      for (;;) {
        const r = await reader.read();
        if (r.done) return;
        yield r.value;
      }
    } finally {
      reader.releaseLock();
    }
  } else {
    yield* source as AsyncIterable<string | Uint8Array>;
  }
}

const decoder = (label: string, fatal = false) =>
  new TextDecoder(label, { ignoreBOM: true, fatal });

/**
 * Chunk-safe byte decoder: call with each chunk, then with no argument at the end. The BOM
 * is decoded as U+FEFF (the parser strips it). Undecided bytes (a possible BOM prefix, or for
 * 'auto' the bytes from the first non-ASCII one) are held until the encoding is known.
 */
function byteDecoder(enc: string): (chunk?: Uint8Array) => string {
  let dec: TextDecoder | undefined;
  let head: Uint8Array = new Uint8Array(0);
  let sniffed = false;
  return (chunk) => {
    const end = !chunk;
    const opts = { stream: !end };
    if (dec) return dec.decode(chunk, opts);
    let b = chunk ?? head;
    if (chunk && head.length) {
      b = new Uint8Array(head.length + chunk.length);
      b.set(head);
      b.set(chunk, head.length);
    }
    if (!sniffed) {
      const b0 = b[0]!;
      if (!end && b.length < 3 && (b.length === 0 || b0 > 0xfd || b0 === 0xef)) {
        head = b;
        return '';
      }
      sniffed = true;
      const label =
        b0 === 0xef && b[1] === 0xbb && b[2] === 0xbf
          ? 'utf-8'
          : b0 === 0xff && b[1] === 0xfe
            ? 'utf-16le'
            : b0 === 0xfe && b[1] === 0xff
              ? 'utf-16be'
              : enc;
      if (label !== 'auto') {
        dec = decoder(label);
        return dec.decode(b, opts);
      }
    }
    // 'auto', undecided: ASCII reads the same in UTF-8 and windows-1252, so it passes through;
    // the first chunk with non-ASCII bytes decides (an incomplete sequence waits for more bytes)
    let i = 0;
    while (i < b.length && b[i]! < 0x80) i++;
    let out = i ? decoder('utf-8').decode(b.subarray(0, i)) : '';
    const rest = b.subarray(i);
    head = rest.slice();
    if (rest.length) {
      try {
        if (decoder('utf-8', true).decode(rest, opts)) dec = decoder('utf-8');
      } catch {
        dec = decoder('windows-1252');
      }
      if (dec) out += dec.decode(rest, opts);
    }
    return out;
  };
}

/**
 * Stream rows from bytes/strings. Byte chunks are decoded chunk-safely (see
 * `CsvRowsOptions.encoding`); string chunks are used as is.
 */
export function csvRows(
  source: ByteOrStringSource,
  o?: CsvStringOptions & CsvRowsOptions,
): AsyncIterable<string[]>;
export function csvRows(source: ByteOrStringSource, o: CsvRowsOptions): AsyncIterable<CsvValue[]>;
export function csvRows(
  source: ByteOrStringSource,
  o: CsvRowsOptions = {},
): AsyncIterable<CsvValue[]> {
  const enc = o.encoding ?? 'utf-8';
  if (enc !== 'auto') decoder(enc); // throws RangeError for an unknown label
  return (async function* () {
    const dec = byteDecoder(enc);
    let pending: CsvValue[][] = [];
    const p = new CsvParser<CsvValue[]>((r) => pending.push(r), o);
    for await (const chunk of chunksOf(source)) {
      p.write(typeof chunk === 'string' ? chunk : dec(chunk));
      if (pending.length) {
        const batch = pending;
        pending = [];
        yield* batch;
      }
    }
    p.write(dec());
    p.end();
    yield* pending;
  })();
}
