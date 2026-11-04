// RFC 4180 push parser. Chunks may split anywhere. Only the unconsumed tail
// (from the start of the current field) is carried across chunks.

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
  /** Strip a leading UTF-8 BOM (default true). */
  bom?: boolean;
  /** Convert unquoted numeric-looking fields to numbers (rows become CsvValue[]). */
  parseNumbers?: boolean;
  /** Convert unquoted ISO 8601 fields to Date (rows become CsvValue[]). */
  parseDates?: boolean;
  /** Max length of a single field in UTF-16 code units (~bytes for ASCII). Default 16 MiB. */
  maxFieldBytes?: number;
  /** Skip lines that are completely empty (default true). */
  skipEmptyLines?: boolean;
}

export type CsvValue = string | number | Date;

const NUM_RE = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const DATE_RE =
  /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

function convert(s: string, nums: boolean, dates: boolean): CsvValue {
  const n = s.length;
  if (n === 0 || n > 40) return s;
  const c = s.charCodeAt(0);
  if (nums && (c === 45 || (c >= 48 && c <= 57)) && NUM_RE.test(s)) {
    // keep long digit strings (ids) lossless
    if (n <= 15 || s.indexOf('.') >= 0 || s.indexOf('e') >= 0 || s.indexOf('E') >= 0) {
      const v = Number(s);
      if (Number.isFinite(v)) return v;
    }
    return s;
  }
  if (dates && c >= 48 && c <= 57) {
    const m = DATE_RE.exec(s);
    if (m) {
      let iso = s.replace(' ', 'T');
      if (iso.length > 10 && m[1] === undefined) iso += 'Z';
      const d = new Date(iso);
      if (!Number.isNaN(d.getTime())) return d;
    }
  }
  return s;
}

export class CsvParser<R extends CsvValue[] = string[]> {
  #onRow: (row: R) => void;
  #delim: number;
  #quote: number;
  #stripBom: boolean;
  #typed: boolean;
  #nums: boolean;
  #dates: boolean;
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
    this.#nums = o.parseNumbers === true;
    this.#dates = o.parseDates === true;
    this.#typed = this.#nums || this.#dates;
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
    const typed = this.#typed;
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
        this.#row.push(typed && !wasQuoted ? convert(val, this.#nums, this.#dates) : val);
        i = after;
        continue;
      }
      // end of record (LF, CR, CRLF or EOF)
      if (this.#row.length === 0 && val.length === 0 && !wasQuoted && this.#skipEmpty) {
        // empty line
      } else {
        this.#row.push(typed && !wasQuoted ? convert(val, this.#nums, this.#dates) : val);
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

/** Parse a whole string. Rows are string[] unless parseNumbers/parseDates is set. */
export function parseCsv(
  text: string,
  o?: CsvReadOptions & { parseNumbers?: false; parseDates?: false },
): string[][];
export function parseCsv(text: string, o: CsvReadOptions): CsvValue[][];
export function parseCsv(text: string, o: CsvReadOptions = {}): CsvValue[][] {
  const out: CsvValue[][] = [];
  const p = new CsvParser<CsvValue[]>((r) => out.push(r), o);
  p.write(text);
  p.end();
  return out;
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

/** Stream rows from bytes/strings. Multi-byte UTF-8 split across chunks is handled. */
export function csvRows(
  source: ByteOrStringSource,
  o?: CsvReadOptions & { parseNumbers?: false; parseDates?: false },
): AsyncIterable<string[]>;
export function csvRows(source: ByteOrStringSource, o: CsvReadOptions): AsyncIterable<CsvValue[]>;
export function csvRows(
  source: ByteOrStringSource,
  o: CsvReadOptions = {},
): AsyncIterable<CsvValue[]> {
  return (async function* () {
    const dec = new TextDecoder('utf-8', { ignoreBOM: true });
    let pending: CsvValue[][] = [];
    const p = new CsvParser<CsvValue[]>((r) => pending.push(r), o);
    for await (const chunk of chunksOf(source)) {
      p.write(typeof chunk === 'string' ? chunk : dec.decode(chunk, { stream: true }));
      if (pending.length) {
        const batch = pending;
        pending = [];
        yield* batch;
      }
    }
    const tail = dec.decode();
    if (tail) p.write(tail);
    p.end();
    yield* pending;
  })();
}
