import type { UnzipDecoder, ZipInputFile, Zippable } from 'fflate';
import {
  zipSync as fflateZipSync,
  strToU8,
  Unzip,
  UnzipInflate,
  UnzipPassThrough,
  Zip,
  ZipDeflate,
  ZipPassThrough,
} from 'fflate';
import type { Limits } from '../core/types.js';
import { Budget, ZipLimitError, ZipPathError } from './limits.js';

export { assertSafePath, ZipLimitError, ZipPathError } from './limits.js';

export interface ZipEntryIn {
  path: string;
  data: Uint8Array | string;
  level?: number;
}

export interface ZipEntryOut {
  path: string;
  chunks: AsyncIterable<Uint8Array>;
}

/** Max compressed bytes fed to the inflater at once; bounds the memory burst of a bomb. */
const SLICE = 16 * 1024;

type Level = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

// ---------------------------------------------------------------- writing

export function zipSync(entries: ZipEntryIn[], level: number): Uint8Array {
  const z: Zippable = {};
  for (const e of entries) {
    const data = typeof e.data === 'string' ? strToU8(e.data) : e.data;
    z[e.path] = [data, { level: (e.level ?? level) as Level }];
  }
  return fflateZipSync(z);
}

/** Entry compressor producing output asynchronously (see `nativeCompressor()`). */
export interface AsyncCompressor {
  /** A zip entry for fflate's Zip (deflate, data descriptor). */
  file(path: string): ZipInputFile & { push(chunk: Uint8Array, final?: boolean): void };
  /** Resolves once at most `maxBytes` of input await compression; undefined = no wait. */
  below(maxBytes: number): Promise<void> | undefined;
  /** Resolves when every entry so far has been compressed (rejects on the first error). */
  idle(): Promise<void>;
  readonly error: unknown;
}

export interface ZipStreamOptions {
  /**
   * Compress level > 0 entries with this async compressor (e.g. `nativeCompressor()`); output
   * then arrives asynchronously — await `idle()` before relying on it. Level-0 entries are
   * still stored synchronously.
   */
  compressor?: AsyncCompressor | undefined;
}

export class ZipStreamWriter {
  readonly #zip: Zip;
  readonly #level: number;
  readonly #native: AsyncCompressor | undefined;

  constructor(
    onData: (chunk: Uint8Array, final: boolean) => void,
    level: number,
    opts?: ZipStreamOptions,
  ) {
    this.#level = level;
    this.#native = opts?.compressor;
    this.#zip = new Zip((err, chunk, final) => {
      if (err) throw err;
      onData(chunk, final);
    });
  }

  /** True when entries are compressed asynchronously by the native deflater. */
  get async(): boolean {
    return this.#native !== undefined;
  }

  /** Add a file. Push data incrementally; pass `final = true` on the last push to close it. */
  file(path: string, level?: number): { push(data: Uint8Array | string, final?: boolean): void } {
    const lvl = level ?? this.#level;
    const n = this.#native;
    if (n?.error !== undefined) throw n.error;
    const f: ZipInputFile & { push(chunk: Uint8Array, final?: boolean): void } =
      lvl === 0
        ? new ZipPassThrough(path)
        : n
          ? n.file(path)
          : new ZipDeflate(path, { level: lvl as Level });
    this.#zip.add(f);
    return {
      push(data, final = false) {
        f.push(typeof data === 'string' ? strToU8(data) : data, final);
      },
    };
  }

  /**
   * Backpressure for async (native) mode: resolves once at most `maxBytes` of pushed input is
   * still waiting for the compressor; undefined when no wait is needed (always, in sync mode).
   */
  ready(maxBytes: number): Promise<void> | undefined {
    return this.#native?.below(maxBytes);
  }

  /** Resolves when all pushed data has been compressed and emitted (immediately in sync mode). */
  idle(): Promise<void> {
    return this.#native ? this.#native.idle() : Promise.resolve();
  }

  end(): void {
    this.#zip.end();
  }
}

// ---------------------------------------------------------------- reading

interface Sink {
  start(path: string): void;
  chunk(data: Uint8Array): void;
  end(): void;
}

/** Streaming unzip driver that enforces limits on actual decompressed output. */
class Driver {
  readonly #unzip: Unzip;
  readonly #sink: Sink;
  #open = false;
  /** Compressed bytes seen for the current entry (counted by the inflate decoder). */
  #compressed = 0;

  constructor(sink: Sink, limits?: Partial<Limits>) {
    this.#sink = sink;
    const budget = new Budget(limits);
    const self = this;
    class CountingInflate implements UnzipDecoder {
      static compression = 8;
      ondata: UnzipDecoder['ondata'] = () => {};
      readonly #inner = new UnzipInflate();
      constructor() {
        this.#inner.ondata = (e, d, f) => this.ondata(e, d, f);
      }
      push(chunk: Uint8Array, final: boolean): void {
        self.#compressed += chunk.length;
        this.#inner.push(chunk, final);
      }
    }
    this.#unzip = new Unzip();
    this.#unzip.register(CountingInflate);
    this.#unzip.register(UnzipPassThrough);
    this.#unzip.onfile = (file) => {
      if (!budget.enter(file.name, file.originalSize)) return; // directory
      this.#open = true;
      this.#compressed = 0;
      let out = 0;
      const name = file.name;
      // Stored entries can't be bombs (output == input) and bypass CountingInflate.
      const checkRatio = file.compression !== 0;
      this.#sink.start(name);
      file.ondata = (err, data, final) => {
        if (err) throw err;
        out += data.length;
        budget.output(name, out, data.length, this.#compressed, checkRatio);
        if (data.length > 0) this.#sink.chunk(data);
        if (final) {
          this.#open = false;
          this.#sink.end();
        }
      };
      file.start();
    };
  }

  /** Push one slice (<= SLICE bytes recommended). */
  push(slice: Uint8Array, final: boolean): void {
    this.#unzip.push(slice, final);
    if (final && this.#open) throw new ZipLimitError('Truncated zip archive');
  }
}

function concat(parts: Uint8Array[], size: number): Uint8Array {
  if (parts.length === 1) return parts[0]!;
  const out = new Uint8Array(size);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

export function unzipSync(data: Uint8Array, limits?: Partial<Limits>): Map<string, Uint8Array> {
  const result = new Map<string, Uint8Array>();
  let name = '';
  let parts: Uint8Array[] = [];
  let size = 0;
  const driver = new Driver(
    {
      start(path) {
        if (result.has(path)) throw new ZipPathError(`Duplicate zip entry: ${path}`);
        name = path;
        parts = [];
        size = 0;
      },
      chunk(d) {
        // fflate may reuse buffers across ondata calls; copy.
        parts.push(d.slice());
        size += d.length;
      },
      end() {
        result.set(name, parts.length === 0 ? new Uint8Array(0) : concat(parts, size));
      },
    },
    limits,
  );
  if (data.length === 0) return result;
  for (let i = 0; i < data.length; i += SLICE) {
    driver.push(data.subarray(i, Math.min(i + SLICE, data.length)), false);
  }
  driver.push(new Uint8Array(0), true);
  return result;
}

type Event = { t: 's'; path: string } | { t: 'c'; data: Uint8Array } | { t: 'e' };

/**
 * Stream entries out of a zip. The `chunks` of an entry must be consumed (or abandoned by
 * moving on: remaining data is skipped) before requesting the next entry; chunks of a previous
 * entry are not valid after the iterator advances. Limits are enforced on actual output.
 */
export async function* unzipStream(
  source: AsyncIterable<Uint8Array>,
  limits?: Partial<Limits>,
): AsyncGenerator<ZipEntryOut, void, undefined> {
  const queue: Event[] = [];
  let head = 0;
  const driver = new Driver(
    {
      start: (path) => void queue.push({ t: 's', path }),
      chunk: (data) => void queue.push({ t: 'c', data: data.slice() }),
      end: () => void queue.push({ t: 'e' }),
    },
    limits,
  );
  const it = source[Symbol.asyncIterator]();
  let cur: Uint8Array | null = null;
  let off = 0;
  let finished = false;

  /** Feed the next slice to the unzipper. Returns false once everything has been fed. */
  const pull = async (): Promise<boolean> => {
    if (finished) return false;
    while (cur === null || off >= cur.length) {
      const r = await it.next();
      if (r.done) {
        finished = true;
        driver.push(new Uint8Array(0), true);
        return true;
      }
      cur = r.value;
      off = 0;
    }
    const end = Math.min(off + SLICE, cur.length);
    const slice = cur.subarray(off, end);
    off = end;
    driver.push(slice, false);
    return true;
  };

  const next = async (): Promise<Event | undefined> => {
    while (head >= queue.length) {
      head = 0;
      queue.length = 0;
      if (!(await pull())) return undefined;
    }
    return queue[head++];
  };

  try {
    for (;;) {
      const ev = await next();
      if (!ev) return;
      if (ev.t !== 's') continue; // leftover data of an abandoned entry
      let done = false;
      const chunks: AsyncIterable<Uint8Array> = {
        [Symbol.asyncIterator]: async function* () {
          while (!done) {
            const e = await next();
            if (!e) throw new ZipLimitError('Truncated zip archive');
            if (e.t === 'c') yield e.data;
            else if (e.t === 'e') done = true;
          }
        },
      };
      yield { path: ev.path, chunks };
    }
  } finally {
    await it.return?.();
  }
}
