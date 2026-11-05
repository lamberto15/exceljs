// Random-access unzip over a Blob / File (web standard; in Node `fs.openAsBlob(path)` gives a
// lazily read file Blob). Only the central directory is read up front; each entry is streamed
// from `blob.slice()` and inflated incrementally, so entry order in the archive does not matter
// and memory stays bounded by one chunk. Limits are enforced on actual output, as everywhere.

import type { Limits } from '../core/types.js';
import { type DirEntry, type LazyOptions, parseDirectory } from './lazy.js';
import { Budget, ZipLimitError } from './limits.js';
import { inflateStream } from './stream.js';

export interface BlobEntry {
  readonly path: string;
  /** Uncompressed size as declared by the central directory (not trusted for limits). */
  readonly size: number;
  readonly compressedSize: number;
  /** 0 = stored, 8 = deflate. */
  readonly compression: number;
  /**
   * Decompressed data, incrementally. `unbounded`: the caller keeps nothing, so only
   * `maxCompressionRatio` applies (not `maxEntryBytes` / `maxTotalBytes`).
   */
  chunks(unbounded?: boolean): AsyncIterable<Uint8Array>;
}

const SLICE = 16 * 1024;
/** Compressed bytes read from the Blob per step. */
const READ = 256 * 1024;

function malformed(what: string): never {
  throw new ZipLimitError(`Malformed zip archive: ${what}`);
}

async function read(blob: Blob, start: number, end: number): Promise<Uint8Array> {
  return new Uint8Array(await blob.slice(start, end).arrayBuffer());
}

/** Read chunks of a ReadableStream without relying on async iteration (Safari lacks it). */
async function* pump(rs: ReadableStream<Uint8Array>): AsyncGenerator<Uint8Array> {
  const r = rs.getReader();
  try {
    for (;;) {
      const x = await r.read();
      if (x.done) return;
      yield x.value;
    }
  } finally {
    r.cancel().catch(() => {});
  }
}

/** Compressed bytes of [start, end) in READ-sized pieces. */
async function* ranges(blob: Blob, start: number, end: number): AsyncGenerator<Uint8Array> {
  for (let p = start; p < end; p += READ) yield read(blob, p, Math.min(p + READ, end));
}

/**
 * Index a zip Blob by its central directory (nothing is inflated). Entry count, names and (unless
 * `deferSizeCheck`) declared sizes are validated; byte / ratio limits apply while inflating.
 */
export async function unzipBlob(
  blob: Blob,
  limits?: Partial<Limits>,
  opts?: LazyOptions,
): Promise<Map<string, BlobEntry>> {
  const budget = new Budget(limits);
  const out = new Map<string, BlobEntry>();
  const n = blob.size;
  if (n === 0) return out;
  const tailLen = Math.min(n, 22 + 0xffff + 20);
  const base = n - tailLen;
  const tail = await read(blob, base, n);
  const dv = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);
  let i = tailLen - 22;
  while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
  if (i < 0) malformed('end of central directory not found (not a zip file?)');
  let count = dv.getUint16(i + 10, true);
  let size = dv.getUint32(i + 12, true);
  let off = dv.getUint32(i + 16, true);
  if (count === 0xffff || size === 0xffffffff || off === 0xffffffff) {
    // ZIP64: locator (20 bytes) right before the EOCD
    const loc = i - 20;
    if (loc >= 0 && dv.getUint32(loc, true) === 0x07064b50) {
      const e64 = Number(dv.getBigUint64(loc + 8, true));
      if (e64 + 56 > n) malformed('bad ZIP64 record');
      const r = await read(blob, e64, e64 + 56);
      const d = new DataView(r.buffer, r.byteOffset, r.byteLength);
      if (d.getUint32(0, true) !== 0x06064b50) malformed('bad ZIP64 record');
      count = Number(d.getBigUint64(32, true));
      size = Number(d.getBigUint64(40, true));
      off = Number(d.getBigUint64(48, true));
    }
  }
  if (off + size > n) malformed('central directory out of bounds');
  const cd =
    off >= base ? tail.subarray(off - base, off - base + size) : await read(blob, off, off + size);
  for (const e of parseDirectory(cd, count, budget, opts?.deferSizeCheck === true)) {
    out.set(e.name, new Entry(blob, budget, e));
  }
  return out;
}

class Entry implements BlobEntry {
  /** Output bytes of this entry already counted against the total (no double counting). */
  private seen = 0;
  readonly path: string;
  readonly size: number;
  readonly compressedSize: number;
  readonly compression: number;

  constructor(
    private readonly blob: Blob,
    private readonly budget: Budget,
    private readonly e: DirEntry,
  ) {
    this.path = e.name;
    this.size = e.usize;
    this.compressedSize = e.csize;
    this.compression = e.method;
  }

  private account(out: number, compressed: number, ratio: boolean, capped: boolean): void {
    if (!capped) {
      this.budget.output(this.path, out, 0, compressed, ratio, false);
      return;
    }
    const fresh = out > this.seen ? out - this.seen : 0;
    if (fresh) this.seen = out;
    this.budget.output(this.path, out, fresh, compressed, ratio);
  }

  async *chunks(unbounded = false): AsyncGenerator<Uint8Array, void, undefined> {
    const { e, blob, path } = this;
    const capped = !unbounded;
    if (e.flags & 1) throw new ZipLimitError(`Encrypted zip entry not supported: ${path}`);
    if (e.method !== 0 && e.method !== 8) {
      throw new ZipLimitError(`Unsupported zip compression method ${e.method}: ${path}`);
    }
    if (capped && e.usize > this.budget.limits.maxEntryBytes) {
      throw new ZipLimitError(`Zip entry too large: ${path}`);
    }
    const h = await read(blob, e.lho, e.lho + 30);
    const hv = new DataView(h.buffer, h.byteOffset, h.byteLength);
    if (h.length < 30 || hv.getUint32(0, true) !== 0x04034b50) {
      malformed(`bad local header for ${path}`);
    }
    const start = e.lho + 30 + hv.getUint16(26, true) + hv.getUint16(28, true);
    const end = start + e.csize;
    if (end > blob.size) throw new ZipLimitError('Truncated zip archive');

    let out = 0;
    if (e.method === 0) {
      for await (const c of ranges(blob, start, end)) {
        out += c.length;
        this.account(out, out, false, capped);
        yield c;
      }
      return;
    }

    let fed = 0;
    if (typeof DecompressionStream === 'function') {
      // off the main thread where the runtime implements it natively (Node: zlib thread pool)
      const counted = blob
        .slice(start, end)
        .stream()
        .pipeThrough(
          new TransformStream<Uint8Array, Uint8Array>({
            transform(c, ctl) {
              fed += c.length;
              ctl.enqueue(c);
            },
          }),
        );
      const inflated = counted.pipeThrough(
        new DecompressionStream(
          'deflate-raw' as CompressionFormat,
        ) as unknown as ReadableWritablePair<Uint8Array, Uint8Array>,
      );
      try {
        for await (const c of pump(inflated)) {
          out += c.length;
          this.account(out, fed, true, capped);
          yield c;
        }
      } catch (err) {
        if (err instanceof ZipLimitError) throw err;
        throw new ZipLimitError(`Malformed zip archive: bad deflate data in ${path}`);
      }
      return;
    }

    const pending: Uint8Array[] = [];
    let done = false;
    const inf = inflateStream((d, final) => {
      if (final) done = true;
      if (d.length === 0) return;
      out += d.length;
      this.account(out, fed, true, capped);
      pending.push(d);
    });
    if (e.csize === 0) inf.push(new Uint8Array(0), true);
    for await (const c of ranges(blob, start, end)) {
      for (let j = 0; j < c.length; j += SLICE) {
        const s = c.subarray(j, Math.min(j + SLICE, c.length));
        fed += s.length;
        inf.push(s, start + fed === end);
        for (let k = 0; k < pending.length; k++) yield pending[k]!;
        pending.length = 0;
      }
    }
    if (!done) throw new ZipLimitError('Truncated zip archive');
  }
}
