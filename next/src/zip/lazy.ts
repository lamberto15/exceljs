// Lazy synchronous unzip: parses the central directory only; entries are inflated on demand,
// either whole (`bytes()`) or incrementally (`chunks()`), so a reader can keep a single part
// decompressed at a time and never materialize a huge sheet XML in full.
// Limits (see limits.ts) are enforced on the ACTUAL inflated output of every access.

import type { Limits } from '../core/types.js';
import { Budget, ZipLimitError, ZipPathError } from './limits.js';
import { inflateStream } from './stream.js';

export interface LazyEntry {
  readonly path: string;
  /** Uncompressed size as declared by the central directory (not trusted for limits). */
  readonly size: number;
  readonly compressedSize: number;
  /** 0 = stored, 8 = deflate. */
  readonly compression: number;
  /**
   * The whole decompressed entry. Inflates on every call (cache it if needed). Stored entries
   * return a view into the archive buffer (no copy).
   */
  bytes(): Uint8Array;
  /**
   * Decompressed data, incrementally. `chunkSize` = compressed bytes inflated per step (default
   * 64 KiB); yielded chunks are typically 5-20x larger for XML. Chunks are fresh arrays
   * (stored entries: views into the archive buffer).
   */
  chunks(chunkSize?: number): Iterable<Uint8Array>;
}

/** Max compressed bytes fed to the inflater at once; bounds the memory burst of a bomb. */
const SLICE = 16 * 1024;
const EMPTY = /* @__PURE__ */ new Uint8Array(0);

function malformed(what: string): never {
  throw new ZipLimitError(`Malformed zip archive: ${what}`);
}

function latin1(b: Uint8Array): string {
  let s = '';
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]!);
  return s;
}

const utf8 = /* @__PURE__ */ new TextDecoder();

/** Locate the end-of-central-directory record; returns [entries, cdOffset, cdSize]. */
function findDirectory(data: Uint8Array, dv: DataView): [number, number, number] {
  const n = data.length;
  const stop = Math.max(0, n - 22 - 0xffff);
  for (let i = n - 22; i >= stop; i--) {
    if (dv.getUint32(i, true) !== 0x06054b50) continue;
    let count = dv.getUint16(i + 10, true);
    let size = dv.getUint32(i + 12, true);
    let off = dv.getUint32(i + 16, true);
    if (count === 0xffff || size === 0xffffffff || off === 0xffffffff) {
      // ZIP64: locator (20 bytes) right before the EOCD
      const loc = i - 20;
      if (loc >= 0 && dv.getUint32(loc, true) === 0x07064b50) {
        const e64 = Number(dv.getBigUint64(loc + 8, true));
        if (e64 + 56 > n || dv.getUint32(e64, true) !== 0x06064b50) malformed('bad ZIP64 record');
        count = Number(dv.getBigUint64(e64 + 32, true));
        size = Number(dv.getBigUint64(e64 + 40, true));
        off = Number(dv.getBigUint64(e64 + 48, true));
      }
    }
    if (off + size > n) malformed('central directory out of bounds');
    return [count, off, size];
  }
  if (n === 0) return [0, 0, 0];
  return malformed('end of central directory not found (not a zip file?)');
}

/**
 * Index a zip archive without decompressing anything. Entry count, names and declared sizes are
 * validated up front; byte/ratio limits are enforced while inflating. Directory entries are
 * skipped. Encrypted entries and methods other than stored/deflate throw on access.
 */
export function unzipLazy(data: Uint8Array, limits?: Partial<Limits>): Map<string, LazyEntry> {
  const budget = new Budget(limits);
  const out = new Map<string, LazyEntry>();
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const [count, cdOff] = findDirectory(data, dv);
  if (count > budget.limits.maxEntries) {
    throw new ZipLimitError(`Too many zip entries (max ${budget.limits.maxEntries})`);
  }
  let p = cdOff;
  for (let k = 0; k < count; k++) {
    if (p + 46 > data.length || dv.getUint32(p, true) !== 0x02014b50) {
      malformed('bad central directory entry');
    }
    const flags = dv.getUint16(p + 8, true);
    const method = dv.getUint16(p + 10, true);
    let csize = dv.getUint32(p + 20, true);
    let usize = dv.getUint32(p + 24, true);
    const nl = dv.getUint16(p + 28, true);
    const xl = dv.getUint16(p + 30, true);
    const cl = dv.getUint16(p + 32, true);
    let lho = dv.getUint32(p + 42, true);
    const nameBytes = data.subarray(p + 46, p + 46 + nl);
    const name = flags & 0x800 ? utf8.decode(nameBytes) : latin1(nameBytes);
    if (usize === 0xffffffff || csize === 0xffffffff || lho === 0xffffffff) {
      // ZIP64 extended information extra field
      let x = p + 46 + nl;
      const xe = x + xl;
      while (x + 4 <= xe) {
        const id = dv.getUint16(x, true);
        const len = dv.getUint16(x + 2, true);
        if (id === 1) {
          let q = x + 4;
          if (usize === 0xffffffff) (usize = Number(dv.getBigUint64(q, true))), (q += 8);
          if (csize === 0xffffffff) (csize = Number(dv.getBigUint64(q, true))), (q += 8);
          if (lho === 0xffffffff) lho = Number(dv.getBigUint64(q, true));
          break;
        }
        x += 4 + len;
      }
    }
    p += 46 + nl + xl + cl;
    if (!budget.enter(name, usize)) continue;
    if (out.has(name)) throw new ZipPathError(`Duplicate zip entry: ${name}`);
    out.set(name, new Entry(data, dv, budget, name, flags, method, csize, usize, lho));
  }
  return out;
}

class Entry implements LazyEntry {
  /** Output bytes of this entry already counted against the total (no double counting). */
  private seen = 0;
  private start = -1;

  constructor(
    private readonly data: Uint8Array,
    private readonly dv: DataView,
    private readonly budget: Budget,
    readonly path: string,
    private readonly flags: number,
    readonly compression: number,
    readonly compressedSize: number,
    readonly size: number,
    private readonly lho: number,
  ) {}

  private raw(): Uint8Array {
    if (this.flags & 1) throw new ZipLimitError(`Encrypted zip entry not supported: ${this.path}`);
    if (this.compression !== 0 && this.compression !== 8) {
      throw new ZipLimitError(
        `Unsupported zip compression method ${this.compression}: ${this.path}`,
      );
    }
    if (this.start < 0) {
      const h = this.lho;
      if (h + 30 > this.data.length || this.dv.getUint32(h, true) !== 0x04034b50) {
        malformed(`bad local header for ${this.path}`);
      }
      this.start = h + 30 + this.dv.getUint16(h + 26, true) + this.dv.getUint16(h + 28, true);
    }
    const end = this.start + this.compressedSize;
    if (end > this.data.length) throw new ZipLimitError('Truncated zip archive');
    return this.data.subarray(this.start, end);
  }

  private account(out: number, compressed: number, ratio: boolean): void {
    const fresh = out > this.seen ? out - this.seen : 0;
    if (fresh) this.seen = out;
    this.budget.output(this.path, out, fresh, compressed, ratio);
  }

  bytes(): Uint8Array {
    const raw = this.raw();
    if (this.compression === 0) {
      this.account(raw.length, raw.length, false);
      return raw;
    }
    const parts: Uint8Array[] = [];
    let n = 0;
    for (const c of this.chunks(1 << 20)) {
      parts.push(c);
      n += c.length;
    }
    if (parts.length === 1) return parts[0]!;
    const buf = new Uint8Array(n);
    let o = 0;
    for (const c of parts) {
      buf.set(c, o);
      o += c.length;
    }
    return buf;
  }

  *chunks(chunkSize = 64 * 1024): Generator<Uint8Array, void, undefined> {
    const raw = this.raw();
    const step = Math.max(SLICE, chunkSize | 0);
    const len = raw.length;
    if (this.compression === 0) {
      for (let i = 0; i < len; i += step) {
        const end = Math.min(i + step, len);
        this.account(end, end, false);
        yield raw.subarray(i, end);
      }
      return;
    }
    const pending: Uint8Array[] = [];
    let out = 0;
    let fed = 0;
    let done = false;
    const inf = inflateStream((d, final) => {
      if (final) done = true;
      if (d.length === 0) return;
      out += d.length;
      this.account(out, fed, true);
      pending.push(d);
    });
    if (len === 0) inf.push(EMPTY, true);
    for (let i = 0; i < len; i += step) {
      const end = Math.min(i + step, len);
      for (let j = i; j < end; j += SLICE) {
        const e = Math.min(j + SLICE, end);
        fed = e;
        inf.push(raw.subarray(j, e), e === len);
      }
      for (let k = 0; k < pending.length; k++) yield pending[k]!;
      pending.length = 0;
    }
    if (!done) throw new ZipLimitError('Truncated zip archive');
  }
}
