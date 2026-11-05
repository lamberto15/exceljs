// Thin adapter over ZipStreamWriter: sequential parts, string chunks encoded with TextEncoder.

import { type AsyncCompressor, ZipStreamWriter } from '../../zip/index.js';

type ZipFile = { push(data: Uint8Array | string, final?: boolean): void };

const encoder = /* @__PURE__ */ new TextEncoder();

/** Max uncompressed bytes queued for the native compressor before writers wait. */
const QUEUE_MAX = 4 << 20;

export class ZipOut {
  private readonly zip: ZipStreamWriter;
  private cur: ZipFile | undefined = undefined;
  private curPath = '';
  /** Entries placed early in the archive by `reserve()` and filled later. */
  private readonly held = new Map<string, ZipFile>();

  /**
   * `compressor`: async entry compressor (async writers pass `nativeCompressor()`); output then
   * arrives asynchronously — use `ready()` for backpressure and await `idle()`.
   */
  constructor(
    onData: (chunk: Uint8Array, final: boolean) => void,
    level: number,
    compressor?: AsyncCompressor,
  ) {
    this.zip = new ZipStreamWriter(onData, level, { compressor });
  }

  /** Wait (async mode only) until little enough input is queued for compression. */
  ready(): Promise<void> | undefined {
    return this.zip.ready(QUEUE_MAX);
  }

  /** Resolves when everything written so far has been compressed and emitted. */
  idle(): Promise<void> {
    return this.zip.idle();
  }

  /**
   * Put entries for `paths` in the archive now and fill them later with `write()`: they come
   * before every entry written afterwards, and the zip holds back the output of later entries
   * until they are complete. Used to store workbook metadata before the worksheets, which
   * streaming readers need, although it is only known after the sheets have been encoded.
   * Every reserved path must be written (and closed) before `end()`.
   */
  reserve(paths: readonly string[]): void {
    for (const p of paths) this.held.set(p, this.zip.file(p));
  }

  /** Append `data` to part `path` (opened on first use); `last` closes the part. */
  write(path: string, data: string | Uint8Array, last: boolean): void {
    const h = this.held.get(path);
    if (h !== undefined) {
      h.push(typeof data === 'string' ? encoder.encode(data) : data, last);
      if (last) this.held.delete(path);
      return;
    }
    let f = this.cur;
    if (f === undefined || this.curPath !== path) {
      if (f !== undefined) throw new Error(`zip part ${this.curPath} was not closed`);
      f = this.zip.file(path);
      this.cur = f;
      this.curPath = path;
    }
    f.push(typeof data === 'string' ? encoder.encode(data) : data, last);
    if (last) this.cur = undefined;
  }

  /** Write a complete part. */
  add(path: string, data: string | Uint8Array): void {
    this.write(path, data, true);
  }

  end(): void {
    if (this.cur !== undefined) throw new Error(`zip part ${this.curPath} was not closed`);
    if (this.held.size > 0) throw new Error(`zip parts never written: ${[...this.held.keys()]}`);
    this.zip.end();
  }
}

export function concatChunks(chunks: readonly Uint8Array[]): Uint8Array {
  let size = 0;
  for (const c of chunks) size += c.length;
  const out = new Uint8Array(size);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}
