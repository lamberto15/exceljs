// Zip entries deflated off the main thread by a native deflater:
// - Node / Bun / Deno: `zlib.createDeflateRaw({ level })` (libuv thread pool), so the
//   `compression` level is honored.
// - Elsewhere (browsers, workers): `CompressionStream('deflate-raw')`, which has no level knob
//   (the browser picks its default, ~6).
// Compression overlaps with XML generation on the main thread. The zip container (local headers
// with data descriptors, CRC-32, central directory) is still fflate's `Zip`: each entry is a
// `ZipPassThrough` subclass whose `process` hands chunks to the native compressor and reports
// compressed output asynchronously (fflate's Zip supports async entries and keeps them ordered).
// Callers use this only for level > 0.

import { ZipPassThrough } from 'fflate';
import { type NodeZlib, nodeZlib } from './node.js';
import type { AsyncCompressor } from './zip.js';

let available: boolean | undefined;

function webDeflateAvailable(): boolean {
  if (available === undefined) {
    try {
      new CompressionStream('deflate-raw' as CompressionFormat);
      available = true;
    } catch {
      available = false;
    }
  }
  return available;
}

/** True when a native deflater exists (`node:zlib` or `CompressionStream('deflate-raw')`). */
export function nativeDeflateAvailable(): boolean {
  return nodeZlib() !== undefined || webDeflateAvailable();
}

/** Tracks input handed to native compressors but not yet consumed (backpressure). */
export class Inflight {
  bytes = 0;
  error: unknown = undefined;
  private waiters: (() => void)[] = [];
  private tail: Promise<void> = Promise.resolve();

  /** Resolves once in-flight input is at most `max` bytes. Rejects on the first error. */
  below(max: number): Promise<void> | undefined {
    if (this.error !== undefined) return Promise.reject(this.error);
    if (this.bytes <= max) return undefined;
    return new Promise<void>((res) => this.waiters.push(res)).then(() => this.below(max));
  }

  /** Resolves when every entry handed out so far has been fully compressed. */
  idle(): Promise<void> {
    return this.tail.then(() => {
      if (this.error !== undefined) throw this.error;
    });
  }

  /** @internal */
  track(p: Promise<void>): void {
    this.tail = Promise.all([this.tail, p]).then(
      () => {},
      (e) => {
        this.fail(e);
      },
    );
  }

  /** @internal */
  done(n: number): void {
    this.bytes -= n;
    this.wake();
  }

  /** @internal */
  fail(e: unknown): void {
    this.error ??= e;
    this.wake();
  }

  private wake(): void {
    const w = this.waiters;
    this.waiters = [];
    for (const f of w) f();
  }
}

const EMPTY = /* @__PURE__ */ new Uint8Array(0);

/** A zip entry compressed with `CompressionStream` (no level). Add it to an fflate `Zip`. */
export class NativeDeflateFile extends ZipPassThrough {
  private readonly w: WritableStreamDefaultWriter<BufferSource>;
  private readonly q: Inflight;

  constructor(path: string, q: Inflight) {
    super(path);
    this.compression = 8;
    this.q = q;
    const cs = new CompressionStream('deflate-raw' as CompressionFormat);
    this.w = cs.writable.getWriter();
    const reader = cs.readable.getReader();
    const pump = async (): Promise<void> => {
      try {
        for (;;) {
          const r = await reader.read();
          if (r.done) break;
          if (r.value.length > 0) this.ondata(null, r.value as Uint8Array<ArrayBuffer>, false);
        }
        this.ondata(null, EMPTY, true);
      } catch (e) {
        q.fail(e);
        throw e;
      }
    };
    q.track(pump());
  }

  protected override process(chunk: Uint8Array<ArrayBuffer>, final: boolean): void {
    const q = this.q;
    const n = chunk.length;
    if (n > 0) {
      q.bytes += n;
      // resolves once the compressor has consumed the chunk
      this.w.write(chunk).then(
        () => q.done(n),
        (e) => q.fail(e),
      );
    }
    if (final) this.w.close().catch((e) => q.fail(e));
  }
}

/** Output chunk size of the zlib deflater (fewer, larger callbacks than the 16 KiB default). */
const ZLIB_CHUNK = 64 * 1024;

/** A zip entry compressed with `zlib.createDeflateRaw({ level })`. Add it to an fflate `Zip`. */
export class ZlibDeflateFile extends ZipPassThrough {
  private readonly d: ReturnType<NodeZlib['createDeflateRaw']>;
  private readonly q: Inflight;

  constructor(path: string, level: number, q: Inflight, z: NodeZlib) {
    super(path);
    this.compression = 8;
    this.q = q;
    const d = z.createDeflateRaw({ level, chunkSize: ZLIB_CHUNK });
    this.d = d;
    q.track(
      new Promise<void>((resolve, reject) => {
        d.on('data', (c) => {
          // plain Uint8Array view of the Buffer (zlib never reuses emitted output)
          const u = new Uint8Array(c.buffer, c.byteOffset, c.byteLength);
          if (u.length > 0) this.ondata(null, u as Uint8Array<ArrayBuffer>, false);
        });
        d.on('end', () => {
          this.ondata(null, EMPTY, true);
          resolve();
        });
        d.on('error', (e) => {
          q.fail(e);
          reject(e);
        });
      }),
    );
  }

  protected override process(chunk: Uint8Array<ArrayBuffer>, final: boolean): void {
    const q = this.q;
    const n = chunk.length;
    if (n > 0) {
      q.bytes += n;
      // the callback runs once zlib has consumed the chunk
      this.d.write(chunk, (e) => (e ? q.fail(e) : q.done(n)));
    }
    if (final) this.d.end();
  }
}

/**
 * An async compressor backed by `node:zlib` (honors the level) or, without it,
 * `CompressionStream('deflate-raw')` (browser default level); undefined when the runtime has
 * neither (callers then fall back to fflate). One per zip.
 */
export function nativeCompressor(): AsyncCompressor | undefined {
  const z = nodeZlib();
  if (z === undefined && !webDeflateAvailable()) return undefined;
  const q = new Inflight();
  return {
    file: z
      ? (path, level) => new ZlibDeflateFile(path, level, q, z)
      : (path) => new NativeDeflateFile(path, q),
    below: (max) => q.below(max),
    idle: () => q.idle(),
    get error() {
      return q.error;
    },
  };
}
