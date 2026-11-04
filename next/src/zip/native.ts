// Zip entries deflated by the platform's native `CompressionStream('deflate-raw')` (Node >= 21.2,
// Deno, Bun, modern browsers). In Node the compression runs on the libuv thread pool, so it
// overlaps with XML generation on the main thread. The zip container (local headers with data
// descriptors, CRC-32, central directory) is still fflate's `Zip`: each entry is a
// `ZipPassThrough` subclass whose `process` hands chunks to the native compressor and reports
// compressed output asynchronously (fflate's Zip supports async entries and keeps them ordered).
// Native deflate has no level knob; callers use this only for level > 0.

import { ZipPassThrough } from 'fflate';

let available: boolean | undefined;

/** True when `CompressionStream('deflate-raw')` exists. */
export function nativeDeflateAvailable(): boolean {
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

/** A zip entry compressed with the native deflater. Add it to an fflate `Zip`. */
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
