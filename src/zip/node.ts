// Node / Bun / Deno built-ins, looked up at runtime with `process.getBuiltinModule` (Node >= 22.3,
// Bun, Deno 2). There is no static `node:*` import, so browser bundles stay free of Node code and
// every caller falls back to Web APIs / fflate when the lookup fails.

/** The subset of a `zlib.DeflateRaw` transform stream used here. */
export interface NodeDeflateStream {
  write(chunk: Uint8Array, cb: (err?: Error | null) => void): boolean;
  end(): void;
  on(event: 'data', fn: (chunk: Uint8Array) => void): unknown;
  on(event: 'end', fn: () => void): unknown;
  on(event: 'error', fn: (err: Error) => void): unknown;
}

/** The subset of `node:zlib` used here. */
export interface NodeZlib {
  createDeflateRaw(opts: { level: number; chunkSize?: number }): NodeDeflateStream;
  deflateRawSync(data: Uint8Array, opts: { level: number; finishFlush?: number }): Uint8Array;
  inflateRawSync(data: Uint8Array, opts?: { maxOutputLength?: number }): Uint8Array;
  /** Node >= 22.2. */
  crc32?: (data: Uint8Array, value?: number) => number;
}

let zlib: NodeZlib | null | undefined;

/** `node:zlib`, or undefined outside Node-compatible runtimes. */
export function nodeZlib(): NodeZlib | undefined {
  if (zlib === undefined) {
    zlib = null;
    try {
      const p = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } })
        .process;
      const z = p?.getBuiltinModule?.('node:zlib') as NodeZlib | undefined;
      if (typeof z?.createDeflateRaw === 'function') zlib = z;
    } catch {
      // not available
    }
  }
  return zlib ?? undefined;
}

/** A running CRC-32 with fflate's `crc()` interface (`p` = update, `d` = digest). */
export interface Crc {
  p(data: Uint8Array): void;
  d(): number;
}

/**
 * A CRC-32 backed by `zlib.crc32` (~5x fflate's byte-at-a-time table), or undefined when the
 * runtime lacks it.
 */
export function nativeCrc(): Crc | undefined {
  const f = nodeZlib()?.crc32;
  if (typeof f !== 'function') return undefined;
  let c = 0;
  return {
    p(d) {
      c = f(d, c);
    },
    d: () => c,
  };
}

/** zlib flush flags (stable values from zlib.h). */
const Z_SYNC_FLUSH = 2;
const Z_FINISH = 4;

/** Uncompressed bytes collected before each synchronous deflate call. */
const SYNC_BATCH = 1 << 20;

/** Plain Uint8Array view of a Buffer. */
const view = (b: Uint8Array): Uint8Array<ArrayBuffer> =>
  new Uint8Array(b.buffer as ArrayBuffer, b.byteOffset, b.byteLength);

/**
 * Synchronous raw deflate built on `zlib.deflateRawSync`: input is collected into ~1 MiB
 * batches; every batch but the last ends with a sync flush (an empty stored block, no BFINAL),
 * so the concatenated outputs form one valid deflate stream. Each batch starts with an empty
 * window, which costs well under 0.1% of output at this batch size. Returns undefined without
 * `node:zlib`.
 */
export function zlibSyncDeflater(
  level: number,
  onData: (data: Uint8Array<ArrayBuffer>, final: boolean) => void,
): { push(data: Uint8Array, final: boolean): void } | undefined {
  const z = nodeZlib();
  if (typeof z?.deflateRawSync !== 'function') return undefined;
  let parts: Uint8Array[] = [];
  let size = 0;
  const run = (final: boolean): void => {
    let buf: Uint8Array;
    if (parts.length === 1) buf = parts[0]!;
    else {
      buf = new Uint8Array(size);
      let o = 0;
      for (const p of parts) {
        buf.set(p, o);
        o += p.length;
      }
    }
    parts = [];
    size = 0;
    onData(
      view(z.deflateRawSync(buf, { level, finishFlush: final ? Z_FINISH : Z_SYNC_FLUSH })),
      final,
    );
  };
  return {
    push(data, final) {
      if (data.length > 0) {
        parts.push(data);
        size += data.length;
      }
      if (final || size >= SYNC_BATCH) run(final);
    },
  };
}

/**
 * Inflate a whole raw-deflate buffer with `zlib.inflateRawSync`, producing at most `max` bytes.
 * Returns undefined without `node:zlib`, and 'over' when the output would exceed `max`.
 * Throws the zlib error for corrupt or truncated data.
 */
export function zlibInflateSync(data: Uint8Array, max: number): Uint8Array | 'over' | undefined {
  const z = nodeZlib();
  if (typeof z?.inflateRawSync !== 'function') return undefined;
  try {
    return view(z.inflateRawSync(data, { maxOutputLength: Math.max(1, max) }));
  } catch (e) {
    if ((e as { code?: string }).code === 'ERR_BUFFER_TOO_LARGE') return 'over';
    throw e;
  }
}
