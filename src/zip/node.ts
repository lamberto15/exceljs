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
