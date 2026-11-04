// Minimal push-style raw DEFLATE streams (thin fflate wrappers), so other modules never import
// fflate directly.

import { Deflate, Inflate } from 'fflate';

export interface ByteStream {
  /** Push input; `final = true` on the last push flushes and ends the stream. */
  push(data: Uint8Array, final?: boolean): void;
}

export type ByteSink = (data: Uint8Array, final: boolean) => void;

type Level = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/** Raw DEFLATE compressor. Output chunks are fresh arrays (safe to keep). */
export function deflateStream(level: number, onData: ByteSink): ByteStream {
  const d = new Deflate({ level: level as Level }, onData);
  return { push: (data, final = false) => d.push(data, final) };
}

/** Raw DEFLATE decompressor. Output chunks are fresh arrays (safe to keep). */
export function inflateStream(onData: ByteSink): ByteStream {
  const i = new Inflate(onData);
  return { push: (data, final = false) => i.push(data, final) };
}
