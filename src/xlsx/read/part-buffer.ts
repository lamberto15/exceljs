// Holds a decompressed zip entry that the streaming reader cannot process yet (its metadata
// arrives later in the archive). Bytes are kept as Uint8Array chunks (off the JS heap, ~1 byte
// per ASCII char); with `compress` they are re-deflated at level 1 (≈5-10x smaller for sheet XML,
// at the cost of one extra deflate + inflate pass).

import { type ByteStream, deflateStream, inflateStream } from '../../zip/index.js';

export class PartBuffer {
  private readonly chunks: Uint8Array[] = [];
  private readonly deflate: ByteStream | undefined;
  /** Bytes held (compressed size when compressing). */
  size = 0;

  constructor(compress: boolean) {
    if (compress) {
      this.deflate = deflateStream(1, (d) => {
        this.chunks.push(d);
        this.size += d.length;
      });
    }
  }

  /** Append data (the caller must not reuse `d` afterwards). */
  push(d: Uint8Array): void {
    if (this.deflate) this.deflate.push(d, false);
    else {
      this.chunks.push(d);
      this.size += d.length;
    }
  }

  end(): void {
    this.deflate?.push(new Uint8Array(0), true);
  }

  /** Decoded text, chunk by chunk; releases buffered data as it goes. */
  *text(): Generator<string, void, undefined> {
    const dec = new TextDecoder();
    const pending: string[] = [];
    const inflate = this.deflate
      ? inflateStream((d) => {
          const s = dec.decode(d, { stream: true });
          if (s) pending.push(s);
        })
      : undefined;
    const chunks = this.chunks;
    for (let i = 0; i < chunks.length; i++) {
      const c = chunks[i]!;
      chunks[i] = EMPTY;
      if (inflate) {
        inflate.push(c, i === chunks.length - 1);
        while (pending.length > 0) yield pending.shift()!;
      } else {
        const s = dec.decode(c, { stream: true });
        if (s) yield s;
      }
    }
    const tail = dec.decode();
    if (tail) yield tail;
    chunks.length = 0;
    this.size = 0;
  }
}

const EMPTY = /* @__PURE__ */ new Uint8Array(0);
