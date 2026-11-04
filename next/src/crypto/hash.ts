// Synchronous SHA-1 / SHA-256 / SHA-384 / SHA-512 for the password spin loop (100k iterations of
// tiny inputs — awaiting a WebCrypto digest per iteration costs seconds). Bulk hashing (HMAC over
// the whole package) goes through WebCrypto instead.

import { iroot, sha512, sha512Consts } from '../core/sha512.js';

export type HashFn = (msg: Uint8Array) => Uint8Array;
export interface Hash {
  fn: HashFn;
  size: number;
  /** WebCrypto name. */
  web: string;
  /** Name written in the agile XML descriptor. */
  xml: string;
}

const rotl = (x: number, n: number): number => (x << n) | (x >>> (32 - n));

/** Merkle–Damgård padding into big-endian 32-bit words (64-byte blocks). */
function words(msg: Uint8Array): Int32Array {
  const n = ((msg.length + 8) >> 6) + 1;
  const w = new Int32Array(n * 16);
  for (let i = 0; i < msg.length; i++) w[i >> 2]! |= msg[i]! << (24 - (i & 3) * 8);
  w[msg.length >> 2]! |= 0x80 << (24 - (msg.length & 3) * 8);
  w[n * 16 - 2] = Math.floor(msg.length / 0x20000000);
  w[n * 16 - 1] = msg.length << 3;
  return w;
}

function bytes(h: ArrayLike<number>): Uint8Array {
  const out = new Uint8Array(h.length * 4);
  const dv = new DataView(out.buffer);
  for (let i = 0; i < h.length; i++) dv.setInt32(i * 4, h[i]!, false);
  return out;
}

export function sha1(msg: Uint8Array): Uint8Array {
  const m = words(msg);
  const h = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0];
  const w = new Int32Array(80);
  for (let b = 0; b < m.length; b += 16) {
    for (let t = 0; t < 80; t++)
      w[t] = t < 16 ? m[b + t]! : rotl(w[t - 3]! ^ w[t - 8]! ^ w[t - 14]! ^ w[t - 16]!, 1);
    let [a, bb, c, d, e] = h as [number, number, number, number, number];
    for (let t = 0; t < 80; t++) {
      const f =
        t < 20
          ? ((bb & c) | (~bb & d)) + 0x5a827999
          : t < 40
            ? (bb ^ c ^ d) + 0x6ed9eba1
            : t < 60
              ? ((bb & c) | (bb & d) | (c & d)) + 0x8f1bbcdc
              : (bb ^ c ^ d) + 0xca62c1d6;
      const tmp = (rotl(a, 5) + f + e + w[t]!) | 0;
      e = d;
      d = c;
      c = rotl(bb, 30);
      bb = a;
      a = tmp;
    }
    h[0] = (h[0]! + a) | 0;
    h[1] = (h[1]! + bb) | 0;
    h[2] = (h[2]! + c) | 0;
    h[3] = (h[3]! + d) | 0;
    h[4] = (h[4]! + e) | 0;
  }
  return bytes(h);
}

export function sha256(msg: Uint8Array): Uint8Array {
  // SHA-256 constants are the high words of SHA-512's.
  const [K, IV] = sha512Consts();
  const m = words(msg);
  const h = Int32Array.from({ length: 8 }, (_, i) => IV[i * 2]!);
  const w = new Int32Array(64);
  const rotr = (x: number, n: number): number => (x >>> n) | (x << (32 - n));
  for (let b = 0; b < m.length; b += 16) {
    for (let t = 0; t < 64; t++) {
      if (t < 16) w[t] = m[b + t]!;
      else {
        const x = w[t - 15]!;
        const y = w[t - 2]!;
        w[t] =
          (rotr(x, 7) ^ rotr(x, 18) ^ (x >>> 3)) +
          (rotr(y, 17) ^ rotr(y, 19) ^ (y >>> 10)) +
          w[t - 7]! +
          w[t - 16]!;
      }
    }
    let a = h[0]!;
    let bb = h[1]!;
    let c = h[2]!;
    let d = h[3]!;
    let e = h[4]!;
    let f = h[5]!;
    let g = h[6]!;
    let hh = h[7]!;
    for (let t = 0; t < 64; t++) {
      const t1 =
        hh + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + K[t * 2]! + w[t]!;
      const t2 = (rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & bb) ^ (a & c) ^ (bb & c));
      hh = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = bb;
      bb = a;
      a = (t1 + t2) | 0;
    }
    const v = [a, bb, c, d, e, f, g, hh];
    for (let i = 0; i < 8; i++) h[i] = h[i]! + v[i]!;
  }
  return bytes(h);
}

let IV384: Int32Array | undefined;
export function sha384(msg: Uint8Array): Uint8Array {
  if (!IV384) {
    // Fractional square roots of the 9th..16th primes.
    IV384 = new Int32Array(16);
    [23, 29, 31, 37, 41, 43, 47, 53].forEach((p, i) => {
      const s = iroot(BigInt(p) << 128n, 2n) & ((1n << 64n) - 1n);
      IV384![i * 2] = Number(s >> 32n) | 0;
      IV384![i * 2 + 1] = Number(s & 0xffffffffn) | 0;
    });
  }
  return sha512(msg, IV384).slice(0, 48);
}

const HASHES: Hash[] = [
  { fn: sha1, size: 20, web: 'SHA-1', xml: 'SHA1' },
  { fn: sha256, size: 32, web: 'SHA-256', xml: 'SHA256' },
  { fn: sha384, size: 48, web: 'SHA-384', xml: 'SHA384' },
  { fn: sha512, size: 64, web: 'SHA-512', xml: 'SHA512' },
];

/** Look up a hash by name; accepts `SHA512`, `SHA-512`, `sha512`, ... */
export function hashByName(name: string): Hash | undefined {
  const n = name.toUpperCase().replace('-', '');
  return HASHES.find((h) => h.xml === n);
}
