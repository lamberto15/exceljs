// Small synchronous SHA-512 (FIPS 180-4) on 32-bit halves. Used for sheet/workbook protection
// password hashes, which sync writers must compute without WebCrypto (async only).
// Round constants are derived on first use (fractional parts of cube / square roots of the
// first primes, via BigInt) instead of shipping a 1.3 KB table.

let K: Int32Array | undefined;
let IV: Int32Array | undefined;

function iroot(n: bigint, k: bigint): bigint {
  // integer k-th root by Newton iteration
  let x = 1n << (BigInt(n.toString(2).length) / k + 1n);
  for (;;) {
    const y = ((k - 1n) * x + n / x ** (k - 1n)) / k;
    if (y >= x) return x;
    x = y;
  }
}

function init(): void {
  const k = new Int32Array(160);
  const iv = new Int32Array(16);
  const mask = (1n << 64n) - 1n;
  let found = 0;
  for (let p = 2; found < 80; p++) {
    let prime = true;
    for (let d = 2; d * d <= p; d++) {
      if (p % d === 0) {
        prime = false;
        break;
      }
    }
    if (!prime) continue;
    const c = iroot(BigInt(p) << 192n, 3n) & mask;
    k[found * 2] = Number(c >> 32n) | 0;
    k[found * 2 + 1] = Number(c & 0xffffffffn) | 0;
    if (found < 8) {
      const s = iroot(BigInt(p) << 128n, 2n) & mask;
      iv[found * 2] = Number(s >> 32n) | 0;
      iv[found * 2 + 1] = Number(s & 0xffffffffn) | 0;
    }
    found++;
  }
  K = k;
  IV = iv;
}

const W = /* @__PURE__ */ new Int32Array(160);

/** SHA-512 digest (64 bytes). */
export function sha512(msg: Uint8Array): Uint8Array {
  if (K === undefined) init();
  const k = K!;
  const H = IV!.slice();
  const len = msg.length;
  const blocks = ((len + 16) >> 7) + 1;
  const buf = new Uint8Array(blocks * 128);
  buf.set(msg);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(buf.length - 8, Math.floor(len / 0x20000000), false);
  dv.setUint32(buf.length - 4, (len << 3) >>> 0, false);
  const w = W;
  for (let b = 0; b < blocks; b++) {
    const off = b * 128;
    for (let t = 0; t < 32; t++) w[t] = dv.getInt32(off + t * 4, false);
    for (let t = 32; t < 160; t += 2) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of w[t-15]; s1 = rotr19 ^ rotr61 ^ shr6 of w[t-2]
      let xh = w[t - 30]!;
      let xl = w[t - 29]!;
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = w[t - 4]!;
      xl = w[t - 3]!;
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l =
        ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      const lo = (s0l >>> 0) + (s1l >>> 0) + (w[t - 13]! >>> 0) + (w[t - 31]! >>> 0);
      w[t] = s0h + s1h + w[t - 14]! + w[t - 32]! + Math.floor(lo / 0x100000000);
      w[t + 1] = lo;
    }
    let ah = H[0]!;
    let al = H[1]!;
    let bh = H[2]!;
    let bl = H[3]!;
    let ch = H[4]!;
    let cl = H[5]!;
    let dh = H[6]!;
    let dl = H[7]!;
    let eh = H[8]!;
    let el = H[9]!;
    let fh = H[10]!;
    let fl = H[11]!;
    let gh = H[12]!;
    let gl = H[13]!;
    let hh = H[14]!;
    let hl = H[15]!;
    for (let t = 0; t < 160; t += 2) {
      // S1 = rotr14 ^ rotr18 ^ rotr41 (e); ch = (e & f) ^ (~e & g)
      const S1h =
        ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l =
        ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      const t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (k[t + 1]! >>> 0) + (w[t + 1]! >>> 0);
      const t1h = hh + S1h + chh + k[t]! + w[t]! + Math.floor(t1l / 0x100000000);
      // S0 = rotr28 ^ rotr34 ^ rotr39 (a); maj = (a & b) ^ (a & c) ^ (b & c)
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      hh = gh;
      hl = gl;
      gh = fh;
      gl = fl;
      fh = eh;
      fl = el;
      let lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x100000000)) | 0;
      el = lo | 0;
      dh = ch;
      dl = cl;
      ch = bh;
      cl = bl;
      bh = ah;
      bl = al;
      lo = (t1l >>> 0) + (S0l >>> 0) + (mjl >>> 0);
      ah = (t1h + S0h + mjh + Math.floor(lo / 0x100000000)) | 0;
      al = lo | 0;
    }
    const v = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
    for (let i = 0; i < 16; i += 2) {
      const lo = (H[i + 1]! >>> 0) + (v[i + 1]! >>> 0);
      H[i] = H[i]! + v[i]! + Math.floor(lo / 0x100000000);
      H[i + 1] = lo;
    }
  }
  const out = new Uint8Array(64);
  const ov = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) ov.setInt32(i * 4, H[i]!, false);
  return out;
}
