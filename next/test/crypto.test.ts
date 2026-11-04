import { describe, expect, it } from 'vitest';
import { CfbError, isCfb, readCfb, writeCfb } from '../src/crypto/cfb.js';

function bytes(n: number, seed = 1): Uint8Array {
  const out = new Uint8Array(n);
  let x = seed;
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0;
    out[i] = x >>> 24;
  }
  return out;
}

describe('cfb', () => {
  it('round-trips mini and regular streams, nested storages', () => {
    const files: [string, Uint8Array][] = [
      ['EncryptionInfo', bytes(1000, 2)],
      ['EncryptedPackage', bytes(20000, 3)],
      ['\x06DataSpaces/Version', bytes(76, 4)],
      ['\x06DataSpaces/DataSpaceMap', bytes(112, 5)],
      ['\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary', bytes(200, 6)],
      ['empty', new Uint8Array(0)],
      ['exact4096', bytes(4096, 7)],
      ['just4095', bytes(4095, 8)],
    ];
    const cfb = writeCfb(files);
    expect(isCfb(cfb)).toBe(true);
    expect(cfb.length % 512).toBe(0);
    const r = readCfb(cfb);
    for (const [p, d] of files) expect(r.read(p)).toEqual(d);
    expect(r.read('missing')).toBeUndefined();
    expect(r.entries.find((e) => e.path === '\x06DataSpaces')?.type).toBe(1);
  });

  it('writes and reads DIFAT for >109 FAT sectors', () => {
    const big = bytes(8 * 1024 * 1024 + 7, 9);
    const cfb = writeCfb([
      ['a', big],
      ['b', bytes(10, 1)],
    ]);
    const dv = new DataView(cfb.buffer);
    expect(dv.getUint32(44, true)).toBeGreaterThan(109); // FAT sectors
    expect(dv.getUint32(72, true)).toBeGreaterThan(0); // DIFAT sectors
    const r = readCfb(cfb);
    expect(r.read('a')).toEqual(big);
    expect(r.read('b')).toEqual(bytes(10, 1));
  });

  it('reads v4 (4096-byte sector) files', () => {
    // Re-layout a v3 file as v4 by hand: same structure, sector size 4096.
    const v3 = writeCfb([
      ['x', bytes(5000, 1)],
      ['y', bytes(100, 2)],
    ]);
    const nSec = v3.length / 512 - 1;
    const v4 = new Uint8Array(4096 * (nSec + 1));
    v4.set(v3.subarray(0, 512));
    const dv = new DataView(v4.buffer);
    dv.setUint16(26, 4, true);
    dv.setUint16(30, 12, true);
    // Each v3 sector goes into its own v4 sector; FAT entries keep their ids (128 per sector, so
    // the FAT stays valid with 4096-byte sectors: 1024 entries per sector, rest free).
    for (let i = 0; i < nSec; i++)
      v4.set(v3.subarray((i + 1) * 512, (i + 2) * 512), (i + 1) * 4096);
    const fatV3 = new DataView(v3.buffer, 512, 512);
    for (let j = 128; j < 1024; j++) dv.setUint32(4096 + j * 4, 0xffffffff, true);
    for (let j = 0; j < 128; j++) dv.setUint32(4096 + j * 4, fatV3.getUint32(j * 4, true), true);
    // x (5000 bytes) spans 10 v3 sectors, which as v4 sectors hold 4096 bytes each: rewrite as
    // a 2-sector chain whose contents are the original payload.
    const r3 = readCfb(v3);
    const ex = r3.entries.find((e) => e.path === 'x')!;
    const payload = r3.read('x')!;
    dv.setUint32(4096 + ex.start * 4, ex.start + 1, true);
    dv.setUint32(4096 + (ex.start + 1) * 4, 0xfffffffe, true);
    v4.set(payload.subarray(0, 4096), (ex.start + 1) * 4096);
    v4.set(payload.subarray(4096), (ex.start + 2) * 4096);
    const r = readCfb(v4);
    expect(r.read('x')).toEqual(payload);
    // mini stream (y) lives in the root's chain: 1 v3 sector → 1 v4 sector, still valid.
    expect(r.read('y')).toEqual(bytes(100, 2));
  });

  it('rejects non-CFB and truncated input with CfbError', () => {
    expect(() => readCfb(new Uint8Array(100))).toThrow(CfbError);
    expect(() => readCfb(bytes(4096))).toThrow(CfbError);
    const cfb = writeCfb([['EncryptedPackage', bytes(10000, 1)]]);
    expect(() => readCfb(cfb.subarray(0, 1536)).read('EncryptedPackage')).toThrow(CfbError);
  });

  it('detects FAT cycles', () => {
    const cfb = writeCfb([['s', bytes(10000, 1)]]);
    const r = readCfb(cfb);
    const e = r.entries[0]!;
    const dv = new DataView(cfb.buffer);
    dv.setUint32(512 + (e.start + 3) * 4, e.start, true); // loop back
    expect(() => readCfb(cfb).read('s')).toThrow(/chain|shorter/);
    // Directory chain cycle: dir sector points at itself.
    const c2 = writeCfb([['s', bytes(10, 1)]]);
    const d2 = new DataView(c2.buffer);
    const dir = d2.getUint32(48, true);
    d2.setUint32(512 + dir * 4, dir, true);
    expect(() => readCfb(c2)).toThrow(CfbError);
  });

  it('enforces the stream size limit', () => {
    const cfb = writeCfb([['s', bytes(10000, 1)]]);
    expect(() => readCfb(cfb, { maxEntryBytes: 5000 })).toThrow(/size limit/);
  });

  it('survives random corruption (only CfbError, bounded time)', () => {
    const base = writeCfb([
      ['EncryptionInfo', bytes(900, 2)],
      ['EncryptedPackage', bytes(9000, 3)],
      ['\x06DataSpaces/Version', bytes(76, 4)],
    ]);
    let seed = 7;
    const rnd = (n: number): number => {
      seed = (seed * 1103515245 + 12345) >>> 0;
      return seed % n;
    };
    const t0 = Date.now();
    for (let iter = 0; iter < 2000; iter++) {
      const d = base.slice(0, iter % 7 === 0 ? rnd(base.length) : base.length);
      for (let k = 0, n = 1 + rnd(8); k < n; k++) {
        // bias corruption towards header, FAT and directory
        const o = rnd(3) === 0 ? rnd(Math.min(d.length, 2048)) : rnd(d.length || 1);
        if (o < d.length) d[o] = rnd(256);
      }
      try {
        const r = readCfb(d);
        for (const e of r.entries) if (e.type === 2) r.read(e.path);
      } catch (e) {
        if (!(e instanceof CfbError)) throw e;
      }
    }
    expect(Date.now() - t0).toBeLessThan(5000);
  });
});
