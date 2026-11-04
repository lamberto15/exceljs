import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { sha512 } from '../src/core/sha512.js';
import { spinHash } from '../src/crypto/agile.js';
import { CfbError, isCfb, readCfb, writeCfb } from '../src/crypto/cfb.js';
import { hashByName, sha1, sha256, sha384 } from '../src/crypto/hash.js';
import {
  decryptXlsx,
  encryptXlsx,
  IntegrityError,
  isEncrypted,
  UnsupportedEncryptionError,
  WrongPasswordError,
} from '../src/crypto/index.js';
import { aesKey, cbcDecrypt, cbcEncrypt } from '../src/crypto/util.js';

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

// ---------------------------------------------------------------------------------------------
// Crypto. The reference implementation below is written independently from MS-OFFCRYPTO on top of
// node:crypto (OpenSSL) primitives — it shares no code with src/crypto except the CFB container.

const nodeHash: Record<string, string> = {
  SHA1: 'sha1',
  SHA256: 'sha256',
  SHA384: 'sha384',
  SHA512: 'sha512',
};
const H = (alg: string, ...parts: Uint8Array[]): Buffer => {
  const h = createHash(nodeHash[alg]!);
  for (const p of parts) h.update(p);
  return h.digest();
};
const u32le = (n: number): Buffer => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n);
  return b;
};
const fitRef = (b: Uint8Array, n: number): Buffer => {
  const out = Buffer.alloc(n, 0x36);
  Buffer.from(b).copy(out, 0, 0, Math.min(n, b.length));
  return out;
};
const padRef = (b: Uint8Array): Buffer => {
  const out = Buffer.alloc(Math.ceil(b.length / 16) * 16);
  Buffer.from(b).copy(out);
  return out;
};
const aes = (enc: boolean, mode: string, key: Uint8Array, iv: Uint8Array | null, d: Uint8Array) => {
  const c = (enc ? createCipheriv : createDecipheriv)(`aes-${key.length * 8}-${mode}`, key, iv);
  c.setAutoPadding(false);
  return Buffer.concat([c.update(d), c.final()]);
};
const BK = {
  vin: Buffer.from('fea7d2763b4b9e79', 'hex'),
  vhash: Buffer.from('d7aa0f6d3061344e', 'hex'),
  key: Buffer.from('146e0be7abacd0d6', 'hex'),
  hkey: Buffer.from('5fb2ad010cb9e1f6', 'hex'),
  hval: Buffer.from('a0677f02b22c8433', 'hex'),
};
function refPwHash(alg: string, salt: Uint8Array, pw: string, spin: number): Buffer {
  let h = H(alg, salt, Buffer.from(pw, 'utf16le'));
  for (let i = 0; i < spin; i++) h = H(alg, u32le(i), h);
  return h;
}

function refAgileEncrypt(zip: Uint8Array, pw: string, alg = 'SHA512', bits = 256, spin = 1000) {
  const kb = bits / 8;
  const hs = createHash(nodeHash[alg]!).digest().length;
  const kdSalt = randomBytes(16);
  const pwSalt = randomBytes(16);
  const secret = randomBytes(kb);
  const segs: Buffer[] = [];
  const size = Buffer.alloc(8);
  size.writeBigUInt64LE(BigInt(zip.length));
  for (let s = 0, o = 0; o < zip.length; s++, o += 4096) {
    const iv = fitRef(H(alg, kdSalt, u32le(s)), 16);
    segs.push(aes(true, 'cbc', secret, iv, padRef(zip.subarray(o, o + 4096))));
  }
  const pkg = Buffer.concat([size, ...segs]);
  const hmacKey = randomBytes(hs);
  const mac = createHmac(nodeHash[alg]!, hmacKey).update(pkg).digest();
  const eHk = aes(true, 'cbc', secret, fitRef(H(alg, kdSalt, BK.hkey), 16), padRef(hmacKey));
  const eHv = aes(true, 'cbc', secret, fitRef(H(alg, kdSalt, BK.hval), 16), padRef(mac));
  const ph = refPwHash(alg, pwSalt, pw, spin);
  const k = (b: Buffer) => fitRef(H(alg, ph, b), kb);
  const vin = randomBytes(16);
  const eVin = aes(true, 'cbc', k(BK.vin), pwSalt, vin);
  const eVh = aes(true, 'cbc', k(BK.vhash), pwSalt, padRef(H(alg, vin)));
  const eKey = aes(true, 'cbc', k(BK.key), pwSalt, padRef(secret));
  const c = (salt: Buffer) =>
    `saltSize="16" blockSize="16" keyBits="${bits}" hashSize="${hs}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${alg}" saltValue="${salt.toString('base64')}"`;
  // Different namespace prefix and attribute quoting than ours, to exercise the parser.
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:pw="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData ${c(kdSalt)}/><dataIntegrity encryptedHmacKey='${eHk.toString('base64')}' encryptedHmacValue="${eHv.toString('base64')}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><pw:encryptedKey spinCount="${spin}" ${c(pwSalt)} encryptedVerifierHashInput="${eVin.toString('base64')}" encryptedVerifierHashValue="${eVh.toString('base64')}" encryptedKeyValue="${eKey.toString('base64')}"/></keyEncryptor></keyEncryptors></encryption>`;
  const info = Buffer.concat([Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0]), Buffer.from(xml)]);
  return writeCfb([
    ['EncryptionInfo', info],
    ['EncryptedPackage', pkg],
  ]);
}

function refAgileDecrypt(file: Uint8Array, pw: string): Buffer {
  const cfb = readCfb(file);
  const info = Buffer.from(cfb.read('EncryptionInfo')!);
  const pkg = Buffer.from(cfb.read('EncryptedPackage')!);
  expect([...info.subarray(0, 8)]).toEqual([4, 0, 4, 0, 0x40, 0, 0, 0]);
  const xml = info.subarray(8).toString('utf8');
  const attr = (el: string, name: string) =>
    new RegExp(`<${el}\\b[^>]*\\s${name}="([^"]*)"`).exec(xml)![1]!;
  const alg = attr('p:encryptedKey', 'hashAlgorithm');
  const kb = Number(attr('p:encryptedKey', 'keyBits')) / 8;
  const pwSalt = Buffer.from(attr('p:encryptedKey', 'saltValue'), 'base64');
  const ph = refPwHash(alg, pwSalt, pw, Number(attr('p:encryptedKey', 'spinCount')));
  const k = (b: Buffer) => fitRef(H(alg, ph, b), kb);
  const d = (key: Buffer, a: string) =>
    aes(false, 'cbc', key, pwSalt, Buffer.from(attr('p:encryptedKey', a), 'base64'));
  const vin = d(k(BK.vin), 'encryptedVerifierHashInput').subarray(0, 16);
  const vh = d(k(BK.vhash), 'encryptedVerifierHashValue');
  const hv = H(alg, vin);
  if (!vh.subarray(0, hv.length).equals(hv)) throw new Error('ref: wrong password');
  const secret = d(k(BK.key), 'encryptedKeyValue').subarray(0, kb);
  const kdSalt = Buffer.from(attr('keyData', 'saltValue'), 'base64');
  const hk = aes(
    false,
    'cbc',
    secret,
    fitRef(H(alg, kdSalt, BK.hkey), 16),
    Buffer.from(attr('dataIntegrity', 'encryptedHmacKey'), 'base64'),
  ).subarray(0, hv.length);
  const hval = aes(
    false,
    'cbc',
    secret,
    fitRef(H(alg, kdSalt, BK.hval), 16),
    Buffer.from(attr('dataIntegrity', 'encryptedHmacValue'), 'base64'),
  ).subarray(0, hv.length);
  if (!createHmac(nodeHash[alg]!, hk).update(pkg).digest().equals(hval))
    throw new Error('ref: hmac mismatch');
  const size = Number(pkg.readBigUInt64LE(0));
  const out: Buffer[] = [];
  for (let s = 0, o = 8; o < pkg.length; s++, o += 4096) {
    const iv = fitRef(H(alg, kdSalt, u32le(s)), 16);
    out.push(aes(false, 'cbc', secret, iv, pkg.subarray(o, o + 4096)));
  }
  return Buffer.concat(out).subarray(0, size);
}

describe('crypto primitives', () => {
  it('sync SHA-1/256/384/512 match node:crypto', () => {
    for (const n of [0, 1, 3, 55, 56, 63, 64, 111, 112, 127, 128, 129, 300, 1000]) {
      const m = bytes(n, n + 1);
      expect(Buffer.from(sha1(m))).toEqual(H('SHA1', m));
      expect(Buffer.from(sha256(m))).toEqual(H('SHA256', m));
      expect(Buffer.from(sha384(m))).toEqual(H('SHA384', m));
      expect(Buffer.from(sha512(m))).toEqual(H('SHA512', m));
    }
  });

  it('unpadded AES-CBC via WebCrypto matches OpenSSL (incl. NIST SP 800-38A F.2.5)', async () => {
    // NIST SP 800-38A F.2.5 CBC-AES256.Encrypt, first two blocks.
    const key = Buffer.from(
      '603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4',
      'hex',
    );
    const iv = Buffer.from('000102030405060708090a0b0c0d0e0f', 'hex');
    const pt = Buffer.from(
      '6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51',
      'hex',
    );
    const ct = Buffer.from(
      'f58c4c04d6e5f1ba779eabfb5f7bfbd69cfc4e967edb808d679f777bc6702c7d',
      'hex',
    );
    const k = await aesKey(key);
    expect(Buffer.from(await cbcEncrypt(k, iv, pt))).toEqual(ct);
    expect(Buffer.from(await cbcDecrypt(k, iv, ct))).toEqual(pt);
    for (const n of [16, 32, 4096, 4112]) {
      const d = bytes(n, n);
      const k2 = randomBytes(16);
      const iv2 = randomBytes(16);
      const c = await cbcEncrypt(await aesKey(k2), iv2, d);
      expect(Buffer.from(c)).toEqual(aes(true, 'cbc', k2, iv2, d));
      expect(await cbcDecrypt(await aesKey(k2), iv2, c)).toEqual(d);
    }
    expect((await cbcDecrypt(k, iv, new Uint8Array(0))).length).toBe(0);
  });

  it('password spin hash matches the reference', () => {
    const salt = bytes(16, 5);
    for (const alg of ['SHA1', 'SHA256', 'SHA384', 'SHA512'])
      expect(Buffer.from(spinHash(hashByName(alg)!, salt, 'pässwörd', 500))).toEqual(
        refPwHash(alg, salt, 'pässwörd', 500),
      );
  });
});

describe('agile encryption', () => {
  const zip = (n: number) => bytes(n, n + 11);

  it('decrypts files from the independent reference encryptor', async () => {
    for (const [alg, bits] of [
      ['SHA512', 256],
      ['SHA1', 128],
      ['SHA256', 192],
      ['SHA384', 256],
    ] as const) {
      const z = zip(10000);
      const f = refAgileEncrypt(z, 'Secret!', alg, bits);
      expect(isEncrypted(f)).toBe(true);
      expect(await decryptXlsx(f, 'Secret!')).toEqual(z);
      await expect(decryptXlsx(f, 'secret!')).rejects.toBeInstanceOf(WrongPasswordError);
    }
  });

  it('reference decryptor reads what we encrypt', async () => {
    for (const [hash, cipher] of [
      ['SHA512', 'AES-256'],
      ['SHA1', 'AES-128'],
    ] as const) {
      const z = zip(9000);
      const f = await encryptXlsx(z, 'pw', { hash, cipher, spinCount: 1000 });
      expect(refAgileDecrypt(f, 'pw')).toEqual(Buffer.from(z));
    }
  });

  it('round-trips segment boundary sizes', async () => {
    for (const n of [0, 1, 15, 16, 4095, 4096, 4097, 8192, 1024 * 1024 + 3]) {
      const z = zip(n);
      const f = await encryptXlsx(z, 'boundary', { spinCount: 10 });
      expect(await decryptXlsx(f, 'boundary')).toEqual(z);
      expect(refAgileDecrypt(f, 'boundary')).toEqual(Buffer.from(z));
    }
  });

  it('writes Excel-shaped EncryptionInfo and DataSpaces', async () => {
    const f = await encryptXlsx(zip(100), 'x', { spinCount: 1 });
    const cfb = readCfb(f);
    expect(cfb.entries.map((e) => e.path).sort()).toEqual(
      [
        '\x06DataSpaces',
        '\x06DataSpaces/DataSpaceInfo',
        '\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace',
        '\x06DataSpaces/DataSpaceMap',
        '\x06DataSpaces/TransformInfo',
        '\x06DataSpaces/TransformInfo/StrongEncryptionTransform',
        '\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary',
        '\x06DataSpaces/Version',
        'EncryptedPackage',
        'EncryptionInfo',
      ].sort(),
    );
    // Byte layout of the constant DataSpaces streams (MS-OFFCRYPTO 2.1.5–2.1.8, 2.2.x).
    expect(Buffer.from(cfb.read('\x06DataSpaces/Version')!).toString('hex')).toBe(
      `3c000000${Buffer.from('Microsoft.Container.DataSpaces', 'utf16le').toString('hex')}010000000100000001000000`,
    );
    const map = Buffer.from(cfb.read('\x06DataSpaces/DataSpaceMap')!);
    expect(map.length).toBe(112);
    expect(map.subarray(0, 24).toString('hex')).toBe(
      '080000000100000068000000010000000000000020000000',
    );
    const prim = Buffer.from(
      cfb.read('\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary')!,
    );
    expect(prim.subarray(0, 12).toString('hex')).toBe('58000000010000004c000000');
    expect(prim.length).toBe(200);
    const xml = new TextDecoder().decode(cfb.read('EncryptionInfo')!.subarray(8));
    expect(xml).toMatch(
      /^<\?xml version="1\.0" encoding="UTF-8" standalone="yes"\?>\r\n<encryption xmlns="http:\/\/schemas\.microsoft\.com\/office\/2006\/encryption" xmlns:p="http:\/\/schemas\.microsoft\.com\/office\/2006\/keyEncryptor\/password" xmlns:c="http:\/\/schemas\.microsoft\.com\/office\/2006\/keyEncryptor\/certificate"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="[^"]+"\/><dataIntegrity /,
    );
  });

  it('rejects tampered ciphertext via the HMAC (unless told to skip)', async () => {
    const z = zip(20000);
    const f = await encryptXlsx(z, 'pw', { spinCount: 10 });
    const cfb = readCfb(f);
    const pe = cfb.entries.find((e) => e.path === 'EncryptedPackage')!;
    // Flip a byte in the last ciphertext segment (regular sectors are contiguous).
    const g = f.slice();
    g[(pe.start + 1) * 512 + 15000]! ^= 1;
    await expect(decryptXlsx(g, 'pw')).rejects.toBeInstanceOf(IntegrityError);
    const loose = await decryptXlsx(g, 'pw', { verifyIntegrity: false });
    expect(loose.length).toBe(z.length);
    expect(loose).not.toEqual(z);
  });

  it('rejects absurd spin counts and unsupported versions', async () => {
    const f = refAgileEncrypt(zip(10), 'pw', 'SHA512', 256, 5);
    await expect(decryptXlsx(f, 'pw', { maxSpinCount: 4 })).rejects.toBeInstanceOf(
      UnsupportedEncryptionError,
    );
    const ext = writeCfb([
      ['EncryptionInfo', new Uint8Array([4, 0, 3, 0, 0, 0, 0, 0])],
      ['EncryptedPackage', new Uint8Array(16)],
    ]);
    await expect(decryptXlsx(ext, 'pw')).rejects.toBeInstanceOf(UnsupportedEncryptionError);
  });

  it('times spinCount 100000 (Excel default)', async () => {
    const z = zip(1024 * 1024);
    let t = performance.now();
    const f = await encryptXlsx(z, 'timing');
    const enc = performance.now() - t;
    t = performance.now();
    expect(await decryptXlsx(f, 'timing')).toEqual(z);
    const dec = performance.now() - t;
    console.log(
      `agile SHA-512 x100000, 1 MB: encrypt ${enc.toFixed(0)} ms, decrypt ${dec.toFixed(0)} ms`,
    );
    expect(enc).toBeLessThan(5000);
  });
});

describe('isEncrypted', () => {
  it('is false for zips, garbage and CFB without EncryptionInfo', () => {
    expect(isEncrypted(new Uint8Array([0x50, 0x4b, 3, 4]))).toBe(false);
    expect(isEncrypted(bytes(2000))).toBe(false);
    expect(isEncrypted(writeCfb([['Workbook', bytes(100)]]))).toBe(false);
  });
});
