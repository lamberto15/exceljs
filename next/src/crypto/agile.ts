// ECMA-376 agile encryption (MS-OFFCRYPTO 2.3.4.10–2.3.4.15): XML EncryptionInfo, password key
// encryptor, 4096-byte AES-CBC segments with per-segment IVs, HMAC data integrity.

import { type Hash, hashByName } from './hash.js';
import {
  aesKey,
  b64,
  cbcDecrypt,
  cbcEncrypt,
  concat,
  equal,
  hmac,
  IntegrityError,
  le32,
  pad16,
  random,
  sized,
  toB64,
  UnsupportedEncryptionError,
  utf16,
  WrongPasswordError,
} from './util.js';

const SEGMENT = 4096;
const BLOCK_VERIFIER_INPUT = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const BLOCK_VERIFIER_HASH = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const BLOCK_KEY = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const BLOCK_HMAC_KEY = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const BLOCK_HMAC_VALUE = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

/**
 * Password hash (2.3.4.11): H0 = H(salt + UTF-16LE(password)), Hn = H(LE32(n-1) + Hn-1).
 * Synchronous hash: 100k awaited WebCrypto digests would take seconds.
 */
export function spinHash(h: Hash, salt: Uint8Array, password: string, spin: number): Uint8Array {
  let H = h.fn(concat(salt, utf16(password)));
  const buf = new Uint8Array(4 + h.size);
  const dv = new DataView(buf.buffer);
  for (let i = 0; i < spin; i++) {
    dv.setUint32(0, i, true);
    buf.set(H, 4);
    H = h.fn(buf);
  }
  return H;
}

/** Key for one block key: H(Hn + blockKey), sized to `bytes` (truncate / pad 0x36). */
const blockKey = (h: Hash, H: Uint8Array, block: number[], bytes: number): Uint8Array =>
  sized(h.fn(concat(H, new Uint8Array(block))), bytes);

/** IV from the keyData salt and a block key or segment index. */
const ivFor = (h: Hash, salt: Uint8Array, extra: Uint8Array): Uint8Array =>
  sized(h.fn(concat(salt, extra)), 16);

type Attrs = Record<string, string>;

/** Attributes of the first element with local name `local` (optionally requiring `has`). */
function element(xml: string, local: string, has?: string): Attrs | undefined {
  for (const m of xml.matchAll(/<([\w.-]+:)?([\w.-]+)(\s[^>]*?)?\/?>/g)) {
    if (m[2] !== local) continue;
    const a: Attrs = {};
    for (const p of (m[3] ?? '').matchAll(/([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g))
      a[p[1]!.replace(/^.*:/, '')] = p[2] ?? p[3]!;
    if (!has || has in a) return a;
  }
  return undefined;
}

interface Cipher {
  hash: Hash;
  keyBytes: number;
  salt: Uint8Array;
}

function cipherOf(a: Attrs | undefined, what: string): Cipher {
  if (!a) throw new IntegrityError(`EncryptionInfo has no ${what}`);
  const hash = hashByName(a.hashAlgorithm ?? '');
  const bits = Number(a.keyBits);
  if (!hash || a.cipherAlgorithm !== 'AES' || ![128, 192, 256].includes(bits))
    throw new UnsupportedEncryptionError(
      `Unsupported agile cipher ${a.cipherAlgorithm}-${a.keyBits} / ${a.hashAlgorithm}`,
    );
  if ((a.cipherChaining ?? 'ChainingModeCBC') !== 'ChainingModeCBC' || Number(a.blockSize) !== 16)
    throw new UnsupportedEncryptionError(`Unsupported cipher chaining ${a.cipherChaining}`);
  return { hash, keyBytes: bits / 8, salt: b64(a.saltValue ?? '') };
}

export interface AgileDecryptOptions {
  /** Check the dataIntegrity HMAC (default true). */
  verifyIntegrity?: boolean;
  /** Reject descriptors asking for more hash iterations than this (default 10,000,000). */
  maxSpinCount?: number;
  /** Maximum decrypted size in bytes. */
  maxBytes?: number;
}

/** Decrypt an agile EncryptedPackage stream given the EncryptionInfo XML. */
export async function agileDecrypt(
  xml: string,
  pkg: Uint8Array,
  password: string,
  opts: AgileDecryptOptions = {},
): Promise<Uint8Array> {
  const ekA = element(xml, 'encryptedKey', 'spinCount');
  const ek = cipherOf(ekA, 'password key encryptor');
  const kd = cipherOf(element(xml, 'keyData'), 'keyData');
  const spin = Number(ekA!.spinCount);
  if (!(spin >= 0 && spin <= (opts.maxSpinCount ?? 1e7) && Number.isInteger(spin)))
    throw new UnsupportedEncryptionError(`Unsupported spinCount ${ekA!.spinCount}`);

  // Password → intermediate key (verifier check first).
  const H = spinHash(ek.hash, ek.salt, password, spin);
  const iv = sized(ek.salt, 16);
  const k = (b: number[]) => aesKey(blockKey(ek.hash, H, b, ek.keyBytes));
  const dec = async (key: Promise<CryptoKey>, attr: string) =>
    cbcDecrypt(await key, iv, b64(ekA![attr] ?? ''));
  const [vin, vhash, secret] = await Promise.all([
    dec(k(BLOCK_VERIFIER_INPUT), 'encryptedVerifierHashInput'),
    dec(k(BLOCK_VERIFIER_HASH), 'encryptedVerifierHashValue'),
    dec(k(BLOCK_KEY), 'encryptedKeyValue'),
  ]);
  const want = ek.hash.fn(vin.subarray(0, ek.salt.length));
  if (vhash.length < want.length || !equal(vhash.subarray(0, want.length), want))
    throw new WrongPasswordError();
  if (secret.length < kd.keyBytes) throw new IntegrityError('Encrypted key too short');
  const key = await aesKey(secret.subarray(0, kd.keyBytes));

  // Data integrity: HMAC over the whole EncryptedPackage stream.
  const di = element(xml, 'dataIntegrity');
  if (di && opts.verifyIntegrity !== false) {
    const h = kd.hash;
    const [hk, hv] = await Promise.all([
      cbcDecrypt(
        key,
        ivFor(h, kd.salt, new Uint8Array(BLOCK_HMAC_KEY)),
        b64(di.encryptedHmacKey ?? ''),
      ),
      cbcDecrypt(
        key,
        ivFor(h, kd.salt, new Uint8Array(BLOCK_HMAC_VALUE)),
        b64(di.encryptedHmacValue ?? ''),
      ),
    ]);
    const mac = await hmac(h.web, hk.subarray(0, h.size), pkg);
    if (hv.length < h.size || !equal(hv.subarray(0, h.size), mac))
      throw new IntegrityError(
        'Encrypted package failed its integrity check (tampered or corrupt)',
      );
  }
  return decryptPackage(pkg, opts.maxBytes ?? Infinity, async (data) => {
    // One WebCrypto call for the whole package with segment 0's IV, then repair the first block
    // of every later segment: CBC chained it to the previous ciphertext block instead of the
    // segment IV, so P = P' XOR C[prev] XOR IV[segment].
    const out = await cbcDecrypt(key, ivFor(kd.hash, kd.salt, le32(0)), data);
    for (let s = 1, o = SEGMENT; o < out.length; s++, o += SEGMENT) {
      const sIv = ivFor(kd.hash, kd.salt, le32(s));
      for (let i = 0; i < 16; i++) out[o + i]! ^= data[o - 16 + i]! ^ sIv[i]!;
    }
    return out;
  });
}

/** Shared EncryptedPackage framing: LE64 size, then block-aligned ciphertext. */
export async function decryptPackage(
  pkg: Uint8Array,
  maxBytes: number,
  decrypt: (data: Uint8Array) => Promise<Uint8Array>,
): Promise<Uint8Array> {
  if (pkg.length < 8) throw new IntegrityError('EncryptedPackage stream too short');
  const dv = new DataView(pkg.buffer, pkg.byteOffset, 8);
  const size = dv.getUint32(4, true) * 0x100000000 + dv.getUint32(0, true);
  const n = Math.ceil(size / 16) * 16;
  if (n > pkg.length - 8) throw new IntegrityError('EncryptedPackage is truncated');
  if (size > maxBytes) throw new IntegrityError('Decrypted package exceeds size limit');
  return (await decrypt(pkg.subarray(8, 8 + n))).subarray(0, size);
}

export interface EncryptOptions {
  /** AES key size. Default 'AES-256'. */
  cipher?: 'AES-128' | 'AES-192' | 'AES-256';
  /** Hash algorithm. Default 'SHA512'. */
  hash?: 'SHA1' | 'SHA256' | 'SHA384' | 'SHA512';
  /** Password hash iterations. Default 100000 (what Excel writes). */
  spinCount?: number;
}

/** Encrypt a package; returns the EncryptionInfo and EncryptedPackage stream bytes. */
export async function agileEncrypt(
  data: Uint8Array,
  password: string,
  opts: EncryptOptions = {},
): Promise<[info: Uint8Array, pkg: Uint8Array]> {
  const h = hashByName(opts.hash ?? 'SHA512');
  const bits = Number((opts.cipher ?? 'AES-256').slice(4));
  const spin = opts.spinCount ?? 100000;
  if (!h || ![128, 192, 256].includes(bits) || !(Number.isInteger(spin) && spin >= 0))
    throw new UnsupportedEncryptionError('Unsupported encryption options');
  const keyBytes = bits / 8;
  const kdSalt = random(16);
  const pwSalt = random(16);
  const secret = random(keyBytes);
  const key = await aesKey(secret);

  // Package: LE64 size + 4096-byte segments, last one zero-padded to the block size.
  const out = new Uint8Array(8 + Math.ceil(data.length / 16) * 16);
  new DataView(out.buffer).setUint32(0, data.length % 0x100000000, true);
  new DataView(out.buffer).setUint32(4, Math.floor(data.length / 0x100000000), true);
  const jobs: Promise<void>[] = [];
  for (let s = 0, o = 0; o < data.length; s++, o += SEGMENT) {
    const seg = pad16(data.subarray(o, o + SEGMENT));
    jobs.push(cbcEncrypt(key, ivFor(h, kdSalt, le32(s)), seg).then((c) => out.set(c, 8 + o)));
    if (jobs.length >= 256) await Promise.all(jobs.splice(0));
  }
  await Promise.all(jobs);

  // Data integrity.
  const hmacKey = random(h.size);
  const mac = await hmac(h.web, hmacKey, out);
  const encHmacKey = await cbcEncrypt(
    key,
    ivFor(h, kdSalt, new Uint8Array(BLOCK_HMAC_KEY)),
    pad16(hmacKey),
  );
  const encHmacValue = await cbcEncrypt(
    key,
    ivFor(h, kdSalt, new Uint8Array(BLOCK_HMAC_VALUE)),
    pad16(mac),
  );

  // Password key encryptor.
  const H = spinHash(h, pwSalt, password, spin);
  const vin = random(16);
  const enc = async (block: number[], plain: Uint8Array) =>
    toB64(await cbcEncrypt(await aesKey(blockKey(h, H, block, keyBytes)), pwSalt, pad16(plain)));
  const common = `saltSize="16" blockSize="16" keyBits="${bits}" hashSize="${h.size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${h.xml}" saltValue="`;
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
    `<keyData ${common}${toB64(kdSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${toB64(encHmacKey)}" encryptedHmacValue="${toB64(encHmacValue)}"/>` +
    '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    `<p:encryptedKey spinCount="${spin}" ${common}${toB64(pwSalt)}" encryptedVerifierHashInput="${await enc(BLOCK_VERIFIER_INPUT, vin)}" encryptedVerifierHashValue="${await enc(BLOCK_VERIFIER_HASH, h.fn(vin))}" encryptedKeyValue="${await enc(BLOCK_KEY, secret)}"/>` +
    '</keyEncryptor></keyEncryptors></encryption>';
  // Version 4.4, flags 0x40 (agile), then the XML.
  return [concat(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml)), out];
}
