// ECMA-376 standard encryption (MS-OFFCRYPTO 2.3.4.5–2.3.4.9) — read only. Excel 2007-era files:
// binary EncryptionInfo (version 3.2/4.2), AES-128/192/256-ECB, SHA-1 with 50,000 iterations.

import { decryptPackage } from './agile.js';
import { sha1 } from './hash.js';
import {
  aesKey,
  cbcDecrypt,
  concat,
  equal,
  IntegrityError,
  le32,
  UnsupportedEncryptionError,
  utf16,
  WrongPasswordError,
} from './util.js';

/** AES-ECB decrypt via CBC: CBC with a zero IV, then undo the chaining XOR block by block. */
async function ecbDecrypt(key: CryptoKey, data: Uint8Array): Promise<Uint8Array> {
  const out = await cbcDecrypt(key, new Uint8Array(16), data);
  for (let i = 16; i < out.length; i++) out[i]! ^= data[i - 16]!;
  return out;
}

/** Key derivation (2.3.4.7). */
export function standardKey(salt: Uint8Array, password: string, keyBytes: number): Uint8Array {
  let h = sha1(concat(salt, utf16(password)));
  const buf = new Uint8Array(24);
  for (let i = 0; i < 50000; i++) {
    buf.set(le32(i));
    buf.set(h, 4);
    h = sha1(buf);
  }
  h = sha1(concat(h, le32(0)));
  const x = (c: number) => {
    const b = new Uint8Array(64).fill(c);
    for (let i = 0; i < 20; i++) b[i]! ^= h[i]!;
    return sha1(b);
  };
  return concat(x(0x36), x(0x5c)).subarray(0, keyBytes);
}

/** Decrypt a standard-encrypted package; `info` is the whole EncryptionInfo stream. */
export async function standardDecrypt(
  info: Uint8Array,
  pkg: Uint8Array,
  password: string,
  maxBytes: number,
): Promise<Uint8Array> {
  if (info.length < 12) throw new IntegrityError('EncryptionInfo too short');
  const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const flags = dv.getUint32(4, true);
  const hdrSize = dv.getUint32(8, true);
  const v = 12 + hdrSize;
  if (hdrSize < 32 || v + 72 > info.length) throw new IntegrityError('Bad EncryptionInfo header');
  const algId = dv.getUint32(20, true);
  const keyBits = dv.getUint32(28, true) || 128;
  // fCryptoAPI without fAES (or the RC4 AlgID) is RC4 CryptoAPI; fExternal is extensible.
  if (flags & 0x10 || !(flags & 0x20) || algId === 0x6801)
    throw new UnsupportedEncryptionError('RC4 CryptoAPI / extensible encryption is not supported');
  if (![128, 192, 256].includes(keyBits) || dv.getUint32(24, true) !== 0x8004)
    throw new UnsupportedEncryptionError(`Unsupported standard encryption (${keyBits}-bit)`);
  const saltSize = dv.getUint32(v, true);
  if (saltSize !== 16) throw new IntegrityError('Bad verifier salt size');
  const salt = info.subarray(v + 4, v + 20);
  const key = await aesKey(standardKey(salt, password, keyBits / 8));
  const verifier = await ecbDecrypt(key, info.subarray(v + 20, v + 36));
  const vhash = await ecbDecrypt(key, info.subarray(v + 40, v + 72));
  if (!equal(vhash.subarray(0, 20), sha1(verifier))) throw new WrongPasswordError();
  return decryptPackage(pkg, maxBytes, (data) => ecbDecrypt(key, data));
}
