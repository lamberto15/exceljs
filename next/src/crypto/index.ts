// `./crypto` entry: password-protected ("Encrypt with Password") OOXML files. The zip package is
// wrapped in an OLE compound file with `EncryptionInfo` + `EncryptedPackage` streams
// (MS-OFFCRYPTO). Kept out of the main entry so it costs nothing unless imported.
//
//   const wb = readXlsx(await decryptXlsx(bytes, 'secret'));
//   const out = await encryptXlsx(writeXlsx(wb), 'secret');

import { DEFAULT_LIMITS, type Limits } from '../core/types.js';
import { agileDecrypt, agileEncrypt, type EncryptOptions } from './agile.js';
import { isCfb, readCfb, writeCfb } from './cfb.js';
import { standardDecrypt } from './standard.js';
import { concat, IntegrityError, UnsupportedEncryptionError, utf16 } from './util.js';

export type { EncryptOptions } from './agile.js';
export { CfbError } from './cfb.js';
export { IntegrityError, UnsupportedEncryptionError, WrongPasswordError } from './util.js';

export interface DecryptOptions {
  /** Verify the agile dataIntegrity HMAC and reject tampered files (default true). */
  verifyIntegrity?: boolean;
  /** Size limits for the container streams and the decrypted package. */
  limits?: Partial<Limits>;
  /** Reject files asking for more password-hash iterations than this (default 10,000,000). */
  maxSpinCount?: number;
}

/** True for an encrypted OOXML container (CFB signature + an `EncryptionInfo` stream). */
export function isEncrypted(data: Uint8Array): boolean {
  if (!isCfb(data)) return false;
  try {
    return readCfb(data).entries.some((e) => e.path === 'EncryptionInfo');
  } catch {
    return false;
  }
}

/**
 * Decrypt a password-protected .xlsx (or any encrypted OOXML file); returns the inner zip bytes.
 * Throws WrongPasswordError, IntegrityError (tampered / corrupt), UnsupportedEncryptionError
 * (RC4, extensible, unknown algorithms) or CfbError (malformed container).
 */
export async function decryptXlsx(
  data: Uint8Array,
  password: string,
  opts: DecryptOptions = {},
): Promise<Uint8Array> {
  const cfb = readCfb(data, opts.limits);
  const info = cfb.read('EncryptionInfo');
  const pkg = cfb.read('EncryptedPackage');
  if (!info || !pkg) throw new IntegrityError('Not an encrypted OOXML file');
  if (info.length < 8) throw new IntegrityError('EncryptionInfo too short');
  const major = info[0]! | (info[1]! << 8);
  const minor = info[2]! | (info[3]! << 8);
  const maxBytes = opts.limits?.maxEntryBytes ?? DEFAULT_LIMITS.maxEntryBytes;
  if (major === 4 && minor === 4) {
    const o = { ...opts, maxBytes };
    return agileDecrypt(new TextDecoder().decode(info.subarray(8)), pkg, password, o);
  }
  if (minor === 2 && major >= 2 && major <= 4)
    return standardDecrypt(info, pkg, password, maxBytes);
  throw new UnsupportedEncryptionError(`Unsupported EncryptionInfo version ${major}.${minor}`);
}

// \x06DataSpaces storage (MS-OFFCRYPTO 2.1, 2.3.4.15 "strong encryption" data space). Constant
// content; Office writes it alongside agile-encrypted packages, so we do too.
function lp(s: string): Uint8Array {
  // UNICODE-LP-P4: u32 byte length, UTF-16LE, padded to 4 bytes
  const b = utf16(s);
  const out = new Uint8Array(4 + Math.ceil(b.length / 4) * 4);
  new DataView(out.buffer).setUint32(0, b.length, true);
  out.set(b, 4);
  return out;
}
function u32s(...v: number[]): Uint8Array {
  const out = new Uint8Array(v.length * 4);
  v.forEach((x, i) => {
    new DataView(out.buffer).setUint32(i * 4, x, true);
  });
  return out;
}
const VERSIONS = u32s(1, 1, 1); // reader / updater / writer 1.0
function dataSpaces(): [string, Uint8Array][] {
  const D = '\x06DataSpaces/';
  const map = concat(lp('EncryptedPackage'), lp('StrongEncryptionDataSpace'));
  const id = lp('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  return [
    [`${D}Version`, concat(lp('Microsoft.Container.DataSpaces'), VERSIONS)],
    [`${D}DataSpaceMap`, concat(u32s(8, 1, 12 + map.length, 1, 0), map)],
    [
      `${D}DataSpaceInfo/StrongEncryptionDataSpace`,
      concat(u32s(8, 1), lp('StrongEncryptionTransform')),
    ],
    [
      `${D}TransformInfo/StrongEncryptionTransform/\x06Primary`,
      concat(
        u32s(8 + id.length, 1),
        id,
        lp('Microsoft.Container.EncryptionTransform'),
        VERSIONS,
        u32s(0, 0, 0, 4), // EncryptionName (empty), block size, cipher mode, reserved
      ),
    ],
  ];
}

/**
 * Encrypt zip bytes (e.g. from `writeXlsx`) with a password using agile encryption — what
 * Excel 2010+ writes: AES-256-CBC, SHA-512, 100,000 spins, HMAC data integrity.
 */
export async function encryptXlsx(
  zipBytes: Uint8Array,
  password: string,
  opts?: EncryptOptions,
): Promise<Uint8Array> {
  const [info, pkg] = await agileEncrypt(zipBytes, password, opts);
  return writeCfb([...dataSpaces(), ['EncryptionInfo', info], ['EncryptedPackage', pkg]]);
}
