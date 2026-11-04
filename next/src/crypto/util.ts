// Byte helpers, errors and the WebCrypto AES-CBC wrappers shared by the agile and standard modes.

export class WrongPasswordError extends Error {
  override name = 'WrongPasswordError';
  constructor() {
    super('Wrong password');
  }
}

/** The file uses an encryption scheme this library does not implement (RC4, extensible, ...). */
export class UnsupportedEncryptionError extends Error {
  override name = 'UnsupportedEncryptionError';
}

/** The encrypted package failed its HMAC (agile `dataIntegrity`) check or is structurally bad. */
export class IntegrityError extends Error {
  override name = 'IntegrityError';
}

export const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
};

export const le32 = (n: number): Uint8Array => new Uint8Array([n, n >>> 8, n >>> 16, n >>> 24]);

/** Truncate to `n` bytes, or pad with 0x36 (MS-OFFCRYPTO key/IV sizing rule). */
export const sized = (b: Uint8Array, n: number): Uint8Array => {
  const out = new Uint8Array(n).fill(0x36);
  out.set(b.subarray(0, n));
  return out;
};

/** Zero-pad to a multiple of 16 bytes (AES block). */
export const pad16 = (b: Uint8Array): Uint8Array => zeroPad(b, Math.ceil(b.length / 16) * 16);
const zeroPad = (b: Uint8Array, n: number): Uint8Array => {
  const out = new Uint8Array(n);
  out.set(b.subarray(0, n));
  return out;
};

export const utf16 = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i * 2] = c;
    out[i * 2 + 1] = c >> 8;
  }
  return out;
};

export const equal = (a: Uint8Array, b: Uint8Array): boolean => {
  let d = a.length ^ b.length;
  for (let i = 0; i < a.length && i < b.length; i++) d |= a[i]! ^ b[i]!;
  return d === 0;
};

const subtle = (): SubtleCrypto => {
  const s = globalThis.crypto?.subtle;
  if (!s) throw new UnsupportedEncryptionError('WebCrypto (crypto.subtle) is not available');
  return s;
};

// WebCrypto wants ArrayBuffer-backed views.
const buf = (b: Uint8Array): Uint8Array<ArrayBuffer> => b as Uint8Array<ArrayBuffer>;

export const aesKey = (raw: Uint8Array): Promise<CryptoKey> =>
  subtle().importKey('raw', buf(raw), 'AES-CBC', false, ['encrypt', 'decrypt']);

/**
 * AES-CBC encrypt without padding (`data.length % 16 === 0`). WebCrypto always appends a PKCS#7
 * block; with block-aligned input that block is a whole extra block, so dropping it leaves the
 * exact unpadded ciphertext.
 */
export const cbcEncrypt = async (
  key: CryptoKey,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> =>
  new Uint8Array(
    await subtle().encrypt({ name: 'AES-CBC', iv: buf(iv) }, key, buf(data)),
    0,
    data.length,
  );

/**
 * AES-CBC decrypt without padding. WebCrypto insists on valid PKCS#7, so we append one crafted
 * block E = AES(C_last XOR 0x10^16): it decrypts to sixteen 0x10 bytes — a full padding block —
 * which WebCrypto strips, returning exactly the plaintext of `data`.
 */
export const cbcDecrypt = async (
  key: CryptoKey,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> => {
  if (data.length % 16) throw new IntegrityError('Ciphertext is not a multiple of the block size');
  if (!data.length) return data;
  const e = await cbcEncrypt(key, data.subarray(data.length - 16), new Uint8Array(16).fill(16));
  return new Uint8Array(
    await subtle().decrypt({ name: 'AES-CBC', iv: buf(iv) }, key, buf(concat(data, e))),
  );
};

export const hmac = async (
  hash: string,
  key: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> => {
  const s = subtle();
  const k = await s.importKey('raw', buf(key), { name: 'HMAC', hash }, false, ['sign']);
  return new Uint8Array(await s.sign('HMAC', k, buf(data)));
};

export const random = (n: number): Uint8Array =>
  globalThis.crypto.getRandomValues(new Uint8Array(n));

export const b64 = (s: string): Uint8Array => {
  try {
    return Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  } catch {
    throw new IntegrityError('Invalid base64 in EncryptionInfo');
  }
};

export const toB64 = (b: Uint8Array): string => {
  let s = '';
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]!);
  return btoa(s);
};
