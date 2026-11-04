// Feature: worksheet protection (<sheetProtection>), ECMA-376 Part 1 §18.3.1.85.
//
//   ws.protection = { password: 'secret', formatColumns: true };   // protect, allow col format
//   ws.protection = undefined;                                     // unprotect
//
// The plaintext `password` is never written: at write time it becomes an agile hash
// (SHA-512, random 16-byte salt, `spinCount` iterations — Excel 2010+ format). A file read back
// carries algorithmName/hashValue/saltValue/spinCount instead (written back unchanged).
// Permission flags mean "users MAY do this while protected" (exceljs semantics); unset flags
// keep Excel's defaults (selecting cells allowed, everything else locked).

import { sha512 } from '../../core/sha512.js';
import type { Worksheet } from '../../core/worksheet.js';
import { type FeatureReader, type FeatureWriter, SheetSlot } from './types.js';
import { attrs, base64, bool, num, unbase64 } from './util.js';

export interface SheetProtection {
  /** Plaintext password, hashed at write time (omit for protection without password). */
  password?: string | undefined;
  /** Precomputed / read hash (used when `password` is not set). */
  algorithmName?: string | undefined;
  hashValue?: string | undefined;
  saltValue?: string | undefined;
  /** Hash iterations (default 100000, like Excel). Lower = faster writes, weaker hash. */
  spinCount?: number | undefined;
  /** Legacy 16-bit password hash (hex), as found in old files. */
  legacyPassword?: string | undefined;
  /** Protect the sheet (default true). */
  sheet?: boolean | undefined;
  objects?: boolean | undefined;
  scenarios?: boolean | undefined;
  selectLockedCells?: boolean | undefined;
  selectUnlockedCells?: boolean | undefined;
  formatCells?: boolean | undefined;
  formatColumns?: boolean | undefined;
  formatRows?: boolean | undefined;
  insertColumns?: boolean | undefined;
  insertRows?: boolean | undefined;
  insertHyperlinks?: boolean | undefined;
  deleteColumns?: boolean | undefined;
  deleteRows?: boolean | undefined;
  sort?: boolean | undefined;
  autoFilter?: boolean | undefined;
  pivotTables?: boolean | undefined;
}

/** Flags whose XML attribute is "locked" (1 = action forbidden); model true = allowed. */
const LOCK_FLAGS = [
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'sort',
  'autoFilter',
  'pivotTables',
] as const;

/**
 * ECMA-376 agile password hash: H0 = SHA512(salt + UTF-16LE(password)),
 * Hn = SHA512(Hn-1 + LE32(n-1)) for n = 1..spinCount; base64.
 */
export function hashPassword(password: string, saltB64: string, spinCount: number): string {
  const salt = unbase64(saltB64);
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    pw[i * 2] = c & 255;
    pw[i * 2 + 1] = c >> 8;
  }
  const first = new Uint8Array(salt.length + pw.length);
  first.set(salt);
  first.set(pw, salt.length);
  let h = sha512(first);
  const buf = new Uint8Array(68);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    buf[64] = i & 255;
    buf[65] = (i >>> 8) & 255;
    buf[66] = (i >>> 16) & 255;
    buf[67] = i >>> 24;
    h = sha512(buf);
  }
  return base64(h);
}

/** Hash cache: protection object → [password, hash attrs] (salt stays stable across writes). */
const hashed = new WeakMap<SheetProtection, [string, number, string, string]>();

function hashOf(p: SheetProtection): [string, string, number] | undefined {
  if (p.password) {
    const spin = Math.max(0, Math.round(p.spinCount ?? 100000));
    let c = hashed.get(p);
    if (!c || c[0] !== p.password || c[1] !== spin) {
      const salt = base64(crypto.getRandomValues(new Uint8Array(16)));
      c = [p.password, spin, salt, hashPassword(p.password, salt, spin)];
      hashed.set(p, c);
    }
    return [c[3], c[2], spin];
  }
  if (p.hashValue && p.saltValue) return [p.hashValue, p.saltValue, p.spinCount ?? 100000];
  return undefined;
}

function emit(ws: Worksheet): string {
  const p = ws.protection;
  if (!p) return '';
  const sheet = p.sheet !== false;
  const a: Record<string, string | number | boolean | undefined> = {};
  if (p.legacyPassword) a.password = p.legacyPassword;
  const h = sheet ? hashOf(p) : undefined;
  if (h) {
    a.algorithmName = p.password ? 'SHA-512' : (p.algorithmName ?? 'SHA-512');
    a.hashValue = h[0];
    a.saltValue = h[1];
    a.spinCount = h[2];
  }
  a.sheet = sheet;
  if (sheet) {
    a.objects = p.objects === false;
    a.scenarios = p.scenarios === false;
  }
  a.selectLockedCells = p.selectLockedCells === false;
  a.selectUnlockedCells = p.selectUnlockedCells === false;
  for (const f of LOCK_FLAGS) if (p[f]) a[f] = '0';
  const s = attrs(a);
  return s ? `<sheetProtection${s}/>` : '';
}

/**
 * Write half of the sheet protection (ws.protection) feature. Always on in the package root; in `./lite` pass it
 * in `features` of writeXlsx / xlsxStream / createXlsxWriter.
 */
export const sheetProtectionWriter: FeatureWriter = {
  name: 'sheet-protection',
  sheet: { emitters: [{ slot: SheetSlot.sheetProtection, emit: (c) => emit(c.ws) }] },
};

/**
 * Read half of the sheet protection (ws.protection) feature. Always on in the package root; in `./lite` pass it
 * in `features` of readXlsx (without it, the elements are kept raw and written back).
 */
export const sheetProtectionReader: FeatureReader = {
  name: 'sheet-protection',
  sheet: {
    elements: ['sheetProtection'],
    open(ctx, _el, a) {
      const p: SheetProtection = {};
      if (!bool(a.sheet)) p.sheet = false;
      if (a.algorithmName) {
        p.algorithmName = a.algorithmName;
        p.hashValue = a.hashValue;
        p.saltValue = a.saltValue;
        p.spinCount = num(a.spinCount);
      }
      if (a.password) p.legacyPassword = a.password;
      if (bool(a.objects)) p.objects = false;
      if (bool(a.scenarios)) p.scenarios = false;
      if (bool(a.selectLockedCells)) p.selectLockedCells = false;
      if (bool(a.selectUnlockedCells)) p.selectUnlockedCells = false;
      for (const f of LOCK_FLAGS) if (a[f] !== undefined && !bool(a[f], true)) p[f] = true;
      ctx.ws.protection = p;
      return undefined;
    },
  },
};
