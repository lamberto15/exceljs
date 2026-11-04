// Tiny helpers for feature modules. Import from here (or src/xml, src/core) — NOT from
// src/xlsx/read/* or src/xlsx/write/* runtime modules: those would drag the whole reader into
// write-only bundles (and vice versa), since every registered feature ships in both.

import { escapeAttr } from '../../xml/index.js';

export { escapeAttr, escapeText } from '../../xml/index.js';
// round-trip: kept drawing anchors + relationship placeholder resolution
export { fillRels, presDrawing, presVml } from '../common/preserved.js';
// rich text (CT_Rst): `<r>` runs writer, `<t>` element, and the <si>/<is>-style run reader
export {
  runsXml,
  type StringItem,
  StringItemBuilder,
  tElem,
} from '../common/rich-text.js';

/** OOXML boolean attribute: "1" / "true" → true, "0" / "false" → false, absent → `dflt`. */
export function bool(v: string | undefined, dflt = false): boolean {
  if (v === undefined) return dflt;
  return v === '1' || v === 'true';
}

/** Numeric attribute, or undefined when absent / not a finite number. */
export function num(v: string | undefined): number | undefined {
  if (v === undefined || v === '') return undefined;
  const n = +v;
  return n - n === 0 ? n : undefined;
}

/**
 * Serialize attributes in the given order, skipping undefined / false. `true` → "1".
 * Returns ' a="1" b="x"' (leading space) or ''.
 */
export function attrs(a: Record<string, string | number | boolean | undefined>): string {
  let s = '';
  for (const k in a) {
    const v = a[k];
    if (v === undefined || v === false) continue;
    s += ` ${k}="${v === true ? '1' : typeof v === 'number' ? v : escapeAttr(v)}"`;
  }
  return s;
}

/** Base64 of bytes (web-standard btoa; no Buffer). */
export function base64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]!);
  return btoa(s);
}

/** Bytes of a base64 string. */
export function unbase64(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
