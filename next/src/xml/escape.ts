// XML escaping for text nodes and attribute values.
//
// Both functions also strip characters that are not allowed in XML 1.0:
// C0 controls other than TAB/LF/CR, U+FFFE, U+FFFF and unpaired surrogates.
// Fast path: input is scanned once with charCodeAt and returned unchanged when
// nothing needs escaping (the overwhelmingly common case for cell data).

/** Index of the first char at or after i that needs escaping or dropping, or n. */
function scanText(s: string, i: number, n: number): number {
  for (; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x20) {
      if (c !== 0x09 && c !== 0x0a && c !== 0x0d) return i;
    } else if (c === 0x26 || c === 0x3c || c === 0x3e) {
      return i;
    } else if (c >= 0xd800) {
      if (c <= 0xdbff) {
        // high surrogate: valid only if followed by a low surrogate
        const d = i + 1 < n ? s.charCodeAt(i + 1) : 0;
        if (d >= 0xdc00 && d <= 0xdfff) {
          i++;
          continue;
        }
        return i;
      }
      if (c <= 0xdfff || c === 0xfffe || c === 0xffff) return i;
    }
  }
  return n;
}

function scanAttr(s: string, i: number, n: number): number {
  for (; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x20 || c === 0x26 || c === 0x3c || c === 0x3e || c === 0x22) return i;
    if (c >= 0xd800) {
      if (c <= 0xdbff) {
        const d = i + 1 < n ? s.charCodeAt(i + 1) : 0;
        if (d >= 0xdc00 && d <= 0xdfff) {
          i++;
          continue;
        }
        return i;
      }
      if (c <= 0xdfff || c === 0xfffe || c === 0xffff) return i;
    }
  }
  return n;
}

function slow(s: string, start: number, attr: boolean): string {
  const n = s.length;
  let out = '';
  let last = 0;
  let i = start;
  while (i < n) {
    const c = s.charCodeAt(i);
    let rep: string | undefined;
    if (c === 0x26) rep = '&amp;';
    else if (c === 0x3c) rep = '&lt;';
    else if (c === 0x3e) rep = '&gt;';
    else if (c === 0x22 && attr) rep = '&quot;';
    else if (c < 0x20) {
      if (c === 0x09) rep = attr ? '&#9;' : undefined;
      else if (c === 0x0a) rep = attr ? '&#10;' : undefined;
      else if (c === 0x0d) rep = attr ? '&#13;' : undefined;
      else rep = ''; // invalid control char: drop
    } else if (c >= 0xd800) {
      if (c <= 0xdbff) {
        const d = i + 1 < n ? s.charCodeAt(i + 1) : 0;
        if (d >= 0xdc00 && d <= 0xdfff) {
          i += 2;
          continue;
        }
        rep = '';
      } else if (c <= 0xdfff || c === 0xfffe || c === 0xffff) rep = '';
    }
    if (rep !== undefined) {
      out += s.slice(last, i) + rep;
      last = i + 1;
    }
    i++;
  }
  return last === 0 ? s : out + s.slice(last);
}

/** Escape text-node content: `& < >`, strip chars invalid in XML 1.0. */
export function escapeText(s: string): string {
  const n = s.length;
  const i = scanText(s, 0, n);
  return i === n ? s : slow(s, i, false);
}

/** Escape a (double-quoted) attribute value: `& < > "` and TAB/LF/CR as char refs; strip invalid chars. */
export function escapeAttr(s: string): string {
  const n = s.length;
  const i = scanAttr(s, 0, n);
  return i === n ? s : slow(s, i, true);
}

// ---- ST_Xstring (cell text): characters XML 1.0 cannot carry are written as `_xHHHH_` (Excel's
// own encoding), so they survive a round trip instead of being stripped; a literal `_xHHHH_`
// in the text is protected as `_x005F_xHHHH_`.

function isHex(c: number): boolean {
  return (c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102);
}

/** Does `_xHHHH_` start at i (s[i] === '_')? */
function xPattern(s: string, i: number): boolean {
  return (
    s.charCodeAt(i + 1) === 120 &&
    s.charCodeAt(i + 6) === 95 &&
    isHex(s.charCodeAt(i + 2)) &&
    isHex(s.charCodeAt(i + 3)) &&
    isHex(s.charCodeAt(i + 4)) &&
    isHex(s.charCodeAt(i + 5))
  );
}

function xHex(c: number): string {
  return `_x${c.toString(16).toUpperCase().padStart(4, '0')}_`;
}

/** Replacement for the char at `i` (advancing `skip` for a valid surrogate pair), or undefined. */
function xRep(s: string, i: number, n: number): string | undefined | null {
  const c = s.charCodeAt(i);
  if (c === 0x26) return '&amp;';
  if (c === 0x3c) return '&lt;';
  if (c === 0x3e) return '&gt;';
  if (c === 0x5f) return xPattern(s, i) ? '_x005F_' : undefined;
  if (c < 0x20) return c === 0x09 || c === 0x0a ? undefined : xHex(c);
  if (c >= 0xd800) {
    if (c <= 0xdbff) {
      const d = i + 1 < n ? s.charCodeAt(i + 1) : 0;
      return d >= 0xdc00 && d <= 0xdfff ? null : xHex(c); // null: valid pair, keep both
    }
    if (c <= 0xdfff || c === 0xfffe || c === 0xffff) return xHex(c);
  }
  return undefined;
}

/**
 * Escape cell text for a text node: like {@link escapeText}, but control characters, CR, U+FFFE/FFFF
 * and unpaired surrogates become `_xHHHH_` instead of being dropped, and a literal `_xHHHH_`
 * becomes `_x005F_xHHHH_`. Same single-scan fast path.
 */
export function escapeXString(s: string): string {
  const n = s.length;
  let out = '';
  let last = 0;
  for (let i = 0; i < n; i++) {
    const c = s.charCodeAt(i);
    // fast skip of the common case
    if (c >= 0x20 && c < 0xd800 && c !== 0x26 && c !== 0x3c && c !== 0x3e && c !== 0x5f) continue;
    const rep = xRep(s, i, n);
    if (rep === null) i++;
    else if (rep !== undefined) {
      out += s.slice(last, i) + rep;
      last = i + 1;
    }
  }
  return last === 0 ? s : out + s.slice(last);
}
