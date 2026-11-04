// A1-style address helpers. Everything here is 0-based: row 0 = "1", col 0 = "A".

/** Number of rows in a worksheet (Excel 2007+). */
export const MAX_ROWS = 1_048_576;
/** Number of columns in a worksheet (Excel 2007+). "XFD" = 16383. */
export const MAX_COLS = 16_384;

/** 0-based, inclusive rectangle. */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** 0-based cell coordinate. */
export interface CellAddress {
  row: number;
  col: number;
}

const CACHE_SIZE = 1024;
const colCache: string[] = [];

function buildCol(col: number): string {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = ((n - 1 - m) / 26) | 0;
  }
  return s;
}

/** 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD". Throws when out of bounds. */
export function encodeCol(col: number): string {
  if (col < CACHE_SIZE && col >= 0) {
    const hit = colCache[col];
    if (hit !== undefined) return hit;
    if (col !== (col | 0)) throw new RangeError(`Invalid column index: ${col}`);
    const s = buildCol(col);
    colCache[col] = s;
    return s;
  }
  if (!(col >= 0 && col < MAX_COLS) || col !== (col | 0)) {
    throw new RangeError(`Column index out of range: ${col}`);
  }
  return buildCol(col);
}

/** "A" → 0, "XFD" → 16383. Case-insensitive. Throws on invalid input. */
export function decodeCol(letters: string): number {
  const n = letters.length;
  if (n === 0 || n > 3) throw new RangeError(`Invalid column: "${letters}"`);
  let v = 0;
  for (let i = 0; i < n; i++) {
    let c = letters.charCodeAt(i);
    if (c >= 97 && c <= 122) c -= 32;
    if (c < 65 || c > 90) throw new RangeError(`Invalid column: "${letters}"`);
    v = v * 26 + (c - 64);
  }
  if (v > MAX_COLS) throw new RangeError(`Column out of range: "${letters}"`);
  return v - 1;
}

/** (row 2, col 1) → "B3". */
export function encodeAddr(row: number, col: number): string {
  if (!(row >= 0 && row < MAX_ROWS) || row !== (row | 0)) {
    throw new RangeError(`Row index out of range: ${row}`);
  }
  return encodeCol(col) + (row + 1);
}

/**
 * "B3" → { row: 2, col: 1 }. Accepts "$B$3", lowercase letters and a sheet prefix ("Sheet1!B3").
 * Throws on invalid / out-of-bounds input.
 */
export function decodeAddr(addr: string): CellAddress {
  const out = { row: 0, col: 0 };
  if (!parseAddrInto(addr, 0, addr.length, out, 3)) {
    throw new RangeError(`Invalid cell address: "${addr}"`);
  }
  return out;
}

/**
 * Parse addr[start, end) into out. `need`: bit 1 = column required, bit 2 = row required.
 * Missing column → col = -1, missing row → row = -1. Returns false on invalid input.
 */
function parseAddrInto(
  s: string,
  start: number,
  end: number,
  out: CellAddress,
  need: number,
): boolean {
  const bang = s.lastIndexOf('!', end - 1);
  let i = bang >= start ? bang + 1 : start;
  if (i < end && s.charCodeAt(i) === 36) i++; // $
  let col = 0;
  let letters = 0;
  while (i < end) {
    let c = s.charCodeAt(i);
    if (c >= 97 && c <= 122) c -= 32;
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    letters++;
    i++;
  }
  if (letters > 3 || col > MAX_COLS) return false;
  if (i < end && s.charCodeAt(i) === 36) i++; // $
  let row = 0;
  let digits = 0;
  while (i < end) {
    const c = s.charCodeAt(i);
    if (c < 48 || c > 57) return false;
    row = row * 10 + (c - 48);
    digits++;
    i++;
  }
  if (digits > 7 || row > MAX_ROWS) return false;
  if (letters === 0) {
    if (need & 1) return false;
    col = 0;
  }
  if (digits === 0) {
    if (need & 2) return false;
    row = 0;
  } else if (row === 0) return false;
  if (letters === 0 && digits === 0) return false;
  out.col = col - 1; // -1 when absent
  out.row = row - 1; // -1 when absent
  return true;
}

/**
 * "A1:C3" → { top: 0, left: 0, bottom: 2, right: 2 } (normalized so top<=bottom, left<=right).
 * Also accepts a single cell ("B2"), whole columns ("A:C"), whole rows ("1:3") and a sheet prefix.
 */
export function decodeRange(ref: string): CellRange {
  const colon = ref.indexOf(':');
  const a = { row: 0, col: 0 };
  const b = { row: 0, col: 0 };
  let ok: boolean;
  if (colon < 0) {
    ok = parseAddrInto(ref, 0, ref.length, a, 3);
    b.row = a.row;
    b.col = a.col;
  } else {
    ok = parseAddrInto(ref, 0, colon, a, 0) && parseAddrInto(ref, colon + 1, ref.length, b, 0);
    if (ok) {
      // both must have the same "shape": both full, both column-only, or both row-only
      const aCol = a.col >= 0;
      const aRow = a.row >= 0;
      if (aCol !== b.col >= 0 || aRow !== b.row >= 0) ok = false;
      if (!aRow) {
        a.row = 0;
        b.row = MAX_ROWS - 1;
      }
      if (!aCol) {
        a.col = 0;
        b.col = MAX_COLS - 1;
      }
    }
  }
  if (!ok) throw new RangeError(`Invalid range: "${ref}"`);
  return {
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  };
}

/** { top: 0, left: 0, bottom: 2, right: 2 } → "A1:C3"; a single cell encodes as "A1". */
export function encodeRange(r: CellRange): string {
  const a = encodeAddr(r.top, r.left);
  if (r.top === r.bottom && r.left === r.right) return a;
  return `${a}:${encodeAddr(r.bottom, r.right)}`;
}
