// Shared-formula expansion: translate the relative references of a master formula by a
// (row, col) offset, as Excel does when it fills a formula down/right.
//
// Handles: A1 / $A$1 / mixed refs, ranges (each side translated), whole-column (A:C) and
// whole-row (1:3) ranges, sheet-qualified refs (Sheet1!A1, 'My Sheet'!A1), and leaves string
// literals, quoted sheet names, structured references ([...]), function names (LOG10(...)) and
// sheet names that look like refs (AB1!A1) untouched. References pushed off the grid become #REF!.

import { encodeCol, MAX_COLS, MAX_ROWS } from '../../core/address.js';

function isIdent(c: number): boolean {
  return (
    (c >= 65 && c <= 90) ||
    (c >= 97 && c <= 122) ||
    (c >= 48 && c <= 57) ||
    c === 36 /* $ */ ||
    c === 95 /* _ */ ||
    c === 46 /* . */ ||
    c === 92 /* \ */ ||
    c > 127
  );
}

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;
const COL = /^(\$?)([A-Za-z]{1,3})$/;
const ROW = /^(\$?)([0-9]{1,7})$/;

function colNum(letters: string): number {
  let v = 0;
  for (let i = 0; i < letters.length; i++) v = v * 26 + ((letters.charCodeAt(i) | 0x20) - 96);
  return v - 1;
}

function shiftCell(tok: string, dr: number, dc: number): string | undefined {
  const m = CELL.exec(tok);
  if (!m) return undefined;
  let c = colNum(m[2]!);
  let r = +m[4]! - 1;
  if (c >= MAX_COLS || r >= MAX_ROWS || r < 0) return undefined;
  if (!m[1]) c += dc;
  if (!m[3]) r += dr;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return '#REF!';
  return `${m[1]}${encodeCol(c)}${m[3]}${r + 1}`;
}

function shiftCol(tok: string, dc: number): string | undefined {
  const m = COL.exec(tok);
  if (!m) return undefined;
  let c = colNum(m[2]!);
  if (c >= MAX_COLS) return undefined;
  if (!m[1]) c += dc;
  if (c < 0 || c >= MAX_COLS) return '#REF!';
  return m[1] + encodeCol(c);
}

function shiftRow(tok: string, dr: number): string | undefined {
  const m = ROW.exec(tok);
  if (!m) return undefined;
  let r = +m[2]! - 1;
  if (r < 0 || r >= MAX_ROWS) return undefined;
  if (!m[1]) r += dr;
  if (r < 0 || r >= MAX_ROWS) return '#REF!';
  return `${m[1]}${r + 1}`;
}

function readIdent(f: string, i: number): number {
  const n = f.length;
  while (i < n && isIdent(f.charCodeAt(i))) i++;
  return i;
}

/** Translate `formula` (no leading '=') by `dr` rows and `dc` columns. */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;
  const f = formula;
  const n = f.length;
  let out = '';
  let last = 0;
  let i = 0;
  while (i < n) {
    const c = f.charCodeAt(i);
    if (c === 34 /* " */ || c === 39 /* ' */) {
      // string literal or quoted sheet name; doubled quote = escaped quote
      let j = i + 1;
      for (;;) {
        const k = f.indexOf(c === 34 ? '"' : "'", j);
        if (k < 0) {
          j = n;
          break;
        }
        if (f.charCodeAt(k + 1) === c) j = k + 2;
        else {
          j = k + 1;
          break;
        }
      }
      i = j;
      continue;
    }
    if (c === 91 /* [ */) {
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        const d = f.charCodeAt(j);
        if (d === 91) depth++;
        else if (d === 93 && --depth === 0) break;
      }
      i = j + 1;
      continue;
    }
    if (!isIdent(c)) {
      i++;
      continue;
    }
    const e = readIdent(f, i);
    const next = f.charCodeAt(e);
    if (next === 40 /* ( */ || next === 33 /* ! */ || next === 91 /* [ */) {
      i = e;
      continue;
    }
    const tok = f.slice(i, e);
    let rep = shiftCell(tok, dr, dc);
    let end = e;
    if (rep === undefined && next === 58 /* : */) {
      // whole-column or whole-row range
      const e2 = readIdent(f, e + 1);
      const tok2 = f.slice(e + 1, e2);
      if (f.charCodeAt(e2) !== 40 && f.charCodeAt(e2) !== 33) {
        let a = shiftCol(tok, dc);
        let b = a !== undefined ? shiftCol(tok2, dc) : undefined;
        if (a === undefined || b === undefined) {
          a = shiftRow(tok, dr);
          b = a !== undefined ? shiftRow(tok2, dr) : undefined;
        }
        if (a !== undefined && b !== undefined) {
          rep = a === '#REF!' || b === '#REF!' ? '#REF!' : `${a}:${b}`;
          end = e2;
        }
      }
    }
    if (rep !== undefined) {
      out += f.slice(last, i) + rep;
      last = end;
    }
    i = end;
  }
  return last === 0 ? f : out + f.slice(last);
}
