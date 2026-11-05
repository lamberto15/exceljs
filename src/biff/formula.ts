// Formula decompiler for the binary formats: a parsed-token stream (Rgce, reverse Polish
// notation, MS-XLS 2.5.198 / MS-XLSB 2.5.97) back to formula text as it appears in .xlsx
// (no leading '='). `.xls` (BIFF8) and `.xlsb` (BIFF12) share the token set; BIFF12 widens
// rows to 32 bits, strings to UTF-16 with a 16-bit length and some reserved fields.
//
// Covered: references (relative / absolute, relative-to-cell RefN / AreaN, whole rows /
// columns, 3-D references through the extern-sheet table, #REF! variants), numbers, strings,
// booleans, errors, array constants, the unary / binary / percent operators, parentheses,
// missing arguments, built-in functions (fixed and variable argument counts, SUM via
// tAttrSum, future functions called through `_xlfn.` names), defined names, spaces (tAttrSpace)
// and the mem-area wrappers. Anything else (structured references, data tables, external
// workbook names, macro-only functions, ...) makes `decodeFormula` return undefined: readers
// then keep the cell's cached value only.

import { encodeCol } from '../core/address.js';
import { ERRORS, latin1, utf16 } from './common.js';
import { ftabArgc, ftabName } from './ftab.js';

export interface FormulaContext {
  /** Token widths of .xlsb (BIFF12) instead of .xls (BIFF8). */
  readonly b12: boolean;
  /** Formula prefix of an extern-sheet (XTI) index, e.g. `Other` or `'My Sheet'`; undefined = deleted sheet. */
  sheet(ixti: number): string | undefined;
  /** Text of a defined name by its 1-based index. */
  name(index: number): string | undefined;
  /** Text of an external / add-in name (`PtgNameX`). */
  nameX?(ixti: number, index: number): string | undefined;
}

const BINARY = ['+', '-', '*', '/', '^', '&', '<', '<=', '=', '>=', '>', '<>', ' ', ',', ':'];

/** Thrown inside the decoder for anything it does not support; caught by decodeFormula. */
class Unsupported extends Error {}

const fail = (): never => {
  throw new Unsupported();
};

/** Number constant as Excel writes it in formula text. */
export function numText(n: number): string {
  const s = String(n);
  return s.indexOf('e') >= 0 ? s.toUpperCase() : s;
}

const quote = (s: string): string => `"${s.replace(/"/g, '""')}"`;

/**
 * Decompile `rgce` (+ its trailing data `rgcb`) of a formula in cell (row, col), both 0-based
 * (needed for RefN / AreaN tokens of shared formulas). Returns undefined when the formula uses
 * a token this decoder does not support, or is malformed.
 */
export function decodeFormula(
  ctx: FormulaContext,
  rgce: Uint8Array,
  rgcb: Uint8Array,
  row: number,
  col: number,
): string | undefined {
  try {
    return run(ctx, rgce, rgcb, row, col);
  } catch (e) {
    if (e instanceof Unsupported || e instanceof RangeError) return undefined;
    throw e;
  }
}

function run(
  ctx: FormulaContext,
  rgce: Uint8Array,
  rgcb: Uint8Array,
  baseRow: number,
  baseCol: number,
): string {
  const b12 = ctx.b12;
  const dv = new DataView(rgce.buffer, rgce.byteOffset, rgce.byteLength);
  const xv = new DataView(rgcb.buffer, rgcb.byteOffset, rgcb.byteLength);
  const u8 = (o: number): number => dv.getUint8(o);
  const u16 = (o: number): number => dv.getUint16(o, true);
  const maxRow = b12 ? 0xfffff : 0xffff;
  const maxCol = b12 ? 0x3fff : 0xff;
  const rw = b12 ? 4 : 2; // row field width
  const st: string[] = [];
  let x = 0; // rgcb position
  let p = 0;
  // pending tAttrSpace text: before the next token / before its '(' / before its ')'
  let space = '';
  let open = '';
  let close = '';
  const pop = (): string => (st.length ? st.pop()! : fail());
  const push = (s: string): void => {
    st.push(space + s);
    space = open = close = '';
  };

  const readRow = (o: number): number => (b12 ? dv.getUint32(o, true) : u16(o));
  /** One cell part: [row text, column text]; `rel` = relative-to-cell (RefN / AreaN). */
  const cellParts = (r: number, c: number, rel: boolean): [string, string, number, number] => {
    const colRel = (c & 0x4000) !== 0;
    const rowRel = (c & 0x8000) !== 0;
    let cc = c & 0x3fff;
    let rr = r;
    if (rel) {
      if (rowRel)
        rr =
          (((baseRow + (b12 ? r | 0 : (r << 16) >> 16)) % (maxRow + 1)) + maxRow + 1) %
          (maxRow + 1);
      if (colRel) {
        const off = b12 ? (cc << 18) >> 18 : (cc << 24) >> 24;
        cc = (((baseCol + off) % (maxCol + 1)) + maxCol + 1) % (maxCol + 1);
      }
    }
    if (cc > maxCol) cc = maxCol;
    return [(rowRel ? '' : '$') + (rr + 1), (colRel ? '' : '$') + encodeCol(cc), rr, cc];
  };
  const ref = (o: number, rel: boolean): string => {
    const [r, c] = cellParts(readRow(o), u16(o + rw), rel);
    return c + r;
  };
  const area = (o: number, rel: boolean): string => {
    const [r1, c1, rr1, cc1] = cellParts(readRow(o), u16(o + 2 * rw), rel);
    const [r2, c2, rr2, cc2] = cellParts(readRow(o + rw), u16(o + 2 * rw + 2), rel);
    if (rr1 === 0 && rr2 === maxRow) return `${c1}:${c2}`;
    if (cc1 === 0 && cc2 === maxCol) return `${r1}:${r2}`;
    return `${c1 + r1}:${c2 + r2}`;
  };
  const sheet = (o: number): string => ctx.sheet(u16(o)) ?? fail();

  while (p < rgce.length) {
    const t = rgce[p++]!;
    const base = t < 0x20 ? t : (t & 0x1f) | 0x20;
    switch (base) {
      case 0x03:
      case 0x04:
      case 0x05:
      case 0x06:
      case 0x07:
      case 0x08:
      case 0x09:
      case 0x0a:
      case 0x0b:
      case 0x0c:
      case 0x0d:
      case 0x0e:
      case 0x0f:
      case 0x10:
      case 0x11: {
        const b = pop();
        const a = pop();
        st.push(a + space + BINARY[base - 3]! + b);
        space = open = close = '';
        break;
      }
      case 0x12:
        push(`+${pop()}`);
        break;
      case 0x13:
        push(`-${pop()}`);
        break;
      case 0x14: {
        const a = pop();
        st.push(`${a + space}%`);
        space = open = close = '';
        break;
      }
      case 0x15:
        push(`${open}(${pop()}${close})`);
        break;
      case 0x16:
        push('');
        break;
      case 0x17: {
        // PtgStr
        let s: string;
        if (b12) {
          const n = u16(p);
          s = utf16(rgce, p + 2, n);
          p += 2 + 2 * n;
          if (p > rgce.length) fail();
        } else {
          const n = u8(p);
          const wide = u8(p + 1) & 1;
          s = wide ? utf16(rgce, p + 2, n) : latin1(rgce, p + 2, n);
          p += 2 + (wide ? 2 * n : n);
        }
        push(quote(s));
        break;
      }
      case 0x19: {
        // PtgAttr
        const f = u8(p);
        const w = u16(p + 1);
        p += 3;
        if (f & 0x04) p += 2 * (w + 1); // tAttrChoose jump table
        if (f & 0x10) push(`SUM${open}(${pop()}${close})`);
        else if (f & 0x40) {
          // tAttrSpace: type in the low byte (spaces / line breaks before the next token, its
          // '(' or its ')'), count in the high byte
          const kind = w & 0xff;
          const ws = (kind & 1 ? '\n' : ' ').repeat(w >> 8);
          if (kind < 2) space += ws;
          else if (kind < 4) open += ws;
          else if (kind < 6) close += ws;
        }
        break;
      }
      case 0x1c:
        push(ERRORS[u8(p++)] ?? fail());
        break;
      case 0x1d:
        push(u8(p++) ? 'TRUE' : 'FALSE');
        break;
      case 0x1e:
        push(String(u16(p)));
        p += 2;
        break;
      case 0x1f:
        push(numText(dv.getFloat64(p, true)));
        p += 8;
        break;
      case 0x20: {
        // PtgArray: the values follow the token stream (PtgExtraArray)
        p += b12 ? 14 : 7;
        let rows: number;
        let cols: number;
        if (b12) {
          rows = xv.getUint32(x, true);
          cols = xv.getUint32(x + 4, true);
          x += 8;
        } else {
          cols = xv.getUint8(x) + 1;
          rows = xv.getUint16(x + 1, true) + 1;
          x += 3;
        }
        if (rows * cols > rgcb.length) fail();
        const out: string[] = [];
        for (let r = 0; r < rows; r++) {
          const line: string[] = [];
          for (let c = 0; c < cols; c++) {
            const k = xv.getUint8(x++);
            if (b12) {
              if (k === 0) {
                line.push(numText(xv.getFloat64(x, true)));
                x += 8;
              } else if (k === 1) {
                const n = xv.getUint16(x, true);
                line.push(quote(utf16(rgcb, x + 2, n)));
                x += 2 + 2 * n;
              } else if (k === 2) line.push(xv.getUint8(x++) ? 'TRUE' : 'FALSE');
              else if (k === 4) {
                line.push(ERRORS[xv.getUint8(x)] ?? fail());
                x += 4;
              } else fail();
            } else if (k === 1) {
              line.push(numText(xv.getFloat64(x, true)));
              x += 8;
            } else if (k === 2) {
              const n = xv.getUint16(x, true);
              const wide = xv.getUint8(x + 2) & 1;
              line.push(quote(wide ? utf16(rgcb, x + 3, n) : latin1(rgcb, x + 3, n)));
              x += 3 + (wide ? 2 * n : n);
            } else if (k === 4) {
              line.push(xv.getUint8(x) ? 'TRUE' : 'FALSE');
              x += 8;
            } else if (k === 16) {
              line.push(ERRORS[xv.getUint8(x)] ?? fail());
              x += 8;
            } else if (k === 0) {
              line.push('');
              x += 8;
            } else fail();
          }
          out.push(line.join(','));
        }
        push(`{${out.join(';')}}`);
        break;
      }
      case 0x21:
      case 0x22: {
        // PtgFunc (fixed argument count) / PtgFuncVar
        let argc: number;
        let idx: number;
        if (base === 0x21) {
          idx = u16(p);
          p += 2;
          argc = ftabArgc(idx);
        } else {
          argc = u8(p) & 0x7f;
          const tab = u16(p + 1);
          p += 3;
          if (tab & 0x8000) fail(); // command-equivalent (macro) function
          idx = tab & 0x7fff;
        }
        if (argc < 0 || argc > st.length) fail();
        const args = st.splice(st.length - argc, argc);
        let fn: string;
        if (idx === 255) {
          // add-in / future function, called through a name: the first argument (the
          // `_xlfn.` prefix is dropped as readXlsx does)
          fn = (args.shift() ?? fail()).replace(/^_xlfn\.(_xlws\.)?/i, '');
        } else fn = ftabName(idx) || fail();
        push(`${fn}${open}(${args.join(',')}${close})`);
        break;
      }
      case 0x23:
        push(ctx.name(b12 ? dv.getUint32(p, true) : u16(p)) ?? fail());
        p += 4;
        break;
      case 0x24:
      case 0x2c:
        push(ref(p, base === 0x2c));
        p += rw + 2;
        break;
      case 0x25:
      case 0x2d:
        push(area(p, base === 0x2d));
        p += 2 * rw + 4;
        break;
      case 0x26:
        // PtgMemArea: its cached areas follow the tokens (PtgExtraMem); the value is the
        // sub-expression that follows
        p += 6;
        if (b12) x += 4 + 16 * xv.getUint32(x, true);
        else x += 2 + 8 * xv.getUint16(x, true);
        break;
      case 0x27:
      case 0x28:
        p += 6;
        break;
      case 0x29:
        p += 2;
        break;
      case 0x2a:
        push('#REF!');
        p += rw + 2;
        break;
      case 0x2b:
        push('#REF!');
        p += 2 * rw + 4;
        break;
      case 0x39:
        push(ctx.nameX?.(u16(p), b12 ? dv.getUint32(p + 2, true) : u16(p + 2)) ?? fail());
        p += 6;
        break;
      case 0x3a:
        push(`${sheet(p)}!${ref(p + 2, false)}`);
        p += 2 + rw + 2;
        break;
      case 0x3b:
        push(`${sheet(p)}!${area(p + 2, false)}`);
        p += 2 + 2 * rw + 4;
        break;
      case 0x3c:
        push(`${sheet(p)}!#REF!`);
        p += 2 + rw + 2;
        break;
      case 0x3d:
        push(`${sheet(p)}!#REF!`);
        p += 2 + 2 * rw + 4;
        break;
      default:
        // PtgExp / PtgTbl (handled by the readers), structured references, ...
        fail();
    }
    if (p > rgce.length) fail();
  }
  if (st.length !== 1) fail();
  return st[0]!;
}
