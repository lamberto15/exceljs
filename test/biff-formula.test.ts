// The binary formula decompiler shared by readXlsb / readXls (src/biff): parsed-token streams
// (MS-XLS / MS-XLSB Rgce) back to formula text, in both token widths.

import { describe, expect, it } from 'vitest';
import { quoteSheet } from '../src/biff/common.js';
import { decodeFormula, type FormulaContext } from '../src/biff/formula.js';
import { ftabArgc, ftabName } from '../src/biff/ftab.js';

describe('function table', () => {
  it('maps indices to the documented names', () => {
    const want: Record<number, string> = {
      0: 'COUNT',
      4: 'SUM',
      27: 'ROUND',
      29: 'INDEX',
      32: 'LEN',
      36: 'AND',
      48: 'TEXT',
      56: 'PV',
      64: 'MATCH',
      65: 'DATE',
      66: 'TIME',
      74: 'NOW',
      76: 'ROWS',
      77: 'COLUMNS',
      78: 'OFFSET',
      82: 'SEARCH',
      83: 'TRANSPOSE',
      86: 'TYPE',
      97: 'ATAN2',
      100: 'CHOOSE',
      101: 'HLOOKUP',
      102: 'VLOOKUP',
      105: 'ISREF',
      109: 'LOG',
      111: 'CHAR',
      115: 'LEFT',
      116: 'RIGHT',
      117: 'EXACT',
      118: 'TRIM',
      119: 'REPLACE',
      120: 'SUBSTITUTE',
      121: 'CODE',
      124: 'FIND',
      125: 'CELL',
      126: 'ISERR',
      131: 'N',
      140: 'DATEVALUE',
      141: 'TIMEVALUE',
      142: 'SLN',
      143: 'SYD',
      144: 'DDB',
      148: 'INDIRECT',
      150: 'CALL',
      162: 'CLEAN',
      163: 'MDETERM',
      164: 'MINVERSE',
      165: 'MMULT',
      167: 'IPMT',
      168: 'PPMT',
      169: 'COUNTA',
      183: 'PRODUCT',
      184: 'FACT',
      189: 'DPRODUCT',
      190: 'ISNONTEXT',
      193: 'STDEVP',
      194: 'VARP',
      195: 'DSTDEVP',
      196: 'DVARP',
      197: 'TRUNC',
      198: 'ISLOGICAL',
      199: 'DCOUNTA',
      204: 'USDOLLAR',
      205: 'FINDB',
      211: 'LENB',
      212: 'ROUNDUP',
      213: 'ROUNDDOWN',
      216: 'RANK',
      219: 'ADDRESS',
      220: 'DAYS360',
      221: 'TODAY',
      222: 'VDB',
      227: 'MEDIAN',
      228: 'SUMPRODUCT',
      229: 'SINH',
      234: 'ATANH',
      235: 'DGET',
      244: 'INFO',
      247: 'DB',
      252: 'FREQUENCY',
      261: 'ERROR.TYPE',
      269: 'AVEDEV',
      270: 'BETADIST',
      288: 'CEILING',
      303: 'SUMXMY2',
      318: 'DEVSQ',
      325: 'LARGE',
      332: 'TINV',
      336: 'CONCATENATE',
      337: 'POWER',
      342: 'RADIANS',
      343: 'DEGREES',
      344: 'SUBTOTAL',
      345: 'SUMIF',
      346: 'COUNTIF',
      347: 'COUNTBLANK',
      350: 'ISPMT',
      351: 'DATEDIF',
      354: 'ROMAN',
      358: 'GETPIVOTDATA',
      359: 'HYPERLINK',
      360: 'PHONETIC',
      361: 'AVERAGEA',
      368: 'BAHTTEXT',
      379: 'RTD',
      380: 'CUBEVALUE',
      384: 'HEX2BIN',
      396: 'IMSUB',
      414: 'SERIESSUM',
      429: 'XIRR',
      449: 'EDATE',
      450: 'EOMONTH',
      464: 'RANDBETWEEN',
      476: 'FVSCHEDULE',
      480: 'IFERROR',
      481: 'COUNTIFS',
      482: 'SUMIFS',
      483: 'AVERAGEIF',
      484: 'AVERAGEIFS',
    };
    const bad: string[] = [];
    for (const [k, v] of Object.entries(want)) {
      if (ftabName(+k) !== v) bad.push(`${k}: ${ftabName(+k)} != ${v}`);
    }
    expect(bad).toEqual([]);
    expect(ftabName(53)).toBe(''); // GOTO: macro sheets only
    expect(ftabName(485)).toBe('');
    expect(ftabArgc(27)).toBe(2);
    expect(ftabArgc(4)).toBe(-1); // SUM is variadic
  });
});

type Part = number | [number, 2 | 4 | 8];

/** Little-endian token bytes: numbers are bytes, [n, 2] / [n, 4] are u16 / u32, [n, 8] f64. */
function bytes(...parts: Part[]): Uint8Array {
  const out: number[] = [];
  for (const p of parts) {
    if (typeof p === 'number') out.push(p);
    else if (p[1] === 8) out.push(...new Uint8Array(new Float64Array([p[0]]).buffer));
    else for (let i = 0; i < p[1]; i++) out.push((p[0] >>> (8 * i)) & 0xff);
  }
  return new Uint8Array(out);
}

const ctx = (b12: boolean): FormulaContext => ({
  b12,
  sheet: (i) => [quoteSheet('Data'), quoteSheet('My Sheet'), 'Data:Other'][i],
  name: (i) => ['Rate', '_xlfn.XLOOKUP', '_xlfn._xlws.SORT'][i - 1],
  nameX: (_x, i) => (i === 1 ? 'ADDIN' : undefined),
});
const REL = 0xc000; // column + row relative
const A1: Part[] = [0x24, [0, 2], [REL, 2]];

describe('decodeFormula', () => {
  const b8 = ctx(false);
  const b12 = ctx(true);
  // the formula is in C5 (row 4, column 2): the base of RefN / AreaN offsets
  const d8 = (...t: Part[]): string | undefined =>
    decodeFormula(b8, bytes(...t), new Uint8Array(0), 4, 2);
  const d12 = (...t: Part[]): string | undefined =>
    decodeFormula(b12, bytes(...t), new Uint8Array(0), 4, 2);

  it('decodes references in both widths', () => {
    expect(d8(0x24, [0, 2], [1 | REL, 2])).toBe('B1');
    expect(d12(0x24, [0, 4], [1 | REL, 2])).toBe('B1');
    expect(d8(0x24, [9, 2], [27, 2])).toBe('$AB$10');
    expect(d12(0x24, [9, 4], [27 | 0x4000, 2])).toBe('AB$10');
    expect(d8(0x25, [0, 2], [3, 2], [REL, 2], [2 | REL, 2])).toBe('A1:C4');
    expect(d12(0x25, [0, 4], [0xfffff, 4], [1 | 0x4000, 2], [1 | 0x4000, 2])).toBe('B:B');
    expect(d8(0x25, [0, 2], [0xffff, 2], [1, 2], [2, 2])).toBe('$B:$C');
    expect(d8(0x25, [4, 2], [4, 2], [0x8000, 2], [0xff | 0x8000, 2])).toBe('5:5');
    expect(d12(0x25, [0, 4], [1, 4], [0, 2], [0x3fff, 2])).toBe('$1:$2');
  });

  it('decodes relative-to-cell references (shared formulas) from the cell position', () => {
    expect(d8(0x2c, [0xffff, 2], [1 | REL, 2])).toBe('D4'); // one row up, one column right
    expect(d12(0x2c, [0xffffffff, 4], [0x3fff | REL, 2])).toBe('B4');
    expect(d8(0x2d, [0, 2], [2, 2], [0xff | REL, 2], [0x4000, 2])).toBe('B5:C$3');
  });

  it('decodes 3-D and error references', () => {
    expect(d8(0x3a, [0, 2], [0, 2], [REL, 2])).toBe('Data!A1');
    expect(d12(0x3b, [1, 2], [0, 4], [1, 4], [0, 2], [1, 2])).toBe("'My Sheet'!$A$1:$B$2");
    expect(d8(0x3a, [2, 2], [0, 2], [0, 2])).toBe('Data:Other!$A$1');
    expect(d8(0x3c, [0, 2], [0, 2], [0, 2])).toBe('Data!#REF!');
    expect(d12(0x2a, [0, 4], [0, 2])).toBe('#REF!');
    expect(d8(0x3a, [7, 2], [0, 2], [0, 2])).toBeUndefined(); // unknown sheet
  });

  it('decodes constants', () => {
    expect(d8(0x1e, [42, 2])).toBe('42');
    expect(d8(0x1f, [0.25, 8])).toBe('0.25');
    expect(d8(0x1f, [1e-7, 8])).toBe('1E-7');
    expect(d8(0x1d, 1)).toBe('TRUE');
    expect(d8(0x1c, 0x2a)).toBe('#N/A');
    expect(d8(0x17, 3, 0, 0x61, 0x22, 0x62)).toBe('"a""b"');
    expect(d8(0x17, 2, 1, 0xe5, 0x65, 0x2c, 0x67)).toBe('"日本"');
    expect(d12(0x17, [2, 2], [0xe9, 2], [0x41, 2])).toBe('"éA"');
  });

  it('decodes operators, parentheses and missing arguments', () => {
    const one: Part[] = [0x1e, [1, 2]];
    const two: Part[] = [0x1e, [2, 2]];
    expect(d8(...one, ...two, 0x03, 0x15, 0x13, 0x14)).toBe('-(1+2)%');
    expect(d8(...one, ...two, 0x0e)).toBe('1<>2');
    expect(d8(...one, ...two, 0x0c, 0x12)).toBe('+1>=2');
    expect(d8(...one, ...two, 0x08)).toBe('1&2');
    expect(d8(...A1, ...A1, 0x10)).toBe('A1,A1');
    expect(d8(...one, 0x16, 0x22, 2, [27, 2])).toBe('ROUND(1,)');
  });

  it('decodes functions: fixed, variable, tAttrSum, names, future functions', () => {
    expect(d8(...A1, 0x21, [32, 2])).toBe('LEN(A1)');
    expect(d8(...A1, ...A1, 0x22, 2, [4, 2])).toBe('SUM(A1,A1)');
    expect(d8(...A1, 0x19, 0x10, [0, 2])).toBe('SUM(A1)');
    expect(d8(0x21, [221, 2])).toBe('TODAY()');
    expect(d8(0x23, [1, 2], [0, 2])).toBe('Rate');
    expect(d12(0x23, [2, 4], 0x24, [0, 4], [REL, 2], 0x22, 2, [255, 2])).toBe('XLOOKUP(A1)');
    expect(d8(0x23, [3, 2], [0, 2], 0x1e, [1, 2], 0x22, 2, [255, 2])).toBe('SORT(1)');
    expect(d8(0x39, [0, 2], [1, 2], [0, 2], 0x22, 1, [255, 2])).toBe('ADDIN()');
    // macro-only (command-equivalent) and unknown functions are not decoded
    expect(d8(0x22, 0, [0x8001, 2])).toBeUndefined();
    expect(d8(0x21, [53, 2])).toBeUndefined();
    expect(d8(0x21, [4, 2])).toBeUndefined(); // SUM has no fixed argument count
  });

  it('skips control tokens and keeps spacing', () => {
    // tAttrIf / tAttrGoto
    expect(
      d8(
        0x19,
        0x01,
        [0, 2],
        ...A1,
        0x19,
        0x02,
        [0, 2],
        0x1e,
        [1, 2],
        0x19,
        0x08,
        [0, 2],
        0x22,
        2,
        [1, 2],
      ),
    ).toBe('IF(A1,1)');
    // tAttrChoose with its jump table
    expect(
      d8(0x1e, [1, 2], 0x19, 0x04, [1, 2], [0, 2], [0, 2], 0x1e, [5, 2], 0x22, 2, [100, 2]),
    ).toBe('CHOOSE(1,5)');
    // spaces before a token and before ')'
    expect(d8(0x19, 0x40, [0x0200, 2], ...A1, 0x19, 0x40, [0x0104, 2], 0x22, 1, [4, 2])).toBe(
      'SUM(  A1 )',
    );
    expect(d8(...A1, ...A1, 0x19, 0x40, [0x0100, 2], 0x03)).toBe('A1 +A1');
    expect(d8(0x19, 0x40, [0x0101, 2], ...A1)).toBe('\nA1');
  });

  it('decodes array constants (.xls and .xlsb layouts)', () => {
    const extra8 = bytes(1, [0, 2], 1, [1, 8], 2, [1, 2], 0, 0x61);
    expect(decodeFormula(b8, bytes(0x40, 0, 0, 0, 0, 0, 0, 0), extra8, 0, 0)).toBe('{1,"a"}');
    const extra8b = bytes(0, [1, 2], 4, 1, 0, 0, 0, 0, 0, 0, 0, 16, 0x2a, 0, 0, 0, 0, 0, 0, 0);
    expect(decodeFormula(b8, bytes(0x40, 0, 0, 0, 0, 0, 0, 0), extra8b, 0, 0)).toBe('{TRUE;#N/A}');
    const extra12 = bytes([2, 4], [2, 4], 0, [1, 8], 1, [1, 2], [0x61, 2], 2, 1, 4, 0x2a, 0, 0, 0);
    expect(decodeFormula(b12, bytes(0x40, ...new Array<number>(14).fill(0)), extra12, 0, 0)).toBe(
      '{1,"a";TRUE,#N/A}',
    );
  });

  it('skips mem-area wrappers', () => {
    const area: Part[] = [0x25, [0, 2], [1, 2], [REL, 2], [REL, 2]];
    expect(d8(0x29, [9, 2], ...area)).toBe('A1:A2'); // PtgMemFunc
    // PtgMemArea: its cached rectangles are in the extra data
    const extra = bytes([1, 2], [0, 2], [1, 2], [0, 2], [0, 2]);
    expect(decodeFormula(b8, bytes(0x26, 0, 0, 0, 0, [9, 2], ...area), extra, 0, 0)).toBe('A1:A2');
  });

  it('returns undefined for unsupported or malformed token streams', () => {
    expect(d8(0x01, [0, 2], [0, 2])).toBeUndefined(); // PtgExp: resolved by the readers
    expect(d12(0x18, 0x19, 0, 0, 0, 0)).toBeUndefined(); // structured reference
    expect(d8(0x24, [0, 2])).toBeUndefined(); // truncated
    expect(d8(0x1e, [1, 2], 0x1e, [1, 2])).toBeUndefined(); // two values left
    expect(d8(0x03)).toBeUndefined(); // operator without operands
    expect(d8(0xff)).toBeUndefined();
  });
});
