// Math, rounding and statistical functions (#107, #108). Every expected value below was checked
// against Excel for Mac (the same workbook evaluated by Excel, compared cell by cell).

import { describe, expect, it } from 'vitest';
import { evaluate } from '../src/formula/index.js';
import { Workbook } from '../src/index.js';

const CASES: [string, unknown][] = [
  ['LOG(100)', 2],
  ['LOG(8,2)', 3],
  ['LOG(-1)', { error: '#NUM!' }],
  ['LOG(10,1)', { error: '#DIV/0!' }],
  ['LN(EXP(2))', 2],
  ['LOG10(0.001)', -3],
  ['EXP(710)', { error: '#NUM!' }],
  ['CEILING(2.5,1)', 3],
  ['CEILING(-2.5,2)', -2],
  ['CEILING(-2.5,-2)', -4],
  ['CEILING(2.5,-2)', { error: '#NUM!' }],
  ['CEILING(0.3,0.1)', 0.3],
  ['CEILING(1.234,0.05)', 1.25],
  ['FLOOR(2.5,1)', 2],
  ['FLOOR(-2.5,2)', -4],
  ['FLOOR(-2.5,-2)', -2],
  ['FLOOR(2.5,-2)', { error: '#NUM!' }],
  ['FLOOR(5,0)', { error: '#DIV/0!' }],
  ['FLOOR(0.7,0.1)', 0.7],
  ['CEILING.MATH(-2.5)', -2],
  ['CEILING.MATH(-2.5,2,1)', -4],
  ['CEILING.MATH(24.3,5)', 25],
  ['CEILING.MATH(-8.1,-2)', -8],
  ['FLOOR.MATH(-2.5)', -3],
  ['FLOOR.MATH(-2.5,2,1)', -2],
  ['FLOOR.MATH(24.3,5)', 20],
  ['FLOOR.MATH(-8.1,2)', -10],
  ['MROUND(10,3)', 9],
  ['MROUND(-10,-3)', -9],
  ['MROUND(1.3,0.2)', 1.4],
  ['MROUND(5,-2)', { error: '#NUM!' }],
  ['MROUND(7.5,5)', 10],
  ['MROUND(2.5,1)', 3],
  ['TRUNC(8.9)', 8],
  ['TRUNC(-8.9)', -8],
  ['TRUNC(3.14159,2)', 3.14],
  ['TRUNC(1234.5,-2)', 1200],
  ['EVEN(1.5)', 2],
  ['EVEN(-1)', -2],
  ['EVEN(0)', 0],
  ['EVEN(2)', 2],
  ['ODD(1.5)', 3],
  ['ODD(-2)', -3],
  ['ODD(0)', 1],
  ['ODD(3)', 3],
  ['QUOTIENT(5,2)', 2],
  ['QUOTIENT(-10,3)', -3],
  ['QUOTIENT(1,0)', { error: '#DIV/0!' }],
  ['GCD(24,36,48)', 12],
  ['GCD(7,0)', 7],
  ['GCD(-1,2)', { error: '#NUM!' }],
  ['LCM(4,6)', 12],
  ['LCM(5,0)', 0],
  ['LCM(3,4,5)', 60],
  ['FACT(5)', 120],
  ['FACT(0)', 1],
  ['FACT(5.9)', 120],
  ['FACT(-1)', { error: '#NUM!' }],
  ['FACT(171)', { error: '#NUM!' }],
  ['COMBIN(8,2)', 28],
  ['COMBIN(10,0)', 1],
  ['COMBIN(4,5)', { error: '#NUM!' }],
  ['PERMUT(8,2)', 56],
  ['PERMUT(3,3)', 6],
  ['SUMSQ(3,4)', 25],
  ['SUMSQ(A1:A3)', 26],
  ['RADIANS(180)', 3.14159265358979],
  ['DEGREES(PI())', 180],
  ['SIN(PI()/6)', 0.5],
  ['COS(0)', 1],
  ['TAN(PI()/4)', 1],
  ['ASIN(1)', 1.5707963267949],
  ['ACOS(2)', { error: '#NUM!' }],
  ['ATAN(1)', 0.785398163397448],
  ['ATAN2(1,1)', 0.785398163397448],
  ['ATAN2(-1,0)', 3.14159265358979],
  ['ATAN2(0,0)', { error: '#DIV/0!' }],
  ['MEDIAN(A1:A10)', 3.5],
  ['MEDIAN(A1:A9)', 4],
  ['MODE(A1:A10)', 3],
  ['MODE.SNGL(7,8,9)', { error: '#N/A' }],
  ['LARGE(A1:A10,2)', 6],
  ['SMALL(A1:A10,3)', 2],
  ['LARGE(A1:A10,11)', { error: '#NUM!' }],
  ['SMALL(A1:A10,1.5)', 1],
  ['RANK(5,A1:A10)', 3],
  ['RANK(5,A1:A10,1)', 7],
  ['RANK.EQ(1,A1:A10)', 9],
  ['RANK.AVG(5,A1:A10)', 3.5],
  ['RANK.AVG(1,A1:A10,1)', 1.5],
  ['RANK(8,A1:A10)', { error: '#N/A' }],
  ['PERCENTILE(A1:A10,0.3)', 2.7],
  ['PERCENTILE.INC(A1:A10,1)', 9],
  ['PERCENTILE.EXC(A1:A10,0.3)', 2.3],
  ['PERCENTILE.EXC(A1:A10,0.05)', { error: '#NUM!' }],
  ['PERCENTILE(A1:A10,1.5)', { error: '#NUM!' }],
  ['QUARTILE(A1:A10,1)', 2.25],
  ['QUARTILE.INC(A1:A10,3)', 5],
  ['QUARTILE.EXC(A1:A10,1)', 1.75],
  ['QUARTILE.EXC(A1:A10,3)', 5.25],
  ['QUARTILE(A1:A10,5)', { error: '#NUM!' }],
  ['STDEV(A1:A10)', 2.46981780704569],
  ['STDEV.S(A1:A10)', 2.46981780704569],
  ['STDEV.P(A1:A10)', 2.343074902772],
  ['STDEVP(A1:A10)', 2.343074902772],
  ['STDEVA(B1:B10)', 2.38671920665766],
  ['VAR(A1:A10)', 6.1],
  ['VAR.S(A1:A10)', 6.1],
  ['VAR.P(A1:A10)', 5.49],
  ['VARP(A1:A10)', 5.49],
  ['VARA(B1:B10)', 5.69642857142857],
  ['STDEV(5)', { error: '#DIV/0!' }],
  ['AVERAGEA(B1:B10)', 1.375],
  ['MAXA(B1:B10)', 7],
  ['MINA(B1:B10)', 0],
  ['MAXIFS(A1:A10,D1:D10,"a")', 6],
  ['MINIFS(A1:A10,D1:D10,"b",C1:C10,">4")', 5],
  ['MAXIFS(A1:A10,D1:D10,"z")', 0],
  ['CORREL(A1:A10,C1:C10)', 0.295702445106008],
  ['PEARSON(A1:A10,C1:C10)', 0.295702445106008],
  ['CORREL(A1:A3,C1:C2)', { error: '#N/A' }],
  ['AGGREGATE(9,0,A1:A10)', 39],
  ['AGGREGATE(9,5,A1:A10)', 35],
  ['AGGREGATE(14,6,A1:A10,2)', 6],
  ['AGGREGATE(12,0,A1:A10)', 3.5],
  ['AGGREGATE(4,6,A1:A10,E1)', 9],
  ['AGGREGATE(4,0,A1:A10,E1)', { error: '#DIV/0!' }],
  ['AGGREGATE(1,1,A1:A10)', 3.88888888888889],
  ['AGGREGATE(16,0,A1:A10,0.4)', 3],
  ['AGGREGATE(19,0,A1:A10,1)', 1.75],
];

describe('math and statistical functions (Excel-checked)', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  [3, 1, 4, 1, 5, 9, 2, 6, 5, 3].forEach((v, i) => {
    ws.setCell(i + 1, 1, v);
  });
  ['x', true, false, 2, null, 'y', 7, null, 1, 0].forEach((v, i) => {
    if (v !== null) ws.setCell(i + 1, 2, v as string | number | boolean);
  });
  [2, 4, 5, 4, 5, 7, 8, 9, 10, 12].forEach((v, i) => {
    ws.setCell(i + 1, 3, v);
  });
  ['a', 'b', 'a', 'b', 'a', 'c', 'c', 'a', 'b', 'a'].forEach((v, i) => {
    ws.setCell(i + 1, 4, v);
  });
  ws.setRowProps(3, { hidden: true });
  ws.setCell('E1', { formula: '1/0' });
  CASES.forEach(([f], i) => {
    ws.setCell(i + 1, 8, { formula: f });
  });
  const report = evaluate(wb);

  it('evaluates without unsupported functions or errors', () => {
    expect(report.unsupported).toEqual([]);
    expect(report.errors).toEqual([]);
  });

  it.each(CASES.map((c, i) => [c[0], c[1], i] as const))('%s = %j', (_f, want, i) => {
    const got = (ws.getValue(i + 1, 8) as { result: unknown }).result;
    if (typeof want === 'number') expect(got).toBeCloseTo(want, 9);
    else expect(got).toEqual(want);
  });
});
