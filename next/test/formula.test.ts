/// <reference types="node" />
// Formula parser + evaluation engine (turbosheet/formula).

import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellKind, encodeAddr, type Formula, Workbook } from '../src/core/index.js';
import {
  createEvaluator,
  evaluate,
  evaluateCell,
  evaluateFormula,
  FormulaParseError,
  type FormulaResult,
  parseFormula,
  supportedFunctions,
  toNumber,
} from '../src/formula/index.js';
import { addDefinedName } from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { writeXlsx } from '../src/xlsx/write/index.js';

const SPEC = new URL('../../spec/integration/data/', import.meta.url).pathname;

function book(): Workbook {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  return wb;
}

const result = (wb: Workbook, sheet: string, addr: string): Formula['result'] =>
  (wb.getWorksheet(sheet)!.getValue(addr) as Formula).result;

describe('parseFormula', () => {
  it('parses literals, references and calls', () => {
    expect(parseFormula('1.5e2')).toEqual({ type: 'number', value: 150 });
    expect(parseFormula('"a""b"')).toEqual({ type: 'string', value: 'a"b' });
    expect(parseFormula('true')).toEqual({ type: 'boolean', value: true });
    expect(parseFormula('#N/A')).toEqual({ type: 'error', value: '#N/A' });
    expect(parseFormula('$B$3')).toEqual({ type: 'cell', sheet: undefined, row: 2, col: 1 });
    expect(parseFormula('A1:B5')).toMatchObject({
      type: 'range',
      top: 0,
      left: 0,
      bottom: 4,
      right: 1,
    });
    expect(parseFormula('B5:A1')).toMatchObject({
      type: 'range',
      top: 0,
      left: 0,
      bottom: 4,
      right: 1,
    });
    expect(parseFormula('A:C')).toMatchObject({
      type: 'range',
      top: 0,
      left: 0,
      bottom: 1048575,
      right: 2,
    });
    expect(parseFormula('$2:$3')).toMatchObject({
      type: 'range',
      top: 1,
      left: 0,
      bottom: 2,
      right: 16383,
    });
    expect(parseFormula("'My Sheet'!A1")).toEqual({
      type: 'cell',
      sheet: 'My Sheet',
      row: 0,
      col: 0,
    });
    expect(parseFormula("'It''s'!A1").type).toBe('cell');
    expect(parseFormula('Sheet1!A1:B2')).toMatchObject({
      type: 'range',
      sheet: 'Sheet1',
      bottom: 1,
      right: 1,
    });
    expect(parseFormula('Sheet1!C:C')).toMatchObject({
      type: 'range',
      sheet: 'Sheet1',
      left: 2,
      right: 2,
    });
    expect(parseFormula('Sheet1!Rate')).toEqual({ type: 'name', name: 'Rate', sheet: 'Sheet1' });
    expect(parseFormula('TaxRate')).toEqual({ type: 'name', name: 'TaxRate' });
    expect(parseFormula('Sheet1!#REF!')).toEqual({ type: 'error', value: '#REF!' });
    expect(parseFormula('_xlfn.CONCAT("a",B1)')).toMatchObject({ type: 'call', name: 'CONCAT' });
    expect(parseFormula('sum()')).toEqual({ type: 'call', name: 'SUM', args: [] });
    expect(parseFormula('IF(A1,,1)')).toMatchObject({
      args: [{ type: 'cell' }, { type: 'missing' }, { type: 'number' }],
    });
    expect(parseFormula('{1,2;3,4}')).toEqual({
      type: 'array',
      rows: [
        [1, 2],
        [3, 4],
      ],
    });
    expect(parseFormula('{-1,"a",TRUE,#N/A}')).toEqual({
      type: 'array',
      rows: [[-1, 'a', true, { error: '#N/A' }]],
    });
    expect(parseFormula('=1')).toEqual({ type: 'number', value: 1 });
    expect(parseFormula(' 1 + 2 ')).toMatchObject({ type: 'binary', op: '+' });
    expect(parseFormula('XFD1048576')).toMatchObject({ type: 'cell', row: 1048575, col: 16383 });
    expect(parseFormula('XFE1')).toEqual({ type: 'name', name: 'XFE1' });
    expect(parseFormula('LOG10(100)')).toMatchObject({ type: 'call', name: 'LOG10' });
  });

  it('builds the AST with Excel precedence', () => {
    const op = (f: string) => parseFormula(f) as { type: string; op?: string; left?: unknown };
    // negation binds tighter than ^: -2^2 = (-2)^2
    expect(parseFormula('-2^2')).toMatchObject({
      type: 'binary',
      op: '^',
      left: { type: 'unary', op: '-', operand: { type: 'number', value: 2 } },
    });
    // left-associative: 2^3^2 = (2^3)^2
    expect(parseFormula('2^3^2')).toMatchObject({ op: '^', left: { op: '^' } });
    expect(parseFormula('1+2*3')).toMatchObject({ op: '+', right: { op: '*' } });
    expect(parseFormula('(1+2)*3')).toMatchObject({ op: '*', left: { op: '+' } });
    expect(parseFormula('10-4-3')).toMatchObject({ op: '-', left: { op: '-' } });
    expect(parseFormula('-50%')).toMatchObject({ type: 'unary', operand: { type: 'percent' } });
    expect(parseFormula('1+2&3')).toMatchObject({ op: '&', left: { op: '+' } });
    expect(parseFormula('1&2=12')).toMatchObject({ op: '=', left: { op: '&' } });
    expect(op('A1:B2*2')).toMatchObject({ op: '*', left: { type: 'range' } });
    expect(parseFormula('2*-3')).toMatchObject({ op: '*', right: { type: 'unary' } });
    expect(parseFormula('@A1:A3')).toMatchObject({ type: 'unary', op: '@' });
    expect(parseFormula('INDEX(A:A,1):A5')).toMatchObject({ type: 'binary', op: ':' });
  });

  it('follows Excel precedence', () => {
    const v = (f: string) => evaluateFormula(book(), f);
    expect(v('-2^2')).toBe(4); // negation binds tighter than ^
    expect(v('0-2^2')).toBe(-4);
    expect(v('2^3^2')).toBe(64); // left-associative
    expect(v('1+2*3')).toBe(7);
    expect(v('(1+2)*3')).toBe(9);
    expect(v('10-4-3')).toBe(3);
    expect(v('12/2/3')).toBe(2);
    expect(v('50%')).toBe(0.5);
    expect(v('-50%')).toBe(-0.5);
    expect(v('2^50%')).toBeCloseTo(Math.SQRT2);
    expect(v('1+2&3')).toBe('33');
    expect(v('1&2=12')).toBe(false); // "12" (text) vs 12: text > number
    expect(v('1&2="12"')).toBe(true);
    expect(v('1<2=TRUE')).toBe(true);
    expect(v('--"3"')).toBe(3);
    expect(v('+"a"')).toBe('a');
    expect(v('2*-3')).toBe(-6);
    expect(v('2^-1')).toBe(0.5);
  });

  it('reports syntax errors with a position', () => {
    for (const [f, msg] of [
      ['1+', 'Unexpected end of formula'],
      ['SUM(1,2', 'Missing ")"'],
      ['(1+2', 'Expected ")"'],
      ['"abc', 'Unterminated string'],
      ['1 2', 'Unexpected "2"'],
      ['{1,2;3}', 'Array rows must have the same length'],
      ['Table1[Col]', 'Structured references are not supported'],
      ['[1]Sheet1!A1', 'External or structured references are not supported'],
      ['Sheet1:Sheet3!A1', '3-D references are not supported'],
      ['(A1,B1)', 'Union references are not supported'],
      ['#FOO!', 'Unknown error literal'],
      ['', 'Empty formula'],
      ['1+*2', 'Unexpected "*"'],
      ["'Sheet1'A1", 'Expected "!" after sheet name'],
      ['1:2:', 'Range operator ":" needs references'],
    ]) {
      expect(() => parseFormula(f!), f).toThrow(FormulaParseError);
      expect(() => parseFormula(f!), f).toThrow(msg!);
    }
    try {
      parseFormula('SUM(1;2)');
    } catch (e) {
      expect(e).toBeInstanceOf(FormulaParseError);
      expect((e as FormulaParseError).position).toBe(5);
      expect((e as Error).message).toContain('position 5');
    }
  });
});

describe('evaluate', () => {
  it('evaluates in dependency order and stores results', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    // written in reverse order of dependency
    ws.setCell('A1', { formula: 'A2*2' });
    ws.setCell('A2', { formula: 'A3+1' });
    ws.setCell('A3', { formula: 'B1+B2' });
    ws.setCell('B1', 1);
    ws.setCell('B2', '2'); // text number
    const rep = evaluate(wb);
    expect(rep).toEqual({ evaluated: 3, circular: [], unsupported: [], errors: [] });
    expect(ws.getValue('A1')).toEqual({ formula: 'A2*2', result: 8 });
    expect(ws.getValue('A3')).toEqual({ formula: 'B1+B2', result: 3 });
  });

  it('writeResults: false leaves the workbook untouched and returns results', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: '1+1', result: 0 });
    const rep = evaluate(wb, { writeResults: false });
    expect(rep.results?.get('Sheet1!A1')).toBe(2);
    expect(ws.getValue('A1')).toEqual({ formula: '1+1', result: 0 });
  });

  it('blank references are 0 / ""; a formula result is never blank', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'Z99' });
    ws.setCell('A2', { formula: 'Z99&"x"' });
    ws.setCell('A3', { formula: 'ISBLANK(A1)' });
    ws.setCell('A4', { formula: 'ISBLANK(Z99)' });
    ws.setCell('A5', { formula: 'IF(TRUE,)' });
    evaluate(wb);
    expect(result(wb, 'Sheet1', 'A1')).toBe(0);
    expect(result(wb, 'Sheet1', 'A2')).toBe('x');
    expect(result(wb, 'Sheet1', 'A3')).toBe(false);
    expect(result(wb, 'Sheet1', 'A4')).toBe(true);
    expect(result(wb, 'Sheet1', 'A5')).toBe(0);
  });

  it('detects circular references', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'B1+1' });
    ws.setCell('B1', { formula: 'C1+1' });
    ws.setCell('C1', { formula: 'A1+1' });
    ws.setCell('D1', { formula: 'D1' }); // self reference
    ws.setCell('E1', { formula: 'A1*2' }); // depends on the cycle
    ws.setCell('F1', { formula: 'SUM(F2:F3)' });
    ws.setCell('F3', { formula: 'F1' });
    ws.setCell('G1', 5);
    ws.setCell('G2', { formula: 'G1*2' }); // unaffected
    const rep = evaluate(wb);
    expect(rep.circular.sort()).toEqual(
      ['Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1', 'Sheet1!D1', 'Sheet1!F1', 'Sheet1!F3'].sort(),
    );
    for (const a of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'F3']) {
      expect(result(wb, 'Sheet1', a), a).toEqual({ error: '#REF!' });
    }
    expect(result(wb, 'Sheet1', 'G2')).toBe(10);
  });

  it('iterative calculation (maxIterations) converges circular references', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    // x = 1 + x/2 converges to 2
    ws.setCell('A1', { formula: '1+A1/2' });
    // B1 = C1/2 + 1, C1 = B1 → 2
    ws.setCell('B1', { formula: 'C1/2+1' });
    ws.setCell('C1', { formula: 'B1' });
    const rep = evaluate(wb, { maxIterations: 100, maxChange: 1e-9 });
    expect(rep.circular).toHaveLength(3);
    expect(result(wb, 'Sheet1', 'A1')).toBeCloseTo(2, 6);
    expect(result(wb, 'Sheet1', 'B1')).toBeCloseTo(2, 6);
    // stops after maxIterations
    ws.setCell('A1', { formula: 'A1+1', result: 0 });
    evaluate(wb, { maxIterations: 5 });
    expect(result(wb, 'Sheet1', 'A1')).toBe(5);
  });

  it('cross-sheet references and quoted sheet names', () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('Data Sheet');
    const b = wb.addWorksheet('Calc');
    a.setCell('A1', 10);
    a.setCell('A2', 20);
    a.setCell('B1', { formula: 'Calc!A1*3' });
    b.setCell('A1', { formula: "SUM('Data Sheet'!A1:A2)" });
    b.setCell('A2', { formula: "'data sheet'!B1+1" }); // sheet names are case-insensitive
    b.setCell('A3', { formula: 'Missing!A1' });
    evaluate(wb);
    expect(result(wb, 'Calc', 'A1')).toBe(30);
    expect(result(wb, 'Data Sheet', 'B1')).toBe(90);
    expect(result(wb, 'Calc', 'A2')).toBe(91);
    expect(result(wb, 'Calc', 'A3')).toEqual({ error: '#REF!' });
  });

  it('defined names: ranges, constants, formulas, sheet scope', () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet('S1');
    const s2 = wb.addWorksheet('S2');
    s1.setCell('A1', 1);
    s1.setCell('A2', 2);
    s1.setCell('A3', { formula: 'A1*10' });
    addDefinedName(wb, 'Data', 'S1!$A$1:$A$3');
    addDefinedName(wb, 'Rate', '0.5');
    addDefinedName(wb, 'Total', 'SUM(Data)*Rate');
    addDefinedName(wb, 'Local', '100', { localSheet: s2 });
    addDefinedName(wb, 'Local', '1');
    wb.definedNames.push({ name: 'Broken', value: 'S1!#REF!' });
    s2.setCell('A1', { formula: 'SUM(Data)' });
    s2.setCell('A2', { formula: 'Total' });
    s2.setCell('A3', { formula: 'Local' });
    s1.setCell('B1', { formula: 'Local' });
    s1.setCell('B2', { formula: 'S2!Local' });
    s1.setCell('B3', { formula: 'Nope+1' });
    s1.setCell('B4', { formula: 'Broken' });
    s1.setCell('B5', { formula: 'INDEX(Data,2)' });
    const rep = evaluate(wb);
    expect(rep.errors).toEqual([]);
    expect(result(wb, 'S2', 'A1')).toBe(13);
    expect(result(wb, 'S2', 'A2')).toBe(6.5);
    expect(result(wb, 'S2', 'A3')).toBe(100);
    expect(result(wb, 'S1', 'B1')).toBe(1);
    expect(result(wb, 'S1', 'B2')).toBe(100);
    expect(result(wb, 'S1', 'B3')).toEqual({ error: '#NAME?' });
    expect(result(wb, 'S1', 'B4')).toEqual({ error: '#REF!' });
    expect(result(wb, 'S1', 'B5')).toBe(2);
  });

  it('unsupported functions → #NAME? and reported once', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'NOPE(1)+1' });
    ws.setCell('A2', { formula: 'NOPE(2)' });
    ws.setCell('A3', { formula: 'IFERROR(OFFSET(A1,1,1),"fallback")' });
    const seen: string[] = [];
    const rep = evaluate(wb, { onUnsupported: (n, a) => seen.push(`${n}@${a}`) });
    expect(rep.unsupported).toEqual(['NOPE', 'OFFSET']);
    expect(seen).toEqual(['NOPE@Sheet1!A1', 'OFFSET@Sheet1!A3']);
    expect(result(wb, 'Sheet1', 'A1')).toEqual({ error: '#NAME?' });
    expect(result(wb, 'Sheet1', 'A3')).toBe('fallback');
  });

  it('unparsable formulas → #NAME? and listed in errors; never throws', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'SUM(1,' });
    ws.setCell('A2', { formula: 'Table1[Col]' });
    ws.setCell('A3', { formula: 'A1' });
    const rep = evaluate(wb);
    expect(rep.errors.map((e) => e.address)).toEqual(['Sheet1!A1', 'Sheet1!A2']);
    expect(rep.errors[0]!.message).toContain('Unexpected end of formula');
    expect(result(wb, 'Sheet1', 'A1')).toEqual({ error: '#NAME?' });
    expect(result(wb, 'Sheet1', 'A3')).toEqual({ error: '#NAME?' });
  });

  it('errors propagate through operators and functions', () => {
    const wb = book();
    const v = (f: string) => evaluateFormula(wb, f);
    expect(v('1/0+NA()')).toEqual({ error: '#DIV/0!' });
    expect(v('NA()+1/0')).toEqual({ error: '#N/A' });
    expect(v('"a"+1')).toEqual({ error: '#VALUE!' });
    expect(v('SUM(1,1/0)')).toEqual({ error: '#DIV/0!' });
    expect(v('#REF!&"x"')).toEqual({ error: '#REF!' });
    expect(v('1/0=1/0')).toEqual({ error: '#DIV/0!' });
    expect(v('(-8)^(1/3)')).toEqual({ error: '#NUM!' });
    expect(v('0^0')).toEqual({ error: '#NUM!' });
    expect(v('0^-1')).toEqual({ error: '#DIV/0!' });
    expect(v('10^400')).toEqual({ error: '#NUM!' });
  });

  it('value semantics: coercion and comparison', () => {
    const wb = book();
    const v = (f: string) => evaluateFormula(wb, f);
    expect(v('"1,000"+1')).toBe(1001);
    expect(v('"50%"*2')).toBe(1);
    expect(v('"$5"+0')).toBe(5);
    expect(v('"2024-01-15"+0')).toBe(45306);
    expect(v('"1/15/2024"+0')).toBe(45306);
    expect(v('"12:00"+0')).toBe(0.5);
    expect(v('TRUE+TRUE')).toBe(2);
    expect(v('"TRUE"+0')).toEqual({ error: '#VALUE!' });
    expect(v('""+0')).toEqual({ error: '#VALUE!' });
    expect(v('"abc"="ABC"')).toBe(true);
    expect(v('"a"<"B"')).toBe(true);
    expect(v('1<"0"')).toBe(true); // number < text
    expect(v('"z"<FALSE')).toBe(true); // text < logical
    expect(v('FALSE<TRUE')).toBe(true);
    expect(v('Z1=0')).toBe(true); // blank = 0
    expect(v('Z1=""')).toBe(true); // blank = ""
    expect(v('Z1=FALSE')).toBe(true);
    expect(v('0.1+0.2=0.3')).toBe(true); // 15-digit comparison
    expect(v('0.1+0.2')).toBe(0.3); // results rounded to 15 significant digits
    expect(v('1/3')).toBe(0.333333333333333);
    expect(v('1/3&""')).toBe('0.333333333333333');
    expect(v('1E+21&""')).toBe('1E+21');
    expect(v('0.0000001&""')).toBe('1E-07');
    expect(v('123456789012345&""')).toBe('123456789012345');
    expect(v('TRUE&1')).toBe('TRUE1');
  });

  it('implicit intersection for ranges in single-value positions', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    for (let r = 1; r <= 3; r++) ws.setCell(r, 1, r * 10);
    ws.setCell('B2', { formula: 'A1:A3*2' }); // row 2 → A2
    ws.setCell('B5', { formula: 'A1:A3*2' }); // no intersection
    ws.setCell('C2', { formula: 'A1:A3' });
    ws.setCell('D2', { formula: 'ABS(A1:A3)' });
    ws.setCell('E2', { formula: 'SUM(A1:A3*2)' }); // legacy: intersection inside SUM
    ws.setCell('F2', { formula: 'SUMPRODUCT(A1:A3*2)' }); // array argument
    ws.setCell('G1', { formula: 'SUM(A1:A3*2)', array: 'G1' }); // array formula
    ws.setCell('H1', { formula: '@A1:A3' });
    evaluate(wb);
    expect(result(wb, 'Sheet1', 'B2')).toBe(40);
    expect(result(wb, 'Sheet1', 'B5')).toEqual({ error: '#VALUE!' });
    expect(result(wb, 'Sheet1', 'C2')).toBe(20);
    expect(result(wb, 'Sheet1', 'D2')).toBe(20);
    expect(result(wb, 'Sheet1', 'E2')).toBe(40);
    expect(result(wb, 'Sheet1', 'F2')).toBe(120);
    expect(result(wb, 'Sheet1', 'G1')).toBe(120);
    expect(result(wb, 'Sheet1', 'H1')).toBe(10);
  });

  it('array formulas spill into their area and feed dependents', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.addRow([1, 10]);
    ws.addRow([2, 20]);
    ws.addRow([3, 30]);
    // as the reader stores it: master holds the formula + area, other cells hold cached values
    ws.setCell('C1', { formula: 'A1:A3*B1:B3', array: 'C1:C3', result: 0 });
    ws.setCell('C2', 0);
    ws.setCell('C3', 0);
    ws.setCell('D1', { formula: 'SUM(C1:C3)' });
    ws.setCell('E1', { formula: 'C3+1' });
    ws.setCell('F1', { formula: 'SUM(IF(A1:A3>1,B1:B3))', array: 'F1' });
    ws.setCell('G1', { formula: '{1,2,3}*2', array: 'G1:I1' });
    evaluate(wb);
    expect(result(wb, 'Sheet1', 'C1')).toBe(10);
    expect(ws.getValue('C2')).toBe(40);
    expect(ws.getValue('C3')).toBe(90);
    expect(result(wb, 'Sheet1', 'D1')).toBe(140);
    expect(result(wb, 'Sheet1', 'E1')).toBe(91);
    expect(result(wb, 'Sheet1', 'F1')).toBe(50);
    expect([ws.getValue('H1'), ws.getValue('I1')]).toEqual([4, 6]);
  });

  it('dates respect wb.date1904', () => {
    for (const d1904 of [false, true]) {
      const wb = new Workbook({ date1904: d1904 });
      const ws = wb.addWorksheet('S');
      ws.setCell('A1', new Date(Date.UTC(2024, 0, 15)));
      ws.setCell('B1', { formula: 'YEAR(A1)*10000+MONTH(A1)*100+DAY(A1)' });
      ws.setCell('B2', { formula: 'DATE(2024,1,15)-A1' });
      ws.setCell('B3', { formula: 'TEXT(A1,"yyyy-mm-dd")' });
      ws.setCell('B4', { formula: 'WEEKDAY(A1)' });
      ws.setCell('B6', { formula: '"2024-01-15"-A1' });
      evaluate(wb);
      expect(result(wb, 'S', 'B1')).toBe(20240115);
      expect(result(wb, 'S', 'B2')).toBe(0);
      expect(result(wb, 'S', 'B3')).toBe('2024-01-15');
      expect(result(wb, 'S', 'B4')).toBe(2); // Monday
      expect(result(wb, 'S', 'B6')).toBe(0);
    }
    const w1904 = new Workbook({ date1904: true });
    w1904.addWorksheet('S');
    expect(evaluateFormula(w1904, 'DATE(1904,1,2)')).toBe(1);
  });

  it('evaluateCell evaluates one cell and its dependencies only', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'A2+1' });
    ws.setCell('A2', { formula: '41' });
    ws.setCell('B1', { formula: '1/0' });
    ws.setCell('C1', 'plain');
    expect(evaluateCell(wb, 'Sheet1', 'A1')).toBe(42);
    expect(ws.getValue('A1')).toEqual({ formula: 'A2+1' }); // not written by default
    expect(evaluateCell(wb, ws, 'C1')).toBe('plain');
    expect(evaluateCell(wb, ws, 'Z9')).toBe(null);
    expect(evaluateCell(wb, ws, 'A1', { writeResults: true })).toBe(42);
    expect(ws.getValue('A2')).toEqual({ formula: '41', result: 41 });
    expect(ws.getValue('B1')).toEqual({ formula: '1/0' });
    expect(() => evaluateCell(wb, 'Nope', 'A1')).toThrow(RangeError);
  });

  it('evaluateFormula: ad-hoc text, sheet, host address, arrays', () => {
    const wb = new Workbook();
    wb.addWorksheet('One').setCell('A1', 1);
    const two = wb.addWorksheet('Two');
    two.addRow([5, 6]);
    two.setCell('A2', { formula: 'A1*100' });
    expect(evaluateFormula(wb, '=A1+1')).toBe(2);
    expect(evaluateFormula(wb, 'A1+1', { sheet: 'Two' })).toBe(6);
    expect(evaluateFormula(wb, 'A2', { sheet: two })).toBe(500); // dependency evaluated
    expect(evaluateFormula(wb, 'A1:B1', { sheet: 'Two' })).toEqual([[5, 6]]);
    expect(evaluateFormula(wb, '{1,2;3,4}*10')).toEqual([
      [10, 20],
      [30, 40],
    ]);
    expect(evaluateFormula(wb, 'ROW()+COLUMN()', { address: 'C5' })).toBe(8);
    expect(evaluateFormula(wb, 'ROW()')).toEqual({ error: '#VALUE!' });
    expect(() => evaluateFormula(wb, 'SUM(')).toThrow(FormulaParseError);
  });

  it('TODAY/NOW use the injectable clock', () => {
    const wb = book();
    const at = new Date(2024, 0, 15, 18, 0, 0); // local wall-clock time
    expect(evaluateFormula(wb, 'TODAY()', { now: () => at })).toBe(45306);
    expect(evaluateFormula(wb, 'NOW()', { now: () => at })).toBe(45306.75);
    const ev = createEvaluator({ now: () => at });
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', { formula: 'NOW()-TODAY()' });
    ev.evaluate(wb);
    expect(result(wb, 'Sheet1', 'A1')).toBe(0.75);
  });

  it('createEvaluator: custom functions per instance, no global mutation', () => {
    const ev = createEvaluator({
      functions: {
        double: (args) => toNumber(args[0]!) * 2,
      },
    });
    ev.registerFunction('Hello', (args) => `hi ${String(args[0])}`);
    ev.registerFunction('BOOM', () => {
      throw { error: '#NUM!' };
    });
    ev.registerFunction('CRASH', () => {
      throw new Error('bug');
    });
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', 21);
    ws.setCell('B1', { formula: 'DOUBLE(A1)' });
    ws.setCell('B2', { formula: 'HELLO("x")' });
    ws.setCell('B3', { formula: 'BOOM()' });
    ws.setCell('B4', { formula: 'CRASH()' });
    ws.setCell('B5', { formula: 'SUM(1,2)' });
    const rep = ev.evaluate(wb);
    expect(result(wb, 'Sheet1', 'B1')).toBe(42);
    expect(result(wb, 'Sheet1', 'B2')).toBe('hi x');
    expect(result(wb, 'Sheet1', 'B3')).toEqual({ error: '#NUM!' });
    expect(result(wb, 'Sheet1', 'B4')).toEqual({ error: '#VALUE!' });
    expect(rep.errors).toEqual([{ address: 'Sheet1!B4', message: 'bug' }]);
    expect(result(wb, 'Sheet1', 'B5')).toBe(3);
    // the default evaluator is unchanged
    expect(evaluateFormula(wb, 'DOUBLE(1)')).toEqual({ error: '#NAME?' });
    expect(supportedFunctions()).not.toContain('DOUBLE');
  });

  it('evaluates only the requested sheets (plus dependencies)', () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    const b = wb.addWorksheet('B');
    a.setCell('A1', { formula: '1+1' });
    a.setCell('A2', { formula: '2+2' });
    b.setCell('A1', { formula: 'A!A1*10' });
    const rep = evaluate(wb, { sheets: ['B'] });
    expect(rep.evaluated).toBe(2);
    expect(result(wb, 'B', 'A1')).toBe(20);
    expect(result(wb, 'A', 'A2')).toBeUndefined();
  });

  it('round trip: evaluate → writeXlsx → readXlsx keeps the results', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.addRow(['Item', 'Qty', 'Price', 'Total']);
    ws.addRow(['a', 2, 1.5, { formula: 'B2*C2' }]);
    ws.addRow(['b', 3, 2.25, { formula: 'B3*C3' }]);
    ws.setCell('D4', { formula: 'SUM(D2:D3)' });
    ws.setCell('E1', { formula: 'IF(D4>5,"big","small")' });
    ws.setCell('E2', { formula: 'D4>5' });
    ws.setCell('E3', { formula: '1/0' });
    evaluate(wb);
    const back = readXlsx(writeXlsx(wb)).getWorksheet()!;
    expect(back.getValue('D4')).toEqual({ formula: 'SUM(D2:D3)', result: 9.75 });
    expect(back.getValue('E1')).toEqual({ formula: 'IF(D4>5,"big","small")', result: 'big' });
    expect(back.getValue('E2')).toEqual({ formula: 'D4>5', result: true });
    expect(back.getValue('E3')).toEqual({ formula: '1/0', result: { error: '#DIV/0!' } });
  });

  it('matches the cached values Excel stored in real files', () => {
    let total = 0;
    let match = 0;
    const mismatches: string[] = [];
    for (const f of readdirSync(SPEC).filter((n) => n.endsWith('.xlsx'))) {
      let wb: Workbook;
      try {
        wb = readXlsx(new Uint8Array(readFileSync(SPEC + f)));
      } catch {
        continue; // invalid fixtures
      }
      const cached = new Map<string, FormulaResult>();
      for (const ws of wb.worksheets) {
        ws.rowData.forEach((row) => {
          if (!row) return;
          for (let i = 0; i < row.n; i++) {
            if (row.kind[i] !== CellKind.Formula) continue;
            const r = (row.obj![i] as Formula).result;
            if (r === undefined) continue;
            const a = `${/^\w+$/.test(ws.name) ? ws.name : `'${ws.name}'`}!${encodeAddr(row.index, row.col[i]!)}`;
            cached.set(a, r);
          }
        });
      }
      if (cached.size === 0) continue;
      const rep = evaluate(wb, { writeResults: false });
      for (const [addr, want] of cached) {
        const got = rep.results!.get(addr);
        total++;
        const ok =
          typeof want === 'number' && typeof got === 'number'
            ? Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want))
            : JSON.stringify(got) === JSON.stringify(want);
        if (ok) match++;
        else
          mismatches.push(
            `${f} ${addr}: got ${JSON.stringify(got)}, Excel ${JSON.stringify(want)}`,
          );
      }
    }
    console.log(`fixture formulas: ${match}/${total} match Excel's cached values`, mismatches);
    expect(total).toBeGreaterThan(200);
    expect(match).toBe(total);
  });

  it('evaluates 100k simple formulas quickly', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    const N = 100_000;
    for (let r = 1; r <= N; r++) ws.addRow([r, r / 2, { formula: `A${r}*2+B${r}` }]);
    const t0 = performance.now();
    const rep = evaluate(wb);
    const ms = performance.now() - t0;
    console.log(`100k formulas: ${ms.toFixed(0)} ms`);
    expect(rep.evaluated).toBe(N);
    expect(result(wb, 'Sheet1', `C${N}`)).toBe(N * 2.5);
    expect(ms).toBeLessThan(2000); // generous for CI; ~0.2 s locally
  });

  it('a 100k-long dependency chain does not overflow the stack', () => {
    const wb = book();
    const ws = wb.getWorksheet()!;
    ws.setCell('A1', 1);
    for (let r = 2; r <= 100_000; r++) ws.setCell(r, 1, { formula: `A${r - 1}+1` });
    evaluate(wb);
    expect(result(wb, 'Sheet1', 'A100000')).toBe(100_000);
  });
});
