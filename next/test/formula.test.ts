// Formula parser (turbosheet/formula).

import { describe, expect, it } from 'vitest';
import { FormulaParseError, parseFormula } from '../src/formula/parser.js';

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

  it('follows Excel precedence', () => {
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
