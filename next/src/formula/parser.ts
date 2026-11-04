// Excel formula parser: text (without the leading '=') → AST.
//
// Precedence, tightest first (as in Excel): ':' (range), unary '-' / '+' / '@', '%' (postfix),
// '^', '*' '/', '+' '-', '&', comparisons ('=' '<>' '<' '>' '<=' '>='). Binary operators are
// left-associative, so 2^3^2 = 64, and negation binds tighter than '^', so -2^2 = 4.
// Whitespace is ignored (the intersection operator is not supported); so are the union operator,
// structured references (Table1[Col]), external workbooks ([1]Sheet1!A1) and 3-D references
// (Sheet1:Sheet3!A1): they raise a FormulaParseError.

import type { CellError, ErrorCode } from '../core/types.js';

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

/** A literal inside an array constant ({1,2;"a",TRUE}). */
export type ArrayLiteral = number | string | boolean | CellError;

/** Binary operators. */
export type BinaryOp =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '>'
  | '<='
  | '>='
  | ':';

/** Parsed formula. Rows and columns are 0-based. */
export type FormulaNode =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean }
  | { readonly type: 'error'; readonly value: ErrorCode }
  | {
      readonly type: 'cell';
      readonly sheet?: string | undefined;
      readonly row: number;
      readonly col: number;
    }
  | {
      /** Whole columns (A:C) span every row, whole rows (1:3) every column. */
      readonly type: 'range';
      readonly sheet?: string | undefined;
      readonly top: number;
      readonly left: number;
      readonly bottom: number;
      readonly right: number;
    }
  | { readonly type: 'name'; readonly name: string; readonly sheet?: string | undefined }
  | { readonly type: 'unary'; readonly op: '-' | '+' | '@'; readonly operand: FormulaNode }
  | { readonly type: 'percent'; readonly operand: FormulaNode }
  | {
      readonly type: 'binary';
      readonly op: BinaryOp;
      readonly left: FormulaNode;
      readonly right: FormulaNode;
    }
  /** `name` is upper-case, without the `_xlfn.` / `_xlws.` prefixes files use. */
  | { readonly type: 'call'; readonly name: string; readonly args: readonly FormulaNode[] }
  | { readonly type: 'array'; readonly rows: readonly (readonly ArrayLiteral[])[] }
  /** An omitted argument: IF(A1,,1). */
  | { readonly type: 'missing' };

/** Thrown by parseFormula; `position` is the 0-based offset in the formula text. */
export class FormulaParseError extends SyntaxError {
  readonly position: number;
  readonly formula: string;
  constructor(message: string, formula: string, position: number) {
    super(`${message} at position ${position} in formula "${formula}"`);
    this.name = 'FormulaParseError';
    this.position = position;
    this.formula = formula;
  }
}

const ERRORS: readonly ErrorCode[] = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
];

/** @internal */
export const CELL_RE: RegExp = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COL_RANGE_RE = /\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![\w.$\\(!:\u0080-￿])/y;
const ROW_RANGE_RE = /\$?(\d{1,7}):\$?(\d{1,7})(?![\w.$\\(!:\u0080-￿])/y;
const NUM_RE = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;

/** @internal "A" → 0 (no validation). */
export function colIndex(s: string): number {
  let v = 0;
  for (let i = 0; i < s.length; i++) v = v * 26 + ((s.charCodeAt(i) | 0x20) - 96);
  return v - 1;
}

function identChar(c: number): boolean {
  return (
    (c >= 65 && c <= 90) ||
    (c >= 97 && c <= 122) ||
    (c >= 48 && c <= 57) ||
    c === 95 ||
    c === 46 ||
    c === 36 ||
    c === 92 ||
    c === 63 ||
    c > 127
  );
}

const PREFIX = /^(?:_xlfn\.|_xlws\.|_xludf\.)+/i;

/** Parse formula text (a leading '=' is allowed). Throws FormulaParseError. */
export function parseFormula(text: string): FormulaNode {
  return new Parser(text).run();
}

class Parser {
  readonly #s: string;
  #p = 0;
  constructor(s: string) {
    this.#s = s;
  }

  run(): FormulaNode {
    const s = this.#s;
    if (s.charCodeAt(0) === 61) this.#p = 1;
    this.#ws();
    if (this.#p >= s.length) this.#fail('Empty formula');
    const n = this.#expr(0);
    this.#ws();
    if (this.#p < s.length) this.#fail(`Unexpected "${s[this.#p]}"`);
    return n;
  }

  #fail(msg: string, at = this.#p): never {
    throw new FormulaParseError(msg, this.#s, at);
  }

  #ws(): void {
    const s = this.#s;
    while (this.#p < s.length) {
      const c = s.charCodeAt(this.#p);
      if (c === 32 || c === 10 || c === 13 || c === 9 || c === 160) this.#p++;
      else break;
    }
  }

  /** Binary operator at the cursor with its binding power, or undefined. */
  #binOp(): [BinaryOp, number] | undefined {
    const s = this.#s;
    const c = s[this.#p];
    switch (c) {
      case '+':
      case '-':
        return [c, 3];
      case '*':
      case '/':
        return [c, 4];
      case '^':
        return [c, 5];
      case '&':
        return [c, 2];
      case '=':
        return [c, 1];
      case ':':
        return [c, 8];
      case '<': {
        const d = s[this.#p + 1];
        return d === '=' || d === '>' ? [`<${d}` as BinaryOp, 1] : ['<', 1];
      }
      case '>':
        return s[this.#p + 1] === '=' ? ['>=', 1] : ['>', 1];
      default:
        return undefined;
    }
  }

  #expr(minBp: number): FormulaNode {
    let left = this.#prefix();
    for (;;) {
      this.#ws();
      if (this.#s[this.#p] === '%') {
        this.#p++;
        left = { type: 'percent', operand: left };
        continue;
      }
      const op = this.#binOp();
      if (!op || op[1] <= minBp) return left;
      const at = this.#p;
      this.#p += op[0].length;
      this.#ws();
      const right = this.#expr(op[1]);
      left =
        op[0] === ':' ? this.#range(left, right, at) : { type: 'binary', op: op[0], left, right };
    }
  }

  /** Fold A1:B2 (both cells) into a range node; anything else stays a ':' operation. */
  #range(a: FormulaNode, b: FormulaNode, at: number): FormulaNode {
    if (a.type === 'cell' && b.type === 'cell' && (b.sheet === undefined || b.sheet === a.sheet)) {
      return {
        type: 'range',
        sheet: a.sheet,
        top: Math.min(a.row, b.row),
        left: Math.min(a.col, b.col),
        bottom: Math.max(a.row, b.row),
        right: Math.max(a.col, b.col),
      };
    }
    const refLike = (n: FormulaNode): boolean =>
      n.type === 'cell' ||
      n.type === 'range' ||
      n.type === 'name' ||
      n.type === 'call' ||
      (n.type === 'binary' && n.op === ':');
    if (!refLike(a) || !refLike(b)) this.#fail('Range operator ":" needs references', at);
    return { type: 'binary', op: ':', left: a, right: b };
  }

  #prefix(): FormulaNode {
    this.#ws();
    const s = this.#s;
    const start = this.#p;
    if (start >= s.length) this.#fail('Unexpected end of formula');
    const c = s.charCodeAt(start);
    // unary
    if (c === 45 || c === 43 || c === 64) {
      this.#p++;
      const operand = this.#expr(6);
      return { type: 'unary', op: s[start] as '-' | '+' | '@', operand };
    }
    if (c === 40) {
      this.#p++;
      const e = this.#expr(0);
      this.#ws();
      if (s[this.#p] === ',') this.#fail('Union references are not supported');
      this.#expect(')');
      return e;
    }
    if (c === 34) return { type: 'string', value: this.#str() };
    if (c === 35) return { type: 'error', value: this.#err() };
    if (c === 123) return this.#array();
    if (c === 39) {
      const sheet = this.#quotedSheet();
      return this.#qualified(sheet);
    }
    if (c === 91) this.#fail('External or structured references are not supported');
    if (identChar(c)) {
      const r = this.#lines(undefined);
      if (r) return r;
    }
    if ((c >= 48 && c <= 57) || c === 46) {
      NUM_RE.lastIndex = start;
      const m = NUM_RE.exec(s);
      if (!m) this.#fail('Invalid number');
      this.#p = NUM_RE.lastIndex;
      return { type: 'number', value: Number(m[0]) };
    }
    if (!identChar(c)) this.#fail(`Unexpected "${s[start]}"`);
    let e = start;
    while (e < s.length && identChar(s.charCodeAt(e))) e++;
    const id = s.slice(start, e);
    this.#p = e;
    const next = s[e];
    if (next === '(') {
      this.#p++;
      return { type: 'call', name: id.replace(PREFIX, '').toUpperCase(), args: this.#args() };
    }
    if (next === '!') return this.#qualified(id);
    if (next === ':' && /^:(?:[\w.]+|'[^']*')!/.test(s.slice(e)))
      this.#fail('3-D references are not supported');
    if (next === '[') this.#fail('Structured references are not supported');
    const cell = this.#cell(undefined, id);
    if (cell) return cell;
    const up = id.toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') return { type: 'boolean', value: up === 'TRUE' };
    return { type: 'name', name: id };
  }

  #cell(sheet: string | undefined, id: string): FormulaNode | undefined {
    const m = CELL_RE.exec(id);
    if (!m) return undefined;
    const col = colIndex(m[1]!);
    const row = +m[2]! - 1;
    if (col > MAX_COL || row < 0 || row > MAX_ROW) return undefined;
    return { type: 'cell', sheet, row, col };
  }

  /** Whole-row (1:3) or whole-column (A:C) range at the cursor. */
  #lines(sheet: string | undefined): FormulaNode | undefined {
    const at = this.#p;
    for (const re of [ROW_RANGE_RE, COL_RANGE_RE]) {
      re.lastIndex = at;
      const m = re.exec(this.#s);
      if (!m) continue;
      this.#p = re.lastIndex;
      const row = re === ROW_RANGE_RE;
      const a = row ? +m[1]! - 1 : colIndex(m[1]!);
      const b = row ? +m[2]! - 1 : colIndex(m[2]!);
      const max = row ? MAX_ROW : MAX_COL;
      if (a < 0 || b < 0 || a > max || b > max)
        this.#fail(row ? 'Row out of range' : 'Column out of range', at);
      const lo = Math.min(a, b);
      const hi = Math.max(a, b);
      return row
        ? { type: 'range', sheet, top: lo, left: 0, bottom: hi, right: MAX_COL }
        : { type: 'range', sheet, top: 0, left: lo, bottom: MAX_ROW, right: hi };
    }
    return undefined;
  }

  /** After a sheet name: expect '!' then a cell, range, name or #REF!. */
  #qualified(sheet: string): FormulaNode {
    const s = this.#s;
    if (s[this.#p] === ':') this.#fail('3-D references are not supported');
    if (s[this.#p] !== '!') this.#fail('Expected "!" after sheet name');
    this.#p++;
    const start = this.#p;
    if (s[start] === '#') return { type: 'error', value: this.#err() };
    const r = this.#lines(sheet);
    if (r) return r;
    let e = start;
    while (e < s.length && identChar(s.charCodeAt(e))) e++;
    if (e === start) this.#fail('Expected a reference after "!"');
    const id = s.slice(start, e);
    this.#p = e;
    return this.#cell(sheet, id) ?? { type: 'name', name: id, sheet };
  }

  #quotedSheet(): string {
    const s = this.#s;
    let out = '';
    let i = this.#p + 1;
    for (;;) {
      const k = s.indexOf("'", i);
      if (k < 0) this.#fail('Unterminated sheet name', this.#p);
      out += s.slice(i, k);
      if (s[k + 1] === "'") {
        out += "'";
        i = k + 2;
      } else {
        this.#p = k + 1;
        break;
      }
    }
    if (out.startsWith('[')) this.#fail('External references are not supported');
    return out;
  }

  #str(): string {
    const s = this.#s;
    let out = '';
    let i = this.#p + 1;
    for (;;) {
      const k = s.indexOf('"', i);
      if (k < 0) this.#fail('Unterminated string', this.#p);
      out += s.slice(i, k);
      if (s[k + 1] === '"') {
        out += '"';
        i = k + 2;
      } else {
        this.#p = k + 1;
        return out;
      }
    }
  }

  #err(): ErrorCode {
    const s = this.#s;
    for (const e of ERRORS) {
      if (s.startsWith(e, this.#p) || s.slice(this.#p, this.#p + e.length).toUpperCase() === e) {
        this.#p += e.length;
        return e;
      }
    }
    return this.#fail('Unknown error literal');
  }

  #expect(ch: string): void {
    this.#ws();
    if (this.#s[this.#p] !== ch) {
      this.#fail(this.#p >= this.#s.length ? `Expected "${ch}" before end` : `Expected "${ch}"`);
    }
    this.#p++;
  }

  #args(): FormulaNode[] {
    const out: FormulaNode[] = [];
    this.#ws();
    if (this.#s[this.#p] === ')') {
      this.#p++;
      return out;
    }
    for (;;) {
      this.#ws();
      const c = this.#s[this.#p];
      if (c === ',' || c === ')') out.push({ type: 'missing' });
      else out.push(this.#expr(0));
      this.#ws();
      const d = this.#s[this.#p];
      if (d === ',') this.#p++;
      else if (d === ')') {
        this.#p++;
        return out;
      } else this.#fail(d === undefined ? 'Missing ")"' : `Expected "," or ")" but found "${d}"`);
    }
  }

  #array(): FormulaNode {
    const s = this.#s;
    this.#p++;
    const rows: ArrayLiteral[][] = [];
    let row: ArrayLiteral[] = [];
    for (;;) {
      this.#ws();
      row.push(this.#literal());
      this.#ws();
      const c = s[this.#p++];
      if (c === ',') continue;
      rows.push(row);
      if (c === ';') {
        if (row.length !== rows[0]!.length)
          this.#fail('Array rows must have the same length', this.#p - 1);
        row = [];
        continue;
      }
      if (c === '}') {
        if (row.length !== rows[0]!.length)
          this.#fail('Array rows must have the same length', this.#p - 1);
        return { type: 'array', rows };
      }
      this.#fail(
        c === undefined ? 'Unterminated array constant' : `Unexpected "${c}" in array`,
        this.#p - 1,
      );
    }
  }

  #literal(): ArrayLiteral {
    const s = this.#s;
    const c = s[this.#p];
    if (c === '"') return this.#str();
    if (c === '#') return { error: this.#err() };
    let sign = 1;
    if (c === '-' || c === '+') {
      if (c === '-') sign = -1;
      this.#p++;
    }
    NUM_RE.lastIndex = this.#p;
    const m = NUM_RE.exec(s);
    if (m) {
      this.#p = NUM_RE.lastIndex;
      return sign * Number(m[0]);
    }
    const w = /TRUE|FALSE/iy;
    w.lastIndex = this.#p;
    const b = w.exec(s);
    if (b && sign === 1) {
      this.#p = w.lastIndex;
      return b[0].toUpperCase() === 'TRUE';
    }
    return this.#fail('Array constants may only hold numbers, strings, booleans and errors');
  }
}
