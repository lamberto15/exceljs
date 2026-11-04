// COUNTIF-style criteria ("<5", ">=1/1/2024", "a*", "<>", "") and wildcard patterns (* ? ~).

import { compare, type FormulaScalar, isError, parseNumber, round15 } from './values.js';

export interface Criterion {
  test(v: FormulaScalar): boolean;
  /** Whether blank cells match (lets callers skip empty areas when false). */
  blank: boolean;
}

const esc = (c: string): string => c.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');

/** Wildcard pattern → case-insensitive RegExp body, or undefined when it has no wildcards. */
export function wildcardSource(p: string): string | undefined {
  if (!/[*?~]/.test(p)) return undefined;
  let s = '';
  for (let i = 0; i < p.length; i++) {
    const ch = p[i]!;
    if (ch === '~' && i + 1 < p.length && '*?~'.includes(p[i + 1]!)) s += esc(p[++i]!);
    else if (ch === '*') s += '[\\s\\S]*';
    else if (ch === '?') s += '[\\s\\S]';
    else s += esc(ch);
  }
  return s;
}

/** Wildcard matcher for exact-match lookups; undefined when the pattern has no wildcards. */
export function wildcard(p: string): RegExp | undefined {
  const s = wildcardSource(p);
  return s === undefined ? undefined : new RegExp(`^${s}$`, 'i');
}

function cmpOp(op: string, c: number): boolean {
  switch (op) {
    case '<':
      return c < 0;
    case '>':
      return c > 0;
    case '<=':
      return c <= 0;
    case '>=':
      return c >= 0;
    case '<>':
      return c !== 0;
    default:
      return c === 0;
  }
}

const numEq = (a: number, b: number): boolean => a === b || round15(a) === round15(b);

/** Build a criterion from a COUNTIF / SUMIF criteria argument. */
export function criterion(c: FormulaScalar): Criterion {
  if (c === null) c = '';
  if (typeof c === 'number') {
    const n = c;
    return {
      test: (v) =>
        typeof v === 'number' ? numEq(v, n) : typeof v === 'string' && parseNumber(v) === n,
      blank: false,
    };
  }
  if (typeof c === 'boolean' || isError(c)) {
    const want = c;
    return {
      test: (v) => v === want || (isError(v) && isError(want) && v.error === want.error),
      blank: false,
    };
  }
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? '';
  const rest = m[2]!;
  const ne = op === '<>';
  if (rest === '') {
    if (op === '') return { test: (v) => v === null || v === '', blank: true };
    if (op === '=') return { test: (v) => v === null, blank: true };
    if (ne) return { test: (v) => v !== null, blank: false };
    return { test: () => false, blank: false };
  }
  const n = parseNumber(rest);
  if (n !== undefined) {
    const eq = op === '' || op === '=';
    return {
      test: (v) => {
        if (typeof v === 'number') return eq ? numEq(v, n) : cmpOp(op, numEq(v, n) ? 0 : v - n);
        if (eq || ne) {
          const hit = typeof v === 'string' && parseNumber(v) === n;
          return ne ? !hit : hit;
        }
        return false;
      },
      blank: ne,
    };
  }
  const up = rest.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') {
    const b = up === 'TRUE';
    return { test: (v) => (ne ? v !== b : op === '' || op === '=' ? v === b : false), blank: ne };
  }
  if (/^#(?:N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(up)) {
    return { test: (v) => (isError(v) && v.error === up) !== ne, blank: ne };
  }
  if (op === '' || op === '=' || ne) {
    const re = wildcard(rest);
    const low = rest.toLowerCase();
    const hit = (v: FormulaScalar): boolean =>
      typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === low);
    return { test: ne ? (v) => !hit(v) : hit, blank: ne };
  }
  return { test: (v) => typeof v === 'string' && cmpOp(op, compare(v, rest)), blank: false };
}
