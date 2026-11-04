// Text functions.

import { wildcardSource } from '../criteria.js';
import type { Builtin } from '../evaluator.js';
import { formatText } from '../format.js';
import { each, type FormulaFunction, S, sc } from '../helpers.js';
import {
  check,
  type FormulaScalar,
  isArr,
  isRef,
  parseNumber,
  toBoolean,
  toNumber,
  toText,
  VALUE,
} from '../values.js';

const MAX = 32767;

function cap(s: string): string {
  if (s.length > MAX) throw VALUE;
  return s;
}

const count = (v: FormulaScalar, d: number): number => {
  const n = v === undefined || v === null ? d : Math.trunc(toNumber(v));
  if (n < 0) throw VALUE;
  return n;
};

function find(caseSensitive: boolean): FormulaFunction {
  return S((needle, hay, start) => {
    const n = toText(needle);
    const h = toText(hay);
    const s = start === undefined || start === null ? 1 : Math.trunc(toNumber(start));
    if (s < 1 || s > h.length + (n === '' ? 1 : 0)) throw VALUE;
    let i: number;
    if (caseSensitive) i = h.indexOf(n, s - 1);
    else {
      const src = wildcardSource(n);
      if (src === undefined) i = h.toLowerCase().indexOf(n.toLowerCase(), s - 1);
      else {
        const re = new RegExp(src, 'gi');
        re.lastIndex = s - 1;
        const m = re.exec(h);
        i = m ? m.index : -1;
      }
    }
    if (i < 0) throw VALUE;
    return i + 1;
  });
}

const CONCAT: FormulaFunction = (args) => {
  let s = '';
  for (const a of args) {
    each(a, (x) => {
      s = cap(s + toText(x));
    });
  }
  return s;
};

const TEXTJOIN: FormulaFunction = (args, ctx) => {
  if (args.length < 3) throw VALUE;
  const d = toText(sc(args[0]!, ctx));
  const skip = toBoolean(sc(args[1]!, ctx));
  const parts: string[] = [];
  let len = 0;
  const add = (x: FormulaScalar): void => {
    const t = toText(x);
    if (skip && t === '') return;
    parts.push(t);
    len += t.length + d.length;
    if (len > MAX + d.length) throw VALUE;
  };
  for (let k = 2; k < args.length; k++) {
    const a = args[k]!;
    if (isRef(a) && !skip) {
      for (let i = 0; i < a.rows; i++) for (let j = 0; j < a.cols; j++) add(a.at(i, j));
    } else if (isRef(a) || isArr(a)) each(a, add);
    else add(a);
  }
  return cap(parts.join(d));
};

export const text: Record<string, Builtin> = {
  CONCAT,
  CONCATENATE: S((...xs) => cap(xs.map(toText).join(''))),
  TEXTJOIN,
  LEFT: S((t, n) => toText(t).slice(0, count(n, 1))),
  RIGHT: S((t, n) => {
    const s = toText(t);
    const k = count(n, 1);
    return k === 0 ? '' : s.slice(-k);
  }),
  MID: S((t, st, n) => {
    const s = toText(t);
    const a = Math.trunc(toNumber(st));
    if (a < 1) throw VALUE;
    return s.substr(a - 1, count(n ?? null, 0));
  }),
  LEN: S((t) => toText(t).length),
  LOWER: S((t) => toText(t).toLowerCase()),
  UPPER: S((t) => toText(t).toUpperCase()),
  PROPER: S((t) =>
    toText(t)
      .toLowerCase()
      .replace(/(^|[^\p{L}])(\p{L})/gu, (_, a: string, b: string) => a + b.toUpperCase()),
  ),
  TRIM: S((t) =>
    toText(t)
      .replace(/^ +| +$/g, '')
      .replace(/ {2,}/g, ' '),
  ),
  SUBSTITUTE: S((t, o, n, inst) => {
    const s = toText(t);
    const old = toText(o);
    const rep = toText(n);
    if (old === '') return s;
    if (inst === undefined || inst === null) return cap(s.split(old).join(rep));
    const k = Math.trunc(toNumber(inst));
    if (k < 1) throw VALUE;
    let i = -1;
    for (let c = 0; c < k; c++) {
      i = s.indexOf(old, i + 1);
      if (i < 0) return s;
    }
    return cap(s.slice(0, i) + rep + s.slice(i + old.length));
  }),
  REPLACE: S((t, st, n, r) => {
    const s = toText(t);
    const a = Math.trunc(toNumber(st));
    if (a < 1) throw VALUE;
    return cap(s.slice(0, a - 1) + toText(r) + s.slice(a - 1 + count(n, 0)));
  }),
  FIND: find(true),
  SEARCH: find(false),
  TEXT: S((v, f) => {
    const fmt = toText(f);
    const x = check(v);
    if (typeof x === 'boolean') return x ? 'TRUE' : 'FALSE';
    if (typeof x === 'string') {
      const n = parseNumber(x);
      return formatText(n ?? x, fmt);
    }
    return formatText(x ?? 0, fmt);
  }),
  VALUE: S((v) => {
    const x = check(v);
    if (typeof x === 'number') return x;
    if (x === null) return 0;
    if (typeof x === 'boolean') throw VALUE;
    const n = parseNumber(x);
    if (n === undefined) throw VALUE;
    return n;
  }),
  EXACT: S((a, b) => toText(a) === toText(b)),
  REPT: S((t, n) => {
    const s = toText(t);
    const k = count(n, 0);
    if (s.length * k > MAX) throw VALUE;
    return s.repeat(k);
  }),
};
