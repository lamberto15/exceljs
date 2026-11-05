// More text functions (CHAR, CODE, UNICHAR, UNICODE, CLEAN, FIXED, DOLLAR, NUMBERVALUE, T,
// TEXTBEFORE, TEXTAFTER, TEXTSPLIT) and information functions (N, TYPE, ISEVEN, ISODD, ISREF,
// ISFORMULA, ERROR.TYPE, ADDRESS, HYPERLINK).

import type { Builtin, EvalContext } from '../evaluator.js';
import { sheetPrefix } from '../evaluator.js';
import { type FormulaFunction, S, sc } from '../helpers.js';
import type { FormulaNode } from '../parser.js';
import {
  FormulaArray,
  type FormulaScalar,
  type FormulaValue,
  isArr,
  isError,
  isRef,
  NA,
  toNumber,
  toText,
  VALUE,
} from '../values.js';

// Windows-1252 code points 128-159 (what CHAR / CODE use for those codes).
const CP1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f' + '\u0090‘’“”•–—˜™š›œ\u009džŸ';

function char(code: number): string {
  const k = Math.trunc(code);
  if (k < 1 || k > 255) throw VALUE;
  return k >= 128 && k <= 159 ? CP1252[k - 128]! : String.fromCharCode(k);
}

function code(s: string): number {
  if (s === '') throw VALUE;
  const c = s.charCodeAt(0);
  if (c < 128 || (c >= 160 && c <= 255)) return c;
  const i = CP1252.indexOf(s[0]!);
  return i >= 0 ? 128 + i : 63; // '?' for characters outside the code page, as Excel
}

/** Round half away from zero at `d` decimals (d may be negative). */
function roundAt(x: number, d: number): number {
  const f = 10 ** d;
  const r = Math.round(Math.abs(x) * f + 1e-9) / f;
  return x < 0 ? -r : r;
}

function grouped(x: number, d: number, commas: boolean): string {
  const r = roundAt(x, d);
  const s = Math.abs(r).toFixed(Math.max(0, d));
  const [int, frac] = s.split('.') as [string, string?];
  const i = commas ? int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : int;
  return (r < 0 ? '-' : '') + i + (frac ? `.${frac}` : '');
}

/** Delimiters argument: one text or an array / range of texts. */
function delims(
  v: FormulaValue | undefined,
  ctx: EvalContext | Parameters<FormulaFunction>[1],
): string[] {
  if (v === undefined || v === null) return [];
  if (isRef(v) || isArr(v)) {
    const out: string[] = [];
    for (let i = 0; i < v.rows; i++) for (let j = 0; j < v.cols; j++) out.push(toText(v.at(i, j)));
    return out;
  }
  return [toText(sc(v, ctx))];
}

/** Positions of the delimiter matches in `s` (left to right, non-overlapping). */
function matches(s: string, ds: string[], ci: boolean): { at: number; len: number }[] {
  const hay = ci ? s.toLowerCase() : s;
  const needles = ci ? ds.map((d) => d.toLowerCase()) : ds;
  const out: { at: number; len: number }[] = [];
  let i = 0;
  while (i <= hay.length) {
    let best: { at: number; len: number } | undefined;
    for (const d of needles) {
      if (d === '') continue;
      const k = hay.indexOf(d, i);
      if (k >= 0 && (!best || k < best.at || (k === best.at && d.length > best.len))) {
        best = { at: k, len: d.length };
      }
    }
    if (!best) break;
    out.push(best);
    i = best.at + best.len;
  }
  return out;
}

const opt = (args: FormulaValue[], i: number, ctx: Parameters<FormulaFunction>[1], dflt: number) =>
  args[i] == null ? dflt : toNumber(sc(args[i]!, ctx));

function textSide(after: boolean): FormulaFunction {
  return (args, ctx) => {
    if (args.length < 2 || args.length > 6) throw VALUE;
    const s = toText(sc(args[0]!, ctx));
    const ds = delims(args[1], ctx);
    const inst = Math.trunc(opt(args, 2, ctx, 1));
    const ci = opt(args, 3, ctx, 0) !== 0;
    const matchEnd = opt(args, 4, ctx, 0) !== 0;
    if (inst === 0 || Math.abs(inst) > s.length + 1) throw VALUE;
    const notFound = (): FormulaValue => (args.length >= 6 ? args[5]! : NA);
    if (ds.length === 0 || ds.every((d) => d === '')) {
      // an empty delimiter matches at the start (or the end when counting from the end)
      return after ? (inst > 0 ? s : '') : inst > 0 ? '' : s;
    }
    const m = matches(s, ds, ci);
    if (matchEnd) {
      if (inst > 0) m.push({ at: s.length, len: 0 });
      else m.unshift({ at: 0, len: 0 });
    }
    const k = inst > 0 ? inst - 1 : m.length + inst;
    const hit = m[k];
    if (!hit || k < 0) return notFound();
    return after ? s.slice(hit.at + hit.len) : s.slice(0, hit.at);
  };
}

const TEXTSPLIT: FormulaFunction = (args, ctx) => {
  if (args.length < 2 || args.length > 6) throw VALUE;
  const s = toText(sc(args[0]!, ctx));
  const cd = delims(args[1], ctx);
  const rd = delims(args[2], ctx);
  const ignoreEmpty = opt(args, 3, ctx, 0) !== 0;
  const ci = opt(args, 4, ctx, 0) !== 0;
  const pad: FormulaScalar = args.length >= 6 ? sc(args[5]!, ctx) : NA;
  if (cd.every((d) => d === '') && rd.every((d) => d === '')) throw VALUE;
  const split = (t: string, ds: string[]): string[] => {
    if (ds.every((d) => d === '')) return [t];
    const out: string[] = [];
    let from = 0;
    for (const m of matches(t, ds, ci)) {
      out.push(t.slice(from, m.at));
      from = m.at + m.len;
    }
    out.push(t.slice(from));
    return ignoreEmpty ? out.filter((x) => x !== '') : out;
  };
  const rows = split(s, rd).map((r) => split(r, cd));
  const kept = rows.filter((r) => r.length > 0);
  if (kept.length === 0) throw VALUE;
  const cols = Math.max(...kept.map((r) => r.length));
  const data: FormulaScalar[] = [];
  for (const r of kept) for (let j = 0; j < cols; j++) data.push(j < r.length ? r[j]! : pad);
  return new FormulaArray(kept.length, cols, data);
};

const ERROR_TYPES: Record<string, number> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CALC!': 14,
};

function colLetters(c: number): string {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

const ISREF = (args: readonly FormulaNode[], ctx: EvalContext): FormulaValue => {
  if (args.length !== 1) throw VALUE;
  return isRef(ctx.safe(args[0]!));
};

export const text2: Record<string, Builtin> = {
  CHAR: S((x) => char(toNumber(x))),
  CODE: S((x) => code(toText(x))),
  UNICHAR: S((x) => {
    const k = Math.trunc(toNumber(x));
    if (k < 1 || k > 0x10ffff || (k >= 0xd800 && k <= 0xdfff)) throw VALUE;
    return String.fromCodePoint(k);
  }),
  UNICODE: S((x) => {
    const s = toText(x);
    if (s === '') throw VALUE;
    return s.codePointAt(0)!;
  }),
  CLEAN: S((x) => {
    let out = '';
    for (const ch of toText(x)) if (ch.charCodeAt(0) > 31) out += ch;
    return out;
  }),
  FIXED: S((x, d, nc) => {
    const k = d === undefined || d === null ? 2 : Math.trunc(toNumber(d));
    if (k > 127) throw VALUE;
    return grouped(toNumber(x), k, !(nc !== undefined && nc !== null && toNumber(nc) !== 0));
  }),
  DOLLAR: S((x, d) => {
    const k = d === undefined || d === null ? 2 : Math.trunc(toNumber(d));
    if (k > 127) throw VALUE;
    const v = toNumber(x);
    const body = grouped(Math.abs(v), k, true);
    return roundAt(v, k) < 0 ? `($${body})` : `$${body}`;
  }),
  NUMBERVALUE: S((x, dec, grp) => {
    const ds = dec === undefined || dec === null ? '.' : (toText(dec)[0] ?? '');
    const gs = grp === undefined || grp === null ? ',' : (toText(grp)[0] ?? '');
    if (ds === '' || ds === gs) throw VALUE;
    let s = toText(x).replace(/\s+/g, '');
    if (s === '') return 0;
    let pct = 0;
    while (s.endsWith('%')) {
      pct++;
      s = s.slice(0, -1);
    }
    const di = s.indexOf(ds);
    if (gs !== '' && di >= 0 && s.indexOf(gs, di) >= 0) throw VALUE;
    if (gs !== '') s = s.split(gs).join('');
    s = s.replace(ds, '.');
    if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) throw VALUE;
    return Number(s) / 100 ** pct;
  }),
  T: (args, ctx) => {
    if (args.length !== 1) throw VALUE;
    const v = sc(args[0]!, ctx);
    if (isError(v)) throw v;
    return typeof v === 'string' ? v : '';
  },
  TEXTBEFORE: textSide(false),
  TEXTAFTER: textSide(true),
  TEXTSPLIT: { array: TEXTSPLIT },
  // information
  N: (args, ctx) => {
    if (args.length !== 1) throw VALUE;
    const v = isRef(args[0]) ? args[0].at(0, 0) : sc(args[0]!, ctx);
    if (isError(v)) throw v;
    return typeof v === 'number' ? v : v === true ? 1 : 0;
  },
  TYPE: (args, ctx) => {
    if (args.length !== 1) throw VALUE;
    const a = args[0]!;
    if (isArr(a)) return 64; // a multi-cell reference intersects, as in Excel
    const v = sc(a, ctx);
    if (isError(v)) return 16;
    if (typeof v === 'string') return 2;
    if (typeof v === 'boolean') return 4;
    return 1;
  },
  ISEVEN: S((x) => {
    if (typeof x === 'boolean') throw VALUE;
    return Math.trunc(toNumber(x)) % 2 === 0;
  }),
  ISODD: S((x) => {
    if (typeof x === 'boolean') throw VALUE;
    return Math.abs(Math.trunc(toNumber(x)) % 2) === 1;
  }),
  ISREF: { lazy: ISREF },
  ISFORMULA: (args) => {
    if (args.length !== 1) throw VALUE;
    const r = args[0];
    if (!isRef(r)) throw VALUE;
    const row = r.sheet.rowData[r.top];
    const i = row ? row.find(r.left) : -1;
    return i >= 0 && row!.kind[i] === 6;
  },
  'ERROR.TYPE': (args, ctx) => {
    if (args.length !== 1) throw VALUE;
    const v = sc(args[0]!, ctx);
    if (!isError(v)) throw NA;
    return ERROR_TYPES[v.error] ?? NA;
  },
  ADDRESS: (args, ctx) => {
    if (args.length < 2 || args.length > 5) throw VALUE;
    const r = Math.trunc(toNumber(sc(args[0]!, ctx)));
    const c = Math.trunc(toNumber(sc(args[1]!, ctx)));
    const abs = args[2] == null ? 1 : Math.trunc(toNumber(sc(args[2], ctx)));
    const a1 = args[3] == null ? true : toNumber(sc(args[3], ctx)) !== 0;
    if (abs < 1 || abs > 4) throw VALUE;
    if (r < 1 || c < 1 || r > 1048576 || c > 16384) throw VALUE;
    const absRow = abs === 1 || abs === 2;
    const absCol = abs === 1 || abs === 3;
    const ref = a1
      ? `${absCol ? '$' : ''}${colLetters(c)}${absRow ? '$' : ''}${r}`
      : `R${absRow ? r : `[${r}]`}C${absCol ? c : `[${c}]`}`;
    if (args[4] == null) return ref;
    const sheet = toText(sc(args[4], ctx));
    return sheet === '' ? `!${ref}` : sheetPrefix(sheet) + ref;
  },
  HYPERLINK: (args, ctx) => {
    if (args.length < 1 || args.length > 2) throw VALUE;
    const v = sc(args[args.length === 2 && args[1] != null ? 1 : 0]!, ctx);
    if (isError(v)) throw v;
    return v ?? 0;
  },
};
