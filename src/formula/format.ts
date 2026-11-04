// Number-format rendering for TEXT(): digit placeholders (0 # ?), decimal point, thousands
// separator and scaling commas, percent, scientific (0.00E+00), literals ("text", \x, $ - + ( )
// : / space), sections (positive;negative;zero;text), General, and dates / times (yyyy yy mmmm
// mmm mm m dddd ddd dd d hh h mm ss AM/PM A/P [h] [m] [s] .0). Colors and locale tags ([Red],
// [$-409]) are ignored; conditions ([>100]) and fractions (# ?/?) are not supported.

import { serialToDate } from '../core/dates.js';
import { roundTo } from './functions/math.js';
import { date1904, numToStr } from './values.js';

const MONTHS =
  'January February March April May June July August September October November December'.split(
    ' ',
  );
const DAYS = 'Sunday Monday Tuesday Wednesday Thursday Friday Saturday'.split(' ');

/** Split a format into sections at ';' outside quotes / brackets / escapes. */
function sections(f: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < f.length; i++) {
    const c = f[i]!;
    if (c === '"') {
      const k = f.indexOf('"', i + 1);
      const e = k < 0 ? f.length : k;
      cur += f.slice(i, e + 1);
      i = e;
    } else if (c === '\\' || c === '_' || c === '*') {
      cur += f.slice(i, i + 2);
      i++;
    } else if (c === '[') {
      const k = f.indexOf(']', i);
      const e = k < 0 ? f.length : k;
      cur += f.slice(i, e + 1);
      i = e;
    } else if (c === ';') {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}

/** Literal-free skeleton used to classify a section. */
function skeleton(s: string): string {
  return s.replace(/"[^"]*"|\\.|_.|\*.|\[(?![hms]+\])[^\]]*\]/gi, '');
}

/** Format a value as TEXT(value, fmt) does. */
export function formatText(v: number | string, fmt: string): string {
  const secs = sections(fmt);
  if (typeof v === 'string') {
    const t = secs[3] ?? (secs.length === 1 && secs[0]!.includes('@') ? secs[0] : undefined);
    if (t === undefined) return v;
    return t.replace(/"([^"]*)"?|\\(.)|_.|\*.|\[[^\]]*\]|@/g, (m, q?: string, e?: string) =>
      m === '@' ? v : (q ?? e ?? (m[0] === '_' ? ' ' : '')),
    );
  }
  let sec = secs[0]!;
  let x = v;
  let minus = v < 0;
  if (v < 0 && secs.length > 1 && secs[1] !== '') {
    sec = secs[1]!;
    x = -v;
    minus = false;
  } else if (v === 0 && secs.length > 2) sec = secs[2]!;
  else if (v < 0) x = -v;
  const sk = skeleton(sec);
  const sign = minus ? '-' : '';
  if (/general/i.test(sk)) return sign + numToStr(x);
  return sign + (/[ymdhs]/i.test(sk.replace(/E[+-]/gi, '')) ? date(x, sec) : number(x, sec));
}

type Tok =
  | { k: 'lit'; s: string }
  | { k: 'd'; c: string }
  | { k: '.' | ',' | '%' | 'E'; s: string };

function number(v: number, sec: string): string {
  const toks: Tok[] = [];
  for (let i = 0; i < sec.length; i++) {
    const c = sec[i]!;
    if (c === '"') {
      const k = sec.indexOf('"', i + 1);
      const e = k < 0 ? sec.length : k;
      toks.push({ k: 'lit', s: sec.slice(i + 1, e) });
      i = e;
    } else if (c === '\\') toks.push({ k: 'lit', s: sec[++i] ?? '' });
    else if (c === '_') {
      toks.push({ k: 'lit', s: ' ' });
      i++;
    } else if (c === '*') i++;
    else if (c === '[') i = Math.max(i, sec.indexOf(']', i));
    else if (c === '0' || c === '#' || c === '?') toks.push({ k: 'd', c });
    else if (c === '.' && !toks.some((t) => t.k === '.')) toks.push({ k: '.', s: c });
    else if (c === ',' || c === '%') toks.push({ k: c, s: c });
    else if ((c === 'E' || c === 'e') && (sec[i + 1] === '+' || sec[i + 1] === '-')) {
      toks.push({ k: 'E', s: sec[++i]! });
    } else toks.push({ k: 'lit', s: c });
  }
  const dot = toks.findIndex((t) => t.k === '.');
  const ex = toks.findIndex((t) => t.k === 'E');
  const intEnd = dot >= 0 ? dot : ex >= 0 ? ex : toks.length;
  const intPh: number[] = [];
  const fracPh: number[] = [];
  const expPh: number[] = [];
  for (let i = 0; i < toks.length; i++) {
    if (toks[i]!.k !== 'd') continue;
    if (i < intEnd) intPh.push(i);
    else if (ex < 0 || i < ex) fracPh.push(i);
    else expPh.push(i);
  }
  // thousands separator: a comma between integer placeholders; commas after the last digit
  // placeholder scale by 1000
  let group = false;
  let scale = 0;
  const lastInt = intPh[intPh.length - 1] ?? -1;
  const lastDigit = Math.max(lastInt, fracPh[fracPh.length - 1] ?? -1);
  for (let i = 0; i < toks.length; i++) {
    if (toks[i]!.k !== ',') continue;
    if (i < lastInt && i > (intPh[0] ?? Infinity)) group = true;
    else if (i > lastDigit && lastDigit >= 0 && (ex < 0 || i < ex)) scale++;
    else toks[i] = { k: 'lit', s: ',' };
  }
  const pct = toks.filter((t) => t.k === '%').length;
  v = (v * 100 ** pct) / 1000 ** scale;
  let e = 0;
  if (ex >= 0) {
    const ints = Math.max(1, intPh.length);
    if (v !== 0) {
      e = Math.floor(Math.log10(v)) - (ints - 1);
      v /= 10 ** e;
      if (roundTo(v, fracPh.length, 0) >= 10 ** ints) {
        v /= 10;
        e++;
      }
    }
  }
  v = roundTo(v, fracPh.length, 0);
  if (v >= 1e21) return numToStr(v);
  const [ip, fp = ''] = v.toFixed(fracPh.length).split('.') as [string, string?];
  const zeroFrom = intPh.findIndex((i) => (toks[i] as { c: string }).c !== '#');
  const minInt = zeroFrom < 0 ? 0 : intPh.length - zeroFrom;
  let digits = ip === '0' ? '' : ip;
  digits = digits.padStart(minInt, '0');
  const out: string[] = toks.map((t) =>
    t.k === 'lit' || t.k === '%' ? t.s : t.k === '.' ? '.' : '',
  );
  if (intPh.length) {
    if (group) {
      out[intPh[0]!] = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    } else {
      const n = intPh.length;
      for (let k = 0; k < n; k++) {
        const pos = intPh[n - 1 - k]!;
        const ch = digits[digits.length - 1 - k];
        const ph = (toks[pos] as { c: string }).c;
        out[pos] = ch ?? (ph === '?' ? ' ' : '');
      }
      out[intPh[0]!] = digits.slice(0, Math.max(0, digits.length - n + 1)) || out[intPh[0]!]!;
    }
  } else if (digits !== '' && dot < 0 && ex < 0) {
    // no placeholders at all: Excel still prints the number before the literals
    out.unshift(digits);
  }
  let trim = true;
  for (let k = fracPh.length - 1; k >= 0; k--) {
    const pos = fracPh[k]!;
    const ph = (toks[pos] as { c: string }).c;
    const ch = fp[k]!;
    if (trim && ch === '0' && ph !== '0') out[pos] = ph === '?' ? ' ' : '';
    else {
      trim = false;
      out[pos] = ch;
    }
  }
  if (ex >= 0) {
    const es = String(Math.abs(e)).padStart(expPh.length, '0');
    out[ex] = `E${e < 0 ? '-' : (toks[ex] as { s: string }).s === '+' ? '+' : ''}${es}`;
    for (const p of expPh) out[p] = '';
  }
  return out.join('');
}

function date(v: number, sec: string): string {
  const toks: string[] = [];
  const re =
    /"([^"]*)"?|\\(.)|_.|\*.|\[(h+|m+|s+)\]|\[[^\]]*\]|(yyyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p)|(\.0+)|([\s\S])/gi;
  for (let m = re.exec(sec); m; m = re.exec(sec)) {
    if (m[1] !== undefined) toks.push(`'${m[1]}`);
    else if (m[2] !== undefined) toks.push(`'${m[2]}`);
    else if (m[0][0] === '_') toks.push("' ");
    else if (m[3] !== undefined) toks.push(`[${m[3].toLowerCase()}]`);
    else if (m[4] !== undefined) toks.push(m[4].toLowerCase());
    else if (m[5] !== undefined) toks.push(m[5]);
    else if (m[6] !== undefined) toks.push(`'${m[6]}`);
  }
  const codes = toks.filter((t) => t[0] !== "'");
  // m / mm after an hour or before a second means minutes
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]!;
    if (t !== 'm' && t !== 'mm') continue;
    const ci = codes.indexOf(t);
    const prev = codes[ci - 1];
    const next = codes[ci + 1];
    if ((prev && /^\[?h/.test(prev)) || (next && /^\[?s/.test(next)))
      toks[i] = t === 'm' ? 'M' : 'MM';
    codes[ci] = '';
  }
  const ampm = toks.some((t) => t === 'am/pm' || t === 'a/p');
  const fracTok = toks.find((t) => t[0] === '.');
  const fd = fracTok ? fracTok.length - 1 : 0;
  let days = Math.floor(v);
  let ms = Math.round((v - days) * 86_400_000);
  if (!fd) ms = Math.round(ms / 1000) * 1000;
  if (ms >= 86_400_000) {
    days++;
    ms -= 86_400_000;
  }
  const d = serialToDate(days, date1904());
  const y = d.getUTCFullYear();
  const mo = d.getUTCMonth();
  const day = d.getUTCDate();
  const H = Math.floor(ms / 3_600_000);
  const M = Math.floor(ms / 60_000) % 60;
  const Sx = Math.floor(ms / 1000) % 60;
  const p2 = (n: number): string => String(n).padStart(2, '0');
  const h12 = ampm ? H % 12 || 12 : H;
  const mn = MONTHS[mo]!;
  const dn = DAYS[d.getUTCDay()]!;
  const total = days * 86400 + ms / 1000;
  const parts: Record<string, string | number> = {
    yyyy: String(y).padStart(4, '0'),
    yy: p2(y % 100),
    y: p2(y % 100),
    mmmmm: mn[0]!,
    mmmm: mn,
    mmm: mn.slice(0, 3),
    mm: p2(mo + 1),
    m: mo + 1,
    dddd: dn,
    ddd: dn.slice(0, 3),
    dd: p2(day),
    d: day,
    hh: p2(h12),
    h: h12,
    MM: p2(M),
    M,
    ss: p2(Sx),
    s: Sx,
    'am/pm': H < 12 ? 'AM' : 'PM',
    'a/p': H < 12 ? 'A' : 'P',
  };
  let out = '';
  for (const t of toks) {
    out +=
      t[0] === "'"
        ? t.slice(1)
        : t[0] === '['
          ? Math.floor(total / (t[1] === 'h' ? 3600 : t[1] === 'm' ? 60 : 1))
          : t[0] === '.'
            ? ((ms % 1000) / 1000).toFixed(fd).slice(1)
            : (parts[t] ?? t);
  }
  return out;
}
