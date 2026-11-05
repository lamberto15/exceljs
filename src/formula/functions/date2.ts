// More date / time functions: TIME, HOUR, MINUTE, SECOND, DATEVALUE, TIMEVALUE, WEEKNUM,
// ISOWEEKNUM, NETWORKDAYS(.INTL), WORKDAY(.INTL), YEARFRAC. Serials are in the workbook's date
// system (wb.date1904).

import { dateToSerial, serialToDate } from '../../core/dates.js';
import type { Builtin } from '../evaluator.js';
import { each, type FormulaFunction, S, sc } from '../helpers.js';
import {
  date1904,
  type FormulaScalar,
  type FormulaValue,
  isError,
  NUM,
  parseNumber,
  toNumber,
  VALUE,
} from '../values.js';

const DAY = 86_400_000;

/** Whole-day serial → UTC Date. */
function toDate(n: number): Date {
  if (n < 0) throw NUM;
  return serialToDate(Math.floor(n), date1904());
}

const toSerial = (d: Date): number => Math.round(dateToSerial(d, date1904()));

/** Day of week of a serial, 0 = Monday … 6 = Sunday. */
const dowMon = (s: number): number => (toDate(s).getUTCDay() + 6) % 7;

/** Seconds of the day of a serial (rounded to the second, as Excel shows them). */
function secs(v: FormulaScalar): number {
  const x = typeof v === 'string' ? textValue(v) : toNumber(v);
  if (x < 0) throw NUM;
  return Math.round((x - Math.floor(x)) * 86400) % 86400;
}

function textValue(t: string): number {
  const x = parseNumber(t);
  if (x === undefined) throw VALUE;
  return x;
}

/**
 * Weekend days of NETWORKDAYS.INTL / WORKDAY.INTL: codes 1-7 (two-day weekends starting
 * Sat/Sun, Sun/Mon, ...), 11-17 (one day, Sunday … Saturday) or a 7-character "0000011" mask
 * starting on Monday. Index 0 = Monday.
 */
function weekend(v: FormulaScalar | undefined): boolean[] {
  if (v === undefined || v === null) return [false, false, false, false, false, true, true];
  if (typeof v === 'string') {
    if (!/^[01]{7}$/.test(v) || v === '1111111') throw VALUE;
    return [...v].map((c) => c === '1');
  }
  const k = Math.trunc(toNumber(v));
  const w = [false, false, false, false, false, false, false];
  if (k >= 1 && k <= 7) {
    // 1: Sat+Sun, 2: Sun+Mon, ... 7: Fri+Sat
    const first = (k + 4) % 7; // Monday-based index of the first weekend day (1 → Saturday)
    w[first] = true;
    w[(first + 1) % 7] = true;
  } else if (k >= 11 && k <= 17) w[(k - 11 + 6) % 7] = true;
  else throw NUM;
  return w;
}

function holidays(v: FormulaValue | undefined): Set<number> {
  const out = new Set<number>();
  if (v === undefined || v === null) return out;
  each(v, (x) => {
    if (isError(x)) throw x;
    if (typeof x === 'number') out.add(Math.floor(x));
    else if (typeof x === 'string') out.add(Math.floor(textValue(x)));
  });
  return out;
}

const dateArg = (v: FormulaScalar): number =>
  Math.floor(typeof v === 'string' ? textValue(v) : toNumber(v));

function networkdays(s: number, e: number, w: boolean[], hol: Set<number>): number {
  const sign = e < s ? -1 : 1;
  const [a, b] = e < s ? [e, s] : [s, e];
  let count = 0;
  let d = dowMon(a);
  for (let x = a; x <= b; x++, d = (d + 1) % 7) if (!w[d] && !hol.has(x)) count++;
  return sign * count;
}

function workday(s: number, days: number, w: boolean[], hol: Set<number>): number {
  if (w.every(Boolean)) throw VALUE;
  let x = s;
  let left = Math.trunc(days);
  const step = left < 0 ? -1 : 1;
  while (left !== 0) {
    x += step;
    if (x < 0) throw NUM;
    if (!w[dowMon(x)] && !hol.has(x)) left -= step;
  }
  return x;
}

const leap = (y: number): boolean => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

function yearfrac(a: number, b: number, basis: number): number {
  if (a > b) [a, b] = [b, a];
  const d1 = toDate(a);
  const d2 = toDate(b);
  const y1 = d1.getUTCFullYear();
  const m1 = d1.getUTCMonth() + 1;
  let dd1 = d1.getUTCDate();
  const y2 = d2.getUTCFullYear();
  const m2 = d2.getUTCMonth() + 1;
  let dd2 = d2.getUTCDate();
  const days = b - a;
  switch (basis) {
    case 0: {
      const lastFeb = (y: number, m: number, d: number) => m === 2 && d === (leap(y) ? 29 : 28);
      if (lastFeb(y1, m1, dd1) && lastFeb(y2, m2, dd2)) dd2 = 30;
      if (lastFeb(y1, m1, dd1)) dd1 = 30;
      if (dd2 === 31 && dd1 >= 30) dd2 = 30;
      if (dd1 === 31) dd1 = 30;
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
    }
    case 1: {
      if (y1 === y2) return days / (leap(y1) ? 366 : 365);
      const withinYear = y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && dd1 >= dd2));
      if (withinYear) {
        const feb29 = (y: number) => Date.UTC(y, 1, 29) / DAY;
        const t1 = Date.UTC(y1, m1 - 1, dd1) / DAY;
        const t2 = Date.UTC(y2, m2 - 1, dd2) / DAY;
        const has = (leap(y1) && t1 <= feb29(y1)) || (leap(y2) && t2 >= feb29(y2));
        return days / (has ? 366 : 365);
      }
      let total = 0;
      for (let y = y1; y <= y2; y++) total += leap(y) ? 366 : 365;
      return days / (total / (y2 - y1 + 1));
    }
    case 2:
      return days / 360;
    case 3:
      return days / 365;
    case 4:
      if (dd1 === 31) dd1 = 30;
      if (dd2 === 31) dd2 = 30;
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
    default:
      throw NUM;
  }
}

/** WEEKNUM types: 1 / 17 Sunday, 2 / 11 Monday, 12-16 Tuesday … Saturday, 21 ISO. */
function weeknum(s: number, type: number): number {
  if (type === 21) return isoweek(s);
  const startDay =
    type === 1 || type === 17
      ? 6
      : type === 2 || type === 11
        ? 0
        : type >= 12 && type <= 16
          ? type - 11
          : -1;
  if (startDay < 0) throw NUM;
  const d = toDate(s);
  const jan1 = toSerial(new Date(Date.UTC(d.getUTCFullYear(), 0, 1)));
  const offset = (dowMon(jan1) - startDay + 7) % 7;
  return Math.floor((s - jan1 + offset) / 7) + 1;
}

function isoweek(s: number): number {
  const thursday = s - dowMon(s) + 3;
  const y = toDate(thursday).getUTCFullYear();
  const jan1 = toSerial(new Date(Date.UTC(y, 0, 1)));
  return Math.floor((thursday - jan1) / 7) + 1;
}

const nwd =
  (intl: boolean): FormulaFunction =>
  (args, ctx) => {
    if (args.length < 2 || args.length > (intl ? 4 : 3)) throw VALUE;
    const s = dateArg(sc(args[0]!, ctx));
    const e = dateArg(sc(args[1]!, ctx));
    const w = intl ? weekend(args[2] == null ? undefined : sc(args[2], ctx)) : weekend(undefined);
    return networkdays(s, e, w, holidays(args[intl ? 3 : 2]));
  };

const wd =
  (intl: boolean): FormulaFunction =>
  (args, ctx) => {
    if (args.length < 2 || args.length > (intl ? 4 : 3)) throw VALUE;
    const s = dateArg(sc(args[0]!, ctx));
    const days = toNumber(sc(args[1]!, ctx));
    const w = intl ? weekend(args[2] == null ? undefined : sc(args[2], ctx)) : weekend(undefined);
    return workday(s, days, w, holidays(args[intl ? 3 : 2]));
  };

export const date2: Record<string, Builtin> = {
  TIME: S((h, m, s) => {
    const H = Math.trunc(toNumber(h));
    const M = Math.trunc(toNumber(m));
    const Sx = Math.trunc(toNumber(s));
    if (H > 32767 || M > 32767 || Sx > 32767) throw NUM;
    const t = H * 3600 + M * 60 + Sx;
    if (t < 0) throw NUM;
    return (t % 86400) / 86400;
  }),
  HOUR: S((v) => Math.floor(secs(v) / 3600)),
  MINUTE: S((v) => Math.floor(secs(v) / 60) % 60),
  SECOND: S((v) => secs(v) % 60),
  DATEVALUE: S((v) => {
    if (typeof v !== 'string') throw VALUE;
    const x = Math.floor(textValue(v));
    if (x < 1) throw VALUE;
    return x;
  }),
  TIMEVALUE: S((v) => {
    if (typeof v !== 'string') throw VALUE;
    const x = textValue(v);
    return x - Math.floor(x);
  }),
  WEEKNUM: S((v, t) =>
    weeknum(dateArg(v), t === undefined || t === null ? 1 : Math.trunc(toNumber(t))),
  ),
  ISOWEEKNUM: S((v) => isoweek(dateArg(v))),
  NETWORKDAYS: nwd(false),
  'NETWORKDAYS.INTL': nwd(true),
  WORKDAY: wd(false),
  'WORKDAY.INTL': wd(true),
  YEARFRAC: S((a, b, basis) =>
    yearfrac(
      dateArg(a),
      dateArg(b),
      basis === undefined || basis === null ? 0 : Math.trunc(toNumber(basis)),
    ),
  ),
};
