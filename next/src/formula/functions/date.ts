// Date functions. Dates are serial numbers in the workbook's date system (wb.date1904).

import { dateToSerial, serialToDate } from '../../core/dates.js';
import type { Builtin } from '../evaluator.js';
import { type FormulaFunction, type FunctionContext, S } from '../helpers.js';
import { date1904, type FormulaScalar, NUM, toNumber, toText, VALUE } from '../values.js';

const DAY = 86_400_000;

/** Serial (integer part) → UTC Date; #NUM! for negative serials. */
function toDate(v: FormulaScalar): Date {
  const n = Math.floor(toNumber(v));
  if (n < 0) throw NUM;
  return serialToDate(n, date1904());
}

function serial(y: number, m: number, d: number): number {
  const s = Math.round(dateToSerial(new Date(Date.UTC(y, m, d)), date1904()));
  if (s < 0 || y > 9999) throw NUM;
  return s;
}

function now(ctx: FunctionContext): number {
  const d = ctx.now();
  const t = Date.UTC(
    d.getFullYear(),
    d.getMonth(),
    d.getDate(),
    d.getHours(),
    d.getMinutes(),
    d.getSeconds(),
    d.getMilliseconds(),
  );
  return dateToSerial(new Date(t), ctx.date1904);
}

const parts = (v: FormulaScalar): [number, number, number] => {
  const d = toDate(v);
  return [d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()];
};

const daysIn = (y: number, m: number): number => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();

const DATEDIF = S((a, b, u) => {
  const s = Math.floor(toNumber(a));
  const e = Math.floor(toNumber(b));
  if (s > e) throw NUM;
  const [sy, sm, sd] = parts(s);
  const [ey, em, ed] = parts(e);
  const months = (ey - sy) * 12 + em - sm - (ed < sd ? 1 : 0);
  switch (toText(u).toUpperCase()) {
    case 'Y':
      return Math.floor(months / 12);
    case 'M':
      return months;
    case 'D':
      return e - s;
    case 'MD':
      return ed >= sd ? ed - sd : daysIn(ey, em - 1) - sd + ed;
    case 'YM':
      return months % 12;
    case 'YD': {
      let y = ey;
      if (em < sm || (em === sm && ed < sd)) y--;
      return Math.round((Date.UTC(ey, em, ed) - Date.UTC(y, sm, sd)) / DAY);
    }
    default:
      throw NUM;
  }
});

const WEEKDAY = S((v, t) => {
  let n = Math.floor(toNumber(v));
  if (n < 0) throw NUM;
  if (date1904()) n += 1462;
  const dow = (((n - 1) % 7) + 7) % 7; // 0 = Sunday (serial 1 = Sunday 1900-01-01)
  const type = t === undefined || t === null ? 1 : Math.trunc(toNumber(t));
  if (type === 1 || type === 17) return dow + 1;
  if (type === 2) return ((dow + 6) % 7) + 1;
  if (type === 3) return (dow + 6) % 7;
  if (type >= 11 && type <= 16) return ((dow - (type - 10) + 7) % 7) + 1;
  throw NUM;
});

const today: FormulaFunction = (_args, ctx) => Math.floor(now(ctx));

export const date: Record<string, Builtin> = {
  DATE: S((y, m, d) => {
    let yy = Math.trunc(toNumber(y));
    if (yy < 0 || yy > 9999) throw NUM;
    if (yy < 1900) yy += 1900;
    return serial(yy, Math.trunc(toNumber(m)) - 1, Math.trunc(toNumber(d)));
  }),
  TODAY: today,
  NOW: (_args, ctx) => now(ctx),
  YEAR: S((v) => parts(v)[0]),
  MONTH: S((v) => parts(v)[1] + 1),
  DAY: S((v) => parts(v)[2]),
  WEEKDAY,
  EDATE: S((v, m) => {
    const [y, mo, d] = parts(v);
    const t = mo + Math.trunc(toNumber(m));
    return serial(y, t, Math.min(d, daysIn(y, t)));
  }),
  EOMONTH: S((v, m) => {
    const [y, mo] = parts(v);
    return serial(y, mo + Math.trunc(toNumber(m)) + 1, 0);
  }),
  DATEDIF,
  DAYS: S((e, s) => {
    const a = Math.floor(toNumber(e));
    const b = Math.floor(toNumber(s));
    if (!Number.isFinite(a - b)) throw VALUE;
    return a - b;
  }),
};
