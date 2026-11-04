// Excel serial date conversion. Excel dates have no time zone. By default conversions are UTC
// based: a JS Date is read as its UTC wall clock (new Date(Date.UTC(2024, 0, 1)) → 45292).
// With `local` the Date's local wall clock is used instead (new Date(2024, 0, 1) → 45292 in
// every time zone), which is what code building dates with `new Date(y, m, d)` expects.
//
// 1900 system: serial 1 = 1900-01-01. Excel pretends 1900 was a leap year, so serial 60 is the
// fictitious 1900-02-29 and every serial >= 61 is "one too high" relative to real day counts.
//   - dateToSerial: dates before 1900-03-01 get the real count (1900-02-28 → 59, 1900-03-01 → 61).
//   - serialToDate: serial 60 maps to 1900-03-01 (just as Date.UTC(1900, 1, 29) rolls over);
//     serials < 60 are shifted so 59 → 1900-02-28 and 1 → 1900-01-01.
// 1904 system: serial 0 = 1904-01-01, no leap-year bug.

const MS_PER_DAY = 86_400_000;
/** Serial of 1970-01-01 in the 1900 system (epoch effectively 1899-12-30 for serials >= 61). */
const EPOCH_1900 = 25_569;
/** Serial of 1970-01-01 in the 1904 system. */
const EPOCH_1904 = 24_107;

/** Convert a Date to an Excel serial number (fractional part = time of day). */
export function dateToSerial(d: Date, date1904 = false, local = false): number {
  const days = (d.getTime() - (local ? d.getTimezoneOffset() * 60_000 : 0)) / MS_PER_DAY;
  if (date1904) return days + EPOCH_1904;
  const s = days + EPOCH_1900;
  // Before 1900-03-01 Excel's count is one less than the real day count from 1899-12-30.
  return s < 61 ? s - 1 : s;
}

/**
 * Convert an Excel serial number to a Date, rounded to the millisecond: its UTC wall clock is
 * the serial's date and time, or its local wall clock with `local`.
 */
export function serialToDate(n: number, date1904 = false, local = false): Date {
  let s = n;
  if (date1904) s -= EPOCH_1904;
  else {
    if (s < 61) s += 1; // undo the fictitious 1900-02-29 for early dates
    s -= EPOCH_1900;
  }
  const u = new Date(Math.round(s * MS_PER_DAY));
  if (!local) return u;
  const d = new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
  d.setHours(u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds(), u.getUTCMilliseconds());
  if (u.getUTCFullYear() < 100) d.setFullYear(u.getUTCFullYear());
  return d;
}
