// Excel serial date conversion. All conversions are UTC based: a JS Date is treated as the
// wall-clock instant in UTC (new Date(Date.UTC(2024, 0, 1)) → 45292).
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
export function dateToSerial(d: Date, date1904 = false): number {
  const days = d.getTime() / MS_PER_DAY;
  if (date1904) return days + EPOCH_1904;
  const s = days + EPOCH_1900;
  // Before 1900-03-01 Excel's count is one less than the real day count from 1899-12-30.
  return s < 61 ? s - 1 : s;
}

/** Convert an Excel serial number to a Date (UTC), rounded to the millisecond. */
export function serialToDate(n: number, date1904 = false): Date {
  let s = n;
  if (date1904) s -= EPOCH_1904;
  else {
    if (s < 61) s += 1; // undo the fictitious 1900-02-29 for early dates
    s -= EPOCH_1900;
  }
  return new Date(Math.round(s * MS_PER_DAY));
}
