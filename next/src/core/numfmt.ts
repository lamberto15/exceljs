// Built-in number formats (ECMA-376 Part 1, §18.8.30) and date-format detection.

/**
 * Built-in format codes by id (locale-independent ones only; ids 27–36 and 50–81 are
 * locale-specific and must be written explicitly if used).
 */
export const BUILTIN_NUM_FMTS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

/** Built-in ids that are dates/times, including the locale-specific ranges. */
function isDateId(id: number): boolean {
  return (
    (id >= 14 && id <= 22) ||
    (id >= 27 && id <= 36) ||
    (id >= 45 && id <= 47) ||
    (id >= 50 && id <= 58)
  );
}

let reverse: Map<string, number> | undefined;

/** Format code → built-in id, or undefined when the code is not built-in. */
export function builtinFormatId(fmt: string): number | undefined {
  if (reverse === undefined) {
    reverse = new Map();
    for (const k in BUILTIN_NUM_FMTS) reverse.set(BUILTIN_NUM_FMTS[k as unknown as number]!, +k);
    // common spelling variants Excel itself treats as the built-ins
    reverse.set('general', 0);
    reverse.set('m/d/yy "h":mm', 22);
    reverse.set('#,##0.00 ;(#,##0.00)', 39);
    reverse.set('#,##0.00 ;[Red](#,##0.00)', 40);
  }
  return reverse.get(fmt);
}

/** Built-in id → format code, or undefined (unknown/locale-specific id). */
export function builtinFormat(id: number): string | undefined {
  return BUILTIN_NUM_FMTS[id];
}

/** True when the built-in format id renders a date/time. */
export function isDateFormatId(id: number): boolean {
  return isDateId(id);
}

/**
 * True when a format code renders a date or time: looks for d, m, y, h, s (and e/g/b era
 * tokens are not needed — real date formats always contain one of those) outside of quoted
 * text, backslash-escaped characters, `_x` / `*x` padding, and `[...]` sections (colors,
 * conditions, locales). Elapsed-time brackets `[h]`, `[mm]`, `[ss]` count as time.
 * A lone `m` is ambiguous (month vs minute) but either way it is a date/time.
 */
export function isDateFormat(fmt: string): boolean {
  const n = fmt.length;
  for (let i = 0; i < n; i++) {
    const c = fmt.charCodeAt(i);
    switch (c) {
      case 34: {
        // "quoted text"
        const end = fmt.indexOf('"', i + 1);
        if (end < 0) return false;
        i = end;
        break;
      }
      case 92: // \x
      case 95: // _x
      case 42: // *x
        i++;
        break;
      case 91: {
        // [ ... ]
        const end = fmt.indexOf(']', i + 1);
        if (end < 0) return false;
        if (end > i + 1 && isElapsed(fmt, i + 1, end)) return true;
        i = end;
        break;
      }
      case 71: // G
      case 103: // g — skip "General"
        if (fmt.substr(i, 7).toLowerCase() === 'general') i += 6;
        break;
      default: {
        const l = c | 32;
        // d m y h s
        if (l === 100 || l === 109 || l === 121 || l === 104 || l === 115) return true;
      }
    }
  }
  return false;
}

function isElapsed(s: string, start: number, end: number): boolean {
  const first = s.charCodeAt(start) | 32;
  if (first !== 104 && first !== 109 && first !== 115) return false;
  for (let i = start + 1; i < end; i++) if ((s.charCodeAt(i) | 32) !== first) return false;
  return true;
}
