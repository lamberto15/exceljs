// The built-in function table.

import type { Builtin } from '../evaluator.js';
import { S } from '../helpers.js';
import { isError, NA } from '../values.js';
import { date } from './date.js';
import { date2 } from './date2.js';
import { dynamic } from './dynamic.js';
import { finance } from './finance.js';
import { logical } from './logical.js';
import { lookup } from './lookup.js';
import { math } from './math.js';
import { math2 } from './math2.js';
import { stats } from './stats.js';
import { text } from './text.js';
import { text2 } from './text2.js';

const info: Record<string, Builtin> = {
  ISBLANK: S((v) => v === null),
  ISNUMBER: S((v) => typeof v === 'number'),
  ISTEXT: S((v) => typeof v === 'string'),
  ISNONTEXT: S((v) => typeof v !== 'string'),
  ISLOGICAL: S((v) => typeof v === 'boolean'),
  ISERROR: S((v) => isError(v)),
  ISNA: S((v) => isError(v) && v.error === '#N/A'),
  ISERR: S((v) => isError(v) && v.error !== '#N/A'),
  NA: () => NA,
  SINGLE: S((v) => v),
};

/** Every built-in function, by upper-case name. */
export const BUILTINS: Record<string, Builtin> = {
  ...math,
  ...math2,
  ...stats,
  ...logical,
  ...lookup,
  ...text,
  ...date,
  ...date2,
  ...finance,
  ...text2,
  ...dynamic,
  ...info,
};
