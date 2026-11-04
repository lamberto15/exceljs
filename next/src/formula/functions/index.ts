// The built-in function table.

import type { Builtin } from '../evaluator.js';
import { S } from '../helpers.js';
import { isError, NA } from '../values.js';
import { logical } from './logical.js';
import { math } from './math.js';

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
  ...logical,
  ...info,
};
