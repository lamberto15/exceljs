// Which formulas are dynamic-array formulas (their results spill). Shared by the writers
// (`cm="1"` + metadata) and the evaluator.

import type { Formula } from '../../core/types.js';

/** Functions whose results spill: such formulas are dynamic-array formulas by default. */
export const SPILL_FN: RegExp =
  /(?:^|[^\w.])(?:_xlfn\.)?(?:_xlws\.)?(?:FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|RANDARRAY|TRANSPOSE|TEXTSPLIT|VSTACK|HSTACK|TOCOL|TOROW|WRAPROWS|WRAPCOLS|TAKE|DROP|CHOOSEROWS|CHOOSECOLS|EXPAND|GROUPBY|PIVOTBY|ANCHORARRAY)\s*\(|\$?[A-Za-z]{1,3}\$?\d+#/i;

/**
 * A dynamic-array formula: `dynamic` when set, else one without `array` that uses a spilling
 * function.
 */
export const isDynamic = (f: Formula): boolean =>
  f.dynamic ?? (f.array === undefined && SPILL_FN.test(f.formula));
