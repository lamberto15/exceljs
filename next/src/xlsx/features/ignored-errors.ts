// Feature: ignored errors (<ignoredErrors>, ECMA-376 §18.3.1.51): switch off Excel's error-check
// triangles for ranges, e.g. "number stored as text" on a column of IDs or zip codes.
//
//   ws.ignoredErrors = [{ sqref: 'A2:A1000', numberStoredAsText: true }];

import { type FeatureReader, type FeatureWriter, SheetSlot } from './types.js';
import { attrs, bool, escapeAttr } from './util.js';

const CHECKS = [
  'evalError',
  'twoDigitTextYear',
  'numberStoredAsText',
  'formula',
  'formulaRange',
  'unlockedFormula',
  'emptyCellReference',
  'listDataValidation',
  'calculatedColumn',
] as const;

/** Error checks Excel skips for `sqref`. */
export interface IgnoredError {
  /** Target cells: space-separated refs ("A2:A100 C2"). */
  sqref: string;
  /** Formula results in an error. */
  evalError?: boolean | undefined;
  /** Text dates with two-digit years. */
  twoDigitTextYear?: boolean | undefined;
  /** Numbers formatted as text or preceded by an apostrophe. */
  numberStoredAsText?: boolean | undefined;
  /** Formula inconsistent with the formulas around it. */
  formula?: boolean | undefined;
  /** Formula omitting adjacent cells. */
  formulaRange?: boolean | undefined;
  /** Unlocked cells containing formulas. */
  unlockedFormula?: boolean | undefined;
  /** Formulas referring to empty cells. */
  emptyCellReference?: boolean | undefined;
  /** Data validation errors. */
  listDataValidation?: boolean | undefined;
  /** Inconsistent calculated-column formula in a table. */
  calculatedColumn?: boolean | undefined;
}

/**
 * Write half of the ignored-errors feature (`ws.ignoredErrors`). Always on in the package root;
 * in `./lite` pass it in `features`.
 */
export const ignoredErrorsWriter: FeatureWriter = {
  name: 'ignored-errors',
  sheet: {
    emitters: [
      {
        slot: SheetSlot.ignoredErrors,
        emit(c) {
          let s = '';
          for (const e of c.ws.ignoredErrors ?? []) {
            const on: Record<string, boolean | undefined> = {};
            for (const k of CHECKS) if (e[k]) on[k] = true;
            if (e.sqref.trim())
              s += `<ignoredError sqref="${escapeAttr(e.sqref.trim())}"${attrs(on)}/>`;
          }
          return s && `<ignoredErrors>${s}</ignoredErrors>`;
        },
      },
    ],
  },
};

/** Read half of the ignored-errors feature (without it the element is kept raw). */
export const ignoredErrorsReader: FeatureReader = {
  name: 'ignored-errors',
  sheet: {
    elements: ['ignoredErrors'],
    open(ctx) {
      const list: IgnoredError[] = [];
      ctx.ws.ignoredErrors = list;
      return {
        open(name, a) {
          if (name !== 'ignoredError' || !a.sqref) return;
          const e: IgnoredError = { sqref: a.sqref };
          for (const k of CHECKS) if (bool(a[k])) e[k] = true;
          list.push(e);
        },
      };
    },
  },
};
