// Feature: data validation — <dataValidations> (ECMA-376 Part 1 §18.3.1.33) and, on read, the
// Excel 2010 form <extLst><ext><x14:dataValidations> (used by Excel when a list source refers
// to another sheet).
//
//   addDataValidation(ws, { sqref: 'B2:B100', type: 'list', values: ['Yes', 'No'] });
//   addDataValidation(ws, { sqref: 'C2:C100', type: 'list', formula1: 'Lists!$A$1:$A$9' });
//   addDataValidation(ws, { sqref: 'D2 F2', type: 'whole', operator: 'between',
//                           formula1: 1, formula2: 10, showErrorMessage: true });
//   addDataValidation(ws, { sqref: 'E2', type: 'date', operator: 'greaterThan',
//                           formula1: new Date(Date.UTC(2024, 0, 1)) });
//
// Formulas are written as given (a leading "=" is dropped); numbers as-is; a Date becomes an
// Excel serial (time of day only, for type 'time'). Read back, formulas are the raw strings
// (a literal list reads as formula1 '"Yes,No"'). Validations are always written in the classic
// form: Excel 2010+ accepts cross-sheet list sources there; the x14 form is read only.

import { dateToSerial } from '../../core/dates.js';
import type { Worksheet } from '../../core/worksheet.js';
import { type FeatureReader, type FeatureWriter, SheetSlot, type SubtreeHandler } from './types.js';
import { attrs, bool, escapeText } from './util.js';

export type DataValidationType =
  | 'list'
  | 'whole'
  | 'decimal'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom'
  | 'any';

export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

/** A formula operand: formula text (no "=" needed), a number, or a Date (→ serial). */
export type DataValidationFormula = string | number | Date;

export interface DataValidation {
  /** Target cells: space-separated refs ("A1:A10 C1"). */
  sqref: string;
  type: DataValidationType;
  /** Comparison for whole/decimal/date/time/textLength (Excel's default: 'between'). */
  operator?: DataValidationOperator | undefined;
  formula1?: DataValidationFormula | undefined;
  formula2?: DataValidationFormula | undefined;
  /**
   * Write-only convenience for type 'list': literal items, written as formula1 '"a,b,c"'
   * (takes precedence over formula1). Items may not contain commas; the joined list may not
   * exceed Excel's 255-character limit (use a range formula for longer lists).
   */
  values?: readonly string[] | undefined;
  allowBlank?: boolean | undefined;
  /** Hide the in-cell dropdown of a list (OOXML `showDropDown="1"`, inverted naming). */
  hideDropDown?: boolean | undefined;
  showInputMessage?: boolean | undefined;
  promptTitle?: string | undefined;
  prompt?: string | undefined;
  showErrorMessage?: boolean | undefined;
  errorStyle?: 'stop' | 'warning' | 'information' | undefined;
  errorTitle?: string | undefined;
  error?: string | undefined;
}

/** Excel's limit for a literal validation list (characters between the quotes). */
const MAX_LIST = 255;

/** `values` → '"a,b,c"'; throws when Excel would reject or misread the list. */
function listFormula(values: readonly string[]): string {
  for (const v of values) {
    if (v.includes(',')) {
      throw new RangeError(
        `Data validation list item ${JSON.stringify(v)} contains a comma; use a range formula instead`,
      );
    }
  }
  const s = values.join(',');
  if (s.length > MAX_LIST) {
    throw new RangeError(
      `Data validation list is ${s.length} characters; Excel allows at most ${MAX_LIST} — put the items in cells and use a range formula (e.g. "Lists!$A$1:$A$100")`,
    );
  }
  return `"${s.replace(/"/g, '""')}"`;
}

function formula(v: DataValidationFormula, type: DataValidationType, date1904: boolean): string {
  if (v instanceof Date) {
    const n = dateToSerial(v, date1904);
    return String(type === 'time' ? n - Math.floor(n) : n);
  }
  if (typeof v === 'number') return String(v);
  return v.charCodeAt(0) === 61 ? v.slice(1) : v; // drop a leading "="
}

/** Add a validation to the sheet (validates `values`); returns it. */
export function addDataValidation(ws: Worksheet, dv: DataValidation): DataValidation {
  if (dv.values) listFormula(dv.values);
  ws.dataValidations ??= [];
  ws.dataValidations.push(dv);
  return dv;
}

function emit(ws: Worksheet, date1904: boolean): string {
  const list = ws.dataValidations;
  if (!list?.length) return '';
  let s = '';
  for (const dv of list) {
    const t = dv.type;
    let body = '';
    const f1 = t === 'list' && dv.values ? listFormula(dv.values) : dv.formula1;
    if (f1 !== undefined) body += `<formula1>${escapeText(formula(f1, t, date1904))}</formula1>`;
    if (dv.formula2 !== undefined) {
      body += `<formula2>${escapeText(formula(dv.formula2, t, date1904))}</formula2>`;
    }
    const a = attrs({
      type: t === 'any' ? undefined : t,
      errorStyle: dv.errorStyle,
      operator: dv.operator,
      allowBlank: dv.allowBlank,
      showDropDown: dv.hideDropDown,
      showInputMessage: dv.showInputMessage,
      showErrorMessage: dv.showErrorMessage,
      errorTitle: dv.errorTitle,
      error: dv.error,
      promptTitle: dv.promptTitle,
      prompt: dv.prompt,
      sqref: dv.sqref,
    });
    s += body ? `<dataValidation${a}>${body}</dataValidation>` : `<dataValidation${a}/>`;
  }
  return `<dataValidations count="${list.length}">${s}</dataValidations>`;
}

export const dataValidationWriter: FeatureWriter = {
  name: 'data-validation',
  sheet: {
    emitters: [{ slot: SheetSlot.dataValidations, emit: (c) => emit(c.ws, c.workbook.date1904) }],
  },
};

/** Extension URI of x14:dataValidations. */
const X14_DV = '{CCE6A557-97BC-4b89-ADB8-1C0A8D0A6E5B}';

/** Parses <dataValidation> elements (classic or x14: formulas in <xm:f>, sqref in <xm:sqref>). */
function handler(ws: Worksheet): SubtreeHandler {
  ws.dataValidations ??= [];
  const list = ws.dataValidations;
  let cur: DataValidation | undefined;
  let cap: 'formula1' | 'formula2' | 'sqref' | undefined;
  let text = '';
  return {
    open(name, a) {
      switch (name) {
        case 'dataValidation': {
          const dv: DataValidation = {
            sqref: a.sqref ?? '',
            type: (a.type ?? 'any') as DataValidationType,
          };
          if (a.operator) dv.operator = a.operator as DataValidationOperator;
          if (bool(a.allowBlank)) dv.allowBlank = true;
          if (bool(a.showDropDown)) dv.hideDropDown = true;
          if (bool(a.showInputMessage)) dv.showInputMessage = true;
          if (bool(a.showErrorMessage)) dv.showErrorMessage = true;
          if (a.errorStyle) dv.errorStyle = a.errorStyle as DataValidation['errorStyle'];
          if (a.errorTitle !== undefined) dv.errorTitle = a.errorTitle;
          if (a.error !== undefined) dv.error = a.error;
          if (a.promptTitle !== undefined) dv.promptTitle = a.promptTitle;
          if (a.prompt !== undefined) dv.prompt = a.prompt;
          cur = dv;
          return;
        }
        case 'formula1':
        case 'formula2':
        case 'sqref':
          cap = name;
          text = '';
      }
    },
    text(t) {
      if (cap) text += t;
    },
    close(name) {
      if (name === 'dataValidation') {
        if (cur) list.push(cur);
        cur = undefined;
      } else if (cap === name && cur) {
        if (name === 'sqref') cur.sqref = text;
        else cur[cap] = text;
        cap = undefined;
      }
    },
  };
}

export const dataValidationReader: FeatureReader = {
  name: 'data-validation',
  sheet: {
    elements: ['dataValidations', 'extLst/ext'],
    open(ctx, el, a) {
      if (el === 'dataValidations' || a.uri === X14_DV) return handler(ctx.ws);
      return undefined;
    },
  },
};
