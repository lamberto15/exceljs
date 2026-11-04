// `turbosheet/formula`: formula parsing and evaluation.
//
//   import { evaluate } from 'turbosheet/formula';
//   const report = evaluate(wb);        // fills Formula.result of every formula cell
//
// Separate entry: the core bundles never include it.

import type { Workbook } from '../core/workbook.js';
import type { Worksheet } from '../core/worksheet.js';
import {
  type EvaluateFormulaOptions,
  type EvaluateOptions,
  type EvaluationReport,
  Evaluator,
  type EvaluatorOptions,
  type FormulaResult,
} from './evaluator.js';
import { BUILTINS } from './functions/index.js';

export {
  type EvaluateFormulaOptions,
  type EvaluateOptions,
  type EvaluationReport,
  Evaluator,
  type EvaluatorOptions,
  type FormulaResult,
} from './evaluator.js';
export type { FormulaFunction, FunctionContext } from './helpers.js';
export {
  type ArrayLiteral,
  type BinaryOp,
  type FormulaNode,
  FormulaParseError,
  parseFormula,
} from './parser.js';
export {
  check,
  compare,
  FormulaArray,
  FormulaRange,
  type FormulaScalar,
  type FormulaValue,
  isError,
  numToStr,
  parseNumber,
  round15,
  scalarValues,
  toBoolean,
  toNumber,
  toText,
} from './values.js';

/** A new evaluator with the built-in functions plus `options.functions`. */
export function createEvaluator(options: EvaluatorOptions = {}): Evaluator {
  return new Evaluator(BUILTINS, options);
}

let shared: Evaluator | undefined;
const def = (): Evaluator => {
  shared ??= createEvaluator();
  return shared;
};

/** Evaluate every formula of the workbook and store the results (see EvaluateOptions). */
export function evaluate(wb: Workbook, options?: EvaluateOptions): EvaluationReport {
  return def().evaluate(wb, options);
}

/** Evaluate one cell (and its dependencies); results are not stored unless `writeResults`. */
export function evaluateCell(
  wb: Workbook,
  sheet: string | Worksheet,
  address: string,
  options?: EvaluateOptions,
): FormulaResult | null {
  return def().evaluateCell(wb, sheet, address, options);
}

/** Evaluate formula text against the workbook; a multi-cell result is a 2-D array. */
export function evaluateFormula(
  wb: Workbook,
  text: string,
  options?: EvaluateFormulaOptions,
): FormulaResult | FormulaResult[][] {
  return def().evaluateFormula(wb, text, options);
}

/** Names of the built-in functions. */
export function supportedFunctions(): string[] {
  return Object.keys(BUILTINS).sort();
}
