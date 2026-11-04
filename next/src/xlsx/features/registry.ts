// The built-in features of the package root (internal; public API lives in ./index.ts). The
// root's writeXlsx / xlsxStream / createXlsxWriter / readXlsx (write/index.ts, read/index.ts)
// run every feature listed here; the `./lite` entry never imports this module, so its users
// get only the features they pass in `options.features`.
//
// To add a feature: create its module in this folder (see README.md), add ONE import line below
// and its two halves to WRITERS / READERS (and export them from ./index.ts). Order here does
// not matter for element order (that comes from SheetSlot / BookSlot); `themeWriter` stays last
// so the theme part and relationship follow the other workbook-level parts.

import { chartsWriter } from './charts.js';
import { commentsReader, commentsWriter } from './comments.js';
import {
  conditionalFormattingReader,
  conditionalFormattingWriter,
} from './conditional-formatting.js';
import { dataValidationReader, dataValidationWriter } from './data-validation.js';
import { ignoredErrorsReader, ignoredErrorsWriter } from './ignored-errors.js';
import { imagesReader, imagesWriter } from './images.js';
import { printSetupReader, printSetupWriter } from './print-setup.js';
import { sheetProtectionReader, sheetProtectionWriter } from './sheet-protection.js';
import { tablesReader, tablesWriter } from './tables.js';
import { themeWriter } from './theme.js';
import type { FeatureReader, FeatureWriter } from './types.js';

export const WRITERS: readonly FeatureWriter[] = [
  chartsWriter,
  commentsWriter,
  conditionalFormattingWriter,
  dataValidationWriter,
  ignoredErrorsWriter,
  imagesWriter,
  printSetupWriter,
  sheetProtectionWriter,
  tablesWriter,
  themeWriter,
];

export const READERS: readonly FeatureReader[] = [
  commentsReader,
  conditionalFormattingReader,
  dataValidationReader,
  ignoredErrorsReader,
  imagesReader,
  printSetupReader,
  sheetProtectionReader,
  tablesReader,
];

/** The built-ins plus the caller's extra features (duplicates dropped). */
export function withAll<T>(all: readonly T[], extra: readonly T[] | undefined): readonly T[] {
  if (!extra || extra.length === 0) return all;
  return [...all, ...extra.filter((f) => !all.includes(f))];
}
