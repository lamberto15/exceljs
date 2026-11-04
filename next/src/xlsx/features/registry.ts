// The registered features (internal; public API lives in ./index.ts). To add a feature: create
// its module in this folder (see README.md), add ONE import line below and its two halves to
// WRITERS / READERS. Order here does not matter (output order comes from SheetSlot / BookSlot).

import { commentsReader, commentsWriter } from './comments.js';
import {
  conditionalFormattingReader,
  conditionalFormattingWriter,
} from './conditional-formatting.js';
import { dataValidationReader, dataValidationWriter } from './data-validation.js';
import { imagesReader, imagesWriter } from './images.js';
import { printSetupReader, printSetupWriter } from './print-setup.js';
import { sheetProtectionReader, sheetProtectionWriter } from './sheet-protection.js';
import { tablesReader, tablesWriter } from './tables.js';
import type { FeatureReader, FeatureWriter } from './types.js';

export const WRITERS: readonly FeatureWriter[] = [
  commentsWriter,
  conditionalFormattingWriter,
  dataValidationWriter,
  imagesWriter,
  printSetupWriter,
  sheetProtectionWriter,
  tablesWriter,
];

export const READERS: readonly FeatureReader[] = [
  commentsReader,
  conditionalFormattingReader,
  dataValidationReader,
  imagesReader,
  printSetupReader,
  sheetProtectionReader,
  tablesReader,
];
