// Optional model fields owned by feature modules (src/xlsx/features/*). They are declared here,
// type-only, and merged into Worksheet / Workbook (interface + class declaration merging in
// worksheet.ts / workbook.ts) so the published .d.ts carries them — TypeScript module
// augmentation from the feature files would not survive declaration bundling.
// Runtime: plain properties, undefined until set; no initializers, nothing to tree-shake.
//
// Adding a feature field: one `import type` line and one documented optional field.

import type { PBook, PSheet } from '../xlsx/common/preserved.js';
import type { Chart } from '../xlsx/features/charts.js';
import type { Note } from '../xlsx/features/comments.js';
import type { ConditionalFormatting } from '../xlsx/features/conditional-formatting.js';
import type { DataValidation } from '../xlsx/features/data-validation.js';
import type { IgnoredError } from '../xlsx/features/ignored-errors.js';
import type { ImagePlacement, WorkbookImage } from '../xlsx/features/images.js';
import type { HeaderFooter, PageSetup } from '../xlsx/features/print-setup.js';
import type { SheetProtection } from '../xlsx/features/sheet-protection.js';
import type { Table } from '../xlsx/features/tables.js';
import type { WorkbookView } from './types.js';

export interface WorksheetFeatureFields {
  /** Page layout / print settings (features/print-setup). */
  pageSetup?: PageSetup | undefined;
  /** Print header & footer (features/print-setup). */
  headerFooter?: HeaderFooter | undefined;
  /** Manual page breaks after these 1-based rows (features/print-setup). */
  rowBreaks?: number[] | undefined;
  /** Manual page breaks after these 1-based columns (features/print-setup). */
  colBreaks?: number[] | undefined;
  /** Excel error checks to switch off for ranges (features/ignored-errors). */
  ignoredErrors?: IgnoredError[] | undefined;
  /** Cell notes (legacy comments), one per cell (features/comments; see addNote). */
  comments?: Note[] | undefined;
  /** Pictures placed on the sheet (features/images; see placeImage). */
  images?: ImagePlacement[] | undefined;
  /** Charts placed on the sheet (features/charts; see addChart). Write only: charts of read files are kept raw. */
  charts?: Chart[] | undefined;
  /** Sheet background picture: an id of workbook.images (features/images). */
  backgroundImage?: number | undefined;
  /** Conditional formatting blocks (features/conditional-formatting; see addConditionalFormatting). */
  conditionalFormatting?: ConditionalFormatting[] | undefined;
  /** Data validations (features/data-validation; see addDataValidation). */
  dataValidations?: DataValidation[] | undefined;
  /** Sheet protection; undefined = unprotected (features/sheet-protection). */
  protection?: SheetProtection | undefined;
  /** Tables / ListObjects (features/tables); use `addTable` to create one with its cells. */
  tables?: Table[] | undefined;
  /**
   * Round-trip data kept by `readXlsx` (preserveUnknown): XML and parts the model does not
   * represent. Opaque; written back by `writeXlsx` / `xlsxStream`. Delete it to drop them.
   */
  preserved?: PSheet | undefined;
}

export interface WorkbookFeatureFields {
  /** Document properties (File > Info / Properties), written to docProps/core.xml & app.xml. */
  title?: string | undefined;
  subject?: string | undefined;
  /** Tags; Excel shows them comma / semicolon separated. */
  keywords?: string | undefined;
  /** Comments. */
  description?: string | undefined;
  category?: string | undefined;
  /** Status, e.g. "Draft". */
  contentStatus?: string | undefined;
  company?: string | undefined;
  manager?: string | undefined;
  /**
   * Calculation settings (ExcelJS `wb.calcProperties`). `fullCalcOnLoad`: Excel recalculates
   * every formula when opening the file. Default: on when the workbook has formulas (their
   * cached results may be missing); set false when they are all current (e.g. after
   * `evaluate`) to spare Excel the recalculation and its "save changes?" prompt.
   */
  calcProperties?: { fullCalcOnLoad?: boolean | undefined } | undefined;
  /** Excel window settings (first entry is written; ExcelJS `wb.views`). */
  views?: WorkbookView[] | undefined;
  /** Pictures stored once per workbook, placed via worksheet.images (features/images; see addImage). */
  images?: WorkbookImage[] | undefined;
  /**
   * Round-trip data kept by `readXlsx` (preserveUnknown): package parts and workbook XML the
   * model does not represent (inspect with `listPreservedParts`). Opaque; written back by
   * `writeXlsx` / `xlsxStream` (not by `createXlsxWriter`). Delete it to drop them.
   */
  preserved?: PBook | undefined;
}
