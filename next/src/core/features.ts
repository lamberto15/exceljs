// Optional model fields owned by feature modules (src/xlsx/features/*). They are declared here,
// type-only, and merged into Worksheet / Workbook (interface + class declaration merging in
// worksheet.ts / workbook.ts) so the published .d.ts carries them — TypeScript module
// augmentation from the feature files would not survive declaration bundling.
// Runtime: plain properties, undefined until set; no initializers, nothing to tree-shake.
//
// Adding a feature field: one `import type` line and one documented optional field.

import type { HeaderFooter, PageSetup } from '../xlsx/features/print-setup.js';
import type { SheetProtection } from '../xlsx/features/sheet-protection.js';

export interface WorksheetFeatureFields {
  /** Page layout / print settings (features/print-setup). */
  pageSetup?: PageSetup | undefined;
  /** Print header & footer (features/print-setup). */
  headerFooter?: HeaderFooter | undefined;
  /** Sheet protection; undefined = unprotected (features/sheet-protection). */
  protection?: SheetProtection | undefined;
}

// biome-ignore lint/suspicious/noEmptyInterface: extension point for workbook-level features
export interface WorkbookFeatureFields {}
