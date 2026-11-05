// Public value & style types shared by every module.
// Keep this file dependency-free: it is the contract between core, xlsx, csv and json (the one
// type-only import below is the feature extension-point interfaces used by the options).

import type { FeatureReader, FeatureWriter } from '../xlsx/features/types.js';

export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA'
  /** A dynamic-array formula cannot spill (its area is blocked). */
  | '#SPILL!'
  /** A dynamic-array calculation error (e.g. FILTER with no match and no fallback). */
  | '#CALC!';

export interface CellError {
  readonly error: ErrorCode;
}

export interface Formula {
  /** Formula text without the leading '='. */
  readonly formula: string;
  /**
   * Cached result, written as <v>; readers return what the file had. Dates are NOT converted:
   * for a date-formatted formula cell the result is the serial number (in the workbook's date
   * system) — convert with `serialToDate(result, wb.date1904)` when needed.
   */
  readonly result?: number | string | boolean | CellError | undefined;
  /** Array formula range, e.g. "A1:B3" (written as t="array"). */
  readonly array?: string | undefined;
  /**
   * Dynamic-array formula (spills, Excel 365): written with Excel's dynamic-array metadata so
   * Excel does not turn it into an implicit-intersection `@` formula. Default: on for formulas
   * using a spilling function (FILTER, SORT, UNIQUE, SEQUENCE, ...) without `array`; `false`
   * turns it off, `true` forces it. `array` (optional) is the last known spill range.
   */
  readonly dynamic?: boolean | undefined;
}

export interface RichTextRun {
  readonly text: string;
  readonly font?: Font | undefined;
}

export interface RichText {
  readonly richText: readonly RichTextRun[];
}

export interface Hyperlink {
  readonly text: string;
  readonly hyperlink: string;
  /** Formatted runs of the link text (their concatenated text equals `text`); kept from files. */
  readonly runs?: readonly RichTextRun[] | undefined;
  readonly tooltip?: string | undefined;
}

/** Everything a cell can hold. `null` / `undefined` = empty. */
export type CellValue =
  | number
  | string
  | boolean
  | Date
  | CellError
  | Formula
  | RichText
  | Hyperlink
  | null
  | undefined;

/**
 * ExcelJS shared-formula value, accepted when SETTING a cell (setCell, addRow, ...): the cell
 * gets the master cell's formula with its relative references moved by the offset from the
 * master (like `fillFormula`), stored and written as a plain `Formula`. The master must already
 * hold a formula. Never returned by getters.
 */
export interface SharedFormulaValue {
  /** Address of the master cell, e.g. "B1". */
  readonly sharedFormula: string;
  readonly result?: Formula['result'];
}

/** What a cell setter accepts: any {@link CellValue}, plus ExcelJS `{ sharedFormula }`. */
export type CellValueInput = CellValue | SharedFormulaValue;

/** Options of `eachRow` / `row.eachCell` / `column.eachCell` / `eachColumnCell` (ExcelJS). */
export interface IterateOptions {
  /**
   * Also visit empty rows / cells, from 1 to the last one: the sheet's `rowCount` for rows, the
   * row's last cell (styled empty cells included) for `row.eachCell`. Default false.
   */
  readonly includeEmpty?: boolean | undefined;
}

/** A workbook-level defined name (named range / constant / formula). */
export interface DefinedName {
  /** e.g. "Rate", or a built-in like "_xlnm.Print_Area". */
  readonly name: string;
  /** Formula text without '=', e.g. "Sheet1!$A$1:$D$10" or "0.2". */
  readonly value: string;
  /** Sheet scope: 0-based index into `workbook.worksheets`; undefined = workbook scope. */
  readonly localSheetId?: number | undefined;
  readonly hidden?: boolean | undefined;
  /** Optional comment shown in Excel's Name Manager. */
  readonly comment?: string | undefined;
}

// ---- styles -------------------------------------------------------------

export interface Color {
  /** ARGB hex, e.g. "FFFF0000". */
  readonly argb?: string | undefined;
  readonly theme?: number | undefined;
  readonly tint?: number | undefined;
  readonly indexed?: number | undefined;
}

export interface Font {
  readonly name?: string | undefined;
  readonly size?: number | undefined;
  readonly bold?: boolean | undefined;
  readonly italic?: boolean | undefined;
  readonly underline?:
    | boolean
    | 'single'
    | 'double'
    | 'singleAccounting'
    | 'doubleAccounting'
    | undefined;
  readonly strike?: boolean | undefined;
  readonly color?: Color | undefined;
  readonly vertAlign?: 'superscript' | 'subscript' | undefined;
  /** Font family class (OOXML `<family>`: 1 roman, 2 swiss, 3 modern, 4 script, 5 decorative). */
  readonly family?: number | undefined;
  /** Theme font scheme: 'minor' / 'major' make Excel use the theme's body / heading font. */
  readonly scheme?: 'major' | 'minor' | 'none' | undefined;
  /** Character set (OOXML `<charset>`, e.g. 0 ANSI, 128 Shift-JIS, 134 GB2312, 222 Thai). */
  readonly charset?: number | undefined;
  /** Outline / shadow (Mac Excel renders them) and condense / extend (legacy spacing flags). */
  readonly outline?: boolean | undefined;
  readonly shadow?: boolean | undefined;
  readonly condense?: boolean | undefined;
  readonly extend?: boolean | undefined;
}

export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dotted'
  | 'dashed'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface BorderEdge {
  readonly style?: BorderStyle | undefined;
  readonly color?: Color | undefined;
}

export interface Border {
  readonly top?: BorderEdge | undefined;
  readonly left?: BorderEdge | undefined;
  readonly bottom?: BorderEdge | undefined;
  readonly right?: BorderEdge | undefined;
  readonly diagonal?: (BorderEdge & { up?: boolean; down?: boolean }) | undefined;
}

export interface PatternFill {
  readonly type: 'pattern';
  readonly pattern: 'none' | 'solid' | 'gray125' | 'darkGray' | 'mediumGray' | 'lightGray' | string;
  readonly fgColor?: Color | undefined;
  readonly bgColor?: Color | undefined;
}

export interface GradientStop {
  /** 0..1 along the gradient. */
  readonly position: number;
  readonly color: Color;
}

/** Gradient fill (ExcelJS shape): linear at `degree`, or radiating from `center` ('path'). */
export interface GradientFill {
  readonly type: 'gradient';
  readonly gradient: 'angle' | 'path';
  /** Angle of a linear gradient in degrees (0 = left to right, 90 = top to bottom). */
  readonly degree?: number | undefined;
  /** Path gradients: the starting point as fractions of the cell (0.5, 0.5 = centre). */
  readonly center?: { readonly left: number; readonly top: number } | undefined;
  readonly stops: readonly GradientStop[];
}

export type Fill = PatternFill | GradientFill;

export interface Alignment {
  readonly horizontal?:
    | 'general'
    | 'left'
    | 'center'
    | 'right'
    | 'fill'
    | 'justify'
    | 'centerContinuous'
    | 'distributed'
    | undefined;
  /** `'middle'` is accepted as an ExcelJS alias of `'center'` (stored as 'center'). */
  readonly vertical?:
    | 'top'
    | 'center'
    | 'bottom'
    | 'justify'
    | 'distributed'
    | 'middle'
    | undefined;
  readonly wrapText?: boolean | undefined;
  readonly shrinkToFit?: boolean | undefined;
  readonly indent?: number | undefined;
  /**
   * Degrees in Excel's encoding: 1..90 counter-clockwise, 91..180 clockwise (= -1..-90), 255 =
   * vertical (stacked) text. Negative degrees and the ExcelJS alias `'vertical'` are accepted
   * and stored in that encoding (-45 → 135, 'vertical' → 255).
   */
  readonly textRotation?: number | 'vertical' | undefined;
}

export interface Protection {
  readonly locked?: boolean | undefined;
  readonly hidden?: boolean | undefined;
}

export interface Style {
  readonly numFmt?: string | undefined;
  readonly font?: Font | undefined;
  readonly fill?: Fill | undefined;
  readonly border?: Border | undefined;
  readonly alignment?: Alignment | undefined;
  readonly protection?: Protection | undefined;
  /**
   * In-cell checkbox (Excel 365, 2024+): boolean cells show a checkbox, TRUE = checked. Older
   * Excel versions show TRUE / FALSE. See `setCheckbox`.
   */
  readonly checkbox?: boolean | undefined;
}

// ---- options / profiles --------------------------------------------------

export type Profile = 'low-memory' | 'balanced' | 'max-speed';

export interface WriteOptions {
  profile?: Profile;
  /**
   * Deflate level 0-9. 0 = store. Async writers in browsers / workers deflate with
   * `CompressionStream`, which has no level: there any level > 0 uses the platform default (~6).
   */
  compression?: number;
  /** 'shared' builds sharedStrings.xml (smaller file); 'inline' writes inlineStr (faster, less RAM). */
  strings?: 'shared' | 'inline';
  /** Use 1904 date system. */
  date1904?: boolean;
  /**
   * Write the default Office theme part (`xl/theme/theme1.xml`, ~0.5 KB compressed). Default
   * true. `false` removes the `themeWriter` feature (in `./lite` it is only written when
   * `themeWriter` is in `features`).
   */
  theme?: boolean;
  /**
   * Feature writers to run. Package root: added to the built-in ones (all enabled; use it for
   * custom features). `./lite`: the ONLY features enabled, e.g. `[tablesWriter, imagesWriter]`.
   */
  features?: readonly FeatureWriter[];
}

/** Excel window settings of a workbook (ExcelJS `wb.views` entry; `<workbookView>`). */
export interface WorkbookView {
  /** 0-based index of the selected tab (default: the first visible sheet with tabSelected). */
  activeTab?: number | undefined;
  /** 0-based index of the first tab shown in the tab bar (scrolled). */
  firstSheet?: number | undefined;
  /** Window position and size in twips (1/20 pt). */
  x?: number | undefined;
  y?: number | undefined;
  width?: number | undefined;
  height?: number | undefined;
  visibility?: 'visible' | 'hidden' | 'veryHidden' | undefined;
}

export interface ReadOptions {
  profile?: Profile;
  /**
   * Hard limits against zip bombs / hostile files. `xlsxRows` applies the byte limits
   * (`maxEntryBytes`, `maxTotalBytes`) only to parts it has to hold in memory, not to sheets it
   * parses as they inflate.
   */
  limits?: Partial<Limits>;
  /**
   * Keep what the model does not represent (charts, shapes, pivot tables, threaded comments,
   * VBA, themes, custom XML, unknown worksheet / workbook elements, ...) so writeXlsx /
   * xlsxStream write it back. Default true. See `listPreservedParts`.
   */
  preserveUnknown?: boolean;
  /** Only load these sheet names. */
  sheets?: readonly string[];
  /**
   * Feature readers to run. Package root: added to the built-in ones (all enabled). `./lite`:
   * the ONLY features modelled, e.g. `[tablesReader]`; elements of the others are kept raw by
   * `preserveUnknown` and written back unchanged. `xlsxRows` ignores features.
   */
  features?: readonly FeatureReader[];
  /** Dates by local wall clock (see WorkbookOptions.localDates); also for xlsxRows. */
  localDates?: boolean | undefined;
  /**
   * Reject sheets with a cell in a row after `maxRows` (ExcelJS `maxRows`): `readXlsx` and
   * `xlsxRows` throw a `SheetLimitError` at the first such cell, before parsing the rest of the
   * sheet. Every `<c>` element counts (empty styled cells too); rows that only carry formatting
   * do not. Unset or 0: no limit.
   */
  maxRows?: number | undefined;
  /** Like `maxRows`, for cells in a column after `maxCols` (ExcelJS `maxCols`). */
  maxCols?: number | undefined;
  /**
   * `xlsxRows` only: called after the last row of each sheet was yielded, with what the sheet
   * XML stores after its cell data (merged ranges as A1 strings, e.g. "A1:C2").
   */
  onSheetEnd?:
    | ((info: { sheet: string; sheetIndex: number; merges: string[] }) => void)
    | undefined;
}

export interface Limits {
  maxEntries: number;
  maxEntryBytes: number;
  maxTotalBytes: number;
  maxCompressionRatio: number;
}

export const DEFAULT_LIMITS: Limits = {
  maxEntries: 10_000,
  maxEntryBytes: 1024 * 1024 * 1024,
  maxTotalBytes: 2 * 1024 * 1024 * 1024,
  maxCompressionRatio: 1000,
};
