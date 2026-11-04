// Public value & style types shared by every module.
// Keep this file dependency-free: it is the contract between core, xlsx, csv and json.

export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA';

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

export interface Fill {
  readonly type: 'pattern';
  readonly pattern: 'none' | 'solid' | 'gray125' | 'darkGray' | 'mediumGray' | 'lightGray' | string;
  readonly fgColor?: Color | undefined;
  readonly bgColor?: Color | undefined;
}

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
  readonly vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed' | undefined;
  readonly wrapText?: boolean | undefined;
  readonly shrinkToFit?: boolean | undefined;
  readonly indent?: number | undefined;
  readonly textRotation?: number | undefined;
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
}

// ---- options / profiles --------------------------------------------------

export type Profile = 'low-memory' | 'balanced' | 'max-speed';

export interface WriteOptions {
  profile?: Profile;
  /** Deflate level 0-9. 0 = store. */
  compression?: number;
  /** 'shared' builds sharedStrings.xml (smaller file); 'inline' writes inlineStr (faster, less RAM). */
  strings?: 'shared' | 'inline';
  /** Use 1904 date system. */
  date1904?: boolean;
  /** Write the default Office theme part (`xl/theme/theme1.xml`, ~0.5 KB compressed). Default true. */
  theme?: boolean;
}

export interface ReadOptions {
  profile?: Profile;
  /** Hard limits against zip bombs / hostile files. */
  limits?: Partial<Limits>;
  /**
   * Keep what the model does not represent (charts, shapes, pivot tables, threaded comments,
   * VBA, themes, custom XML, unknown worksheet / workbook elements, ...) so writeXlsx /
   * xlsxStream write it back. Default true. See `listPreservedParts`.
   */
  preserveUnknown?: boolean;
  /** Only load these sheet names. */
  sheets?: readonly string[];
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
