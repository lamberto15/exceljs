// Feature extension points. A feature is a self-contained module (see README.md) exporting a
// `FeatureWriter` and a `FeatureReader`. The writers and the reader drive the features they are
// given (`options.features`; the package root adds every built-in from ./registry.ts) through the
// interfaces below; feature data lives on the model as optional Worksheet / Workbook fields
// declared in src/core/features.ts.

import type { StyleRegistry } from '../../core/styles.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import type { XmlAttrs, XmlHandler } from '../../xml/index.js';
import type { Relationship } from '../common/r-opc.js';

/**
 * CT_Worksheet child order (ECMA-376 Part 1, 18.3.1.99). Every element of a worksheet part is
 * produced by an emitter bound to one of these slots; emitters run sorted by slot, so element
 * order is guaranteed globally. Slots below `sheetData` form the head (written before the
 * rows — in the streaming writer that is when the sheet starts), the rest the tail (written
 * after the rows / at commit).
 */
export const SheetSlot = {
  sheetPr: 0,
  dimension: 1,
  sheetViews: 2,
  sheetFormatPr: 3,
  cols: 4,
  sheetData: 5,
  sheetCalcPr: 6,
  sheetProtection: 7,
  protectedRanges: 8,
  scenarios: 9,
  autoFilter: 10,
  sortState: 11,
  dataConsolidate: 12,
  customSheetViews: 13,
  mergeCells: 14,
  phoneticPr: 15,
  conditionalFormatting: 16,
  dataValidations: 17,
  hyperlinks: 18,
  printOptions: 19,
  pageMargins: 20,
  pageSetup: 21,
  headerFooter: 22,
  rowBreaks: 23,
  colBreaks: 24,
  customProperties: 25,
  cellWatches: 26,
  ignoredErrors: 27,
  smartTags: 28,
  drawing: 29,
  legacyDrawing: 30,
  legacyDrawingHF: 31,
  picture: 32,
  oleObjects: 33,
  controls: 34,
  webPublishItems: 35,
  tableParts: 36,
  extLst: 37,
} as const;
export type SheetSlot = (typeof SheetSlot)[keyof typeof SheetSlot];

/**
 * Child order inside `<sheetPr>` (CT_SheetPr). `sheetPr` is a container slot: emitters bound to
 * it return CHILD elements (ordered by `order`) and the framework wraps them in one `<sheetPr>`.
 */
export const SheetPrOrder = { tabColor: 0, outlinePr: 1, pageSetUpPr: 2 } as const;

/**
 * CT_Workbook child order (18.2.27). Built-ins: workbookPr, bookViews, sheets, definedNames,
 * calcPr. Features add e.g. workbookProtection, pivotCaches, extLst.
 */
export const BookSlot = {
  fileVersion: 0,
  fileSharing: 1,
  workbookPr: 2,
  workbookProtection: 3,
  bookViews: 4,
  sheets: 5,
  functionGroups: 6,
  externalReferences: 7,
  definedNames: 8,
  calcPr: 9,
  oleSize: 10,
  customWorkbookViews: 11,
  pivotCaches: 12,
  smartTagPr: 13,
  smartTagTypes: 14,
  webPublishing: 15,
  fileRecoveryPr: 16,
  webPublishObjects: 17,
  extLst: 18,
} as const;
export type BookSlot = (typeof BookSlot)[keyof typeof BookSlot];

// ---------------------------------------------------------------- writing

/** Package-level services shared by sheet and workbook write contexts. */
export interface PackageWriteCtx {
  /**
   * Add a zip part (path without leading '/', e.g. "xl/comments1.xml"). From a sheet emitter it
   * is written right after the sheet part (in every writer, including the streaming one); from
   * `book.write` together with the workbook-level parts. `contentType` registers an Override.
   */
  addPart(path: string, data: string | Uint8Array, contentType?: string): void;
  /** Register a Default content type by file extension (e.g. "png", "image/png"); idempotent. */
  addDefault(ext: string, contentType: string): void;
  /**
   * Next 1-based number of a part family, unique within the package: `nextIndex('comments')`
   * → 1, 2, ... (for "xl/comments1.xml", "xl/comments2.xml"). Families are free-form strings.
   */
  nextIndex(family: string): number;
}

export interface SheetWriteCtx extends PackageWriteCtx {
  readonly ws: Worksheet;
  readonly workbook: Workbook;
  readonly styles: StyleRegistry;
  /** 0-based position of the sheet in the package. */
  readonly index: number;
  /** Zip path of the sheet part, e.g. "xl/worksheets/sheet1.xml". */
  readonly path: string;
  /** True in `createXlsxWriter` (the Worksheet holds metadata only, never cells). */
  readonly streaming: boolean;
  /** True when this sheet is the selected tab. */
  readonly tabSelected: boolean;
  /**
   * Add a relationship from this sheet; returns its id ("rId3"). `type`: a short name of the
   * officeDocument relationships namespace ("comments", "vmlDrawing", "table", "drawing",
   * "hyperlink", ...) or a full URI. `target`: relative to the sheet (e.g. "../comments1.xml")
   * or an absolute URL with `external = true`.
   */
  addRel(type: string, target: string, external?: boolean): string;
}

/**
 * Produces one worksheet element (or, for container slots `sheetPr` / `extLst`, one child:
 * `<tabColor/>`-like elements resp. `<ext uri="{...}" xmlns:x14="...">…</ext>`).
 * Return '' when there is nothing to write. Runs once per sheet: head slots when the sheet
 * starts, tail slots after the rows.
 */
export interface SheetEmitter {
  readonly slot: SheetSlot;
  /** Order among emitters of a container slot (SheetPrOrder for sheetPr). Default 0. */
  readonly order?: number;
  emit(ctx: SheetWriteCtx): string;
}

export interface BookWriteCtx extends PackageWriteCtx {
  readonly workbook: Workbook;
  readonly styles: StyleRegistry;
  /** Add a relationship from the workbook part (same conventions as SheetWriteCtx.addRel). */
  addRel(type: string, target: string, external?: boolean): string;
}

/** Produces one workbook.xml element (container slot: extLst). */
export interface BookEmitter {
  readonly slot: BookSlot;
  readonly order?: number;
  emit(ctx: BookWriteCtx): string;
}

// ---------------------------------------------------------------- reading

export interface PackageReadCtx {
  readonly workbook: Workbook;
  /** Decoded text of a zip part (lazy inflate); undefined when absent. Case-insensitive. */
  readText(path: string): string | undefined;
  /** Raw bytes of a zip part; undefined when absent. */
  readBytes(path: string): Uint8Array | undefined;
  /** Relationships of any part (keyed by id), e.g. a drawing's rels. */
  relsOf(path: string): ReadonlyMap<string, Relationship> | undefined;
}

export interface SheetReadCtx extends PackageReadCtx {
  readonly ws: Worksheet;
  /** Zip path of the sheet part. */
  readonly path: string;
  /** The sheet's relationships (resolve `r:id` attributes). */
  readonly rels: ReadonlyMap<string, Relationship> | undefined;
  /** File cellXfs index (the `s` / `style` attribute) → registry xf index. */
  xf(s: string | number): number;
  /** File dxf index (`dxfId`) → registry dxf index (see StyleRegistry.internDxf). */
  dxf(id: string | number): number;
  /** Per-sheet scratch space, e.g. `ctx.scratch.set(myFeature, state)` in open, read in finish. */
  readonly scratch: Map<unknown, unknown>;
}

/**
 * Receives a delegated subtree. The element's own open tag goes to `SheetReader.open`; its
 * descendants' events and its direct text go here (local names, prefixes stripped); `end` is
 * called when the element closes. Attribute objects are shared and only valid during the call:
 * copy what you keep.
 */
export interface SubtreeHandler extends XmlHandler {
  end?(): void;
}

export interface SheetReader {
  /**
   * Worksheet elements this feature consumes, as paths below `<worksheet>`: "sheetProtection",
   * "pageSetup", or one level deeper inside sheetPr: "sheetPr/pageSetUpPr". Never elements of
   * sheetData (the cell hot loop does not dispatch to features).
   */
  readonly elements?: readonly string[];
  /** A consumed element opened. Return a handler for its subtree (or undefined for leaves). */
  open?(ctx: SheetReadCtx, el: string, attrs: XmlAttrs): SubtreeHandler | undefined;
  /** After the sheet part is parsed; load related parts (via ctx.rels + readText) here. */
  finish?(ctx: SheetReadCtx): void;
}

// ---------------------------------------------------------------- feature

/**
 * Write half of a feature. Kept separate from the read half so write-only bundles do not ship
 * feature readers (and vice versa): `./lite` users pass `<name>Writer` to the writers and
 * `<name>Reader` to readXlsx; registry.ts lists the built-ins in WRITERS / READERS.
 */
export interface FeatureWriter {
  /** Unique name (diagnostics). */
  readonly name: string;
  readonly sheet?: {
    readonly emitters?: readonly SheetEmitter[];
    /** Extra namespace declarations for `<worksheet>`, e.g. ['xmlns:x14ac="..."']. */
    readonly namespaces?: readonly string[];
  };
  readonly book?: {
    readonly emitters?: readonly BookEmitter[];
    /** Workbook-level parts/rels, run once after all sheets were written. */
    write?(ctx: BookWriteCtx): void;
  };
}

/** Read half of a feature (DOM mode, `readXlsx`, only; `xlsxRows` ignores features). */
export interface FeatureReader {
  readonly name: string;
  readonly sheet?: SheetReader;
  /** After every sheet was read: workbook-level data / parts. */
  book?(ctx: PackageReadCtx): void;
}
