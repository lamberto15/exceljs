// Round-trip preservation model (ReadOptions.preserveUnknown): package parts and XML the model
// does not represent, kept by `readXlsx` and re-emitted by `writeXlsx` / `xlsxStream`.
//
// - Parts are kept as lazy raw bytes (`data()` inflates from the source archive on demand) with
//   their content type and their own relationships (original ids, so their XML is untouched).
// - XML fragments (worksheet / workbook / drawing children nobody models) are re-serialized at
//   read time; every relationship-id attribute inside is replaced by a placeholder
//   "\u0001<k>\u0001" (k = index into the fragment owner's `rels`). U+0001 cannot occur in XML
//   1.0, so the writer resolves placeholders to freshly allocated rIds with one regex replace.
// - The writer decides output names (`PPart.out`) once per write, renaming a part only when its
//   path collides with one the writer generates, and writes every part reachable from the live
//   model (parts only reachable from a removed sheet are dropped).
//
// - The writer side (../write/preserve.ts) is reached through `w` on the kept data: readXlsx
//   attaches it, so bundles that never read a file (write-only) never include it.
//
// Internal: the shapes below are not public API (wb.preserved is opaque to users).

import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import type { Table } from '../features/tables.js';
import type { PackageParts } from '../write/package.js';

/** Writer-side preservation (../write/preserve.ts `presWrite`), attached by the reader. */
export interface PresWrite {
  /** Name the reachable kept parts for this write (before the first sheet). */
  plan(wb: Workbook, pkg: PackageParts): PPart[] | undefined;
  /** workbook.xml sheetIds. */
  ids(wb: Workbook): number[];
  /**
   * Merge a sheet's kept fragments of slots [lo, hi] into the emitted elements `out` (by slot)
   * and sheetPr children `pr` ([order, xml]).
   */
  frags(
    ws: Worksheet,
    path: string,
    add: AddRel,
    lo: number,
    hi: number,
    out: string[],
    pr: [number, string][],
  ): void;
  /** Add a sheet's implicit relationships (after its tail). */
  implicit(ws: Worksheet, path: string, add: AddRel): void;
  /** Write the planned parts; workbook.xml pieces. */
  book(
    wb: Workbook,
    pkg: PackageParts,
    parts: readonly PPart[],
    add: AddRel,
    ids: readonly number[],
  ): {
    readonly frags: [number, string, number][];
    /** Kept non-worksheet `<sheet>`s by position (see BookInfo.extraSheets). */
    readonly sheets: string[];
    /** Worksheet index → position in `<sheets>`. */
    readonly at: number[] | undefined;
    readonly root: PRel[];
  };
}

/** A preserved zip part. */
export interface PPart {
  /** Original path (no leading slash). */
  readonly path: string;
  /** Content-type Override of the part, if it had one. */
  readonly ct: string | undefined;
  /** Content-type Default of its extension (when it had no Override). */
  readonly def: string | undefined;
  /** Raw bytes (lazy). */
  readonly data: () => Uint8Array;
  /** The part's own relationships (written back with their original ids). */
  readonly rels: readonly POwnRel[];
  /** Output path, decided per write. */
  out?: string | undefined;
}

/** A relationship from a preserved part (original id kept). */
export interface POwnRel {
  readonly id: string;
  readonly type: string;
  /** Preserved part, or an external target (URL) kept verbatim. */
  readonly target: PPart | string;
}

/** A relationship from a part we regenerate (sheet, workbook, drawing): fresh id per write. */
export interface PRel {
  readonly type: string;
  /** Preserved part, or an external target (URL). */
  readonly target: PPart | string;
}

/** A worksheet / workbook child element, raw (placeholders index the owner's `rels`). */
export interface PFrag {
  /** SheetSlot / BookSlot of the element (container slots: sheetPr, extLst → a child). */
  readonly slot: number;
  /** Order inside a container slot. */
  readonly order: number;
  readonly xml: string;
  /** True when the fragment holds relationship placeholders. */
  readonly linked: boolean;
}

/** Non-picture anchors of a sheet's drawing, merged into the drawing the images writer emits. */
export interface PDrawing {
  readonly xml: string;
  readonly rels: readonly PRel[];
  /** Highest cNvPr id among the kept anchors (new pictures are numbered above it). */
  readonly maxId: number;
}

/** Non-note shapes of a legacy VML that also held notes: merged into the regenerated VML. */
export interface PVml {
  /** Shape types + shapes, raw (VML namespaces v / o / x as on the regenerated root). */
  readonly xml: string;
  /** `o:idmap` shape-id blocks the shapes use (kept: form controls refer to shape ids). */
  readonly blocks: readonly number[];
}

/**
 * A modelled table whose part has relationships of its own (a query table): while the model
 * table is unchanged, its original part is written instead of a regenerated one.
 */
export interface PTable {
  readonly table: Table;
  /** JSON of `table` when read. */
  readonly key: string;
  /** Sheet → table part relationship. */
  readonly rel: PRel;
  /** The part's table `id` (kept unique against generated tables). */
  readonly id: number;
}

/** A hyperlink of a formula cell, kept raw (see SheetCapture.formulaLink). */
export interface PLink {
  /** Single cell, "A1". */
  readonly ref: string;
  /** The cell's Formula object when read: the link is written while the cell still holds it. */
  readonly formula: unknown;
  /** `<hyperlink .../>` with relationship placeholders. */
  readonly xml: string;
}

export interface PSheet {
  /** Sheet name when read: renaming the sheet drops its linked preserved content. */
  readonly name: string;
  /** sheetId in the source workbook.xml (kept: slicer caches / pivots refer to it). */
  readonly sheetId: number;
  readonly frags: readonly PFrag[];
  /** Fragment placeholder targets + implicit relationships (see `implicit`). */
  readonly rels: readonly PRel[];
  /** Indices into `rels` of relationships no element references (pivotTable, threadedComments). */
  readonly implicit: readonly number[];
  /** Raw `<sheetPr>` attributes (codeName, filterMode, ...), leading space. */
  readonly prAttrs: string;
  readonly drawing: PDrawing | undefined;
  readonly vml?: PVml | undefined;
  readonly tables?: readonly PTable[] | undefined;
  readonly links?: readonly PLink[] | undefined;
  /** Writer hooks. */
  readonly w: PresWrite;
}

/** Kept formula-cell hyperlinks whose cell still holds the formula it had when read. */
export function liveLinks(ws: Worksheet): PLink[] {
  return ws.preserved?.links?.filter((l) => ws.getValue(l.ref) === l.formula) ?? [];
}

/**
 * A non-worksheet sheet (chartsheet, dialogsheet, macrosheet) kept in `<sheets>`. Tab order:
 * written right after the worksheet it followed in the file (`prev`), or, when that one is gone,
 * right before the one it preceded (`next`), else last.
 */
export interface PForeignSheet {
  readonly name: string;
  readonly sheetId: number;
  readonly state: string;
  readonly rel: PRel;
  readonly prev: Worksheet | undefined;
  readonly next: Worksheet | undefined;
}

export interface PBook {
  /** workbook.xml children nobody models. */
  readonly frags: readonly PFrag[];
  readonly rels: readonly PRel[];
  readonly implicit: readonly number[];
  /** Root (_rels/.rels) relationships besides workbook / core / app properties. */
  readonly root: readonly PRel[];
  readonly sheets: readonly PForeignSheet[];
  /** Raw `<workbookPr>` / `<calcPr>` attributes (leading space; date1904 / fullCalcOnLoad excluded). */
  readonly prAttrs: string;
  readonly calcAttrs: string | undefined;
  /** Main part content type when it must be kept (macro-enabled workbooks). */
  readonly mainCt: string | undefined;
  /** styles.xml pieces: raw `<tableStyles>` (dxfIds remapped) and `<colors>` + `<extLst>`. */
  readonly tableStyles: string | undefined;
  readonly stylesTail: string;
  /** Named cell styles: raw `<cellStyleXfs>` (ids remapped) and `<cellStyles>`. */
  readonly named: readonly [string, string] | undefined;
  /** Part-family numbers taken by preserved content (e.g. "vmlShapeBlock1"). */
  readonly reserve: readonly string[];
  /** Writer hooks. */
  readonly w: PresWrite;
}

/** Writer hooks of a workbook read with preserveUnknown (undefined: nothing kept). */
export function presOf(wb: Workbook): PresWrite | undefined {
  return wb.preserved?.w ?? wb.worksheets.find((s) => s.preserved)?.preserved?.w;
}

/** Path relative to the directory of `from` ("xl/worksheets/sheet1.xml", "xl/charts/c.xml" → "../charts/c.xml"). */
export function relTarget(from: string, to: string): string {
  const a = from.split('/');
  a.pop();
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}

/** Relationship placeholders "\x01<k>\x01". */
// biome-ignore lint/suspicious/noControlCharactersInRegex: U+0001 delimits placeholders
export const PH: RegExp = /\x01(\d+)\x01/g;

export type AddRel = (type: string, target: string, external?: boolean) => string;

/** Add one preserved relationship from part `from`; returns its new id. */
export function addPRel(r: PRel, from: string, add: AddRel): string {
  const t = r.target;
  return typeof t === 'string' ? add(r.type, t, true) : add(r.type, relTarget(from, t.out!));
}

/** Resolve placeholders: each distinct relationship gets one new id from `add`. */
export function fillRels(xml: string, rels: readonly PRel[], from: string, add: AddRel): string {
  const ids: string[] = [];
  return xml.replace(PH, (_, k: string) => {
    ids[+k] ??= addPRel(rels[+k]!, from, add);
    return ids[+k]!;
  });
}

/** Kept non-note VML shapes of a sheet, unless the sheet was renamed since it was read. */
export function presVml(ws: Worksheet): PVml | undefined {
  const p = ws.preserved;
  return p && p.name === ws.name ? p.vml : undefined;
}

/** The kept part of a table read from the file, while the model table is unchanged. */
export function presTable(ws: Worksheet, t: Table): PTable | undefined {
  const k = ws.preserved?.tables?.find((x) => x.table === t);
  return k && JSON.stringify(t) === k.key ? k : undefined;
}

/** Kept drawing anchors of a sheet, unless the sheet was renamed since it was read. */
export function presDrawing(ws: Worksheet): PDrawing | undefined {
  const p = ws.preserved;
  return p && p.name === ws.name ? p.drawing : undefined;
}
