// Writer side of round-trip preservation (model: ../common/preserved.ts, captured by
// ../read/preserve.ts). Kept small: it ships in the write-only bundle.
//
// - planPreserved (before the first sheet): collects the preserved parts reachable from the live
//   model, names them (original path unless the writer generates that path: then the trailing
//   number is bumped) and reserves their basenames so `nextIndex` families skip them.
// - sheet fragments are merged into the emitted elements (sheetFrags, from sheet.ts runEmitters)
//   and implicit sheet relationships are added in the tail (sheetImplicit); workbook pieces by
//   bookParts (presBook).
// The writer core reaches all of this only through `presWrite`, which readXlsx attaches to the
// kept data (`wb.preserved.w`, `ws.preserved.w`): write-only bundles do not include this module.

import { type CellRange, decodeRange } from '../../core/address.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { escapeAttr } from '../../xml/index.js';
import {
  type AddRel,
  addPRel,
  fillRels,
  liveLinks,
  type PBook,
  type PForeignSheet,
  type PFrag,
  PH,
  type PPart,
  type PRel,
  type PresWrite,
  presTable,
  relTarget,
} from '../common/preserved.js';
import { SheetSlot } from '../features/types.js';
import { type PackageParts, relsXml } from './package.js';

/**
 * Slots whose kept element (and the parts it links) gives way to model data of the owning
 * feature: the model wins, nothing is written twice. Raw elements only coexist with model
 * fields when a workbook read without a feature (`./lite`) is written with it.
 */
const OWNER: Partial<Record<number, (ws: Worksheet) => unknown>> = {
  [SheetSlot.legacyDrawing]: (ws) => ws.comments?.length,
  [SheetSlot.drawing]: (ws) => ws.images?.length || ws.charts?.length,
  [SheetSlot.picture]: (ws) => ws.backgroundImage !== undefined,
  [SheetSlot.tableParts]: (ws) => ws.tables?.length,
  [SheetSlot.pageSetup]: (ws) => ws.pageSetup,
};

const FILTER_COL = /<((?:\w+:)?filterColumn)\b[^>]*?\scolId="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/\1>)/g;
const SORT_STATE = /<((?:\w+:)?sortState)\b(?:[^>]*\/>|[\s\S]*?<\/\1>)/;

/**
 * Kept `<autoFilter>` with its criteria for the model's range `ref`. When the range changed,
 * filter columns are kept only if the range still starts in the same column and is wide enough
 * (`colId` is relative to it); the sort state (absolute ranges) is dropped.
 */
function filterXml(kept: string, ref: string): string {
  const old = / ref="([^"]*)"/.exec(kept)?.[1];
  if (old === ref) return kept;
  const bare = `<autoFilter ref="${escapeAttr(ref)}"/>`;
  let a: CellRange;
  let b: CellRange;
  try {
    a = decodeRange(old ?? '');
    b = decodeRange(ref);
  } catch {
    return bare;
  }
  if (a.left !== b.left) return bare;
  return kept
    .replace(/ ref="[^"]*"/, ` ref="${escapeAttr(ref)}"`)
    .replace(FILTER_COL, (x, _t, id: string) => (+id <= b.right - b.left ? x : ''))
    .replace(SORT_STATE, '');
}

/** First `uri` attribute of a kept `<ext>` element. */
const EXT_URI = /^<[^>]*?\suri="([^"]*)"/;

/**
 * Merge the sheet's kept fragments of slots [lo, hi]: a kept element is written only when no
 * emitter produced its slot; in the container slots, only children no emitter produced (same
 * sheetPr child order, same extLst `uri`).
 */
function sheetFrags(
  ws: Worksheet,
  path: string,
  add: AddRel,
  lo: number,
  hi: number,
  out: string[],
  pr: [number, string][],
): void {
  const ps = ws.preserved!;
  for (const f of ps.frags) {
    const k = f.slot;
    if (k < lo || k > hi || !fragLive(f, ws)) continue;
    if (k === SheetSlot.sheetFormatPr) {
      // kept attributes the model does not cover (zeroHeight, thickTop, baseColWidth, ...),
      // added to the built-in element (ws.properties covers the rest)
      out[k] = (out[k] ?? '<sheetFormatPr/>').replace(/\/>$/, f.xml.slice(14));
      continue;
    }
    if (k === SheetSlot.autoFilter) {
      // criteria of the model's autoFilter (none when it was removed)
      if (out[k]) out[k] = filterXml(f.xml, ws.autoFilter!);
      continue;
    }
    if (k === SheetSlot.sheetPr) {
      if (f.order !== 9 && pr.some((p) => p[0] === f.order)) continue;
    } else if (k === SheetSlot.extLst) {
      const uri = EXT_URI.exec(f.xml)?.[1];
      if (uri && out[k]?.includes(`uri="${uri}"`)) continue;
    } else if (out[k]) continue;
    const x = fillRels(f.xml, ps.rels, path, add);
    if (k === SheetSlot.sheetPr) pr.push([f.order, x]);
    else out[k] = (out[k] ?? '') + x;
  }
  const h = SheetSlot.hyperlinks;
  const links = h >= lo && h <= hi ? liveLinks(ws) : [];
  if (links.length) {
    const x = fillRels(links.map((l) => l.xml).join(''), ps.rels, path, add);
    const cur = out[h];
    out[h] = cur
      ? cur.replace('</hyperlinks>', `${x}</hyperlinks>`)
      : `<hyperlinks>${x}</hyperlinks>`;
  }
}

/** Whether a sheet fragment is written (not after a rename when linked; not when the model owns its slot). */
export function fragLive(f: PFrag, ws: Worksheet): boolean {
  return !(f.linked && ws.preserved!.name !== ws.name) && !OWNER[f.slot]?.(ws);
}

const COMMENTS = /\/comments$/;

/**
 * Implicit sheet relationships still valid: raw legacy notes give way to modelled ones;
 * threaded comments need legacy notes (modelled or raw).
 */
function implicitRels(ws: Worksheet): PRel[] {
  const p = ws.preserved;
  if (!p || p.name !== ws.name) return [];
  const notes = !!ws.comments?.length;
  const raw = !notes && p.implicit.some((k) => COMMENTS.test(p.rels[k]!.type));
  const out: PRel[] = [];
  for (const k of p.implicit) {
    const r = p.rels[k]!;
    if (COMMENTS.test(r.type) ? !notes : !/threadedComment/i.test(r.type) || notes || raw)
      out.push(r);
  }
  return out;
}

/** Add the sheet's implicit relationships (pivot tables, threaded comments, ...). */
function sheetImplicit(ws: Worksheet, path: string, add: AddRel): void {
  for (const r of implicitRels(ws)) addPRel(r, path, add);
}

/**
 * workbook.xml sheetIds: the ids read from the file where still unique (other parts refer to
 * them), else the worksheet's own id or a fresh one.
 */
function sheetIds(wb: Workbook): number[] {
  const used = new Set<number>();
  const sheets = wb.worksheets;
  const out = sheets.map((ws) => {
    const id = ws.preserved?.sheetId;
    if (id === undefined || used.has(id)) return 0;
    used.add(id);
    return id;
  });
  let next = Math.max(0, ...used, ...sheets.map((w) => w.id));
  return out.map((id, i) => {
    if (id) return id;
    const own = sheets[i]!.id;
    const v = used.has(own) ? ++next : own;
    used.add(v);
    return v;
  });
}

/** Foreign sheets (chartsheets, ...) whose name no worksheet took meanwhile. */
function foreign(wb: Workbook, pb: PBook): PForeignSheet[] {
  const names = new Set(wb.worksheets.map((w) => w.name.toLowerCase()));
  return pb.sheets.filter((s) => !names.has(s.name.toLowerCase()));
}

/** Preserved parts reachable from the live model (what the next write would emit). */
function reachable(wb: Workbook, pkg?: PackageParts): PPart[] {
  const pb = wb.preserved;
  const list: PPart[] = [];
  const seen = new Set<PPart>();
  const visit = (t: PPart | string): void => {
    if (typeof t === 'string' || seen.has(t)) return;
    seen.add(t);
    list.push(t);
    for (const r of t.rels) visit(r.target);
  };
  const used = (xml: string, rels: readonly PRel[]): void => {
    for (const m of xml.matchAll(PH)) visit(rels[+m[1]!]!.target);
  };
  if (pb) {
    for (const f of pb.frags) used(f.xml, pb.rels);
    for (const k of pb.implicit) visit(pb.rels[k]!.target);
    for (const r of pb.root) visit(r.target);
    for (const s of foreign(wb, pb)) visit(s.rel.target);
    if (pkg) for (const r of pb.reserve) pkg.reserve(r);
  }
  for (const ws of wb.worksheets) {
    const p = ws.preserved;
    if (!p) continue;
    for (const f of p.frags) if (fragLive(f, ws)) used(f.xml, p.rels);
    for (const l of liveLinks(ws)) used(l.xml, p.rels);
    for (const r of implicitRels(ws)) visit(r.target);
    if (p.drawing && p.name === ws.name) used(p.drawing.xml, p.drawing.rels);
    for (const t of ws.tables ?? []) {
      const k = presTable(ws, t);
      if (!k) continue;
      visit(k.rel.target);
      pkg?.reserve(`table${k.id}`); // generated tables number their ids around it
    }
  }
  return list;
}

/** A package part kept from the file a workbook was read from (see `listPreservedParts`). */
export interface PreservedPart {
  /** Zip path in the source file. */
  readonly path: string;
  readonly contentType: string | undefined;
}

/**
 * Parts `readXlsx` kept because the model does not represent them (charts, pivot tables,
 * threaded comments, VBA, themes, ...) and that the next `writeXlsx` / `xlsxStream` would write
 * back: parts reachable only from removed or renamed sheets are not listed (they are dropped).
 * Worksheet XML elements kept in place (rowBreaks, ignoredErrors, ...) are not parts.
 */
export function listPreservedParts(wb: Workbook): PreservedPart[] {
  return reachable(wb).map((p) => ({ path: p.path, contentType: p.ct ?? p.def }));
}

/** Name the reachable preserved parts for this write; undefined when there is nothing kept. */
function planPreserved(wb: Workbook, pkg: PackageParts): PPart[] | undefined {
  if (!wb.preserved && !wb.worksheets.some((w) => w.preserved)) return undefined;
  const list = reachable(wb, pkg);
  // names the writer generates outside nextIndex families (workbook-level parts, styles, docProps
  // are owned by the model, never kept)
  const taken = new Set<string>();
  for (let i = 0; i < wb.worksheets.length; i++) taken.add(`xl/worksheets/sheet${i + 1}.xml`);
  const imgs = wb.images ?? [];
  for (let i = 0; i < imgs.length; i++) taken.add(`xl/media/image${i + 1}.${imgs[i]!.extension}`);
  for (const p of list) {
    let out = p.path;
    const m = /^(.*?)\d*(\.[^./]*)?$/.exec(out)!;
    for (let k = 1; taken.has(out.toLowerCase()); k++) out = `${m[1]}${k}${m[2] ?? ''}`;
    taken.add(out.toLowerCase());
    p.out = out;
    pkg.reserve(out.slice(out.lastIndexOf('/') + 1).replace(/\.[^.]*$/, ''));
  }
  return list;
}

/** Write the planned parts (+ their rels and content types); workbook.xml pieces. */
function presBook(
  wb: Workbook,
  pkg: PackageParts,
  parts: readonly PPart[],
  add: AddRel,
  ids: readonly number[],
): ReturnType<PresWrite['book']> {
  for (const p of parts) {
    const out = p.out!;
    pkg.addPart(out, p.data(), p.ct);
    if (!p.ct && p.def) pkg.addDefault(out.slice(out.lastIndexOf('.') + 1), p.def);
    if (p.rels.length > 0) {
      pkg.addPart(
        out.replace(/[^/]*$/, '_rels/$&.rels'),
        relsXml(
          p.rels.map((r) => {
            const t = r.target;
            const external = typeof t === 'string';
            return { ...r, target: external ? t : relTarget(out, t.out!), external };
          }),
        ),
      );
    }
  }
  const pb = wb.preserved;
  const frags: [number, string, number][] = [];
  const ws = wb.worksheets;
  const n = ws.length;
  const sheets: string[] = [];
  let at: number[] | undefined;
  if (pb) {
    const from = 'xl/workbook.xml';
    for (const f of pb.frags) frags.push([f.slot, fillRels(f.xml, pb.rels, from, add), f.order]);
    for (const k of pb.implicit) addPRel(pb.rels[k]!, from, add);
    const used = new Set(ids);
    let next = Math.max(0, ...ids, ...pb.sheets.map((s) => s.sheetId));
    const count: number[] = [];
    for (const s of foreign(wb, pb)) {
      const id = used.has(s.sheetId) ? ++next : s.sheetId;
      used.add(id);
      // original neighbours: after the previous worksheet, else before the next one, else last
      const p = s.prev ? ws.indexOf(s.prev) : -1;
      const q = s.next ? ws.indexOf(s.next) : -1;
      const pos = p >= 0 ? p + 1 : q >= 0 ? q : n;
      count[pos] = (count[pos] ?? 0) + 1;
      sheets[pos] = `${sheets[pos] ?? ''}<sheet name="${escapeAttr(s.name)}" sheetId="${id}"${
        s.state !== 'visible' ? ` state="${s.state}"` : ''
      } r:id="${addPRel(s.rel, from, add)}"/>`;
    }
    if (count.length) {
      at = [];
      let k = 0;
      for (let i = 0; i < n; i++) {
        k += count[i] ?? 0;
        at.push(i + k);
      }
    }
    if (pb.mainCt) pkg.overrides.push(['xl/workbook.xml', pb.mainCt]);
  }
  return { frags, sheets, at, root: pb ? [...pb.root] : [] };
}

/** The writer-side hooks (attached to kept data by readXlsx; see the header). */
export const presWrite: PresWrite = {
  plan: planPreserved,
  ids: sheetIds,
  frags: sheetFrags,
  implicit: sheetImplicit,
  book: presBook,
};
