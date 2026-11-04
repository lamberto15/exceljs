// Writer side of round-trip preservation (model: ../common/preserved.ts, captured by
// ../read/preserve.ts). Kept small: it ships in the write-only bundle.
//
// - planPreserved (before the first sheet): collects the preserved parts reachable from the live
//   model, names them (original path unless the writer generates that path: then the trailing
//   number is bumped) and reserves their basenames so `nextIndex` families skip them.
// - sheet fragments are merged by sheet.ts (runEmitters) and implicit sheet relationships are
//   added in the tail (sheetImplicit); workbook pieces by bookParts (presBook).

import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { escapeAttr } from '../../xml/index.js';
import {
  type AddRel,
  addPRel,
  fillRels,
  type PBook,
  type PForeignSheet,
  type PFrag,
  PH,
  type PPart,
  type PRel,
  relTarget,
} from '../common/preserved.js';
import { SheetSlot } from '../features/types.js';
import { type PackageParts, relsXml } from './package.js';

/** Whether a sheet fragment is written (legacy VML: the comments feature wins when it has notes). */
export function fragLive(f: PFrag, ws: Worksheet): boolean {
  return (
    !(f.linked && ws.preserved!.name !== ws.name) &&
    !(f.slot === SheetSlot.legacyDrawing && ws.comments?.length)
  );
}

/** Implicit sheet relationships still valid (threaded comments need the legacy notes). */
function implicitRels(ws: Worksheet): PRel[] {
  const p = ws.preserved;
  if (!p || p.name !== ws.name) return [];
  const out: PRel[] = [];
  for (const k of p.implicit) {
    const r = p.rels[k]!;
    if (!/threadedComment/i.test(r.type) || ws.comments?.length) out.push(r);
  }
  return out;
}

/** Add the sheet's implicit relationships (pivot tables, threaded comments, ...). */
export function sheetImplicit(ws: Worksheet, path: string, add: AddRel): void {
  for (const r of implicitRels(ws)) addPRel(r, path, add);
}

/**
 * workbook.xml sheetIds: the ids read from the file where still unique (other parts refer to
 * them), else the worksheet's own id or a fresh one.
 */
export function sheetIds(wb: Workbook): number[] {
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
    for (const r of implicitRels(ws)) visit(r.target);
    if (p.drawing && p.name === ws.name) used(p.drawing.xml, p.drawing.rels);
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
export function planPreserved(wb: Workbook, pkg: PackageParts): PPart[] | undefined {
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

export interface PresBookOut {
  /** workbook.xml fragments as [BookSlot, xml, order]. */
  readonly frags: [number, string, number][];
  /** Extra `<sheet>` elements (foreign sheets). */
  readonly sheets: string;
  /** Extra root relationships (to append to _rels/.rels, ids from rId4). */
  readonly root: PRel[];
}

/** Write the planned parts (+ their rels and content types); workbook.xml pieces. */
export function presBook(
  wb: Workbook,
  pkg: PackageParts,
  parts: readonly PPart[],
  add: AddRel,
  ids: readonly number[],
): PresBookOut {
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
  let sheets = '';
  if (pb) {
    const from = 'xl/workbook.xml';
    for (const f of pb.frags) frags.push([f.slot, fillRels(f.xml, pb.rels, from, add), f.order]);
    for (const k of pb.implicit) addPRel(pb.rels[k]!, from, add);
    const used = new Set(ids);
    let next = Math.max(0, ...ids, ...pb.sheets.map((s) => s.sheetId));
    for (const s of foreign(wb, pb)) {
      const id = used.has(s.sheetId) ? ++next : s.sheetId;
      used.add(id);
      sheets += `<sheet name="${escapeAttr(s.name)}" sheetId="${id}"${
        s.state !== 'visible' ? ` state="${s.state}"` : ''
      } r:id="${addPRel(s.rel, from, add)}"/>`;
    }
    if (pb.mainCt) pkg.overrides.push(['xl/workbook.xml', pb.mainCt]);
  }
  return { frags, sheets, root: pb ? [...pb.root] : [] };
}
