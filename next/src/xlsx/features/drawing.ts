// The sheet's drawing part (xl/drawings/drawingN.xml, `<drawing r:id>`), shared by every feature
// that places anchors on a sheet (images, charts) and by the anchors readXlsx kept from the file
// (ws.preserved.drawing: charts, shapes, ... nobody models). Features contribute through
// `FeatureWriter.sheet.anchors` (types.ts DrawingAnchors) and list `drawingEmitter` in their
// emitters; the writer runs the emitter once per sheet however many features list it.

import { decodeAddr, decodeRange } from '../../core/address.js';
import { escapeAttr } from '../../xml/index.js';
import type { CellPosition, ImageAnchor, ImageSize } from './images.js';
import type { DrawingWriteCtx, SheetEmitter, SheetWriteCtx } from './types.js';
import { SheetSlot } from './types.js';
import { attrs, fillRels, presDrawing } from './util.js';

/**
 * Anchor as users pass it (placeImage, addChart): a range ("B2:D6" covers those cells), a
 * top-left cell + pixel size, or an explicit {@link ImageAnchor}.
 */
export type AnchorInput = string | ImageAnchor | { tl: string | CellPosition; ext: ImageSize };

/** Normalize an {@link AnchorInput}. */
export function toAnchor(anchor: AnchorInput): ImageAnchor {
  if (typeof anchor === 'string') {
    const r = decodeRange(anchor);
    return { from: { col: r.left, row: r.top }, to: { col: r.right + 1, row: r.bottom + 1 } };
  }
  if ('tl' in anchor) {
    const tl = anchor.tl;
    return { from: typeof tl === 'string' ? decodeAddr(tl) : tl, ext: anchor.ext };
  }
  return anchor;
}

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
/** XML declaration of generated parts. */
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Pixels → EMU (1 px = 9525 EMU). */
export const emu = (px: number | undefined): number => Math.round((px ?? 0) * 9525);

/** Package relationships part. */
export function relsPart(rels: string): string {
  return `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

/** One `<Relationship>` element (`type`: short officeDocument name or full URI). */
export function relElem(id: string, type: string, target: string, ext?: boolean): string {
  return `<Relationship Id="${id}" Type="${type.includes(':') ? '' : `${NS_R}/`}${escapeAttr(type)}" Target="${escapeAttr(target)}"${ext ? ' TargetMode="External"' : ''}/>`;
}

function cellXml(tag: string, p: CellPosition): string {
  return `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${emu(p.colOff)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${emu(p.rowOff)}</xdr:rowOff></xdr:${tag}>`;
}

/**
 * Anchor element around `content` (a pic / graphicFrame): `<xdr:twoCellAnchor>` (from + to),
 * `<xdr:oneCellAnchor>` (from + ext) or `<xdr:absoluteAnchor>` (pos + ext).
 */
export function anchorXml(a: ImageAnchor, editAs: string | undefined, content: string): string {
  let tag: string;
  let pos: string;
  if ('pos' in a) {
    tag = 'absoluteAnchor';
    pos = `<xdr:pos x="${emu(a.pos.x)}" y="${emu(a.pos.y)}"/>`;
  } else {
    tag = 'to' in a ? 'twoCellAnchor' : 'oneCellAnchor';
    pos = cellXml('from', a.from);
  }
  if ('to' in a) pos += cellXml('to', a.to);
  else pos += `<xdr:ext cx="${emu(a.ext.width)}" cy="${emu(a.ext.height)}"/>`;
  return `<xdr:${tag}${tag === 'twoCellAnchor' ? attrs({ editAs }) : ''}>${pos}${content}<xdr:clientData/></xdr:${tag}>`;
}

function drawing(c: SheetWriteCtx): string {
  const n = c.nextIndex('drawing');
  const path = `xl/drawings/drawing${n}.xml`;
  const kept = presDrawing(c.ws);
  let shape = (kept?.maxId ?? 0) + 1;
  let rels = '';
  let rid = 0;
  const rel = (type: string, target: string, ext?: boolean): string => {
    const id = `rId${++rid}`;
    rels += relElem(id, type, target, ext);
    return id;
  };
  const d: DrawingWriteCtx = { ...c, drawingPath: path, drawingRel: rel, shapeId: () => ++shape };
  let body = '';
  for (const a of c.anchors) if (a.has(c.ws)) body += a.emit(d);
  if (kept) body += fillRels(kept.xml, kept.rels, path, rel);
  c.addPart(
    path,
    `${XML_DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${body}</xdr:wsDr>`,
    CT_DRAWING,
  );
  c.addPart(`xl/drawings/_rels/drawing${n}.xml.rels`, relsPart(rels));
  return `<drawing r:id="${c.addRel('drawing', `../drawings/drawing${n}.xml`)}"/>`;
}

/**
 * The `<drawing>` worksheet element and its part: every enabled feature's anchors (in `order`),
 * then the anchors kept from the file the sheet was read from (relationships renumbered).
 */
export const drawingEmitter: SheetEmitter = {
  slot: SheetSlot.drawing,
  emit: (c) => (c.anchors.some((a) => a.has(c.ws)) || presDrawing(c.ws) ? drawing(c) : ''),
};
