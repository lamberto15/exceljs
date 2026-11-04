// Feature: pictures — workbook media (xl/media/imageN.ext) placed on sheets through a DrawingML
// part per sheet (xl/drawings/drawingN.xml, <drawing r:id>), plus the sheet background
// (<picture r:id>).
//
//   const id = addImage(wb, { data: pngBytes, extension: 'png' });   // or { base64, extension }
//   placeImage(ws, id, 'B2:D6');                                      // covers B2..D6
//   placeImage(ws, id, { tl: 'F2', ext: { width: 120, height: 80 } }, { hyperlink: 'https://x' });
//   placeImage(ws, id, { from: { col: 1, row: 8, colOff: 10 }, to: { col: 4, row: 12 } });
//   ws.backgroundImage = id;
//
// Units: cells are 0-based ({ col: 1, row: 1 } = B2); offsets, sizes and absolute positions are
// pixels (1 px = 9525 EMU; values read from files may be fractional).
// `wb.images` holds each picture once (addImage dedupes identical bytes; the reader dedupes
// identical media across sheets); every entry is written as one media part, referenced by any
// number of placements on any sheets. Charts, shapes, group shapes, linked pictures inside a
// drawing are not modelled: readXlsx keeps them raw (ws.preserved, see read/preserve.ts) and
// this writer appends them after the pictures, renumbering their relationships (until the
// sheet is renamed).

import { decodeAddr, decodeRange } from '../../core/address.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, parseXml } from '../../xml/index.js';
import { type FeatureReader, type FeatureWriter, SheetSlot, type SheetWriteCtx } from './types.js';
import { attrs, escapeAttr, fillRels, num, presDrawing, unbase64 } from './util.js';

export type ImageExtension =
  | 'png'
  | 'jpeg'
  | 'jpg'
  | 'gif'
  | 'bmp'
  | 'tiff'
  | 'svg'
  | 'emf'
  | 'wmf'
  | 'webp';

/** A picture stored once per workbook (`wb.images`). */
export interface WorkbookImage {
  /** Id used by placements (`ImagePlacement.imageId`, `ws.backgroundImage`). */
  id: number;
  data: Uint8Array;
  /** File extension, lower case (decides the media part name and content type). */
  extension: ImageExtension | (string & {});
}

/** Cell-relative position: 0-based cell + offset into it in pixels. */
export interface CellPosition {
  col: number;
  row: number;
  colOff?: number | undefined;
  rowOff?: number | undefined;
}

export interface ImageSize {
  /** Pixels. */
  width: number;
  height: number;
}

/**
 * Where a picture sits: `{ from, to }` = two-cell anchor (moves/resizes with cells per
 * `editAs`), `{ from, ext }` = one-cell anchor (fixed size), `{ pos, ext }` = absolute.
 */
export type ImageAnchor =
  | { from: CellPosition; to: CellPosition }
  | { from: CellPosition; ext: ImageSize }
  | { pos: { x: number; y: number }; ext: ImageSize };

/** A picture placed on a sheet (`ws.images`). */
export interface ImagePlacement {
  imageId: number;
  anchor: ImageAnchor;
  /** Two-cell anchors only: how the picture follows cell moves/resizes (default twoCell). */
  editAs?: 'oneCell' | 'twoCell' | 'absolute' | undefined;
  /** Click target (URL, or "#Sheet2!A1" for an in-workbook location). */
  hyperlink?: string | undefined;
  tooltip?: string | undefined;
  /** Alt text. */
  description?: string | undefined;
  /** Shape name (default "Picture N"). */
  name?: string | undefined;
}

export interface PlaceImageOptions {
  editAs?: ImagePlacement['editAs'];
  hyperlink?: string | undefined;
  tooltip?: string | undefined;
  description?: string | undefined;
  name?: string | undefined;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Find or add a picture in `wb.images` (identical bytes + extension are stored once). */
function intern(wb: Workbook, data: Uint8Array, extension: string): number {
  wb.images ??= [];
  const list = wb.images;
  let max = -1;
  for (const im of list) {
    if (im.extension === extension && sameBytes(im.data, data)) return im.id;
    if (im.id > max) max = im.id;
  }
  list.push({ id: max + 1, data, extension });
  return max + 1;
}

/**
 * Add a picture to the workbook; returns its id. `base64` may be a data URL. Identical pictures
 * are stored once (the existing id is returned).
 */
export function addImage(
  wb: Workbook,
  img: {
    data?: Uint8Array | undefined;
    base64?: string | undefined;
    extension: WorkbookImage['extension'];
  },
): number {
  const data =
    img.data ??
    (img.base64 !== undefined ? unbase64(img.base64.replace(/^data:[^,]*,/, '')) : undefined);
  if (!data) throw new TypeError('addImage: data or base64 is required');
  return intern(wb, data, img.extension.toLowerCase());
}

function cellPos(v: string | CellPosition): CellPosition {
  return typeof v === 'string' ? decodeAddr(v) : v;
}

/**
 * Place a workbook picture on a sheet. `anchor`: a range ("B2:D6" — the picture covers those
 * cells; "B2" covers one cell), `{ tl, ext }` (top-left cell + pixel size), or an
 * {@link ImageAnchor}. Returns the stored placement.
 */
export function placeImage(
  ws: Worksheet,
  imageId: number,
  anchor: string | ImageAnchor | { tl: string | CellPosition; ext: ImageSize },
  opts: PlaceImageOptions = {},
): ImagePlacement {
  let a: ImageAnchor;
  if (typeof anchor === 'string') {
    const r = decodeRange(anchor);
    a = { from: { col: r.left, row: r.top }, to: { col: r.right + 1, row: r.bottom + 1 } };
  } else if ('tl' in anchor) a = { from: cellPos(anchor.tl), ext: anchor.ext };
  else a = anchor;
  const p: ImagePlacement = { imageId, anchor: a };
  for (const k of ['editAs', 'hyperlink', 'tooltip', 'description', 'name'] as const) {
    if (opts[k] !== undefined) (p as unknown as Record<string, unknown>)[k] = opts[k];
  }
  ws.images ??= [];
  ws.images.push(p);
  return p;
}

// ---------------------------------------------------------------- writing

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const MIME: Record<string, string> = {
  jpg: 'image/jpeg',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const emu = (px: number | undefined): number => Math.round((px ?? 0) * 9525);

/** Media part name of a workbook image (by position in wb.images), relative to xl/. */
function mediaOf(wb: Workbook, id: number): string {
  const list = wb.images ?? [];
  for (let i = 0; i < list.length; i++) {
    if (list[i]!.id === id) return `media/image${i + 1}.${list[i]!.extension}`;
  }
  throw new Error(`Image ${id} not found in workbook.images`);
}

function cellXml(tag: string, p: CellPosition): string {
  return `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${emu(p.colOff)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${emu(p.rowOff)}</xdr:rowOff></xdr:${tag}>`;
}

/**
 * The sheet's drawing part: its pictures, then the anchors kept from the file it was read from
 * (charts, shapes, ...: see read/preserve.ts) with their relationships renumbered.
 */
function drawing(c: SheetWriteCtx, list: readonly ImagePlacement[]): string {
  const wb = c.workbook;
  const n = c.nextIndex('drawing');
  const kept = presDrawing(c.ws);
  const base = kept?.maxId ?? 0;
  let rels = '';
  let rid = 0;
  const media = new Map<number, string>();
  const rel = (type: string, target: string, ext?: boolean): string => {
    const id = `rId${++rid}`;
    rels += `<Relationship Id="${id}" Type="${type.includes(':') ? '' : `${NS_R}/`}${escapeAttr(type)}" Target="${escapeAttr(target)}"${ext ? ' TargetMode="External"' : ''}/>`;
    return id;
  };
  let body = '';
  for (let i = 0; i < list.length; i++) {
    const p = list[i]!;
    let embed = media.get(p.imageId);
    if (embed === undefined) {
      embed = rel('image', `../${mediaOf(wb, p.imageId)}`, false);
      media.set(p.imageId, embed);
    }
    const a = p.anchor;
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
    let link = '';
    if (p.hyperlink) {
      const internal = p.hyperlink.charCodeAt(0) === 35;
      link = `<a:hlinkClick r:id="${rel('hyperlink', p.hyperlink, !internal)}"${attrs({ tooltip: p.tooltip })}/>`;
    }
    body += `<xdr:${tag}${tag === 'twoCellAnchor' ? attrs({ editAs: p.editAs }) : ''}>${pos}<xdr:pic><xdr:nvPicPr><xdr:cNvPr${attrs({ id: base + i + 2, name: p.name ?? `Picture ${i + 1}`, descr: p.description })}>${link}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${embed}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:${tag}>`;
  }
  const path = `xl/drawings/drawing${n}.xml`;
  if (kept) body += fillRels(kept.xml, kept.rels, path, rel);
  c.addPart(
    path,
    `${XML_DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${body}</xdr:wsDr>`,
    CT_DRAWING,
  );
  c.addPart(
    `xl/drawings/_rels/drawing${n}.xml.rels`,
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
  );
  return `<drawing r:id="${c.addRel('drawing', `../drawings/drawing${n}.xml`)}"/>`;
}

export const imagesWriter: FeatureWriter = {
  name: 'images',
  sheet: {
    emitters: [
      {
        slot: SheetSlot.drawing,
        emit: (c) =>
          c.ws.images?.length || presDrawing(c.ws) ? drawing(c, c.ws.images ?? []) : '',
      },
      {
        slot: SheetSlot.picture,
        emit: (c) =>
          c.ws.backgroundImage === undefined
            ? ''
            : `<picture r:id="${c.addRel('image', `../${mediaOf(c.workbook, c.ws.backgroundImage)}`)}"/>`,
      },
    ],
  },
  book: {
    // every USED workbook picture once, however many sheets place it; media nobody places (nor
    // uses as a background) is skipped. Part names keep the position in wb.images, so they are
    // stable whatever the other sheets use.
    write(c) {
      const list = c.workbook.images ?? [];
      const used = new Set<number>();
      for (const ws of c.workbook.worksheets) {
        for (const p of ws.images ?? []) used.add(p.imageId);
        if (ws.backgroundImage !== undefined) used.add(ws.backgroundImage);
      }
      for (let i = 0; i < list.length; i++) {
        const im = list[i]!;
        if (!used.has(im.id)) continue;
        const ext = im.extension;
        c.addPart(`xl/media/image${i + 1}.${ext}`, im.data);
        c.addDefault(ext, MIME[ext] ?? `image/${ext}`);
      }
    },
  },
};

// ---------------------------------------------------------------- reading

/** Per workbook being read: media part path → image id. */
const mediaIds = /* @__PURE__ */ new WeakMap<Workbook, Map<string, number>>();

function imageFor(
  wb: Workbook,
  path: string,
  bytes: (p: string) => Uint8Array | undefined,
): number | undefined {
  let m = mediaIds.get(wb);
  if (!m) {
    m = new Map();
    mediaIds.set(wb, m);
  }
  const hit = m.get(path);
  if (hit !== undefined) return hit;
  const data = bytes(path);
  if (!data) return undefined;
  const dot = path.lastIndexOf('.');
  const id = intern(wb, data, dot < 0 ? 'png' : path.slice(dot + 1).toLowerCase());
  m.set(path, id);
  return id;
}

function relKind(type: string): string {
  return type.slice(type.lastIndexOf('/') + 1);
}

const px = (v: string | undefined): number => (num(v) ?? 0) / 9525;

function rid(a: Record<string, string>, local: string): string | undefined {
  for (const k in a) if (k === local || k.endsWith(`:${local}`)) return a[k];
  return undefined;
}

interface RawPic {
  tag: string;
  editAs?: string | undefined;
  from?: CellPosition;
  to?: CellPosition;
  ext?: ImageSize;
  pos?: { x: number; y: number };
  embed?: string | undefined;
  link?: string | undefined;
  tooltip?: string | undefined;
  descr?: string | undefined;
  name?: string | undefined;
}

function parseDrawing(xml: string): RawPic[] {
  const out: RawPic[] = [];
  let cur: RawPic | undefined;
  let depth = 0; // depth below the anchor element
  let pic = false; // inside a direct <pic> child of the anchor
  let cell: CellPosition | undefined;
  let field = '';
  let text = '';
  parseXml(xml, {
    open(name, a) {
      const ln = localName(name);
      if (!cur) {
        if (ln === 'twoCellAnchor' || ln === 'oneCellAnchor' || ln === 'absoluteAnchor') {
          cur = { tag: ln, editAs: a.editAs };
          depth = 0;
        }
        return;
      }
      depth++;
      if (depth === 1) {
        if (ln === 'from' || ln === 'to') cell = { col: 0, row: 0 };
        else if (ln === 'ext') cur.ext = { width: px(a.cx), height: px(a.cy) };
        else if (ln === 'pos') cur.pos = { x: px(a.x), y: px(a.y) };
        else if (ln === 'pic') pic = true;
      } else if (depth === 2 && cell) {
        field = ln;
        text = '';
      } else if (pic) {
        if (ln === 'cNvPr') {
          cur.descr = a.descr;
          cur.name = a.name;
        } else if (ln === 'hlinkClick' && depth === 4) {
          cur.link = rid(a, 'id');
          cur.tooltip = a.tooltip;
        } else if (ln === 'blip') cur.embed = rid(a, 'embed');
      }
    },
    text(t) {
      if (field) text += t;
    },
    close(name) {
      if (!cur) return;
      const ln = localName(name);
      if (depth === 0) {
        if (cur.embed) out.push(cur);
        cur = undefined;
        pic = false;
        return;
      }
      if (depth === 2 && cell && field) {
        const v = +text || 0;
        if (field === 'col') cell.col = v;
        else if (field === 'row') cell.row = v;
        else if (field === 'colOff' && v) cell.colOff = v / 9525;
        else if (field === 'rowOff' && v) cell.rowOff = v / 9525;
        field = '';
      } else if (depth === 1) {
        if (cell && (ln === 'from' || ln === 'to')) cur[ln] = cell;
        cell = undefined;
        if (ln === 'pic') pic = false;
      }
      depth--;
    },
  });
  return out;
}

const DRAWING_KEY = {};

export const imagesReader: FeatureReader = {
  name: 'images',
  sheet: {
    elements: ['drawing', 'picture'],
    open(ctx, el, a) {
      const r = rid(a, 'id');
      if (r) ctx.scratch.set(el === 'drawing' ? DRAWING_KEY : imagesReader, r);
      return undefined;
    },
    finish(ctx) {
      const wb = ctx.workbook;
      const bytes = ctx.readBytes;
      const bg = ctx.rels?.get(ctx.scratch.get(imagesReader) as string);
      if (bg && !bg.external) {
        const id = imageFor(wb, bg.path, bytes);
        if (id !== undefined) ctx.ws.backgroundImage = id;
      }
      const rel = ctx.rels?.get(ctx.scratch.get(DRAWING_KEY) as string);
      if (!rel || rel.external || relKind(rel.type) !== 'drawing') return;
      const xml = ctx.readText(rel.path);
      if (!xml) return;
      const drels = ctx.relsOf(rel.path);
      const list: ImagePlacement[] = [];
      for (const r of parseDrawing(xml)) {
        const m = drels?.get(r.embed!);
        if (!m || m.external) continue; // linked (not embedded) pictures are skipped
        const imageId = imageFor(wb, m.path, bytes);
        if (imageId === undefined) continue;
        let anchor: ImageAnchor | undefined;
        if (r.tag === 'absoluteAnchor') {
          if (r.pos && r.ext) anchor = { pos: r.pos, ext: r.ext };
        } else if (r.from) {
          if (r.tag === 'twoCellAnchor') {
            if (r.to) anchor = { from: r.from, to: r.to };
          } else if (r.ext) anchor = { from: r.from, ext: r.ext };
        }
        if (!anchor) continue;
        const p: ImagePlacement = { imageId, anchor };
        if (r.tag === 'twoCellAnchor' && r.editAs) p.editAs = r.editAs as ImagePlacement['editAs'];
        const h = r.link ? drels?.get(r.link) : undefined;
        if (h) {
          p.hyperlink = h.external ? h.target : `#${h.target.replace(/^#/, '')}`;
          if (r.tooltip) p.tooltip = r.tooltip;
        }
        if (r.descr) p.description = r.descr;
        if (r.name && !/^Picture \d+$/.test(r.name)) p.name = r.name;
        list.push(p);
      }
      if (list.length > 0) ctx.ws.images = list;
    },
  },
};
