// Feature: cell comments ("notes" in Excel 365) — xl/commentsN.xml + the legacy VML drawing
// (xl/drawings/vmlDrawingN.vml) Excel needs to display them, linked from <legacyDrawing r:id>.
//
//   addNote(ws, 'B2', 'Check this');                                   // plain text
//   addNote(ws, 'C3', { text: { richText: [{ text: 'Bold', font: { bold: true } }] },
//                       author: 'Ann', visible: true, width: 200, height: 100 });
//   ws.comments = [{ ref: 'D4', text: 'x' }];                         // or set the list directly
//
// Model: `ws.comments` is a plain array of `Note` (one per cell; `addNote` replaces an existing
// note on the same cell). An array (not a Map) keeps the model plain data / JSON friendly and in
// file order; look one up with `ws.comments?.find((n) => n.ref === 'B2')`.
// Box size is in pixels; the VML anchor is derived from it and the sheet's actual column widths /
// row heights (defaults: 64 x 20 px). On read, visibility and size come from the VML part when present
// (a missing VML part is tolerated). Excel 365 threaded comments (xl/threadedComments) are not
// modelled: Excel also writes a legacy note for each thread ("[Threaded comment] ..."), which
// is what is read here; the threaded part itself is ignored (and not written back).

import { decodeAddr, decodeCol, encodeAddr } from '../../core/address.js';
import type { RichText } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, parseXml } from '../../xml/index.js';
import { type FeatureReader, type FeatureWriter, SheetSlot, type SheetWriteCtx } from './types.js';
import { escapeText, num, runsXml, StringItemBuilder, tElem } from './util.js';

export interface Note {
  /** Cell address, "B2". */
  ref: string;
  /** Plain text or rich text runs. */
  text: string | RichText;
  /** Author name (comments.xml authors list). */
  author?: string | undefined;
  /** Always show the note (default false: shown on hover). */
  visible?: boolean | undefined;
  /** Box width in pixels (default 144). */
  width?: number | undefined;
  /** Box height in pixels (default 79). */
  height?: number | undefined;
}

/** Add (or replace) the note of a cell; returns the stored note. */
export function addNote(
  ws: Worksheet,
  ref: string,
  note: string | RichText | Omit<Note, 'ref'>,
): Note {
  const a = decodeAddr(ref);
  const r = encodeAddr(a.row, a.col);
  const n: Note =
    typeof note === 'string' || 'richText' in note ? { ref: r, text: note } : { ...note, ref: r };
  ws.comments ??= [];
  const list = ws.comments;
  const i = list.findIndex((x) => x.ref === r);
  if (i >= 0) list[i] = n;
  else list.push(n);
  return n;
}

const W = 144;
const H = 79;
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

/** Points with at most 2 decimals. */
const pt = (px: number): string => `${Math.round(px * 75) / 100}pt`;

function commentsXml(notes: readonly Note[], cells: string[]): string {
  const authors: string[] = [];
  let list = '';
  for (let i = 0; i < notes.length; i++) {
    const n = notes[i]!;
    const au = n.author ?? '';
    let id = authors.indexOf(au);
    if (id < 0) id = authors.push(au) - 1;
    const t = n.text;
    // runs even for plain text: exceljs only reads <r> runs
    const body = typeof t === 'string' ? `<r>${tElem(t)}</r>` : runsXml(t.richText);
    list += `<comment ref="${cells[i]}" authorId="${id}"><text>${body}</text></comment>`;
  }
  let a = '';
  for (const au of authors) a += `<author>${escapeText(au)}</author>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${NS_MAIN}"><authors>${a}</authors><commentList>${list}</commentList></comments>`;
}

/** Column width in characters → pixels (Calibri 11, max digit width 7 px; Excel's formula). */
const colWidthPx = (chars: number): number =>
  Math.trunc(((256 * chars + Math.trunc(128 / 7)) / 256) * 7);

const MAX_WALK = 1 << 20;

/** Walk `size` px from offset `off` into item `i`; returns the item and offset where it ends. */
function walk(i: number, size: number, px: (i: number) => number): [number, number] {
  let rest = size;
  for (let n = 0; n < MAX_WALK; n++) {
    const p = px(i);
    if (rest < p) break;
    rest -= p;
    i++;
  }
  return [i, Math.round(rest)];
}

/** Pixels from the sheet origin to the start of item `i`. */
function offset(i: number, px: (i: number) => number): number {
  let s = 0;
  for (let k = 0; k < i; k++) s += px(k);
  return s;
}

function vmlXml(c: SheetWriteCtx, notes: readonly Note[], cells: string[]): string {
  const cols = c.ws.columnData;
  const rows = c.ws.rowData;
  const colPx = (i: number): number => {
    const d = cols[i];
    return d?.hidden ? 0 : colWidthPx(d?.width ?? 9.140625);
  };
  const rowPx = (i: number): number => {
    const d = rows[i];
    return d?.hidden ? 0 : d?.height ? Math.round((d.height * 96) / 72) : 20;
  };
  // shape ids come in blocks of 1024 per <o:idmap> entry, unique within the package
  const blocks: number[] = [];
  for (let k = 0; k * 1023 < notes.length; k++) blocks.push(c.nextIndex('vmlShapeBlock'));
  let s = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${blocks.join(',')}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`;
  for (let i = 0; i < notes.length; i++) {
    const n = notes[i]!;
    const { row, col } = decodeAddr(cells[i]!);
    const w = n.width ?? W;
    const h = n.height ?? H;
    const id = blocks[(i / 1023) | 0]! * 1024 + (i % 1023) + 1;
    const vis = n.visible ? 'visible' : 'hidden';
    // anchor: LeftCol, LeftOff(px), TopRow, TopOff, RightCol, RightOff, BottomRow, BottomOff;
    // the box starts one column right of the cell, one row above it (like Excel's own notes)
    const l = col + 1;
    const t = Math.max(row - 1, 0);
    const [r, ro] = walk(l, 15 + w, colPx);
    const [b, bo] = walk(t, 10 + h, rowPx);
    const anchor = `${l}, 15, ${t}, 10, ${r}, ${ro}, ${b}, ${bo}`;
    s += `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:${pt(offset(l, colPx) + 15)};margin-top:${pt(offset(t, rowPx) + 10)};width:${pt(w)};height:${pt(h)};z-index:${i + 1};visibility:${vis}" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column>${n.visible ? '<x:Visible/>' : ''}</x:ClientData></v:shape>`;
  }
  return `${s}</xml>`;
}

function emit(c: SheetWriteCtx): string {
  const notes = c.ws.comments;
  if (!notes || notes.length === 0) return '';
  const cells = notes.map((n) => {
    const a = decodeAddr(n.ref);
    return encodeAddr(a.row, a.col);
  });
  const n = c.nextIndex('comments');
  const v = c.nextIndex('vmlDrawing');
  c.addPart(`xl/comments${n}.xml`, commentsXml(notes, cells), CT_COMMENTS);
  c.addPart(`xl/drawings/vmlDrawing${v}.vml`, vmlXml(c, notes, cells));
  c.addDefault('vml', CT_VML);
  c.addRel('comments', `../comments${n}.xml`);
  return `<legacyDrawing r:id="${c.addRel('vmlDrawing', `../drawings/vmlDrawing${v}.vml`)}"/>`;
}

export const commentsWriter: FeatureWriter = {
  name: 'comments',
  sheet: { emitters: [{ slot: SheetSlot.legacyDrawing, emit }] },
};

// ---------------------------------------------------------------- reading

function relKind(type: string): string {
  return type.slice(type.lastIndexOf('/') + 1);
}

function readComments(xml: string): Note[] {
  const notes: Note[] = [];
  const authors: string[] = [];
  const b = new StringItemBuilder();
  let au: string | undefined;
  let cur: { ref: string; authorId: number } | undefined;
  let inText = false;
  parseXml(xml, {
    open(name, a) {
      const ln = localName(name);
      if (inText) b.open(ln, a);
      else if (ln === 'author') au = '';
      else if (ln === 'comment' && a.ref) cur = { ref: a.ref, authorId: num(a.authorId) ?? -1 };
      else if (ln === 'text' && cur) {
        inText = true;
        b.reset();
      }
    },
    text(t) {
      if (inText) b.onText(t);
      else if (au !== undefined) au += t;
    },
    close(name) {
      const ln = localName(name);
      if (inText) {
        if (ln === 'text') inText = false;
        else b.close(ln);
      } else if (ln === 'author' && au !== undefined) {
        authors.push(au);
        au = undefined;
      } else if (ln === 'comment' && cur) {
        const n: Note = { ref: cur.ref, text: b.result() };
        const author = authors[cur.authorId];
        if (author) n.author = author;
        notes.push(n);
        cur = undefined;
      }
    },
  });
  return notes;
}

interface Shape {
  visible: boolean;
  width?: number | undefined;
  height?: number | undefined;
}

const PX_PER: Record<string, number> = { pt: 4 / 3, px: 1, in: 96, cm: 96 / 2.54, mm: 9.6 / 2.54 };

function cssPx(style: string, prop: string): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([\\d.]+)\\s*(pt|px|in|cm|mm)?`, 'i').exec(
    style,
  );
  if (!m) return undefined;
  return Math.round(+m[1]! * (PX_PER[(m[2] ?? 'px').toLowerCase()] ?? 1) * 100) / 100;
}

/** VML note shapes keyed by "row,col" (0-based). */
function readVml(xml: string): Map<string, Shape> {
  const out = new Map<string, Shape>();
  let style = '';
  let note = false;
  let visible = false;
  let field = '';
  let row = '';
  let col = '';
  parseXml(xml, {
    open(name, a) {
      const ln = localName(name);
      if (ln === 'shape') {
        style = a.style ?? '';
        note = false;
        visible = false;
        row = '';
        col = '';
      } else if (ln === 'ClientData') note = a.ObjectType === 'Note';
      else if (ln === 'Visible') visible = true;
      else if (ln === 'Row' || ln === 'Column') field = ln;
    },
    text(t) {
      if (field === 'Row') row += t;
      else if (field === 'Column') col += t;
    },
    close(name) {
      const ln = localName(name);
      if (ln === 'Row' || ln === 'Column') field = '';
      else if (ln === 'shape' && note && row && col) {
        out.set(`${+row},${+col}`, {
          visible: visible || /visibility\s*:\s*visible/i.test(style),
          width: cssPx(style, 'width'),
          height: cssPx(style, 'height'),
        });
      }
    },
  });
  return out;
}

export const commentsReader: FeatureReader = {
  name: 'comments',
  sheet: {
    finish(ctx) {
      if (!ctx.rels) return;
      let notes: Note[] | undefined;
      let vml: string | undefined;
      for (const r of ctx.rels.values()) {
        if (r.external) continue;
        const k = relKind(r.type);
        if (k === 'comments') {
          const xml = ctx.readText(r.path);
          if (xml) notes = readComments(xml);
        } else if (k === 'vmlDrawing') vml ??= ctx.readText(r.path);
      }
      if (!notes || notes.length === 0) return;
      let shapes: Map<string, Shape> | undefined;
      try {
        shapes = vml ? readVml(vml) : undefined;
      } catch {
        // legacy VML is often sloppy; notes still load without it
      }
      if (shapes) {
        for (const n of notes) {
          const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(n.ref);
          const s = m && shapes.get(`${+m[2]! - 1},${decodeCol(m[1]!)}`);
          if (!s) continue;
          if (s.visible) n.visible = true;
          if (s.width !== undefined && Math.abs(s.width - W) > 0.5) n.width = s.width;
          if (s.height !== undefined && Math.abs(s.height - H) > 0.5) n.height = s.height;
        }
      }
      ctx.ws.comments = notes;
    },
  },
};
