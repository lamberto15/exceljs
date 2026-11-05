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
// (a missing VML part is tolerated).
// Threaded comments (Excel 365 conversations: xl/threadedComments/threadedCommentN.xml, authors
// in xl/persons/person.xml) are `ws.threadedComments` (one thread per cell, with replies) and
// `wb.persons`. For each thread the writer also emits the legacy note Excel expects
// ("[Threaded comment] ...", author "tc={thread id}") — older Excel versions show it; the reader
// drops those generated notes from `ws.comments`. A cell holds a note or a thread, not both.
// The VML part is regenerated from the notes. Other shapes of a read VML (form controls) are
// kept: merged into the regenerated VML when the sheet had notes (ws.preserved.vml, see
// read/preserve.ts), else the whole part is kept with its <legacyDrawing>.

import { decodeAddr, decodeCol, encodeAddr } from '../../core/address.js';
import type { RichText } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import { localName, parseXml } from '../../xml/index.js';
import {
  type FeatureReader,
  type FeatureWriter,
  type SheetReadCtx,
  SheetSlot,
  type SheetWriteCtx,
} from './types.js';
import { escapeText, num, presVml, runsXml, StringItemBuilder, tElem } from './util.js';

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
  /**
   * Text margins (ExcelJS shape). `insetmode: 'custom'` uses `inset` = [left, top, right,
   * bottom] in centimetres; 'auto' (default) lets Excel choose.
   */
  margins?:
    | {
        insetmode?: 'auto' | 'custom' | undefined;
        inset?: number[] | undefined;
      }
    | undefined;
  /**
   * Note protection, effective when the sheet is protected: `locked` (the note box) and
   * `lockText` (its text). Excel's default for both is true.
   */
  protection?: { locked?: boolean | undefined; lockText?: boolean | undefined } | undefined;
  /**
   * How the box follows cell edits: 'absolute' (default, like Excel: don't move or size with
   * cells), 'oneCells' (move but don't size), 'twoCells' (move and size).
   */
  editAs?: 'twoCells' | 'oneCells' | 'absolute' | undefined;
}

/** A reply in a threaded comment. */
export interface ThreadedReply {
  text: string;
  /** Display name of the author (a `wb.persons` entry; added when missing). */
  author: string;
  date?: Date | undefined;
  /** GUID kept from the file ("{...}"); generated when missing. */
  id?: string | undefined;
}

/** A threaded comment (Excel 365 conversation) on a cell. */
export interface ThreadedComment {
  /** Cell address, "B2". */
  ref: string;
  text: string;
  /** Display name of the author (a `wb.persons` entry; added when missing). */
  author: string;
  date?: Date | undefined;
  /** The thread is resolved. */
  done?: boolean | undefined;
  replies?: ThreadedReply[] | undefined;
  /** GUID kept from the file ("{...}"); generated when missing. */
  id?: string | undefined;
}

/** An author of threaded comments (xl/persons/person.xml). */
export interface Person {
  displayName: string;
  /** GUID ("{...}"). */
  id: string;
  /** Account id, e.g. an e-mail address (default: the display name). */
  userId?: string | undefined;
  /** Account provider: "AD", "Windows Live", "PeoplePicker", "None" (default). */
  providerId?: string | undefined;
}

/**
 * Add a threaded comment to a cell (replacing a thread already there); returns it. `replies`
 * are appended in order. A note on the same cell is an error when writing.
 */
export function addThreadedComment(
  ws: Worksheet,
  ref: string,
  comment: Omit<ThreadedComment, 'ref'>,
): ThreadedComment {
  const a = decodeAddr(ref);
  const t: ThreadedComment = { ...comment, ref: encodeAddr(a.row, a.col) };
  ws.threadedComments ??= [];
  const list = ws.threadedComments;
  const i = list.findIndex((x) => x.ref === t.ref);
  if (i >= 0) list[i] = t;
  else list.push(t);
  return t;
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
  const p = c.ws.properties;
  const defW = p.defaultColWidth ?? 9.140625;
  const defH = p.defaultRowHeight ? Math.round((p.defaultRowHeight * 96) / 72) : 20;
  const colPx = (i: number): number => {
    const d = cols[i];
    return d?.hidden ? 0 : colWidthPx(d?.width ?? defW);
  };
  const rowPx = (i: number): number => {
    const d = rows[i];
    return d?.hidden ? 0 : d?.height ? Math.round((d.height * 96) / 72) : defH;
  };
  // shape ids come in blocks of 1024 per <o:idmap> entry, unique within the package; shapes
  // kept from the read file (form controls) keep theirs (controls refer to them)
  const blocks: number[] = [];
  for (let k = 0; k * 1023 < notes.length; k++) blocks.push(c.nextIndex('vmlShapeBlock'));
  const kept = presVml(c.ws);
  let s = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${[...blocks, ...(kept?.blocks ?? [])].join(',')}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${kept?.xml ?? ''}`;
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
    const m = n.margins;
    // inset in mm (ExcelJS: centimetres x 10)
    const inset = m?.inset?.length
      ? ` inset="${m.inset.map((x) => `${Math.round(x * 1000) / 100}mm`).join(',')}"`
      : '';
    const pr = n.protection;
    const tf = (k: string, v: boolean | undefined): string =>
      v === undefined ? '' : `<x:${k}>${v ? 'True' : 'False'}</x:${k}>`;
    s += `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:${pt(offset(l, colPx) + 15)};margin-top:${pt(offset(t, rowPx) + 10)};width:${pt(w)};height:${pt(h)};z-index:${i + 1};visibility:${vis}" fillcolor="#ffffe1" o:insetmode="${m?.insetmode === 'custom' ? 'custom' : 'auto'}"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"${inset}><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note">${n.editAs === 'twoCells' ? '' : n.editAs === 'oneCells' ? '<x:SizeWithCells/>' : '<x:MoveWithCells/><x:SizeWithCells/>'}<x:Anchor>${anchor}</x:Anchor>${tf('Locked', pr?.locked)}<x:AutoFill>False</x:AutoFill>${tf('LockText', pr?.lockText)}<x:Row>${row}</x:Row><x:Column>${col}</x:Column>${n.visible ? '<x:Visible/>' : ''}</x:ClientData></v:shape>`;
  }
  return `${s}</xml>`;
}

const NS_TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const REL_TC = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const REL_PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const CT_TC = 'application/vnd.ms-excel.threadedcomments+xml';
const CT_PERSON = 'application/vnd.ms-excel.person+xml';

/** A stable GUID ("{XXXXXXXX-...}") derived from `seed` (deterministic output). */
function guid(seed: string): string {
  let hex = '';
  for (let k = 0; k < 4; k++) {
    let h = 0x811c9dc5 ^ (k * 0x9e3779b9);
    for (let i = 0; i < seed.length; i++) {
      h ^= seed.charCodeAt(i);
      h = Math.imul(h, 0x01000193);
    }
    hex += (h >>> 0).toString(16).padStart(8, '0');
  }
  hex = hex.toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}}`;
}

/** The person id of an author: from `wb.persons`, else derived from the name. */
function personId(c: { workbook: Worksheet['workbook'] }, name: string): string {
  return c.workbook.persons?.find((p) => p.displayName === name)?.id ?? guid(`person:${name}`);
}

const threadId = (ws: Worksheet, t: ThreadedComment): string =>
  t.id ?? guid(`tc:${ws.name}:${t.ref}`);

/** Excel's timestamp form: local time without zone, two decimals. */
const stamp = (d: Date): string => `${d.toISOString().slice(0, 19)}.00`;

const LEGACY =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ';

/** The legacy notes Excel writes alongside threads (older versions show them). */
function threadNotes(ws: Worksheet): Note[] {
  return (ws.threadedComments ?? []).map((t) => {
    let text = LEGACY + t.text;
    for (const r of t.replies ?? []) text += `\nReply:\n    ${r.text}`;
    return { ref: t.ref, text, author: `tc=${threadId(ws, t)}` };
  });
}

function threadsXml(c: SheetWriteCtx, threads: readonly ThreadedComment[]): string {
  let s = '';
  const one = (
    ref: string,
    text: string,
    author: string,
    date: Date | undefined,
    id: string,
    extra: string,
  ) => {
    s += `<threadedComment ref="${ref}"${date ? ` dT="${stamp(date)}"` : ''} personId="${personId(c, author)}" id="${id}"${extra}><text>${escapeText(text)}</text></threadedComment>`;
  };
  for (const t of threads) {
    const a = decodeAddr(t.ref);
    const ref = encodeAddr(a.row, a.col);
    const root = threadId(c.ws, t);
    one(ref, t.text, t.author, t.date, root, t.done ? ' done="1"' : '');
    (t.replies ?? []).forEach((r, k) => {
      one(
        ref,
        r.text,
        r.author,
        r.date,
        r.id ?? guid(`tc:${c.ws.name}:${t.ref}:${k}`),
        ` parentId="${root}"`,
      );
    });
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="${NS_TC}" xmlns:x="${NS_MAIN}">${s}</ThreadedComments>`;
}

function emit(c: SheetWriteCtx): string {
  const threads = c.ws.threadedComments ?? [];
  const own = c.ws.comments ?? [];
  for (const t of threads) {
    const a = decodeAddr(t.ref);
    if (
      own.some((n) => {
        const b = decodeAddr(n.ref);
        return a.row === b.row && a.col === b.col;
      })
    ) {
      throw new Error(
        `Worksheet '${c.ws.name}': cell ${t.ref} has both a note and a threaded comment (Excel allows one)`,
      );
    }
  }
  const notes = threads.length ? [...own, ...threadNotes(c.ws)] : own;
  if (threads.length) {
    const n = c.nextIndex('threadedComment');
    c.addPart(`xl/threadedComments/threadedComment${n}.xml`, threadsXml(c, threads), CT_TC);
    c.addRel(REL_TC, `../threadedComments/threadedComment${n}.xml`);
  }
  // without notes, a VML is still needed for shapes kept next to the read file's notes
  if (notes.length === 0 && !presVml(c.ws)) return '';
  const cells = notes.map((n) => {
    const a = decodeAddr(n.ref);
    return encodeAddr(a.row, a.col);
  });
  const v = c.nextIndex('vmlDrawing');
  if (notes.length > 0) {
    const n = c.nextIndex('comments');
    c.addPart(`xl/comments${n}.xml`, commentsXml(notes, cells), CT_COMMENTS);
    c.addRel('comments', `../comments${n}.xml`);
  }
  c.addPart(`xl/drawings/vmlDrawing${v}.vml`, vmlXml(c, notes, cells));
  c.addDefault('vml', CT_VML);
  return `<legacyDrawing r:id="${c.addRel('vmlDrawing', `../drawings/vmlDrawing${v}.vml`)}"/>`;
}

/**
 * Write half of the cell notes (ws.comments) feature. Always on in the package root; in `./lite` pass it
 * in `features` of writeXlsx / xlsxStream / createXlsxWriter.
 */
export const commentsWriter: FeatureWriter = {
  name: 'comments',
  sheet: { emitters: [{ slot: SheetSlot.legacyDrawing, emit }] },
  book: {
    write(c) {
      // the persons part: kept persons plus one per new threaded-comment author
      const wb = c.workbook;
      const list: Person[] = [...(wb.persons ?? [])];
      for (const ws of wb.worksheets) {
        for (const t of ws.threadedComments ?? []) {
          for (const name of [t.author, ...(t.replies ?? []).map((r) => r.author)]) {
            if (!list.some((p) => p.displayName === name)) {
              list.push({ displayName: name, id: guid(`person:${name}`) });
            }
          }
        }
      }
      if (!wb.worksheets.some((ws) => ws.threadedComments?.length)) return;
      let s = '';
      for (const p of list) {
        s += `<person displayName="${escapeText(p.displayName).replace(/"/g, '&quot;')}" id="${p.id}" userId="${escapeText(p.userId ?? p.displayName).replace(/"/g, '&quot;')}" providerId="${escapeText(p.providerId ?? 'None').replace(/"/g, '&quot;')}"/>`;
      }
      c.addPart(
        'xl/persons/person.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<personList xmlns="${NS_TC}" xmlns:x="${NS_MAIN}">${s}</personList>`,
        CT_PERSON,
      );
      c.addRel(REL_PERSON, 'persons/person.xml');
    },
  },
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
  margins?: Note['margins'];
  protection?: Note['protection'];
  editAs?: Note['editAs'];
}

const PX_PER: Record<string, number> = { pt: 4 / 3, px: 1, in: 96, cm: 96 / 2.54, mm: 9.6 / 2.54 };

function cssPx(style: string, prop: string): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([\\d.]+)\\s*(pt|px|in|cm|mm)?`, 'i').exec(
    style,
  );
  if (!m) return undefined;
  return Math.round(+m[1]! * (PX_PER[(m[2] ?? 'px').toLowerCase()] ?? 1) * 100) / 100;
}

/** VML inset ("2.5mm,1.3mm,...") → centimetres. */
function insetCm(v: string): number[] | undefined {
  const out = v.split(',').map((x) => {
    const px = cssPx(`w:${x}`, 'w');
    return px === undefined ? Number.NaN : Math.round((px * 254) / 96) / 100;
  });
  return out.length === 4 && out.every((x) => x - x === 0) ? out : undefined;
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
  let mode = '';
  let inset: number[] | undefined;
  let prot: Record<string, boolean> = {};
  let flag = '';
  let pos = 0;
  parseXml(xml, {
    open(name, a) {
      const ln = localName(name);
      if (ln === 'shape') {
        style = a.style ?? '';
        note = false;
        visible = false;
        row = '';
        col = '';
        mode = a['o:insetmode'] ?? '';
        inset = undefined;
        prot = {};
        pos = 0;
      } else if (ln === 'textbox') inset = a.inset ? insetCm(a.inset) : undefined;
      else if (ln === 'ClientData') note = a.ObjectType === 'Note';
      else if (ln === 'Visible') visible = true;
      else if (ln === 'MoveWithCells') pos |= 1;
      else if (ln === 'SizeWithCells') pos |= 2;
      else if (ln === 'Row' || ln === 'Column') field = ln;
      else if (ln === 'Locked' || ln === 'LockText') {
        field = ln;
        flag = '';
      }
    },
    text(t) {
      if (field === 'Row') row += t;
      else if (field === 'Column') col += t;
      else if (field) flag += t;
    },
    close(name) {
      const ln = localName(name);
      if (ln === 'Locked' || ln === 'LockText') {
        const f = flag.trim().toLowerCase();
        if (f) prot[ln === 'Locked' ? 'locked' : 'lockText'] = f !== 'false' && f !== 'f';
        field = '';
      } else if (ln === 'Row' || ln === 'Column') field = '';
      else if (ln === 'shape' && note && row && col) {
        const s: Shape = {
          visible: visible || /visibility\s*:\s*visible/i.test(style),
          width: cssPx(style, 'width'),
          height: cssPx(style, 'height'),
        };
        if (mode === 'custom' || inset) {
          s.margins = mode === 'custom' || mode === 'auto' ? { insetmode: mode } : {};
          if (inset) s.margins.inset = inset;
        }
        if (Object.keys(prot).length) s.protection = prot;
        // present = does NOT move / size with cells (Excel's note default: both)
        if (pos !== 3) s.editAs = pos & 2 ? 'oneCells' : 'twoCells';
        out.set(`${+row},${+col}`, s);
      }
    },
  });
  return out;
}

/**
 * Read half of the cell notes (ws.comments) feature. Always on in the package root; in `./lite` pass it
 * in `features` of readXlsx (without it, the elements are kept raw and written back).
 */
/** Persons of the package (xl/persons/person.xml via the workbook's relationships). */
function readPersons(ctx: SheetReadCtx): Person[] {
  const root = [...(ctx.relsOf('')?.values() ?? [])].find(
    (r) => relKind(r.type) === 'officeDocument',
  );
  const rel =
    root && [...(ctx.relsOf(root.path)?.values() ?? [])].find((r) => relKind(r.type) === 'person');
  const xml = rel && !rel.external ? ctx.readText(rel.path) : undefined;
  const out: Person[] = [];
  if (!xml) return out;
  parseXml(xml, {
    open(name, a) {
      if (localName(name) !== 'person' || !a.id) return;
      const p: Person = { displayName: a.displayName ?? '', id: a.id };
      if (a.userId !== undefined && a.userId !== p.displayName) p.userId = a.userId;
      if (a.providerId !== undefined && a.providerId !== 'None') p.providerId = a.providerId;
      out.push(p);
    },
  });
  return out;
}

const parseStamp = (s: string | undefined): Date | undefined => {
  const m = s && /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(\.\d+)?/.exec(s);
  return m
    ? new Date(
        Date.UTC(
          +m[1]!,
          +m[2]! - 1,
          +m[3]!,
          +m[4]!,
          +m[5]!,
          +m[6]!,
          Math.round(+(m[7] ?? 0) * 1000),
        ),
      )
    : undefined;
};

function readThreads(xml: string, persons: Person[]): ThreadedComment[] {
  const byId = new Map<string, ThreadedComment>();
  const out: ThreadedComment[] = [];
  let cur: { a: Record<string, string>; text: string } | undefined;
  let inText = false;
  const name = (pid: string | undefined) => persons.find((p) => p.id === pid)?.displayName ?? '';
  parseXml(xml, {
    open(n, a) {
      const ln = localName(n);
      if (ln === 'threadedComment') cur = { a: { ...a }, text: '' };
      else if (ln === 'text' && cur) inText = true;
    },
    text(t) {
      if (inText && cur) cur.text += t;
    },
    close(n) {
      const ln = localName(n);
      if (ln === 'text') inText = false;
      else if (ln === 'threadedComment' && cur) {
        const { a, text } = cur;
        cur = undefined;
        const date = parseStamp(a.dT);
        if (a.parentId) {
          const root = byId.get(a.parentId);
          if (!root) return;
          const r: ThreadedReply = { text, author: name(a.personId), id: a.id };
          if (date) r.date = date;
          if (root.replies) root.replies.push(r);
          else root.replies = [r];
          return;
        }
        if (!a.ref) return;
        const t: ThreadedComment = { ref: a.ref, text, author: name(a.personId), id: a.id };
        if (date) t.date = date;
        if (a.done === '1' || a.done === 'true') t.done = true;
        if (a.id) byId.set(a.id, t);
        out.push(t);
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
      let threads: ThreadedComment[] | undefined;
      for (const r of ctx.rels.values()) {
        if (r.external) continue;
        const k = relKind(r.type);
        if (k === 'comments') {
          const xml = ctx.readText(r.path);
          if (xml) notes = readComments(xml);
        } else if (k === 'vmlDrawing') vml ??= ctx.readText(r.path);
        else if (k === 'threadedComment') {
          const xml = ctx.readText(r.path);
          if (xml) {
            const wb = ctx.workbook;
            wb.persons ??= readPersons(ctx);
            threads = readThreads(xml, wb.persons);
          }
        }
      }
      if (threads?.length) {
        ctx.ws.threadedComments = threads;
        // the legacy notes Excel wrote for the threads are regenerated from them
        const at = new Set(threads.map((t) => t.ref));
        notes = (notes ?? []).filter((n) => !(n.author?.startsWith('tc=') && at.has(n.ref)));
        if (notes.length === 0) {
          ctx.ws.comments = [];
          return;
        }
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
          if (s.margins) n.margins = s.margins;
          if (s.protection) n.protection = s.protection;
          if (s.editAs) n.editAs = s.editAs;
        }
      }
      ctx.ws.comments = notes;
    },
  },
};
