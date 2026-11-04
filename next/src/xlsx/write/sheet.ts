// Worksheet XML: head (everything before <sheetData>), the row/cell encoder (hot loop) and the
// tail (everything after </sheetData>). The cell encoder `encodeRow` is the single
// implementation shared by writeXlsx, xlsxStream and the streaming writer.
//
// Head and tail are produced by SheetEmitters ordered by CT_Worksheet slot (features/types.ts):
// the built-in ones below plus every registered feature's, so element order is global.
//
// Rich text is always written inline (<is><r>…</r></is>) — the shared string table only holds
// plain strings. Hyperlink cells are written as their text (shared or inline per options) and
// recorded in `ctx.links` for the <hyperlinks> section and the sheet rels part.

import { encodeAddr, encodeCol, encodeRange } from '../../core/address.js';
import type { SharedStrings } from '../../core/sharedstrings.js';
import type { CellError, Formula, Hyperlink, RichText } from '../../core/types.js';
import type { Column, RowData, Worksheet, WorksheetView } from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';
import { fillRels } from '../common/preserved.js';
import { runsXml } from '../common/rich-text.js';
import {
  type FeatureWriter,
  type SheetEmitter,
  SheetPrOrder,
  SheetSlot,
  type SheetWriteCtx,
} from '../features/types.js';
import { type PackageParts, type PartRec, RelList, relsXml } from './package.js';
import {
  colorAttrs,
  type LinkRec,
  NS_MAIN,
  NS_R,
  sheetPath,
  sheetRelsPath,
  tElem,
  XML_DECL,
} from './parts.js';
import { fragLive, sheetImplicit } from './preserve.js';

/** Mutable per-sheet encoding state (also the features' SheetWriteCtx). */
export interface SheetCtx extends SheetWriteCtx {
  /** Shared string table; undefined = inline strings. */
  readonly sst: SharedStrings | undefined;
  /** Hyperlinks collected while encoding (emitted in the sheet tail). */
  readonly links: LinkRec[];
  /** Set when any formula cell was written. */
  formulas: boolean;
  /** Sheet relationships (written to the rels part after the sheet). */
  readonly rels: RelList;
  /** Used range ("A1:C10"); omitted when undefined (streaming). */
  dimension: string | undefined;
  /** Max row outline level (sheetFormatPr outlineLevelRow); 0 = unknown/none. */
  outlineLevelRow: number;
  tabSelected: boolean;
  /** Emitters of this write (built-in + enabled features). */
  readonly plan: SheetPlan;
}

export interface NewSheetCtx {
  readonly plan: SheetPlan;
  readonly sst: SharedStrings | undefined;
  readonly pkg: PackageParts;
  readonly ws: Worksheet;
  readonly index: number;
  readonly streaming: boolean;
}

export function newSheetCtx(o: NewSheetCtx): SheetCtx {
  const pkg = o.pkg;
  const rels = new RelList();
  const ws = o.ws;
  return {
    sst: o.sst,
    links: [],
    formulas: false,
    rels,
    dimension: undefined,
    outlineLevelRow: 0,
    tabSelected: false,
    ws,
    workbook: ws.workbook,
    styles: ws.workbook.styles,
    index: o.index,
    path: sheetPath(o.index),
    streaming: o.streaming,
    plan: o.plan,
    addRel: (type, target, external) => rels.add(type, target, external),
    addPart: (path, data, ct) => pkg.addPart(path, data, ct),
    addDefault: (ext, ct) => pkg.addDefault(ext, ct),
    nextIndex: (f) => pkg.nextIndex(f),
  };
}

// ---- caches (module level; contents are deterministic, so sharing is safe) -------------

const COLS: string[] = [];
const S_ATTR: string[] = [];

function colLetters(c: number): string {
  const hit = COLS[c];
  if (hit !== undefined) return hit;
  const s = encodeCol(c);
  COLS[c] = s;
  return s;
}

function sAttr(xf: number): string {
  const hit = S_ATTR[xf];
  if (hit !== undefined) return hit;
  const s = ` s="${xf}"`;
  if (xf < 65536) S_ATTR[xf] = s;
  return s;
}

function strCell(ref: string, s: string, v: string, sst: SharedStrings | undefined): string {
  return sst !== undefined
    ? `<c r="${ref}"${s} t="s"><v>${sst.intern(v)}</v></c>`
    : `<c r="${ref}"${s} t="inlineStr"><is>${tElem(v)}</is></c>`;
}

function richCell(ref: string, s: string, rt: RichText): string {
  return `<c r="${ref}"${s} t="inlineStr"><is>${runsXml(rt.richText)}</is></c>`;
}

function formulaCell(ref: string, s: string, f: Formula, ctx: SheetCtx): string {
  ctx.formulas = true;
  let t = '';
  let v = '';
  const res = f.result;
  if (res !== undefined && res !== null) {
    switch (typeof res) {
      case 'number':
        if (res - res === 0) v = `<v>${res}</v>`;
        else {
          t = ' t="e"';
          v = '<v>#NUM!</v>';
        }
        break;
      case 'string':
        t = ' t="str"';
        v = `<v>${escapeText(res)}</v>`;
        break;
      case 'boolean':
        t = ' t="b"';
        v = res ? '<v>1</v>' : '<v>0</v>';
        break;
      case 'object':
        if ('error' in res) {
          t = ' t="e"';
          v = `<v>${escapeText((res as CellError).error)}</v>`;
        }
        break;
    }
  }
  const fa = f.array ? ` t="array" ref="${escapeAttr(f.array)}"` : '';
  return `<c r="${ref}"${s}${t}><f${fa}>${escapeText(f.formula)}</f>${v}</c>`;
}

/**
 * Encode one row as `<row …>…</row>`. `r0` is the absolute 0-based row index (streaming writers
 * buffer rows at relative positions). Returns '' for a row with neither cells nor properties.
 */
export function encodeRow(row: RowData, r0: number, ctx: SheetCtx): string {
  const n = row.n;
  const rs = String(r0 + 1);
  let out = `<row r="${rs}"`;
  let props = false;
  if (row.height > 0) {
    out += ` ht="${row.height}" customHeight="1"`;
    props = true;
  }
  if (row.hidden) {
    out += ' hidden="1"';
    props = true;
  }
  if (row.outlineLevel > 0) {
    out += ` outlineLevel="${row.outlineLevel}"`;
    props = true;
  }
  if (row.xf !== 0) {
    out += ` s="${row.xf}" customFormat="1"`;
    props = true;
  }
  if (n === 0) return props ? `${out}/>` : '';
  out += '>';
  const col = row.col;
  const kind = row.kind;
  const num = row.num;
  const style = row.style;
  const obj = row.obj;
  const sst = ctx.sst;
  for (let i = 0; i < n; i++) {
    const ref = colLetters(col[i]!) + rs;
    const xf = style[i]!;
    const s = xf === 0 ? '' : sAttr(xf);
    switch (kind[i]) {
      case 1: // Number
      case 4: // Date (serial)
        out += `<c r="${ref}"${s}><v>${num[i]}</v></c>`;
        break;
      case 2: // String
        out += strCell(ref, s, obj![i] as string, sst);
        break;
      case 3: // Boolean
        out += `<c r="${ref}"${s} t="b"><v>${num[i] ? 1 : 0}</v></c>`;
        break;
      case 5: // Error
        out += `<c r="${ref}"${s} t="e"><v>${escapeText(obj![i] as string)}</v></c>`;
        break;
      case 6: // Formula
        out += formulaCell(ref, s, obj![i] as Formula, ctx);
        break;
      case 7: // RichText
        out += richCell(ref, s, obj![i] as RichText);
        break;
      case 8: {
        // Hyperlink
        const h = obj![i] as Hyperlink;
        out += strCell(ref, s, h.text ?? '', sst);
        if (h.hyperlink) ctx.links.push({ ref, target: h.hyperlink, tooltip: h.tooltip });
        break;
      }
      default: // Empty, style only
        out += `<c r="${ref}"${s}/>`;
    }
  }
  return `${out}</row>`;
}

// ---- head & tail -----------------------------------------------------------------------

function viewXml(v: WorksheetView | undefined, selected: boolean): string {
  let s = '<sheetView workbookViewId="0"';
  if (v?.rightToLeft) s += ' rightToLeft="1"';
  if (selected) s += ' tabSelected="1"';
  if (v?.showRowColHeaders === false) s += ' showRowColHeaders="0"';
  if (v?.showGridLines === false) s += ' showGridLines="0"';
  if (v?.zoomScale && v.zoomScale !== 100) s += ` zoomScale="${v.zoomScale}"`;
  const x = v?.xSplit ?? 0;
  const y = v?.ySplit ?? 0;
  const active = v?.activeCell;
  const sel = active ? ` activeCell="${escapeAttr(active)}" sqref="${escapeAttr(active)}"` : '';
  if (v && (v.state === 'frozen' || v.state === 'split') && (x > 0 || y > 0)) {
    const pane = x > 0 && y > 0 ? 'bottomRight' : y > 0 ? 'bottomLeft' : 'topRight';
    const frozen = v.state === 'frozen';
    const tl = v.topLeftCell ?? (frozen ? encodeAddr(y, x) : undefined);
    s += `><pane${x > 0 ? ` xSplit="${x}"` : ''}${y > 0 ? ` ySplit="${y}"` : ''}${
      tl ? ` topLeftCell="${escapeAttr(tl)}"` : ''
    } activePane="${pane}"${frozen ? ' state="frozen"' : ''}/><selection pane="${pane}"${sel}/></sheetView>`;
    return s;
  }
  return sel ? `${s}><selection${sel}/></sheetView>` : `${s}/>`;
}

function sameCol(a: Column, b: Column): boolean {
  return (
    a.width === b.width &&
    a.hidden === b.hidden &&
    a.outlineLevel === b.outlineLevel &&
    a.styleIndex === b.styleIndex
  );
}

function colsXml(cols: readonly (Column | undefined)[]): string {
  let s = '';
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (!c) continue;
    if (c.width === undefined && !c.hidden && !c.outlineLevel && !c.styleIndex) continue;
    let j = i;
    while (j + 1 < cols.length && cols[j + 1] && sameCol(c, cols[j + 1]!)) j++;
    s += `<col min="${i + 1}" max="${j + 1}" width="${c.width ?? 9.140625}"${
      c.width !== undefined ? ' customWidth="1"' : ''
    }${c.hidden ? ' hidden="1"' : ''}${c.outlineLevel ? ` outlineLevel="${c.outlineLevel}"` : ''}${
      c.styleIndex ? ` style="${c.styleIndex}"` : ''
    }/>`;
    i = j;
  }
  return s ? `<cols>${s}</cols>` : '';
}

function colOutline(cols: readonly (Column | undefined)[]): number {
  let outline = 0;
  for (const c of cols) if (c && c.outlineLevel > outline) outline = c.outlineLevel;
  return outline;
}

/** A sheet emitter as the writer sees it (feature emitters get the full SheetCtx). */
export interface Emitter {
  readonly slot: number;
  readonly order?: number | undefined;
  emit(ctx: SheetCtx): string;
}

/** Built-in worksheet elements (features add the rest, see features/types.ts SheetSlot). */
const BUILTIN: readonly Emitter[] = [
  {
    slot: SheetSlot.sheetPr,
    order: SheetPrOrder.tabColor,
    emit: (c) => (c.ws.tabColor ? `<tabColor${colorAttrs(c.ws.tabColor)}/>` : ''),
  },
  {
    slot: SheetSlot.dimension,
    emit: (c) => (c.dimension ? `<dimension ref="${c.dimension}"/>` : ''),
  },
  {
    slot: SheetSlot.sheetViews,
    emit: (c) => `<sheetViews>${viewXml(c.ws.views[0], c.tabSelected)}</sheetViews>`,
  },
  {
    slot: SheetSlot.sheetFormatPr,
    emit: (c) => {
      let s = '<sheetFormatPr defaultRowHeight="15"';
      if (c.outlineLevelRow) s += ` outlineLevelRow="${c.outlineLevelRow}"`;
      const oc = colOutline(c.ws.columnData);
      if (oc) s += ` outlineLevelCol="${oc}"`;
      return `${s}/>`;
    },
  },
  { slot: SheetSlot.cols, emit: (c) => colsXml(c.ws.columnData) },
  {
    slot: SheetSlot.autoFilter,
    emit: (c) => (c.ws.autoFilter ? `<autoFilter ref="${escapeAttr(c.ws.autoFilter)}"/>` : ''),
  },
  {
    slot: SheetSlot.mergeCells,
    emit: (c) => {
      const merges = c.ws.mergeRanges;
      if (merges.length === 0) return '';
      let s = `<mergeCells count="${merges.length}">`;
      for (const m of merges) s += `<mergeCell ref="${encodeRange(m)}"/>`;
      return `${s}</mergeCells>`;
    },
  },
  {
    slot: SheetSlot.hyperlinks,
    emit: (c) => {
      if (c.links.length === 0) return '';
      let s = '<hyperlinks>';
      for (const l of c.links) {
        const tip = l.tooltip !== undefined ? ` tooltip="${escapeAttr(l.tooltip)}"` : '';
        s +=
          l.target.charCodeAt(0) === 35
            ? `<hyperlink ref="${l.ref}" location="${escapeAttr(l.target.slice(1))}"${tip}/>`
            : `<hyperlink ref="${l.ref}" r:id="${c.addRel('hyperlink', l.target, true)}"${tip}/>`;
      }
      return `${s}</hyperlinks>`;
    },
  },
];

/** Per-write emitter list and `<worksheet>` start tag (see sheetPlan). */
export interface SheetPlan {
  readonly writers: readonly FeatureWriter[];
  readonly emitters: readonly Emitter[];
  readonly open: string;
}

/** Built-in + feature emitters sorted by (slot, order); once per write (never per sheet/cell). */
export function sheetPlan(writers: readonly FeatureWriter[]): SheetPlan {
  const all: Emitter[] = [...BUILTIN];
  let ns = '';
  for (const f of writers) {
    for (const e of f.sheet?.emitters ?? []) all.push(e as SheetEmitter as Emitter);
    for (const n of f.sheet?.namespaces ?? []) if (!ns.includes(n)) ns += ` ${n}`;
  }
  all.sort((a, b) => a.slot - b.slot || (a.order ?? 0) - (b.order ?? 0));
  return {
    writers,
    emitters: all,
    open: `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"${ns}>`,
  };
}

/**
 * Run the emitters of slots [lo, hi] plus the sheet's preserved fragments (a preserved element
 * is written only when no emitter produced its slot); container slots (sheetPr, extLst) are
 * wrapped once.
 */
function runEmitters(ctx: SheetCtx, lo: number, hi: number): string {
  const out: string[] = [];
  const pr: [number, string][] = [];
  for (const e of ctx.plan.emitters) {
    if (e.slot < lo || e.slot > hi) continue;
    const x = e.emit(ctx);
    if (x === '') continue;
    if (e.slot === SheetSlot.sheetPr) pr.push([e.order ?? 0, x]);
    else out[e.slot] = (out[e.slot] ?? '') + x;
  }
  const ws = ctx.ws;
  const ps = ws.preserved;
  if (ps) {
    for (const f of ps.frags) {
      const k = f.slot;
      if (k < lo || k > hi || !fragLive(f, ws)) continue;
      if (k !== SheetSlot.sheetPr && k !== SheetSlot.extLst && out[k]) continue;
      const x = fillRels(f.xml, ps.rels, ctx.path, ctx.addRel);
      if (k === SheetSlot.sheetPr) pr.push([f.order, x]);
      else out[k] = (out[k] ?? '') + x;
    }
  }
  const pa = lo === 0 ? (ps?.prAttrs ?? '') : '';
  if (pr.length || pa) {
    pr.sort((a, b) => a[0] - b[0]);
    out[0] = pr.length
      ? `<sheetPr${pa}>${pr.map((p) => p[1]).join('')}</sheetPr>`
      : `<sheetPr${pa}/>`;
  }
  const ext = out[SheetSlot.extLst];
  if (ext) out[SheetSlot.extLst] = `<extLst>${ext}</extLst>`;
  return out.join('');
}

/** Everything up to and including `<sheetData>`. Set ctx.dimension/outlineLevelRow/tabSelected first. */
export function sheetHead(ctx: SheetCtx): string {
  const body = runEmitters(ctx, 0, SheetSlot.sheetData - 1);
  return `${ctx.plan.open}${body}<sheetData>`;
}

/** `</sheetData>` through `</worksheet>`. Allocates the sheet's relationships (see ctx.rels). */
export function sheetTail(ctx: SheetCtx): string {
  const s = `</sheetData>${runEmitters(ctx, SheetSlot.sheetData + 1, SheetSlot.extLst)}</worksheet>`;
  sheetImplicit(ctx.ws, ctx.path, ctx.addRel);
  return s;
}

/**
 * Parts to write right after the sheet part: queued feature parts, then the sheet's rels part
 * when it has relationships.
 */
export function afterSheetParts(ctx: SheetCtx, pkg: PackageParts): PartRec[] {
  const out = pkg.take();
  if (ctx.rels.recs.length > 0) {
    out.push({ path: sheetRelsPath(ctx.index), data: relsXml(ctx.rels.recs) });
  }
  return out;
}
