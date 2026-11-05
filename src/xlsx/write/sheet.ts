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
import { fontMetrics } from '../../core/styles.js';
import type { CellError, Formula, Hyperlink, RichText } from '../../core/types.js';
import type {
  Column,
  RowData,
  SharedGroup,
  Worksheet,
  WorksheetView,
} from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';
import { runsXml } from '../common/rich-text.js';
import { isDynamic } from '../common/spill.js';
import { addXlfn } from '../common/xlfn.js';
import {
  type DrawingAnchors,
  type FeatureWriter,
  SheetPrOrder,
  SheetSlot,
  type SheetWriteCtx,
} from '../features/types.js';
import { translateFormula } from '../read/formula.js';
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

/** Mutable per-sheet encoding state (also the features' SheetWriteCtx). */
export interface SheetCtx extends SheetWriteCtx {
  /** Shared string table; undefined = inline strings. */
  readonly sst: SharedStrings | undefined;
  /** Hyperlinks collected while encoding (emitted in the sheet tail). */
  readonly links: LinkRec[];
  /** Set when any formula cell was written. */
  formulas: boolean;
  /** Shared-formula groups of the rows being encoded (see {@link sharedFormulas}). */
  shared: SharedWrite | undefined;
  /** Next shared-formula index (`si`) of this sheet. */
  si: number;
  /** Sheet relationships (written to the rels part after the sheet). */
  readonly rels: RelList;
  /** Used range ("A1:C10"); omitted when undefined (streaming). */
  dimension: string | undefined;
  /** Max row outline level (sheetFormatPr outlineLevelRow); 0 = unknown/none. */
  outlineLevelRow: number;
  tabSelected: boolean;
  /** Emitters of this write (built-in + enabled features). */
  readonly plan: SheetPlan;
  /** The package being written (shared by the sheets of one write). */
  readonly pkg: PackageParts;
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
    shared: undefined,
    si: 0,
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
    anchors: o.plan.anchors,
    addRel: (type, target, external) => rels.add(type, target, external),
    addPart: (path, data, ct) => pkg.addPart(path, data, ct),
    addDefault: (ext, ct) => pkg.addDefault(ext, ct),
    nextIndex: (f) => pkg.nextIndex(f),
    pkg,
  };
}


// ---- shared formulas ------------------------------------------------------------------

interface SharedRun {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
  /** The master's formula text. */
  readonly formula: string;
  readonly si: number;
}

/** Shared-formula groups of a run of rows, consumed in ascending row order by encodeRow. */
export interface SharedWrite {
  /** Sorted by top row. */
  readonly runs: SharedRun[];
  next: number;
  /** Runs covering the current row. */
  active: SharedRun[];
  row: number;
}



/**
 * Shared-formula state for encoding `ws`'s rows (`rowBase`: absolute index of `ws.rowData[0]`).
 * A group is kept when its master cell (top-left) still holds a plain formula and it overlaps
 * no later group; undefined when there is nothing to share.
 */
export function sharedFormulas(ws: Worksheet, ctx: SheetCtx, rowBase = 0): SharedWrite | undefined {
  const gs = ws._sharedGroups;
  if (gs === undefined || gs.length === 0) return undefined;
  const kept: SharedGroup[] = [];
  const runs: SharedRun[] = [];
  for (let k = gs.length - 1; k >= 0; k--) {
    const g: SharedGroup = gs[k]!;
    if (
      kept.some(
        (o) => o.top <= g.bottom && o.bottom >= g.top && o.left <= g.right && o.right >= g.left,
      )
    ) {
      continue;
    }
    const row: RowData | undefined = ws.rowData[g.top - rowBase];
    const i = row ? row.find(g.left) : -1;
    if (i < 0 || row!.kind[i] !== 6) continue;
    const f = row!.obj![i] as Formula;
    if (f.array !== undefined || isDynamic(f)) continue;
    kept.push(g);
    runs.push({
      top: g.top,
      left: g.left,
      bottom: g.bottom,
      right: g.right,
      formula: f.formula,
      si: 0,
    });
  }
  if (runs.length === 0) return undefined;
  runs.sort((a, b) => a.top - b.top || a.left - b.left);
  for (let k = 0; k < runs.length; k++) (runs[k] as { si: number }).si = ctx.si++;
  return { runs, next: 0, active: [], row: -1 };
}

/** The run cell (r0, c0) is a member of: master (its top-left) or a follower; else undefined. */
function sharedRole(sh: SharedWrite, r0: number, c0: number, f: Formula): SharedRun | undefined {
  if (sh.row !== r0) {
    sh.row = r0;
    sh.active = sh.active.filter((g) => g.bottom >= r0);
    while (sh.next < sh.runs.length && sh.runs[sh.next]!.top <= r0) {
      const g = sh.runs[sh.next++]!;
      if (g.bottom >= r0) sh.active.push(g);
    }
  }
  const act = sh.active;
  for (let k = 0; k < act.length; k++) {
    const g = act[k]!;
    if (c0 < g.left || c0 > g.right) continue;
    if (r0 === g.top && c0 === g.left) return g;
    if (f.array !== undefined || f.dynamic === true) return undefined;
    return translateFormula(g.formula, r0 - g.top, c0 - g.left) === f.formula ? g : undefined;
  }
  return undefined;
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

function formulaCell(
  ref: string,
  s: string,
  f: Formula,
  ctx: SheetCtx,
  r0: number,
  c0: number,
): string {
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
  const sh = ctx.shared;
  const g = sh === undefined ? undefined : sharedRole(sh, r0, c0, f);
  if (g !== undefined) {
    if (r0 !== g.top || c0 !== g.left)
      return `<c r="${ref}"${s}${t}><f t="shared" si="${g.si}"/>${v}</c>`;
    const sref = encodeRange({ top: g.top, left: g.left, bottom: g.bottom, right: g.right });
    return `<c r="${ref}"${s}${t}><f t="shared" ref="${sref}" si="${g.si}">${escapeText(addXlfn(f.formula))}</f>${v}</c>`;
  }
  const dyn = isDynamic(f);
  if (dyn) ctx.pkg.dynamicArrays = true;
  const fa = dyn
    ? ` t="array" ref="${escapeAttr(f.array ?? ref)}"`
    : f.array
      ? ` t="array" ref="${escapeAttr(f.array)}"`
      : '';
  return `<c r="${ref}"${s}${t}${dyn ? ' cm="1"' : ''}><f${fa}>${escapeText(addXlfn(f.formula))}</f>${v}</c>`;
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
  if (row.collapsed) {
    out += ' collapsed="1"';
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
        out += formulaCell(ref, s, obj![i] as Formula, ctx, r0, col[i]!);
        break;
      case 7: // RichText
        out += richCell(ref, s, obj![i] as RichText);
        break;
      case 8: {
        // Hyperlink
        const h = obj![i] as Hyperlink;
        out += h.runs?.length
          ? richCell(ref, s, { richText: h.runs })
          : strCell(ref, s, h.text ?? '', sst);
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

function viewXml(v: WorksheetView | undefined, i: number, selected: boolean): string {
  let s = `<sheetView workbookViewId="${v?.workbookViewId ?? i}"`;
  if (v?.rightToLeft) s += ' rightToLeft="1"';
  if (selected) s += ' tabSelected="1"';
  if (v?.showRowColHeaders === false) s += ' showRowColHeaders="0"';
  if (v?.showGridLines === false) s += ' showGridLines="0"';
  if (v?.showRuler === false) s += ' showRuler="0"';
  if (v?.style === 'pageLayout' || v?.style === 'pageBreakPreview') s += ` view="${v.style}"`;
  if (v?.zoomScale && v.zoomScale !== 100) s += ` zoomScale="${v.zoomScale}"`;
  if (v?.zoomScaleNormal) s += ` zoomScaleNormal="${v.zoomScaleNormal}"`;
  const x = v?.xSplit ?? 0;
  const y = v?.ySplit ?? 0;
  const active = v?.activeCell;
  const sel = active ? ` activeCell="${escapeAttr(active)}" sqref="${escapeAttr(active)}"` : '';
  if (v && (v.state === 'frozen' || v.state === 'split') && (x > 0 || y > 0)) {
    const pane =
      v.activePane ?? (x > 0 && y > 0 ? 'bottomRight' : y > 0 ? 'bottomLeft' : 'topRight');
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
    a.collapsed === b.collapsed &&
    a.styleIndex === b.styleIndex
  );
}

function colsXml(cols: readonly (Column | undefined)[], defW = 9.140625): string {
  let s = '';
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (!c) continue;
    if (c.width === undefined && !c.hidden && !c.outlineLevel && !c.collapsed && !c.styleIndex) {
      continue;
    }
    let j = i;
    while (j + 1 < cols.length && cols[j + 1] && sameCol(c, cols[j + 1]!)) j++;
    s += `<col min="${i + 1}" max="${j + 1}" width="${c.width ?? defW}"${
      c.width !== undefined ? ' customWidth="1"' : ''
    }${c.hidden ? ' hidden="1"' : ''}${c.outlineLevel ? ` outlineLevel="${c.outlineLevel}"` : ''}${
      c.collapsed ? ' collapsed="1"' : ''
    }${c.styleIndex ? ` style="${c.styleIndex}"` : ''}/>`;
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
    emit: (c) => {
      const vs = c.ws.views;
      let s = '';
      for (let i = 0; i < (vs.length || 1); i++) s += viewXml(vs[i], i, c.tabSelected);
      return `<sheetViews>${s}</sheetViews>`;
    },
  },
  {
    slot: SheetSlot.sheetFormatPr,
    emit: (c) => {
      // explicit ws.properties win; outline levels: the higher of the property and the data's
      const p = c.ws.properties;
      const d = fontMetrics(c.styles.defaultFont).rowPt;
      const h = p.defaultRowHeight ?? d;
      let s = `<sheetFormatPr${p.defaultColWidth ? ` defaultColWidth="${p.defaultColWidth}"` : ''} defaultRowHeight="${h}"`;
      if (h !== d) s += ' customHeight="1"';
      const or = Math.max(c.outlineLevelRow, p.outlineLevelRow ?? 0);
      if (or) s += ` outlineLevelRow="${or}"`;
      const oc = Math.max(colOutline(c.ws.columnData), p.outlineLevelCol ?? 0);
      if (oc) s += ` outlineLevelCol="${oc}"`;
      if (p.dyDescent !== undefined) s += ` x14ac:dyDescent="${p.dyDescent}"`;
      return `${s}/>`;
    },
  },
  {
    slot: SheetSlot.cols,
    emit: (c) => colsXml(c.ws.columnData, c.ws.properties.defaultColWidth),
  },
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

/**
 * Worksheet model fields owned by a feature, as "field:FeatureWriter.name". Writing without
 * that feature would silently drop the data, so it throws instead (a property check per sheet;
 * no feature code is referenced). The public export is `<camelCase name>Writer`.
 */
const OWNED: readonly string[] = [
  'tables:tables',
  'comments:comments',
  'images:images',
  'charts:charts',
  'backgroundImage:images',
  'conditionalFormatting:conditional-formatting',
  'dataValidations:data-validation',
  'protection:sheet-protection',
  'pageSetup:print-setup',
  'headerFooter:print-setup',
  'rowBreaks:print-setup',
  'colBreaks:print-setup',
  'ignoredErrors:ignored-errors',
];

function checkFeatures(ws: Worksheet, plan: SheetPlan): void {
  // drawing anchors (charts, shapes) kept by a read WITH the images feature: merged by the shared
  // drawing emitter (images, charts)
  const drawing = ws.preserved?.drawing && ws.preserved.name === ws.name && !plan.anchors.length;
  for (const o of drawing ? [...OWNED, 'preserved:images'] : OWNED) {
    const [field, name] = o.split(':') as [keyof Worksheet, string];
    const v = ws[field];
    if (v == null || (Array.isArray(v) && v.length === 0) || plan.names.has(name)) continue;
    throw new Error(
      `Worksheet '${ws.name}': ws.${field} is set but the '${name}' feature is not enabled; pass features: [${name.replace(/-(.)/g, (_, c: string) => c.toUpperCase())}Writer]`,
    );
  }
}

/** Per-write emitter list and `<worksheet>` start tag (see sheetPlan). */
export interface SheetPlan {
  readonly writers: readonly FeatureWriter[];
  /** FeatureWriter names of `writers`. */
  readonly names: ReadonlySet<string>;
  readonly emitters: readonly Emitter[];
  /** Drawing anchor contributions, sorted by order. */
  readonly anchors: readonly DrawingAnchors[];
  readonly open: string;
}

/** Built-in + feature emitters sorted by (slot, order); once per write (never per sheet/cell). */
export function sheetPlan(writers: readonly FeatureWriter[]): SheetPlan {
  const all: Emitter[] = [...BUILTIN];
  const anchors: DrawingAnchors[] = [];
  let ns = '';
  for (const f of writers) {
    // shared emitters (the drawing) are listed by several features: run once
    for (const e of f.sheet?.emitters ?? []) if (!all.includes(e)) all.push(e as Emitter);
    for (const n of f.sheet?.namespaces ?? []) if (!ns.includes(n)) ns += ` ${n}`;
    anchors.push(...(f.sheet?.anchors ?? []));
  }
  const by = (a: { order?: number | undefined }, b: { order?: number | undefined }): number =>
    (a.order ?? 0) - (b.order ?? 0);
  all.sort((a, b) => a.slot - b.slot || by(a, b));
  return {
    writers,
    names: new Set(writers.map((w) => w.name)),
    emitters: all,
    anchors: anchors.sort(by),
    open: `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"${ns}>`,
  };
}

/**
 * Run the emitters of slots [lo, hi] plus the sheet's preserved fragments (a preserved element
 * is written only when no emitter produced its slot); container slots (sheetPr, extLst) are
 * wrapped once.
 */
function runEmitters(ctx: SheetCtx, lo: number, hi: number): string {
  checkFeatures(ctx.ws, ctx.plan);
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
  ps?.w.frags(ws, ctx.path, ctx.addRel, lo, hi, out, pr);
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
  const open = ctx.plan.open;
  // x14ac:dyDescent (sheetFormatPr) is an Office 2010 extension: Ignorable for older readers
  return `${ctx.ws.properties.dyDescent !== undefined ? `${open.slice(0, -1)} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">` : open}${body}<sheetData>`;
}

/** `</sheetData>` through `</worksheet>`. Allocates the sheet's relationships (see ctx.rels). */
export function sheetTail(ctx: SheetCtx): string {
  const s = `</sheetData>${runEmitters(ctx, SheetSlot.sheetData + 1, SheetSlot.extLst)}</worksheet>`;
  ctx.ws.preserved?.w.implicit(ctx.ws, ctx.path, ctx.addRel);
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
