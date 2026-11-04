// Worksheet XML: head (sheetPr … <sheetData>), the row/cell encoder (hot loop) and the tail
// (</sheetData> autoFilter mergeCells hyperlinks pageMargins). The cell encoder `encodeRow` is
// the single implementation shared by writeXlsx, xlsxStream and the streaming writer.
//
// Rich text is always written inline (<is><r>…</r></is>) — the shared string table only holds
// plain strings. Hyperlink cells are written as their text (shared or inline per options) and
// recorded in `ctx.links` for the <hyperlinks> section and the sheet rels part.

import { encodeAddr, encodeCol, encodeRange } from '../../core/address.js';
import type { SharedStrings } from '../../core/sharedstrings.js';
import type { CellError, Formula, Hyperlink, RichText, RichTextRun } from '../../core/types.js';
import type { Column, RowData, Worksheet, WorksheetView } from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';
import { colorAttrs, fontBody, type LinkRec, NS_MAIN, NS_R, tElem, XML_DECL } from './parts.js';

/** Mutable per-sheet encoding state. */
export interface SheetCtx {
  /** Shared string table; undefined = inline strings. */
  readonly sst: SharedStrings | undefined;
  /** Hyperlinks collected while encoding (emitted in the sheet tail). */
  readonly links: LinkRec[];
  /** Set when any formula cell was written. */
  formulas: boolean;
}

export function newSheetCtx(sst: SharedStrings | undefined): SheetCtx {
  return { sst, links: [], formulas: false };
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
  let out = `<c r="${ref}"${s} t="inlineStr"><is>`;
  const runs: readonly RichTextRun[] = rt.richText;
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i]!;
    const f = run.font;
    out += f
      ? `<r><rPr>${fontBody(f, 'rFont')}</rPr>${tElem(run.text)}</r>`
      : `<r>${tElem(run.text)}</r>`;
  }
  return `${out}</is></c>`;
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

// ---- head ------------------------------------------------------------------------------

export interface SheetHeadOptions {
  /** Used range ("A1:C10"); omitted when undefined (streaming). */
  readonly dimension?: string | undefined;
  readonly tabSelected: boolean;
  /** Max row outline level (sheetFormatPr outlineLevelRow); 0 = unknown/none. */
  readonly outlineLevelRow?: number | undefined;
}

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

function colsXml(cols: readonly (Column | undefined)[]): { xml: string; outline: number } {
  let s = '';
  let outline = 0;
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (!c) continue;
    if (c.width === undefined && !c.hidden && !c.outlineLevel && !c.styleIndex) continue;
    let j = i;
    while (j + 1 < cols.length && cols[j + 1] && sameCol(c, cols[j + 1]!)) j++;
    if (c.outlineLevel > outline) outline = c.outlineLevel;
    s += `<col min="${i + 1}" max="${j + 1}" width="${c.width ?? 9.140625}"${
      c.width !== undefined ? ' customWidth="1"' : ''
    }${c.hidden ? ' hidden="1"' : ''}${c.outlineLevel ? ` outlineLevel="${c.outlineLevel}"` : ''}${
      c.styleIndex ? ` style="${c.styleIndex}"` : ''
    }/>`;
    i = j;
  }
  return { xml: s ? `<cols>${s}</cols>` : '', outline };
}

/** Everything up to and including `<sheetData>`. */
export function sheetHead(ws: Worksheet, o: SheetHeadOptions): string {
  let s = `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`;
  if (ws.tabColor) s += `<sheetPr><tabColor${colorAttrs(ws.tabColor)}/></sheetPr>`;
  if (o.dimension) s += `<dimension ref="${o.dimension}"/>`;
  s += `<sheetViews>${viewXml(ws.views[0], o.tabSelected)}</sheetViews>`;
  const cols = colsXml(ws.columnData);
  s += '<sheetFormatPr defaultRowHeight="15"';
  if (o.outlineLevelRow) s += ` outlineLevelRow="${o.outlineLevelRow}"`;
  if (cols.outline) s += ` outlineLevelCol="${cols.outline}"`;
  return `${s}/>${cols.xml}<sheetData>`;
}

/** `</sheetData>` through `</worksheet>`. Hyperlink r:ids match {@link sheetRelsXml}. */
export function sheetTail(ws: Worksheet, ctx: SheetCtx): string {
  let s = '</sheetData>';
  if (ws.autoFilter) s += `<autoFilter ref="${escapeAttr(ws.autoFilter)}"/>`;
  const merges = ws.mergeRanges;
  if (merges.length > 0) {
    s += `<mergeCells count="${merges.length}">`;
    for (const m of merges) s += `<mergeCell ref="${encodeRange(m)}"/>`;
    s += '</mergeCells>';
  }
  if (ctx.links.length > 0) {
    s += '<hyperlinks>';
    let id = 0;
    for (const l of ctx.links) {
      const tip = l.tooltip !== undefined ? ` tooltip="${escapeAttr(l.tooltip)}"` : '';
      s +=
        l.target.charCodeAt(0) === 35
          ? `<hyperlink ref="${l.ref}" location="${escapeAttr(l.target.slice(1))}"${tip}/>`
          : `<hyperlink ref="${l.ref}" r:id="rId${++id}"${tip}/>`;
    }
    s += '</hyperlinks>';
  }
  return `${s}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

/** True when the sheet needs a rels part (external hyperlinks). */
export function hasExternalLinks(ctx: SheetCtx): boolean {
  for (const l of ctx.links) if (l.target.charCodeAt(0) !== 35) return true;
  return false;
}
