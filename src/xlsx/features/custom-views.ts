// Feature: custom views (View > Custom Views) — `<customWorkbookViews>` in workbook.xml (one
// named view each, with the window it restores) and `<customSheetViews>` in each worksheet (the
// sheet's settings in that view, linked by `guid`).
//
//   addCustomView(wb, 'Summary');                        // every sheet as it is now
//   wb.customViews; ws.customSheetViews;                 // read from / written to files
//
// Model: `wb.customViews` (name, guid, window) and `ws.customSheetViews` (per view: zoom,
// gridlines, hidden state, view mode, frozen pane / selection, print settings ...). The
// elements a sheet view holds besides its attributes (pane, selection, page breaks, margins,
// page setup, header / footer, autoFilter) are kept as XML (`children`). Hidden rows / columns,
// filters and print areas of a view are stored by Excel as defined names
// ("Z_{GUID}_.wvu.Rows", ...), which round-trip with `wb.definedNames`.

import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { escapeAttr, escapeText, localName, parseXml, type XmlAttrs } from '../../xml/index.js';
import {
  BookSlot,
  type FeatureReader,
  type FeatureWriter,
  type PackageReadCtx,
  SheetSlot,
} from './types.js';

/** A custom view of the workbook (`<customWorkbookView>`). */
export interface CustomView {
  name: string;
  /** GUID ("{...}") linking the sheets' `customSheetViews`. */
  guid: string;
  /** Window size in twips (default 16000 x 9000). */
  windowWidth?: number | undefined;
  windowHeight?: number | undefined;
  /** `sheetId` of the active sheet (Worksheet.id; default the first sheet). */
  activeSheetId?: number | undefined;
  /** The view includes print settings (default true in Excel's dialog). */
  includePrintSettings?: boolean | undefined;
  /** The view includes hidden rows, columns and filter settings (default true). */
  includeHiddenRowCol?: boolean | undefined;
  /** Other `<customWorkbookView>` attributes, as written in the file (xWindow, maximized, ...). */
  attributes?: Record<string, string> | undefined;
}

/** A sheet's settings in one custom view (`<customSheetView>`). */
export interface CustomSheetView {
  /** GUID of the `wb.customViews` entry. */
  guid: string;
  /** Zoom (%, default 100). */
  scale?: number | undefined;
  showGridLines?: boolean | undefined;
  showRowCol?: boolean | undefined;
  showFormulas?: boolean | undefined;
  showPageBreaks?: boolean | undefined;
  zeroValues?: boolean | undefined;
  outlineSymbols?: boolean | undefined;
  showRuler?: boolean | undefined;
  fitToPage?: boolean | undefined;
  printArea?: boolean | undefined;
  filter?: boolean | undefined;
  showAutoFilter?: boolean | undefined;
  hiddenRows?: boolean | undefined;
  hiddenColumns?: boolean | undefined;
  filterUnique?: boolean | undefined;
  state?: 'visible' | 'hidden' | 'veryHidden' | undefined;
  view?: 'normal' | 'pageBreakPreview' | 'pageLayout' | undefined;
  topLeftCell?: string | undefined;
  colorId?: number | undefined;
  /** Child elements as XML (`<pane/>`, `<selection/>`, `<pageMargins/>`, `<pageSetup/>`, ...). */
  children?: string | undefined;
}

const BOOLS = [
  'showGridLines',
  'showRowCol',
  'showFormulas',
  'showPageBreaks',
  'zeroValues',
  'outlineSymbols',
  'showRuler',
  'fitToPage',
  'printArea',
  'filter',
  'showAutoFilter',
  'hiddenRows',
  'hiddenColumns',
  'filterUnique',
] as const;

/** A stable GUID derived from a seed. */
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

/**
 * Add a custom view named `name` that restores every sheet as it is now (zoom, gridlines,
 * frozen panes from `ws.views[0]`, hidden state); returns it. Replaces a view of the same name.
 */
export function addCustomView(
  wb: Workbook,
  name: string,
  options: Partial<Omit<CustomView, 'name'>> = {},
): CustomView {
  const old = wb.customViews?.find((v) => v.name.toLowerCase() === name.toLowerCase());
  const id = options.guid ?? old?.guid ?? guid(`view:${name}`);
  const v: CustomView = { name, guid: id, ...options };
  wb.customViews = [...(wb.customViews ?? []).filter((x) => x !== old), v];
  for (const ws of wb.worksheets) {
    const w = ws.views[0];
    const sv: CustomSheetView = { guid: id };
    if (w?.zoomScale) sv.scale = w.zoomScale;
    if (w?.showGridLines === false) sv.showGridLines = false;
    if (w?.showRowColHeaders === false) sv.showRowCol = false;
    if (w?.style) sv.view = w.style;
    if (ws.state !== 'visible') sv.state = ws.state;
    if (w && (w.state === 'frozen' || w.state === 'split') && (w.xSplit || w.ySplit)) {
      const x = w.xSplit ?? 0;
      const y = w.ySplit ?? 0;
      const pane = x && y ? 'bottomRight' : y ? 'bottomLeft' : 'topRight';
      sv.children = `<pane${x ? ` xSplit="${x}"` : ''}${y ? ` ySplit="${y}"` : ''}${w.topLeftCell ? ` topLeftCell="${escapeAttr(w.topLeftCell)}"` : ''} activePane="${pane}"${w.state === 'frozen' ? ' state="frozen"' : ''}/><selection pane="${pane}"/>`;
    }
    ws.customSheetViews = [...(ws.customSheetViews ?? []).filter((x) => x.guid !== id), sv];
  }
  return v;
}

// ---------------------------------------------------------------- writing

function sheetViewXml(v: CustomSheetView): string {
  let a = ` guid="${escapeAttr(v.guid)}"`;
  if (v.scale !== undefined) a += ` scale="${v.scale}"`;
  if (v.colorId !== undefined) a += ` colorId="${v.colorId}"`;
  for (const k of BOOLS) if (v[k] !== undefined) a += ` ${k}="${v[k] ? 1 : 0}"`;
  if (v.state && v.state !== 'visible') a += ` state="${v.state}"`;
  if (v.view && v.view !== 'normal') a += ` view="${v.view}"`;
  if (v.topLeftCell) a += ` topLeftCell="${escapeAttr(v.topLeftCell)}"`;
  // pageMargins is required by Excel's own files; keep what the view has, else add defaults
  const kids = v.children ?? '';
  const margins = /<pageMargins\b/.test(kids)
    ? ''
    : '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
  return `<customSheetView${a}>${kids}${margins}</customSheetView>`;
}

function bookViewXml(wb: Workbook, v: CustomView): string {
  let a = ` name="${escapeAttr(v.name)}" guid="${escapeAttr(v.guid)}"`;
  const extra = v.attributes ?? {};
  for (const [k, x] of Object.entries(extra)) a += ` ${k}="${escapeAttr(x)}"`;
  if (v.includePrintSettings === false) a += ' includePrintSettings="0"';
  if (v.includeHiddenRowCol === false) a += ' includeHiddenRowCol="0"';
  a += ` windowWidth="${v.windowWidth ?? 16000}" windowHeight="${v.windowHeight ?? 9000}"`;
  a += ` activeSheetId="${v.activeSheetId ?? wb.worksheets[0]?.id ?? 1}"`;
  return `<customWorkbookView${a}/>`;
}

/**
 * Write half of the custom views feature (wb.customViews, ws.customSheetViews). Always on in the
 * package root; in `./lite` pass it in `features`.
 */
export const customViewsWriter: FeatureWriter = {
  name: 'custom-views',
  sheet: {
    emitters: [
      {
        slot: SheetSlot.customSheetViews,
        emit(c) {
          const list = c.ws.customSheetViews;
          if (!list?.length) return '';
          return `<customSheetViews>${list.map(sheetViewXml).join('')}</customSheetViews>`;
        },
      },
    ],
  },
  book: {
    emitters: [
      {
        slot: BookSlot.customWorkbookViews,
        emit(c) {
          const list = c.workbook.customViews;
          if (!list?.length) return '';
          return `<customWorkbookViews>${list.map((v) => bookViewXml(c.workbook, v)).join('')}</customWorkbookViews>`;
        },
        owns: (wb) => wb.customViews !== undefined,
      },
    ],
  },
};

// ---------------------------------------------------------------- reading

const flag = (v: string): boolean => v === '1' || v === 'true';

/** Serializes a delegated subtree back to XML (local names; the attributes as read). */
function attrsXml(a: XmlAttrs): string {
  let s = '';
  for (const k in a) s += ` ${k}="${escapeAttr(a[k]!)}"`;
  return s;
}

function sheetView(a: XmlAttrs): CustomSheetView {
  const v: CustomSheetView = { guid: a.guid ?? '' };
  if (a.scale !== undefined) v.scale = +a.scale;
  if (a.colorId !== undefined) v.colorId = +a.colorId;
  for (const k of BOOLS) if (a[k] !== undefined) v[k] = flag(a[k]);
  if (a.state === 'hidden' || a.state === 'veryHidden' || a.state === 'visible') v.state = a.state;
  if (a.view === 'pageBreakPreview' || a.view === 'pageLayout' || a.view === 'normal')
    v.view = a.view;
  if (a.topLeftCell) v.topLeftCell = a.topLeftCell;
  return v;
}

const KNOWN = new Set([
  'name',
  'guid',
  'windowWidth',
  'windowHeight',
  'activeSheetId',
  'includePrintSettings',
  'includeHiddenRowCol',
]);

function readBookViews(ctx: PackageReadCtx): void {
  const root = [...(ctx.relsOf('')?.values() ?? [])].find((r) =>
    r.type.endsWith('/officeDocument'),
  );
  const xml = root ? ctx.readText(root.path) : undefined;
  const m =
    xml && /<(?:\w+:)?customWorkbookViews\b[\s\S]*?<\/(?:\w+:)?customWorkbookViews>/.exec(xml);
  if (!m) return;
  const out: CustomView[] = [];
  parseXml(m[0], {
    open(name, a) {
      if (localName(name) !== 'customWorkbookView' || !a.name || !a.guid) return;
      const v: CustomView = { name: a.name, guid: a.guid };
      if (a.windowWidth !== undefined) v.windowWidth = +a.windowWidth;
      if (a.windowHeight !== undefined) v.windowHeight = +a.windowHeight;
      if (a.activeSheetId !== undefined) v.activeSheetId = +a.activeSheetId;
      if (a.includePrintSettings !== undefined)
        v.includePrintSettings = flag(a.includePrintSettings);
      if (a.includeHiddenRowCol !== undefined) v.includeHiddenRowCol = flag(a.includeHiddenRowCol);
      const rest: Record<string, string> = {};
      for (const k in a) if (!KNOWN.has(k) && !k.startsWith('xmlns')) rest[k] = a[k]!;
      if (Object.keys(rest).length) v.attributes = rest;
      out.push(v);
    },
  });
  ctx.workbook.customViews = out;
}

/**
 * Read half of the custom views feature. Always on in the package root; in `./lite` pass it in
 * `features` of readXlsx (without it, the elements are kept raw and written back).
 */
export const customViewsReader: FeatureReader = {
  name: 'custom-views',
  sheet: {
    elements: ['customSheetViews'],
    open(ctx) {
      const ws: Worksheet = ctx.ws;
      const list: CustomSheetView[] = [];
      let cur: CustomSheetView | undefined;
      let depth = 0;
      let kids = '';
      ws.customSheetViews = list;
      return {
        open(name, a, selfClosing) {
          const ln = localName(name);
          if (depth === 0 && ln === 'customSheetView') {
            cur = sheetView(a);
            kids = '';
            if (selfClosing) {
              list.push(cur);
              cur = undefined;
            } else depth = 1;
            return;
          }
          if (!cur) return;
          kids += `<${ln}${attrsXml(a)}${selfClosing ? '/>' : '>'}`;
          if (!selfClosing) depth++;
        },
        text(t) {
          if (cur && depth > 1) kids += escapeText(t);
        },
        close(name) {
          if (!cur) return;
          depth--;
          if (depth === 0) {
            if (kids) cur.children = kids;
            list.push(cur);
            cur = undefined;
          } else kids += `</${localName(name)}>`;
        },
      };
    },
  },
  book: readBookViews,
};
