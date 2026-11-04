// Feature: print setup — <printOptions>, <pageMargins>, <pageSetup>, <headerFooter> and
// <sheetPr><pageSetUpPr fitToPage/>.
//
//   ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1,
//                    fitToHeight: 0, margins: { left: 0.5 }, horizontalCentered: true };
//   ws.headerFooter = { oddHeader: '&C&BReport', oddFooter: '&CPage &P of &N' };
//
// pageMargins is always written (Excel's defaults 0.7/0.7/0.75/0.75/0.3/0.3 fill missing
// values). Print areas / print titles are defined names: use workbook.definedNames
// ('_xlnm.Print_Area' / '_xlnm.Print_Titles' with localSheetId).

import type { Worksheet } from '../../core/worksheet.js';
import {
  type FeatureReader,
  type FeatureWriter,
  SheetPrOrder,
  SheetSlot,
  type SubtreeHandler,
} from './types.js';
import { attrs, bool, escapeText, num } from './util.js';

export interface PageMargins {
  left?: number | undefined;
  right?: number | undefined;
  top?: number | undefined;
  bottom?: number | undefined;
  header?: number | undefined;
  footer?: number | undefined;
}

export interface PageSetup {
  margins?: PageMargins | undefined;
  orientation?: 'default' | 'portrait' | 'landscape' | undefined;
  /** Excel paper size code (1 = Letter, 9 = A4, ...). */
  paperSize?: number | undefined;
  horizontalDpi?: number | undefined;
  verticalDpi?: number | undefined;
  pageOrder?: 'downThenOver' | 'overThenDown' | undefined;
  blackAndWhite?: boolean | undefined;
  draft?: boolean | undefined;
  cellComments?: 'none' | 'asDisplayed' | 'atEnd' | undefined;
  errors?: 'displayed' | 'blank' | 'dash' | 'NA' | undefined;
  /** Zoom percentage (10-400); ignored by Excel when fitToPage is set. */
  scale?: number | undefined;
  /** Fit to `fitToWidth` x `fitToHeight` pages (sheetPr/pageSetUpPr). */
  fitToPage?: boolean | undefined;
  /** Pages wide (0 = automatic). */
  fitToWidth?: number | undefined;
  /** Pages tall (0 = automatic). */
  fitToHeight?: number | undefined;
  firstPageNumber?: number | undefined;
  usePrinterDefaults?: boolean | undefined;
  copies?: number | undefined;
  // printOptions
  horizontalCentered?: boolean | undefined;
  verticalCentered?: boolean | undefined;
  /** Print gridlines. */
  showGridLines?: boolean | undefined;
  /** Print row and column headings. */
  showRowColHeaders?: boolean | undefined;
}

export interface HeaderFooter {
  differentFirst?: boolean | undefined;
  differentOddEven?: boolean | undefined;
  oddHeader?: string | undefined;
  oddFooter?: string | undefined;
  evenHeader?: string | undefined;
  evenFooter?: string | undefined;
  firstHeader?: string | undefined;
  firstFooter?: string | undefined;
}

const HF_PARTS = [
  'oddHeader',
  'oddFooter',
  'evenHeader',
  'evenFooter',
  'firstHeader',
  'firstFooter',
] as const;

function printOptions(ws: Worksheet): string {
  const p = ws.pageSetup;
  if (!p) return '';
  const s = attrs({
    headings: p.showRowColHeaders,
    gridLines: p.showGridLines,
    horizontalCentered: p.horizontalCentered,
    verticalCentered: p.verticalCentered,
  });
  return s ? `<printOptions${s}/>` : '';
}

function pageMargins(ws: Worksheet): string {
  const m = ws.pageSetup?.margins;
  return `<pageMargins${attrs({
    left: m?.left ?? 0.7,
    right: m?.right ?? 0.7,
    top: m?.top ?? 0.75,
    bottom: m?.bottom ?? 0.75,
    header: m?.header ?? 0.3,
    footer: m?.footer ?? 0.3,
  })}/>`;
}

function pageSetup(ws: Worksheet): string {
  const p = ws.pageSetup;
  if (!p) return '';
  const s = attrs({
    paperSize: p.paperSize,
    scale: p.scale,
    firstPageNumber: p.firstPageNumber,
    fitToWidth: p.fitToWidth,
    fitToHeight: p.fitToHeight,
    pageOrder: p.pageOrder,
    orientation: p.orientation,
    usePrinterDefaults: p.usePrinterDefaults,
    blackAndWhite: p.blackAndWhite,
    draft: p.draft,
    cellComments: p.cellComments,
    useFirstPageNumber: p.firstPageNumber !== undefined,
    errors: p.errors,
    horizontalDpi: p.horizontalDpi,
    verticalDpi: p.verticalDpi,
    copies: p.copies,
  });
  return s ? `<pageSetup${s}/>` : '';
}

function headerFooter(ws: Worksheet): string {
  const h = ws.headerFooter;
  if (!h) return '';
  let body = '';
  for (const k of HF_PARTS) {
    const v = h[k];
    if (v) body += `<${k}>${escapeText(v)}</${k}>`;
  }
  const a = attrs({ differentOddEven: h.differentOddEven, differentFirst: h.differentFirst });
  return body || a ? `<headerFooter${a}>${body}</headerFooter>` : '';
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

function setup(ws: Worksheet): Mutable<PageSetup> {
  ws.pageSetup ??= {};
  return ws.pageSetup;
}

export const printSetupWriter: FeatureWriter = {
  name: 'print-setup',
  sheet: {
    emitters: [
      {
        slot: SheetSlot.sheetPr,
        order: SheetPrOrder.pageSetUpPr,
        emit: (c) => (c.ws.pageSetup?.fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ''),
      },
      { slot: SheetSlot.printOptions, emit: (c) => printOptions(c.ws) },
      { slot: SheetSlot.pageMargins, emit: (c) => pageMargins(c.ws) },
      { slot: SheetSlot.pageSetup, emit: (c) => pageSetup(c.ws) },
      { slot: SheetSlot.headerFooter, emit: (c) => headerFooter(c.ws) },
    ],
  },
};

export const printSetupReader: FeatureReader = {
  name: 'print-setup',
  sheet: {
    elements: ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'sheetPr/pageSetUpPr'],
    open(ctx, el, a): SubtreeHandler | undefined {
      const ws = ctx.ws;
      switch (el) {
        case 'sheetPr/pageSetUpPr':
          if (bool(a.fitToPage)) setup(ws).fitToPage = true;
          return undefined;
        case 'printOptions': {
          const p = setup(ws);
          if (bool(a.headings)) p.showRowColHeaders = true;
          if (bool(a.gridLines)) p.showGridLines = true;
          if (bool(a.horizontalCentered)) p.horizontalCentered = true;
          if (bool(a.verticalCentered)) p.verticalCentered = true;
          return undefined;
        }
        case 'pageMargins':
          setup(ws).margins = {
            left: num(a.left) ?? 0.7,
            right: num(a.right) ?? 0.7,
            top: num(a.top) ?? 0.75,
            bottom: num(a.bottom) ?? 0.75,
            header: num(a.header) ?? 0.3,
            footer: num(a.footer) ?? 0.3,
          };
          return undefined;
        case 'pageSetup': {
          const p = setup(ws);
          const n = (k: string): number | undefined => num(a[k]);
          if (a.paperSize !== undefined) p.paperSize = n('paperSize');
          if (a.scale !== undefined) p.scale = n('scale');
          if (a.fitToWidth !== undefined) p.fitToWidth = n('fitToWidth');
          if (a.fitToHeight !== undefined) p.fitToHeight = n('fitToHeight');
          if (bool(a.useFirstPageNumber) && a.firstPageNumber !== undefined) {
            p.firstPageNumber = n('firstPageNumber');
          }
          if (a.pageOrder) p.pageOrder = a.pageOrder as PageSetup['pageOrder'];
          if (a.orientation) p.orientation = a.orientation as PageSetup['orientation'];
          if (bool(a.usePrinterDefaults)) p.usePrinterDefaults = true;
          if (bool(a.blackAndWhite)) p.blackAndWhite = true;
          if (bool(a.draft)) p.draft = true;
          if (a.cellComments) p.cellComments = a.cellComments as PageSetup['cellComments'];
          if (a.errors) p.errors = a.errors as PageSetup['errors'];
          if (a.horizontalDpi !== undefined) p.horizontalDpi = n('horizontalDpi');
          if (a.verticalDpi !== undefined) p.verticalDpi = n('verticalDpi');
          if (a.copies !== undefined) p.copies = n('copies');
          return undefined;
        }
        case 'headerFooter': {
          const h: Mutable<HeaderFooter> = {};
          if (bool(a.differentFirst)) h.differentFirst = true;
          if (bool(a.differentOddEven)) h.differentOddEven = true;
          ws.headerFooter = h;
          let cur: (typeof HF_PARTS)[number] | undefined;
          let text = '';
          return {
            open(name) {
              cur = (HF_PARTS as readonly string[]).includes(name)
                ? (name as typeof cur)
                : undefined;
              text = '';
            },
            text(t) {
              if (cur) text += t;
            },
            close() {
              if (cur) h[cur] = text;
              cur = undefined;
            },
          };
        }
      }
      return undefined;
    },
  },
};
