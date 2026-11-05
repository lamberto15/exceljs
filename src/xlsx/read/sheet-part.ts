// Worksheet XML reader (the hot loop). One handler serves both modes:
//  - DOM mode (`ws` given): cells go through ws.setCellRaw; rows/cols/merges/views/autoFilter/
//    hyperlinks are applied to the worksheet.
//  - Row mode (`onRow` given): each <row> is converted to a sparse 0-based CellValue[] and
//    handed to the callback when the row closes; sheet-level features are ignored
//    (hyperlinks live after <sheetData>, so they cannot be attached to streamed rows).

import { decodeRange, encodeAddr, MAX_ROWS } from '../../core/address.js';
import { dateToSerial, serialToDate } from '../../core/dates.js';
import type { StyleRegistry } from '../../core/styles.js';
import type { CellValue, ErrorCode, Formula, Hyperlink, RichText } from '../../core/types.js';
import {
  CellKind,
  type Worksheet,
  type WorksheetProperties,
  type WorksheetView,
} from '../../core/worksheet.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import type { Relationship } from '../common/r-opc.js';
import { stripXlfn } from '../common/xlfn.js';
import type { SheetReadCtx, SheetReader, SubtreeHandler } from '../features/types.js';
import { translateFormula } from './formula.js';
import { type StringItem, StringItemBuilder } from './strings-part.js';
import { parseColor } from './styles-part.js';
import { attrBool, attrLocal, decodeXString, SheetLimitError } from './xml-util.js';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

export interface SheetContext {
  readonly strings: readonly StringItem[];
  /** File cellXfs index → registry xf; undefined when the file has no styles part. */
  readonly xfMap: Uint32Array | undefined;
  readonly styles: StyleRegistry;
  readonly date1904: boolean;
  /** Row mode: Dates by local wall clock. */
  readonly localDates?: boolean | undefined;
  /** The sheet's relationships (hyperlink targets). */
  readonly rels?: Map<string, Relationship> | undefined;
  /** Feature readers by element path ("pageSetup", "sheetPr/pageSetUpPr"); DOM mode only. */
  readonly features?: ReadonlyMap<string, SheetReader> | undefined;
  readonly featureCtx?: SheetReadCtx | undefined;
  /** Round-trip capture of elements nobody models (DOM mode, preserveUnknown). */
  readonly capture?: SheetCapture | undefined;
  /** Sheet name (error messages). */
  readonly name?: string | undefined;
  /** ReadOptions.maxRows / maxCols. */
  readonly maxRows?: number | undefined;
  readonly maxCols?: number | undefined;
}

/** Raw subtree receiver: full (prefixed) names; `end()` when the element closes. */
export interface RawHandler {
  open(name: string, attrs: XmlAttrs, selfClosing: boolean): void;
  close(name: string): void;
  text(text: string): void;
  end(): void;
}

/**
 * Receives what the sheet parser does not model (outside sheetData): the root's attributes
 * (namespace scope), `<sheetPr>` attributes, the attributes of consumed elements (to track which
 * relationships they reference) and whole unknown elements (`start`, by path below <worksheet>:
 * "rowBreaks", "sheetPr/outlinePr", "extLst/ext").
 */
export interface SheetCapture {
  root(a: XmlAttrs): void;
  sheetPr(a: XmlAttrs): void;
  /** `<sheetFormatPr>` attributes (kept when they say more than the writer's defaults). */
  formatPr(a: XmlAttrs): void;
  ref(a: XmlAttrs): void;
  /**
   * A hyperlink on a formula cell (a cell holds one kind, so it is not modelled): kept raw and
   * written back while the cell still holds `formula` (the same object).
   */
  formulaLink?(ref: string, formula: unknown, h: HyperlinkAttrs): void;
  start(path: string, name: string, a: XmlAttrs, selfClosing: boolean): RawHandler;
}

/** Worksheet children the core parser handles itself. */
const BUILTIN = new Set([
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'autoFilter',
  'mergeCells',
  'hyperlinks',
  'extLst',
]);

/** Placeholder subtree handler for consumed leaf elements. */
const LEAF: SubtreeHandler = {};

export type RowCallback = (r0: number, values: CellValue[]) => void;

enum Cap {
  None = 0,
  V = 1,
  F = 2,
}

interface ColRange {
  min: number;
  max: number;
  width: number | undefined;
  hidden: boolean;
  outlineLevel: number;
  collapsed: boolean;
  xf: number;
}

/** `<hyperlink>` attributes (relationship id resolved by the capture). */
export interface HyperlinkAttrs {
  rId: string | undefined;
  location: string | undefined;
  tooltip: string | undefined;
  display: string | undefined;
}

interface HyperlinkRec extends HyperlinkAttrs {
  ref: string;
}

/** Above this many columns, a <col> range is only applied up to the last used column. */
const MAX_COL_RANGE = 1024;

/** Parse an ISO-8601 date/time (t="d"). Values without a zone designator are taken as UTC. */
function parseIsoDate(s: string): Date | undefined {
  let t = s.trim();
  if (t.length > 10 && t.indexOf('T') > 0 && !/(Z|[+-]\d\d:?\d\d)$/i.test(t)) t += 'Z';
  else if (/^\d\d:\d\d/.test(t)) t = `1899-12-30T${t}Z`;
  const d = new Date(t);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

export class SheetHandler implements XmlHandler {
  private readonly ctx: SheetContext;
  private readonly ws: Worksheet | undefined;
  private readonly onRow: RowCallback | undefined;
  /** Row mode: merged ranges seen (A1 refs). */
  readonly merges: string[] = [];
  private readonly strings: readonly StringItem[];
  private readonly xfMap: Uint32Array | undefined;
  private readonly styles: StyleRegistry;
  /** maxRows / maxCols (Infinity when unset). */
  private readonly maxR: number;
  private readonly maxC: number;

  private depth = 0;
  private inData = false;
  private inViews = false;
  private row = -1;
  private col = -1;

  // current cell
  private cRow = 0;
  private cCol = 0;
  private cType: string | undefined;
  private cXf = 0;
  private cap: Cap = Cap.None;
  private vText = '';
  private hasV = false;
  private fText: string | undefined;
  private fType: string | undefined;
  /** The cell has cell metadata (cm): a dynamic-array formula. */
  private cMeta = false;
  private fRef: string | undefined;
  private fSi: string | undefined;
  private inIs = false;
  private hadIs = false;
  private readonly is = new StringItemBuilder();

  // row mode
  private values: CellValue[] | undefined;

  // DOM-only collections
  private readonly shared = new Map<string, { formula: string; r: number; c: number }>();
  private readonly cols: ColRange[] = [];
  private readonly links: HyperlinkRec[] = [];
  private readonly views: Mutable<WorksheetView>[] = [];
  private maxCol = -1;
  /** Cells stored in the current row / in the previous non-empty row (DOM mode sizing hint). */
  private rowStored = 0;
  private hint = 0;

  // feature delegation (outside sheetData only)
  private readonly fmap: ReadonlyMap<string, SheetReader> | undefined;
  private sub: SubtreeHandler | undefined;
  private subDepth = 0;
  private readonly capt: SheetCapture | undefined;
  private raw: RawHandler | undefined;
  private rawDepth = 0;
  /** Local name of the current depth-2 element (for "parent/child" feature paths). */
  private el2 = '';

  constructor(ctx: SheetContext, ws: Worksheet | undefined, onRow?: RowCallback) {
    this.ctx = ctx;
    this.ws = ws;
    this.onRow = onRow;
    this.strings = ctx.strings;
    this.xfMap = ctx.xfMap;
    this.styles = ctx.styles;
    this.maxR = ctx.maxRows! > 0 ? ctx.maxRows! : Infinity;
    this.maxC = ctx.maxCols! > 0 ? ctx.maxCols! : Infinity;
    this.fmap = ws !== undefined && ctx.features?.size ? ctx.features : undefined;
    this.capt = ws !== undefined ? ctx.capture : undefined;
  }

  open(name: string, a: XmlAttrs, _selfClosing: boolean): void {
    const ln = name.indexOf(':') < 0 ? name : localName(name);
    const d = ++this.depth;
    if (this.inData) {
      switch (ln) {
        case 'c':
          this.openCell(a);
          return;
        case 'v':
          this.cap = Cap.V;
          return;
        case 'f':
          this.cap = Cap.F;
          this.fText = '';
          this.fType = a.t;
          this.fRef = a.ref;
          this.fSi = a.si;
          return;
        case 'row':
          this.openRow(a);
          return;
        case 'is':
          this.inIs = true;
          this.hadIs = true;
          this.is.reset();
          return;
        default:
          if (this.inIs) this.is.open(ln, a);
          return;
      }
    }
    if (this.raw !== undefined) {
      this.raw.open(name, a, _selfClosing);
      return;
    }
    const cap = this.capt;
    if (this.sub !== undefined) {
      if (cap) cap.ref(a);
      this.sub.open?.(ln, a, _selfClosing);
      return;
    }
    if (this.fmap !== undefined && d <= 3) {
      const key = d === 2 ? ln : `${this.el2}/${ln}`;
      const r = this.fmap.get(key);
      if (r !== undefined) {
        const h = r.open?.(this.ctx.featureCtx!, key, a);
        // an "extLst/ext" nobody claims (all readers declined its uri) stays unknown
        if (h !== undefined || key !== 'extLst/ext' || !cap) {
          if (cap) cap.ref(a);
          this.sub = h ?? LEAF;
          this.subDepth = d;
          return;
        }
      }
    }
    if (cap !== undefined) {
      if (d === 1) cap.root(a);
      else if (
        d === 2
          ? !BUILTIN.has(ln)
          : d === 3 && ((this.el2 === 'sheetPr' && ln !== 'tabColor') || this.el2 === 'extLst')
      ) {
        this.raw = cap.start(d === 2 ? ln : `${this.el2}/${ln}`, name, a, _selfClosing);
        this.rawDepth = d;
        this.raw.open(name, a, _selfClosing);
        return;
      } else {
        if (d === 2 && ln === 'sheetPr') cap.sheetPr(a);
        else if (d === 2 && ln === 'sheetFormatPr') cap.formatPr(a);
        cap.ref(a);
      }
    }
    if (d === 2) {
      this.el2 = ln;
      switch (ln) {
        case 'sheetData':
          this.inData = true;
          return;
        case 'sheetViews':
          this.inViews = true;
          return;
        case 'sheetFormatPr':
          if (this.ws) readFormatPr(this.ws.properties, a);
          return;
        case 'autoFilter':
          if (this.ws && a.ref) this.ws.autoFilter = a.ref;
          // criteria (filterColumn, sortState) are kept raw next to the modelled range
          if (cap && !_selfClosing) {
            this.raw = cap.start(ln, name, a, false);
            this.rawDepth = d;
            this.raw.open(name, a, false);
          }
          return;
      }
      return;
    }
    if (this.ws === undefined) {
      if (ln === 'mergeCell' && a.ref) this.merges.push(a.ref);
      return;
    }
    switch (ln) {
      case 'col':
        if (d === 3) this.addCol(a);
        return;
      case 'mergeCell':
        if (a.ref) {
          try {
            this.ws.mergeCells(a.ref);
          } catch {
            // single-cell / overlapping / malformed merges are dropped
          }
        }
        return;
      case 'hyperlink':
        if (a.ref) {
          this.links.push({
            ref: a.ref,
            rId: attrLocal(a, 'id'),
            location: a.location,
            tooltip: a.tooltip,
            display: a.display,
          });
        }
        return;
      case 'tabColor': {
        const c = parseColor(a);
        if (c) this.ws.tabColor = c;
        return;
      }
    }
    if (this.inViews) this.openView(ln, d, a);
  }

  close(name: string): void {
    const d = this.depth--;
    if (this.inData) {
      const ln = name.indexOf(':') < 0 ? name : localName(name);
      switch (ln) {
        case 'c':
          this.closeCell();
          return;
        case 'v':
        case 'f':
          this.cap = Cap.None;
          return;
        case 'row':
          if (this.values !== undefined) {
            this.onRow!(this.row, this.values);
            this.values = undefined;
          }
          return;
        case 'is':
          this.inIs = false;
          return;
        case 'sheetData':
          this.inData = false;
          return;
        default:
          if (this.inIs) this.is.close(ln);
          return;
      }
    }
    if (this.raw !== undefined) {
      if (d === this.rawDepth) {
        this.raw.close(name);
        this.raw.end();
        this.raw = undefined;
      } else this.raw.close(name);
      return;
    }
    if (this.sub !== undefined) {
      if (d === this.subDepth) {
        this.sub.end?.();
        this.sub = undefined;
      } else this.sub.close?.(name.indexOf(':') < 0 ? name : localName(name));
      return;
    }
    if (d === 2 && this.inViews) {
      this.inViews = false;
    }
  }

  text(t: string): void {
    if (this.cap === Cap.V) {
      this.vText = this.hasV ? this.vText + t : t;
      this.hasV = true;
    } else if (this.cap === Cap.F) this.fText += t;
    else if (this.inIs) this.is.onText(t);
    else if (this.raw !== undefined) this.raw.text(t);
    else if (this.sub !== undefined) this.sub.text?.(t);
  }

  /** Apply deferred sheet-level data (DOM mode). Call after the last chunk. */
  finish(): void {
    const ws = this.ws;
    if (!ws) return;
    for (const c of this.cols) {
      let max = c.max;
      if (max - c.min >= MAX_COL_RANGE) max = Math.max(c.min, Math.min(max, this.maxCol));
      for (let i = c.min; i <= max; i++) {
        const col = ws.getColumn(i + 1);
        if (c.width !== undefined) col.width = c.width;
        if (c.hidden) col.hidden = true;
        if (c.outlineLevel) col.outlineLevel = c.outlineLevel;
        if (c.collapsed) col.collapsed = true;
        if (c.xf) col.styleIndex = c.xf;
      }
    }
    if (this.views.length) ws.views = this.views;
    for (const h of this.links) this.applyLink(ws, h);
  }

  // ---- rows & cells ----

  private openRow(a: XmlAttrs): void {
    const r = a.r;
    this.row = r !== undefined ? +r - 1 : this.row + 1;
    this.col = -1;
    if (this.onRow) return;
    if (this.rowStored > 0) this.hint = this.rowStored;
    this.rowStored = 0;
    const ht = a.ht;
    const hidden = a.hidden !== undefined && attrBool(a.hidden);
    const ol = a.outlineLevel !== undefined ? +a.outlineLevel : 0;
    const collapsed = a.collapsed !== undefined && attrBool(a.collapsed);
    const height = ht !== undefined && attrBool(a.customHeight) ? +ht : 0;
    const s = a.s !== undefined && attrBool(a.customFormat) ? this.xf(a.s) : 0;
    if ((height || hidden || ol || collapsed || s) && this.row >= 0) {
      this.ws!.setRowProps(this.row + 1, {
        height: height > 0 ? height : undefined,
        hidden,
        outlineLevel: ol,
        collapsed,
        styleIndex: s,
      });
    }
  }

  private xf(s: string): number {
    const m = this.xfMap;
    if (m === undefined) return 0;
    return m[+s] ?? 0;
  }

  private openCell(a: XmlAttrs): void {
    const r = a.r;
    if (r !== undefined) {
      // inline A1 decode: letters → col, digits → row
      let c = 0;
      let rr = 0;
      let i = 0;
      const n = r.length;
      for (; i < n; i++) {
        const ch = r.charCodeAt(i);
        if (ch >= 65 && ch <= 90) c = c * 26 + ch - 64;
        else if (ch >= 97 && ch <= 122) c = c * 26 + ch - 96;
        else if (ch !== 36) break;
      }
      for (; i < n; i++) {
        const ch = r.charCodeAt(i);
        if (ch >= 48 && ch <= 57) rr = rr * 10 + ch - 48;
        else if (ch !== 36) break;
      }
      if (c > 0 && rr > 0) {
        this.cCol = c - 1;
        if (rr - 1 !== this.row) {
          if (this.values !== undefined) {
            this.onRow!(this.row, this.values);
            this.values = undefined;
          }
          this.row = rr - 1;
        }
      } else this.cCol = this.col + 1;
    } else this.cCol = this.col + 1;
    if (this.row < 0) this.row = 0;
    this.cRow = this.row;
    this.col = this.cCol;
    if (this.row >= this.maxR || this.cCol >= this.maxC) this.overLimit();
    this.cType = a.t;
    this.cMeta = a.cm !== undefined;
    const s = a.s;
    this.cXf = s !== undefined && s !== '0' ? this.xf(s) : 0;
    this.hasV = false;
    this.vText = '';
    this.fText = undefined;
    this.fType = undefined;
    this.inIs = false;
    this.hadIs = false;
  }

  private closeCell(): void {
    let kind: number = CellKind.Empty;
    let num = 0;
    let obj: unknown;
    const v = this.vText;
    const xf = this.cXf;
    switch (this.cType) {
      case undefined:
      case 'n':
        if (this.hasV) {
          num = +v;
          if (num - num === 0) {
            kind = xf !== 0 && this.styles.isDateXf(xf) ? CellKind.Date : CellKind.Number;
          } else {
            kind = CellKind.String;
            obj = v;
            num = 0;
          }
        }
        break;
      case 's':
        if (this.hasV) {
          const item = this.strings[+v];
          if (typeof item === 'string') {
            kind = CellKind.String;
            obj = item;
          } else if (item !== undefined) {
            kind = CellKind.RichText;
            obj = item;
          }
        }
        break;
      case 'str':
        if (this.hasV) {
          kind = CellKind.String;
          obj = decodeXString(v);
        } else if (this.fText !== undefined) {
          kind = CellKind.String;
          obj = '';
        }
        break;
      case 'inlineStr': {
        if (!this.hadIs) {
          if (this.hasV) {
            kind = CellKind.String;
            obj = decodeXString(v);
          }
          break;
        }
        const item = this.is.result();
        if (typeof item === 'string') {
          kind = CellKind.String;
          obj = item;
        } else {
          kind = CellKind.RichText;
          obj = item;
        }
        break;
      }
      case 'b':
        if (this.hasV) {
          kind = CellKind.Boolean;
          num = v === '1' || v === 'true' ? 1 : 0;
        }
        break;
      case 'e':
        if (this.hasV) {
          kind = CellKind.Error;
          obj = v;
        }
        break;
      case 'd':
        if (this.hasV) {
          const dt = parseIsoDate(v);
          if (dt) {
            kind = CellKind.Date;
            num = dateToSerial(dt, this.ctx.date1904);
          } else {
            kind = CellKind.String;
            obj = v;
          }
        }
        break;
      default:
        if (this.hasV) {
          kind = CellKind.String;
          obj = v;
        }
    }

    if (this.fText !== undefined) {
      const formula = this.formulaText();
      if (formula !== undefined) {
        const f: Mutable<Formula> = { formula };
        if (kind !== CellKind.Empty) f.result = this.resultOf(kind, num, obj);
        if (this.fType === 'array' && this.fRef) f.array = this.fRef;
        if (this.cMeta) f.dynamic = true;
        kind = CellKind.Formula;
        num = 0;
        obj = f;
      }
    }

    const c0 = this.cCol;
    if (this.onRow) {
      if (kind === CellKind.Empty) return;
      if (this.values === undefined) this.values = [];
      this.values[c0] = this.valueOf(kind, num, obj);
      return;
    }
    if (kind === CellKind.Empty && xf === 0) return;
    if (c0 > this.maxCol) this.maxCol = c0;
    const ws = this.ws!;
    if (this.rowStored++ === 0 && this.hint > 4) {
      // pre-size the row from the previous row's cell count (avoids ~5 regrowths per row)
      const r0 = this.cRow;
      if (ws.rowData[r0] === undefined && r0 >= 0 && r0 < MAX_ROWS) ws._ensureRow(r0, this.hint);
    }
    ws.setCellRaw(this.cRow, c0, kind, num, obj, xf);
  }

  private overLimit(): never {
    const name = this.ctx.name ?? '';
    if (this.row >= this.maxR) throw new SheetLimitError(name, 'maxRows', this.maxR, this.row + 1);
    throw new SheetLimitError(name, 'maxCols', this.maxC, this.cCol + 1);
  }

  private formulaText(): string | undefined {
    const text = stripXlfn(this.fText!);
    if (this.fType === 'shared' && this.fSi !== undefined) {
      if (text !== '') {
        this.shared.set(this.fSi, { formula: text, r: this.cRow, c: this.cCol });
        return text;
      }
      const m = this.shared.get(this.fSi);
      return m ? translateFormula(m.formula, this.cRow - m.r, this.cCol - m.c) : undefined;
    }
    return text !== '' ? text : undefined;
  }

  private resultOf(kind: number, num: number, obj: unknown): Formula['result'] {
    switch (kind) {
      case CellKind.Number:
      case CellKind.Date:
        return num;
      case CellKind.Boolean:
        return num !== 0;
      case CellKind.Error:
        return { error: obj as ErrorCode };
      case CellKind.RichText:
        return flatten(obj as RichText);
      default:
        return obj as string;
    }
  }

  private valueOf(kind: number, num: number, obj: unknown): CellValue {
    switch (kind) {
      case CellKind.Number:
        return num;
      case CellKind.Date:
        return serialToDate(num, this.ctx.date1904, this.ctx.localDates);
      case CellKind.Boolean:
        return num !== 0;
      case CellKind.Error:
        return { error: obj as ErrorCode };
      default:
        return obj as CellValue;
    }
  }

  // ---- columns / views / hyperlinks (DOM mode) ----

  private addCol(a: XmlAttrs): void {
    const min = a.min !== undefined ? +a.min - 1 : -1;
    const max = a.max !== undefined ? +a.max - 1 : min;
    if (!(min >= 0 && max >= min && max < 16384)) return;
    const width = a.width !== undefined ? +a.width : undefined;
    const rec: ColRange = {
      min,
      max,
      width: width !== undefined && width - width === 0 ? width : undefined,
      hidden: attrBool(a.hidden),
      outlineLevel: a.outlineLevel !== undefined ? +a.outlineLevel || 0 : 0,
      collapsed: attrBool(a.collapsed),
      xf: a.style !== undefined ? this.xf(a.style) : 0,
    };
    if (rec.width === undefined && !rec.hidden && !rec.outlineLevel && !rec.collapsed && !rec.xf) {
      return;
    }
    this.cols.push(rec);
  }

  private openView(ln: string, d: number, a: XmlAttrs): void {
    if (ln === 'sheetView' && d === 3) {
      const v: Mutable<WorksheetView> = { state: 'normal' };
      // one view per Excel window; the id is kept when it is not the view's position
      if (a.workbookViewId !== undefined && +a.workbookViewId !== this.views.length) {
        v.workbookViewId = +a.workbookViewId || 0;
      }
      if (attrBool(a.tabSelected)) v.tabSelected = true;
      if (a.showGridLines !== undefined && !attrBool(a.showGridLines)) v.showGridLines = false;
      if (a.showRowColHeaders !== undefined && !attrBool(a.showRowColHeaders)) {
        v.showRowColHeaders = false;
      }
      if (a.zoomScale !== undefined && +a.zoomScale !== 100) v.zoomScale = +a.zoomScale;
      if (attrBool(a.rightToLeft)) v.rightToLeft = true;
      this.views.push(v);
      return;
    }
    const v = this.views[this.views.length - 1];
    if (!v || d !== 4) return;
    if (ln === 'pane') {
      const frozen = a.state === 'frozen' || a.state === 'frozenSplit';
      v.state = frozen ? 'frozen' : 'split';
      if (a.xSplit !== undefined && +a.xSplit) v.xSplit = +a.xSplit;
      if (a.ySplit !== undefined && +a.ySplit) v.ySplit = +a.ySplit;
      if (a.topLeftCell) v.topLeftCell = a.topLeftCell;
    } else if (ln === 'selection') {
      if (a.activeCell && v.activeCell === undefined) v.activeCell = a.activeCell;
    }
  }

  private applyLink(ws: Worksheet, h: HyperlinkRec): void {
    let target: string | undefined;
    if (h.rId !== undefined) {
      const rel = this.ctx.rels?.get(h.rId);
      if (rel) target = rel.target;
    }
    const loc = h.location;
    if (loc) {
      if (target === undefined) target = `#${loc}`;
      // exceljs writes internal links both as rel target "#Sheet!A1" and as location
      else if (!target.startsWith('#') && !target.endsWith(`#${loc}`)) target += `#${loc}`;
    }
    if (target === undefined) return;
    let range: ReturnType<typeof decodeRange>;
    try {
      range = decodeRange(h.ref);
    } catch {
      return;
    }
    const rows = ws.rowData;
    for (let r = range.top; r <= range.bottom; r++) {
      const row = rows[r];
      for (let c = range.left; c <= range.right; c++) {
        const i = row ? row.find(c) : -1;
        let text: string;
        let runs: RichText['richText'] | undefined;
        let xf = 0;
        if (i >= 0) {
          const k = row!.kind[i]!;
          if (k === CellKind.Formula) {
            this.capt?.formulaLink?.(encodeAddr(r, c), row!.obj![i], h);
            continue;
          }
          if (k === CellKind.Hyperlink) continue;
          xf = row!.style[i]!;
          const v = ws._valueAt(row!, i);
          text = displayText(v);
          if (k === CellKind.RichText) runs = (v as RichText).richText;
        } else if (range.top === range.bottom && range.left === range.right) {
          text = '';
        } else continue;
        const link: Mutable<Hyperlink> = { text, hyperlink: target };
        if (runs) link.runs = runs;
        if (h.tooltip !== undefined) link.tooltip = h.tooltip;
        ws.setCellRaw(r, c, CellKind.Hyperlink, 0, link, xf);
      }
    }
  }
}

/**
 * `<sheetFormatPr>` → ws.properties. customHeight is implied by defaultRowHeight; the outline
 * levels are not kept (writers derive them from the rows / columns, so a stale one is dropped).
 */
function readFormatPr(p: Mutable<WorksheetProperties>, a: XmlAttrs): void {
  for (const k in a) {
    const v = +a[k]!;
    const ln = k.slice(k.indexOf(':') + 1);
    if (v - v !== 0) continue;
    if (ln === 'dyDescent' || (v > 0 && (ln === 'defaultRowHeight' || ln === 'defaultColWidth'))) {
      p[ln] = v;
    }
  }
}

function flatten(rt: RichText): string {
  let s = '';
  for (const r of rt.richText) s += r.text;
  return s;
}

function displayText(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if ('richText' in v) return flatten(v);
  if ('error' in v) return v.error;
  return '';
}
