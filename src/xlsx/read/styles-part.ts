// styles.xml reader: numFmts, fonts, fills, borders, cellXfs → StyleRegistry.
// Produces `xfMap`: file cellXfs index → registry xf index (cells use `xfMap[s]`).
// dxfs (differential formats for conditional formatting / tables) → `dxfMap` (file dxf index →
// registry dxf index, see `applyDxfs`).
// Named styles (cellStyleXfs / cellStyles) are not modelled: with `apply(reg, true)` the cellXfs
// keep their `xfId` and read/preserve.ts keeps both lists raw (ids remapped through
// fontMap / fillMap / borderMap). Not read (yet): gradient fills (mapped to no fill), indexedColors / mruColors overrides.

import type { StyleRegistry } from '../../core/styles.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  Fill,
  Font,
  Protection,
  Style,
} from '../../core/types.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { cleanFont, fontChild, parseColor } from '../common/rich-text.js';
import { attrBool } from './xml-util.js';

export { cleanFont, fontChild, parseColor };

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

interface RawXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  xfId: number;
  alignment?: Alignment;
  protection?: Protection;
  checkbox?: boolean;
}

const EDGES = new Set(['left', 'right', 'top', 'bottom', 'diagonal', 'start', 'end']);

/** SAX handler for styles.xml. Call `apply(registry)` after parsing. */
export class StylesHandler implements XmlHandler {
  readonly numFmts: { id: number; code: string }[] = [];
  readonly fonts: Font[] = [];
  readonly fills: (Fill | undefined)[] = [];
  readonly borders: Border[] = [];
  readonly xfs: RawXf[] = [];
  readonly dxfs: Style[] = [];
  /** File font / fill / border index → registry index (set by `apply`). */
  readonly fontMap: number[] = [];
  readonly fillMap: number[] = [];
  readonly borderMap: number[] = [];
  /** cellStyleXfs / cellStyles entries (named styles beyond "Normal" when either is > 1). */
  styleXfs = 0;
  cellStyles = 0;
  private dxf: Mutable<Style> | undefined;
  private dxfPart = '';

  private section = '';
  private font: Mutable<Font> | undefined;
  private fill: Mutable<Fill> | undefined;
  private inFill = false;
  private border: Mutable<Border> | undefined;
  private edge: string | undefined;
  private edgeVal: Mutable<BorderEdge> | undefined;
  private xf: RawXf | undefined;
  private depth = 0;
  private sectionDepth = 0;

  open(name: string, a: XmlAttrs, _sc: boolean): void {
    const ln = localName(name);
    this.depth++;
    if (this.depth === 2) {
      this.section = ln;
      this.sectionDepth = 2;
      return;
    }
    if (this.depth < 2) return;
    if (this.depth === 3) {
      if (ln === 'xf' && this.section === 'cellStyleXfs') this.styleXfs++;
      else if (ln === 'cellStyle') this.cellStyles++;
    }
    switch (this.section) {
      case 'numFmts':
        if (ln === 'numFmt' && a.numFmtId !== undefined && a.formatCode !== undefined) {
          this.numFmts.push({ id: +a.numFmtId, code: a.formatCode });
        }
        break;
      case 'fonts':
        if (this.depth === 3 && ln === 'font') this.font = {};
        else if (this.font && this.depth === 4) fontChild(this.font, ln, a);
        break;
      case 'fills':
        if (this.depth === 3 && ln === 'fill') {
          this.inFill = true;
          this.fill = undefined;
        } else if (this.inFill && ln === 'patternFill') {
          this.fill = { type: 'pattern', pattern: a.patternType ?? 'none' };
        } else if (this.fill && (ln === 'fgColor' || ln === 'bgColor')) {
          const c = parseColor(a);
          if (c) this.fill[ln] = c;
        }
        break;
      case 'borders':
        if (this.depth === 3 && ln === 'border') {
          this.border = {};
          if (attrBool(a.diagonalUp) || attrBool(a.diagonalDown)) {
            this.border.diagonal = {
              ...(attrBool(a.diagonalUp) ? { up: true } : {}),
              ...(attrBool(a.diagonalDown) ? { down: true } : {}),
            };
          }
        } else if (this.border && this.depth === 4 && EDGES.has(ln)) {
          this.edge = ln === 'start' ? 'left' : ln === 'end' ? 'right' : ln;
          this.edgeVal = a.style !== undefined ? { style: a.style as BorderStyle } : undefined;
        } else if (this.edgeVal && this.depth === 5 && ln === 'color') {
          const c = parseColor(a);
          if (c) this.edgeVal.color = c;
        }
        break;
      case 'cellXfs':
        if (this.depth === 3 && ln === 'xf') {
          this.xf = {
            numFmtId: a.numFmtId !== undefined ? +a.numFmtId : 0,
            fontId: a.fontId !== undefined ? +a.fontId : 0,
            fillId: a.fillId !== undefined ? +a.fillId : 0,
            borderId: a.borderId !== undefined ? +a.borderId : 0,
            xfId: a.xfId !== undefined ? +a.xfId || 0 : 0,
          };
        } else if (this.xf && ln === 'alignment') {
          const al: Mutable<Alignment> = {};
          if (a.horizontal) al.horizontal = a.horizontal as Alignment['horizontal'];
          if (a.vertical) al.vertical = a.vertical as Alignment['vertical'];
          if (attrBool(a.wrapText)) al.wrapText = true;
          if (attrBool(a.shrinkToFit)) al.shrinkToFit = true;
          if (a.indent !== undefined && +a.indent) al.indent = +a.indent;
          if (a.textRotation !== undefined && +a.textRotation) al.textRotation = +a.textRotation;
          this.xf.alignment = al;
        } else if (this.xf && ln === 'xfComplement') {
          this.xf.checkbox = true; // feature property bag: in-cell checkbox
        } else if (this.xf && ln === 'protection') {
          const p: Mutable<Protection> = {};
          if (a.locked !== undefined) p.locked = attrBool(a.locked);
          if (a.hidden !== undefined) p.hidden = attrBool(a.hidden);
          this.xf.protection = p;
        }
        break;
      case 'dxfs':
        this.openDxf(ln, a);
        break;
    }
  }

  private openDxf(ln: string, a: XmlAttrs): void {
    const d = this.depth;
    if (d === 3) {
      if (ln === 'dxf') this.dxf = {};
      return;
    }
    const x = this.dxf;
    if (!x) return;
    if (d === 4) {
      this.dxfPart = ln;
      switch (ln) {
        case 'font':
          x.font = {};
          break;
        case 'numFmt':
          if (a.formatCode !== undefined) x.numFmt = a.formatCode;
          break;
        case 'alignment': {
          const al: Mutable<Alignment> = {};
          if (a.horizontal) al.horizontal = a.horizontal as Alignment['horizontal'];
          if (a.vertical) al.vertical = a.vertical as Alignment['vertical'];
          if (attrBool(a.wrapText)) al.wrapText = true;
          if (attrBool(a.shrinkToFit)) al.shrinkToFit = true;
          if (a.indent !== undefined && +a.indent) al.indent = +a.indent;
          if (a.textRotation !== undefined && +a.textRotation) al.textRotation = +a.textRotation;
          x.alignment = al;
          break;
        }
        case 'border':
          x.border = {};
          break;
        case 'protection': {
          const p: Mutable<Protection> = {};
          if (a.locked !== undefined) p.locked = attrBool(a.locked);
          if (a.hidden !== undefined) p.hidden = attrBool(a.hidden);
          x.protection = p;
          break;
        }
      }
      return;
    }
    switch (this.dxfPart) {
      case 'font':
        if (d === 5) fontChild(x.font as Mutable<Font>, ln, a);
        break;
      case 'fill':
        if (ln === 'patternFill') {
          // dxf fills may omit patternType (Excel then means solid with bgColor)
          x.fill = { type: 'pattern', pattern: a.patternType ?? 'solid' };
        } else if (x.fill && (ln === 'fgColor' || ln === 'bgColor')) {
          const c = parseColor(a);
          if (c) (x.fill as Mutable<Fill>)[ln] = c;
        }
        break;
      case 'border': {
        const b = x.border as Mutable<Border>;
        if (d === 5 && EDGES.has(ln)) {
          const e = ln === 'start' ? 'left' : ln === 'end' ? 'right' : ln;
          if (a.style !== undefined) b[e as 'left'] = { style: a.style as BorderStyle };
          this.edge = e;
        } else if (d === 6 && ln === 'color' && this.edge) {
          const cur = b[this.edge as 'left'];
          const c = parseColor(a);
          if (cur && c) b[this.edge as 'left'] = { ...cur, color: c };
        }
        break;
      }
    }
  }

  close(name: string): void {
    const d = this.depth--;
    if (d === this.sectionDepth) {
      this.section = '';
      return;
    }
    const ln = localName(name);
    switch (this.section) {
      case 'fonts':
        if (d === 3 && this.font) {
          this.fonts.push(cleanFont(this.font));
          this.font = undefined;
        }
        break;
      case 'fills':
        if (d === 3 && ln === 'fill') {
          this.fills.push(this.fill);
          this.inFill = false;
          this.fill = undefined;
        }
        break;
      case 'borders':
        if (d === 4 && this.border && this.edge) {
          const e = this.edgeVal;
          if (e) {
            const b = this.border;
            if (this.edge === 'diagonal') b.diagonal = { ...b.diagonal, ...e };
            else b[this.edge as 'left' | 'right' | 'top' | 'bottom'] = e;
          }
          this.edge = undefined;
          this.edgeVal = undefined;
        } else if (d === 3 && this.border) {
          this.borders.push(this.border);
          this.border = undefined;
        }
        break;
      case 'cellXfs':
        if (d === 3 && this.xf) {
          this.xfs.push(this.xf);
          this.xf = undefined;
        }
        break;
      case 'dxfs':
        if (d === 3 && this.dxf) {
          if (this.dxf.font) this.dxf.font = cleanFont(this.dxf.font as Mutable<Font>);
          this.dxfs.push(this.dxf);
          this.dxf = undefined;
        } else if (d === 4) this.dxfPart = '';
        break;
    }
  }

  /** Register dxfs; returns file dxf index → registry dxf index. Call after `apply`. */
  applyDxfs(reg: StyleRegistry): Uint32Array {
    const out = new Uint32Array(this.dxfs.length);
    for (let i = 0; i < this.dxfs.length; i++) out[i] = reg.internDxf(this.dxfs[i]!);
    return out;
  }

  /**
   * Register everything into the registry; returns file xf index → registry xf index.
   * `named`: keep each xf's named-style link (`xfId`), when the named styles are kept too.
   */
  apply(reg: StyleRegistry, named = false): Uint32Array {
    for (const nf of this.numFmts) reg.registerNumFmt(nf.id, nf.code);
    const fontMap = this.fontMap;
    for (let i = 0; i < this.fonts.length; i++) {
      if (i === 0) {
        reg.setDefaultFont(this.fonts[0]!);
        fontMap[0] = 0;
      } else fontMap[i] = reg.addFont(this.fonts[i]!);
    }
    const fillMap = this.fillMap;
    for (let i = 0; i < this.fills.length; i++) {
      const f = this.fills[i];
      fillMap[i] = f ? reg.addFill(f) : 0;
    }
    const borderMap = this.borderMap;
    for (let i = 0; i < this.borders.length; i++) borderMap[i] = reg.addBorder(this.borders[i]!);
    const out = new Uint32Array(Math.max(this.xfs.length, 1));
    for (let i = 0; i < this.xfs.length; i++) {
      const x = this.xfs[i]!;
      out[i] = reg.internXf({
        numFmtId: x.numFmtId,
        fontId: fontMap[x.fontId] ?? 0,
        fillId: fillMap[x.fillId] ?? 0,
        borderId: borderMap[x.borderId] ?? 0,
        alignment: x.alignment,
        protection: x.protection,
        xfId: named ? x.xfId : 0,
        checkbox: x.checkbox,
      });
    }
    return out;
  }
}
