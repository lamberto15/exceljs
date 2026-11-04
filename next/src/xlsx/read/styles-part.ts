// styles.xml reader: numFmts, fonts, fills, borders, cellXfs → StyleRegistry.
// Produces `xfMap`: file cellXfs index → registry xf index (cells use `xfMap[s]`).
// Not read (yet): dxfs (conditional-format styles), cellStyleXfs / cellStyles (named styles),
// gradient fills (mapped to no fill), indexedColors / mruColors overrides.

import type { StyleRegistry } from '../../core/styles.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  Color,
  Fill,
  Font,
  Protection,
} from '../../core/types.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { attrBool } from './xml-util.js';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

export function parseColor(a: XmlAttrs): Color | undefined {
  const c: Mutable<Color> = {};
  if (a.rgb !== undefined) c.argb = a.rgb.length === 6 ? `FF${a.rgb}` : a.rgb;
  if (a.theme !== undefined) c.theme = +a.theme;
  if (a.indexed !== undefined) c.indexed = +a.indexed;
  if (a.tint !== undefined && +a.tint !== 0) c.tint = +a.tint;
  if (
    a.auto !== undefined &&
    c.argb === undefined &&
    c.theme === undefined &&
    c.indexed === undefined
  ) {
    return undefined;
  }
  return c.argb === undefined && c.theme === undefined && c.indexed === undefined ? undefined : c;
}

/**
 * Incremental font builder for `<font>` (styles) and `<rPr>` (rich text runs): feed child
 * elements with `fontChild(font, localName, attrs)`.
 */
export function fontChild(f: Mutable<Font>, ln: string, a: XmlAttrs): void {
  switch (ln) {
    case 'b':
      f.bold = attrBool(a.val, true);
      break;
    case 'i':
      f.italic = attrBool(a.val, true);
      break;
    case 'strike':
      f.strike = attrBool(a.val, true);
      break;
    case 'u': {
      const v = a.val ?? 'single';
      if (v !== 'none') f.underline = v as Font['underline'];
      break;
    }
    case 'sz':
      if (a.val !== undefined) f.size = +a.val;
      break;
    case 'name':
    case 'rFont':
      if (a.val !== undefined) f.name = a.val;
      break;
    case 'color': {
      const c = parseColor(a);
      if (c) f.color = c;
      break;
    }
    case 'vertAlign':
      if (a.val === 'superscript' || a.val === 'subscript') f.vertAlign = a.val;
      break;
  }
}

/** Remove false flags so fonts compare cleanly (b val="0" → no bold key). */
export function cleanFont(f: Mutable<Font>): Font {
  if (f.bold === false) delete f.bold;
  if (f.italic === false) delete f.italic;
  if (f.strike === false) delete f.strike;
  return f;
}

interface RawXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: Alignment;
  protection?: Protection;
}

const EDGES = new Set(['left', 'right', 'top', 'bottom', 'diagonal', 'start', 'end']);

/** SAX handler for styles.xml. Call `apply(registry)` after parsing. */
export class StylesHandler implements XmlHandler {
  readonly numFmts: { id: number; code: string }[] = [];
  readonly fonts: Font[] = [];
  readonly fills: (Fill | undefined)[] = [];
  readonly borders: Border[] = [];
  readonly xfs: RawXf[] = [];

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
        } else if (this.xf && ln === 'protection') {
          const p: Mutable<Protection> = {};
          if (a.locked !== undefined) p.locked = attrBool(a.locked);
          if (a.hidden !== undefined) p.hidden = attrBool(a.hidden);
          this.xf.protection = p;
        }
        break;
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
    }
  }

  /** Register everything into the registry; returns file xf index → registry xf index. */
  apply(reg: StyleRegistry): Uint32Array {
    for (const nf of this.numFmts) reg.registerNumFmt(nf.id, nf.code);
    const fontMap: number[] = [];
    for (let i = 0; i < this.fonts.length; i++) {
      if (i === 0) {
        reg.setDefaultFont(this.fonts[0]!);
        fontMap[0] = 0;
      } else fontMap[i] = reg.addFont(this.fonts[i]!);
    }
    const fillMap: number[] = [];
    for (let i = 0; i < this.fills.length; i++) {
      const f = this.fills[i];
      fillMap[i] = f ? reg.addFill(f) : 0;
    }
    const borderMap: number[] = [];
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
      });
    }
    return out;
  }
}
