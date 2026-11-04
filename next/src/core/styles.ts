// StyleRegistry: interns `Style` objects into deduped cellXfs (xf index 0 = default style)
// backed by deduped font / fill / border / numFmt tables, ready for styles.xml serialization.
//
// Dedupe uses hand-built canonical keys (fixed field order, normalized defaults), so
// { font: { bold: true, name: 'Arial' } } and { font: { name: 'Arial', bold: true } } share an xf.
// intern() also caches by object identity (WeakMap): style objects are treated as immutable
// once interned — create a new object instead of mutating one you already used.
//
// Writer view (all read-only, index = id used in styles.xml):
//   fonts[]    — fonts[0] is the default font (Calibri 11, theme color 1)
//   fills[]    — fills[0] = pattern none, fills[1] = gray125 (both required by Excel)
//   borders[]  — borders[0] = empty border
//   numFmts[]  — custom formats only ({ id >= 164 (or reader-provided), formatCode })
//   xfs[]      — cellXfs; xfs[0] = all zero ids
// Partial fonts are completed with the default font's name and size (so { bold: true }
// renders as Calibri 11 bold instead of an underspecified font).

import { builtinFormat, builtinFormatId, isDateFormat, isDateFormatId } from './numfmt.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  Color,
  Fill,
  Font,
  Protection,
  Style,
} from './types.js';

/** One cellXfs entry. Ids index into the registry tables. */
export interface Xf {
  readonly numFmtId: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly alignment?: Alignment | undefined;
  readonly protection?: Protection | undefined;
}

export interface NumFmt {
  readonly id: number;
  readonly formatCode: string;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

const DEFAULT_FONT: Font = { name: 'Calibri', size: 11, color: { theme: 1 } };
const FIRST_CUSTOM_NUMFMT = 164;

// ---- canonical keys + normalizers ------------------------------------------

function colorKey(c: Color | undefined): string {
  if (!c) return '';
  return `${c.argb ?? ''}/${c.theme ?? ''}/${c.tint ?? ''}/${c.indexed ?? ''}`;
}

function normColor(c: Color | undefined): Color | undefined {
  if (!c) return undefined;
  const o: Mutable<Color> = {};
  if (c.argb !== undefined) o.argb = c.argb.toUpperCase();
  if (c.theme !== undefined) o.theme = c.theme;
  if (c.tint !== undefined) o.tint = c.tint;
  if (c.indexed !== undefined) o.indexed = c.indexed;
  return o.argb === undefined && o.theme === undefined && o.indexed === undefined ? undefined : o;
}

function normFont(f: Font, base: Font): Font {
  const o: Mutable<Font> = {};
  const name = f.name ?? base.name;
  const size = f.size ?? base.size;
  if (name !== undefined) o.name = name;
  if (size !== undefined) o.size = size;
  if (f.bold) o.bold = true;
  if (f.italic) o.italic = true;
  if (f.underline) o.underline = f.underline === true ? 'single' : f.underline;
  if (f.strike) o.strike = true;
  const color = normColor(f.color);
  if (color) o.color = color;
  if (f.vertAlign) o.vertAlign = f.vertAlign;
  return o;
}

function fontKey(f: Font): string {
  return `${f.name ?? ''}|${f.size ?? ''}|${f.bold ? 1 : 0}${f.italic ? 1 : 0}${f.strike ? 1 : 0}|${
    f.underline ?? ''
  }|${f.vertAlign ?? ''}|${colorKey(f.color)}`;
}

function normFill(f: Fill): Fill {
  const o: Mutable<Fill> = { type: 'pattern', pattern: f.pattern || 'none' };
  const fg = normColor(f.fgColor);
  const bg = normColor(f.bgColor);
  if (fg) o.fgColor = fg;
  if (bg) o.bgColor = bg;
  return o;
}

function fillKey(f: Fill): string {
  return `${f.pattern}|${colorKey(f.fgColor)}|${colorKey(f.bgColor)}`;
}

function normEdge(e: BorderEdge | undefined): BorderEdge | undefined {
  if (!e || !e.style) return undefined;
  const color = normColor(e.color);
  return color ? { style: e.style, color } : { style: e.style };
}

function normBorder(b: Border): Border {
  const o: Mutable<Border> = {};
  const top = normEdge(b.top);
  const left = normEdge(b.left);
  const bottom = normEdge(b.bottom);
  const right = normEdge(b.right);
  if (top) o.top = top;
  if (left) o.left = left;
  if (bottom) o.bottom = bottom;
  if (right) o.right = right;
  const d = b.diagonal;
  if (d && d.style && (d.up || d.down)) {
    const diag: Mutable<BorderEdge> & { up?: boolean; down?: boolean } = { style: d.style };
    const color = normColor(d.color);
    if (color) diag.color = color;
    if (d.up) diag.up = true;
    if (d.down) diag.down = true;
    o.diagonal = diag;
  }
  return o;
}

function edgeKey(e: BorderEdge | undefined): string {
  return e ? `${e.style}:${colorKey(e.color)}` : '';
}

function borderKey(b: Border): string {
  const d = b.diagonal;
  return `${edgeKey(b.top)}|${edgeKey(b.left)}|${edgeKey(b.bottom)}|${edgeKey(b.right)}|${
    d ? `${edgeKey(d)}${d.up ? 'u' : ''}${d.down ? 'd' : ''}` : ''
  }`;
}

function normAlignment(a: Alignment | undefined): Alignment | undefined {
  if (!a) return undefined;
  const o: Mutable<Alignment> = {};
  let any = false;
  if (a.horizontal !== undefined && a.horizontal !== 'general') {
    o.horizontal = a.horizontal;
    any = true;
  }
  if (a.vertical !== undefined && a.vertical !== 'bottom') {
    o.vertical = a.vertical;
    any = true;
  }
  if (a.wrapText) {
    o.wrapText = true;
    any = true;
  }
  if (a.shrinkToFit) {
    o.shrinkToFit = true;
    any = true;
  }
  if (a.indent) {
    o.indent = a.indent;
    any = true;
  }
  if (a.textRotation) {
    o.textRotation = a.textRotation;
    any = true;
  }
  return any ? o : undefined;
}

function alignKey(a: Alignment | undefined): string {
  if (!a) return '';
  return `${a.horizontal ?? ''}/${a.vertical ?? ''}/${a.wrapText ? 1 : 0}${a.shrinkToFit ? 1 : 0}/${
    a.indent ?? ''
  }/${a.textRotation ?? ''}`;
}

function normProtection(p: Protection | undefined): Protection | undefined {
  if (!p) return undefined;
  // Excel defaults: locked = true, hidden = false
  const locked = p.locked !== false;
  const hidden = p.hidden === true;
  if (locked && !hidden) return undefined;
  return { locked, hidden };
}

function protKey(p: Protection | undefined): string {
  return p ? `${p.locked ? 1 : 0}${p.hidden ? 1 : 0}` : '';
}

// ---- registry ----------------------------------------------------------------

export class StyleRegistry {
  private readonly _fonts: Font[] = [DEFAULT_FONT];
  private readonly _fills: Fill[] = [
    { type: 'pattern', pattern: 'none' },
    { type: 'pattern', pattern: 'gray125' },
  ];
  private readonly _borders: Border[] = [{}];
  private readonly _numFmts: NumFmt[] = [];
  private readonly _xfs: Xf[] = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }];

  private readonly _fontIdx = new Map<string, number>([[fontKey(DEFAULT_FONT), 0]]);
  private readonly _fillIdx = new Map<string, number>([
    ['none||', 0],
    ['gray125||', 1],
  ]);
  private readonly _borderIdx = new Map<string, number>([[borderKey({}), 0]]);
  private readonly _numFmtIdx = new Map<string, number>();
  private readonly _numFmtById = new Map<number, string>();
  private _nextNumFmt = FIRST_CUSTOM_NUMFMT;
  private readonly _xfIdx = new Map<string, number>([['0,0,0,0,,', 0]]);
  private readonly _byObject = new WeakMap<Style, number>();
  private readonly _styleCache: (Style | undefined)[] = [];
  private readonly _dateXf: (boolean | undefined)[] = [];
  private readonly _withFmt = new Map<string, number>();

  get fonts(): readonly Font[] {
    return this._fonts;
  }
  get fills(): readonly Fill[] {
    return this._fills;
  }
  get borders(): readonly Border[] {
    return this._borders;
  }
  /** Custom number formats (built-in ids are not listed). */
  get numFmts(): readonly NumFmt[] {
    return this._numFmts;
  }
  /** cellXfs; index = the `s` attribute of a cell. */
  get xfs(): readonly Xf[] {
    return this._xfs;
  }
  get defaultFont(): Font {
    return this._fonts[0]!;
  }

  /** Style → xf index (0 for undefined / empty style). */
  intern(style: Style | undefined): number {
    if (!style) return 0;
    const hit = this._byObject.get(style);
    if (hit !== undefined) return hit;
    const xf = this.internXf({
      numFmtId: this.numFmtId(style.numFmt),
      fontId: style.font ? this.addFont(style.font) : 0,
      fillId: style.fill ? this.addFill(style.fill) : 0,
      borderId: style.border ? this.addBorder(style.border) : 0,
      alignment: style.alignment,
      protection: style.protection,
    });
    this._byObject.set(style, xf);
    return xf;
  }

  /** Register (dedupe) a raw xf record, e.g. from a reader. Ids must already exist. */
  internXf(raw: Xf): number {
    const alignment = normAlignment(raw.alignment);
    const protection = normProtection(raw.protection);
    const key = `${raw.numFmtId},${raw.fontId},${raw.fillId},${raw.borderId},${alignKey(alignment)},${protKey(protection)}`;
    const hit = this._xfIdx.get(key);
    if (hit !== undefined) return hit;
    if (
      raw.fontId >= this._fonts.length ||
      raw.fillId >= this._fills.length ||
      raw.borderId >= this._borders.length
    ) {
      throw new RangeError('internXf: font/fill/border id not registered');
    }
    const xf: Mutable<Xf> = {
      numFmtId: raw.numFmtId,
      fontId: raw.fontId,
      fillId: raw.fillId,
      borderId: raw.borderId,
    };
    if (alignment) xf.alignment = alignment;
    if (protection) xf.protection = protection;
    const idx = this._xfs.length;
    this._xfs.push(xf);
    this._xfIdx.set(key, idx);
    return idx;
  }

  /** Font → font id (deduped). Partial fonts inherit the default font's name/size. */
  addFont(font: Font): number {
    const f = normFont(font, this._fonts[0]!);
    const key = fontKey(f);
    let idx = this._fontIdx.get(key);
    if (idx === undefined) {
      idx = this._fonts.length;
      this._fonts.push(f);
      this._fontIdx.set(key, idx);
    }
    return idx;
  }

  /** Fill → fill id (deduped; pattern 'none' without colors → 0). */
  addFill(fill: Fill): number {
    const f = normFill(fill);
    const key = fillKey(f);
    let idx = this._fillIdx.get(key);
    if (idx === undefined) {
      idx = this._fills.length;
      this._fills.push(f);
      this._fillIdx.set(key, idx);
    }
    return idx;
  }

  /** Border → border id (deduped; edges without a style are dropped). */
  addBorder(border: Border): number {
    const b = normBorder(border);
    const key = borderKey(b);
    let idx = this._borderIdx.get(key);
    if (idx === undefined) {
      idx = this._borders.length;
      this._borders.push(b);
      this._borderIdx.set(key, idx);
    }
    return idx;
  }

  /** Format code → numFmt id: 0 for General/empty, built-in id, or a custom id (>= 164). */
  numFmtId(code: string | undefined): number {
    if (!code) return 0;
    let id = this._numFmtIdx.get(code);
    if (id !== undefined) return id;
    id = builtinFormatId(code);
    if (id === undefined) {
      id = this._nextNumFmt++;
      this._numFmts.push({ id, formatCode: code });
      this._numFmtById.set(id, code);
    }
    this._numFmtIdx.set(code, id);
    return id;
  }

  /**
   * Reader hook: register a numFmt with the id the file uses. Returns the id. Ids < 164 that
   * are not standard built-ins (locale formats) are listed in `numFmts` so they are re-written.
   */
  registerNumFmt(id: number, code: string): number {
    if (builtinFormat(id) === code) return id;
    if (!this._numFmtById.has(id)) {
      this._numFmts.push({ id, formatCode: code });
      this._numFmtById.set(id, code);
      if (!this._numFmtIdx.has(code)) this._numFmtIdx.set(code, id);
      if (id >= this._nextNumFmt) this._nextNumFmt = id + 1;
    }
    return id;
  }

  /** Reader hook: replace font 0 (the workbook default font). */
  setDefaultFont(font: Font): void {
    const f = normFont(font, {});
    this._fontIdx.delete(fontKey(this._fonts[0]!));
    this._fonts[0] = f;
    this._fontIdx.set(fontKey(f), 0);
  }

  /** Format code for a numFmt id (custom or built-in), undefined if unknown. */
  numFmtCode(id: number): string | undefined {
    return this._numFmtById.get(id) ?? builtinFormat(id);
  }

  /** True when the xf's number format renders dates (cached; useful for readers). */
  isDateXf(xf: number): boolean {
    let v = this._dateXf[xf];
    if (v === undefined) {
      const rec = this._xfs[xf];
      if (!rec) return false;
      const id = rec.numFmtId;
      const code = this._numFmtById.get(id);
      v = code !== undefined ? isDateFormat(code) : isDateFormatId(id);
      this._dateXf[xf] = v;
    }
    return v;
  }

  /** xf with its numFmt replaced by `code` if it currently has General (id 0); cached. */
  withDefaultNumFmt(xf: number, code: string): number {
    const rec = this._xfs[xf];
    if (!rec || rec.numFmtId !== 0) return xf;
    const key = `${xf}|${code}`;
    let out = this._withFmt.get(key);
    if (out === undefined) {
      out = this.internXf({ ...rec, numFmtId: this.numFmtId(code) });
      this._withFmt.set(key, out);
    }
    return out;
  }

  /** xf index → Style (cached, shared object: do not mutate). Unknown index → {}. */
  get(xf: number): Style {
    const cached = this._styleCache[xf];
    if (cached) return cached;
    const rec = this._xfs[xf];
    const s: Mutable<Style> = {};
    if (rec) {
      if (rec.numFmtId !== 0) {
        const code = this.numFmtCode(rec.numFmtId);
        if (code !== undefined) s.numFmt = code;
      }
      if (rec.fontId !== 0) s.font = this._fonts[rec.fontId];
      if (rec.fillId !== 0) s.fill = this._fills[rec.fillId];
      if (rec.borderId !== 0) s.border = this._borders[rec.borderId];
      if (rec.alignment) s.alignment = rec.alignment;
      if (rec.protection) s.protection = rec.protection;
      this._styleCache[xf] = s;
      this._byObject.set(s, xf);
    }
    return s;
  }
}
