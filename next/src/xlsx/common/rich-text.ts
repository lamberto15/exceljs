// Rich text / font (de)serialization shared by the cell writer/reader and features (comments,
// ...). Side-effect free: bundles keep only what they import.

import type { Color, Font, RichText, RichTextRun } from '../../core/types.js';
import { escapeAttr, escapeXString, type XmlAttrs } from '../../xml/index.js';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

// ---- writing

/** `<t>` element, with xml:space="preserve" only when the text has leading/trailing whitespace. */
export function tElem(v: string): string {
  const n = v.length;
  if (n === 0) return '<t></t>';
  const a = v.charCodeAt(0);
  const z = v.charCodeAt(n - 1);
  const e = escapeXString(v);
  return a === 32 || a === 9 || a === 10 || a === 13 || z === 32 || z === 9 || z === 10 || z === 13
    ? `<t xml:space="preserve">${e}</t>`
    : `<t>${e}</t>`;
}

export function colorAttrs(c: Color): string {
  let s = '';
  // ST_UnsignedIntHex is 8 hex digits: an RGB-only value gets an opaque alpha (as readers assume)
  if (c.argb !== undefined)
    s += ` rgb="${escapeAttr(c.argb.length === 6 ? `FF${c.argb}` : c.argb)}"`;
  else if (c.theme !== undefined) s += ` theme="${c.theme}"`;
  else if (c.indexed !== undefined) s += ` indexed="${c.indexed}"`;
  if (c.tint !== undefined && c.tint !== 0) s += ` tint="${c.tint}"`;
  return s;
}

/** Font children in Excel's order. `nameTag` is 'name' in styles.xml and 'rFont' in rich text. */
export function fontBody(f: Font, nameTag: 'name' | 'rFont'): string {
  let s = '';
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.underline) {
    const u = f.underline === true ? 'single' : f.underline;
    s += u === 'single' ? '<u/>' : `<u val="${u}"/>`;
  }
  if (f.vertAlign) s += `<vertAlign val="${f.vertAlign}"/>`;
  if (f.size !== undefined) s += `<sz val="${f.size}"/>`;
  if (f.color) s += `<color${colorAttrs(f.color)}/>`;
  if (f.name !== undefined) s += `<${nameTag} val="${escapeAttr(f.name)}"/>`;
  return s;
}

/** `<r>` runs of a rich text value (CT_Rst body). */
export function runsXml(runs: readonly RichTextRun[]): string {
  let out = '';
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i]!;
    const f = run.font;
    out += f
      ? `<r><rPr>${fontBody(f, 'rFont')}</rPr>${tElem(run.text)}</r>`
      : `<r>${tElem(run.text)}</r>`;
  }
  return out;
}

// ---- reading

function attrBool(v: string | undefined, dflt = false): boolean {
  if (v === undefined) return dflt;
  return v === '1' || v === 'true' || v === 'on';
}

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

/**
 * Decode ST_Xstring escapes: `_xHHHH_` → UTF-16 code unit (e.g. `_x000D_` → CR);
 * `_x005F_` escapes a literal underscore. Fast path when no `_x` is present.
 */
export function decodeXString(s: string): string {
  let i = s.indexOf('_x');
  if (i < 0) return s;
  let out = '';
  let last = 0;
  while (i >= 0) {
    if (i + 7 <= s.length && s.charCodeAt(i + 6) === 95 && isHex4(s, i + 2)) {
      out += s.slice(last, i) + String.fromCharCode(parseInt(s.slice(i + 2, i + 6), 16));
      last = i + 7;
      i = s.indexOf('_x', last);
    } else {
      i = s.indexOf('_x', i + 1);
    }
  }
  return out + s.slice(last);
}

function isHex4(s: string, p: number): boolean {
  for (let k = p; k < p + 4; k++) {
    const c = s.charCodeAt(k);
    if (!((c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102))) return false;
  }
  return true;
}

export type StringItem = string | RichText;

/**
 * Builds one string item from the children of `<si>` / `<is>`:
 *  - `<t>` text (plain),
 *  - `<r><rPr>…</rPr><t>…</t></r>` runs → RichText when any run carries an `<rPr>`, else the
 *    runs are concatenated into a plain string,
 *  - `<rPh>` phonetic runs and `<phoneticPr>` are skipped.
 * `_xHHHH_` escapes are decoded.
 */
export class StringItemBuilder {
  private text = '';
  private runs: RichTextRun[] | undefined;
  private rich = false;
  private inT = false;
  private skip = 0;
  private runFont: Mutable<Font> | undefined;
  private inRun = false;
  private inRPr = false;
  private runText = '';

  reset(): void {
    this.text = '';
    this.runs = undefined;
    this.rich = false;
    this.inT = false;
    this.skip = 0;
    this.inRun = false;
    this.inRPr = false;
  }

  open(ln: string, a: XmlAttrs): void {
    if (this.skip > 0) {
      this.skip++;
      return;
    }
    switch (ln) {
      case 't':
        this.inT = true;
        return;
      case 'r':
        this.inRun = true;
        this.runFont = undefined;
        this.runText = '';
        return;
      case 'rPr':
        if (this.inRun) {
          this.inRPr = true;
          this.runFont = {};
          this.rich = true;
        }
        return;
      case 'rPh':
      case 'phoneticPr':
        this.skip = 1;
        return;
      default:
        if (this.inRPr && this.runFont) fontChild(this.runFont, ln, a);
    }
  }

  close(ln: string): void {
    if (this.skip > 0) {
      this.skip--;
      return;
    }
    switch (ln) {
      case 't':
        this.inT = false;
        return;
      case 'rPr':
        this.inRPr = false;
        return;
      case 'r':
        if (this.inRun) {
          this.inRun = false;
          const run: Mutable<RichTextRun> = { text: decodeXString(this.runText) };
          if (this.runFont) run.font = cleanFont(this.runFont);
          if (this.runs === undefined) this.runs = [run];
          else this.runs.push(run);
          this.text += this.runText;
        }
        return;
    }
  }

  onText(t: string): void {
    if (!this.inT || this.skip > 0) return;
    if (this.inRun) this.runText += t;
    else this.text += t;
  }

  result(): StringItem {
    if (this.rich && this.runs) return { richText: this.runs };
    return decodeXString(this.text);
  }
}
