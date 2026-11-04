// sharedStrings.xml reader and the string-item (<si> / <is>) builder shared with inline strings.

import type { Font, RichText, RichTextRun } from '../../core/types.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { cleanFont, fontChild } from './styles-part.js';
import { decodeXString } from './xml-util.js';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

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

/** SAX handler for sharedStrings.xml; items in file order. */
export class SharedStringsHandler implements XmlHandler {
  readonly items: StringItem[] = [];
  private readonly b = new StringItemBuilder();
  private inSi = false;

  open(name: string, a: XmlAttrs): void {
    const ln = localName(name);
    if (ln === 'si') {
      this.inSi = true;
      this.b.reset();
    } else if (this.inSi) this.b.open(ln, a);
  }

  close(name: string): void {
    const ln = localName(name);
    if (ln === 'si' && this.inSi) {
      this.inSi = false;
      this.items.push(this.b.result());
    } else if (this.inSi) this.b.close(ln);
  }

  text(t: string): void {
    if (this.inSi) this.b.onText(t);
  }
}
