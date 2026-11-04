// sharedStrings.xml reader. The string-item (<si> / <is>) builder lives in common/rich-text.ts
// (shared with inline strings and features such as comments).

import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { type StringItem, StringItemBuilder } from '../common/rich-text.js';

export { type StringItem, StringItemBuilder };

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
