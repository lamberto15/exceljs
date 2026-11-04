// Small XML helpers shared by the part parsers.

import { type XmlAttrs, type XmlHandler, XmlTokenizer } from '../../xml/index.js';

/** Thrown for structurally invalid XML inside an XLSX part. */
export class XlsxParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'XlsxParseError';
  }
}

/**
 * Well-formedness guard around a handler: balanced/matching tags, a single root element and
 * no non-whitespace text outside it. The tokenizer itself is lenient; this turns garbage into a
 * clear error instead of silently producing an empty sheet. Cost: one array push/pop per element.
 */
export class StrictXml implements XmlHandler {
  private readonly stack: string[] = [];
  private depth = 0;
  private roots = 0;
  private readonly inner: XmlHandler;
  private readonly part: string;
  private readonly tok: XmlTokenizer;

  constructor(inner: XmlHandler, part: string) {
    this.inner = inner;
    this.part = part;
    this.tok = new XmlTokenizer(this);
  }

  open(name: string, attrs: XmlAttrs, selfClosing: boolean): void {
    if (this.depth === 0 && ++this.roots > 1) this.fail('multiple root elements');
    this.stack[this.depth++] = name;
    this.inner.open?.(name, attrs, selfClosing);
  }

  close(name: string): void {
    if (this.depth === 0) this.fail(`unexpected closing tag </${name}>`);
    const top = this.stack[--this.depth];
    if (top !== name) this.fail(`mismatched closing tag </${name}>, expected </${top}>`);
    this.inner.close?.(name);
  }

  text(t: string): void {
    if (this.depth === 0) {
      for (let i = 0; i < t.length; i++) {
        if (t.charCodeAt(i) > 32) this.fail('text data outside of root element');
      }
      return;
    }
    this.inner.text?.(t);
  }

  write(chunk: string): void {
    try {
      this.tok.write(chunk);
    } catch (e) {
      throw this.wrap(e);
    }
  }

  end(): void {
    try {
      this.tok.end();
    } catch (e) {
      throw this.wrap(e);
    }
    if (this.depth !== 0)
      this.fail(`unexpected end of document inside <${this.stack[this.depth - 1]}>`);
    if (this.roots === 0) this.fail('no root element');
  }

  private wrap(e: unknown): unknown {
    if (e instanceof XlsxParseError || !(e instanceof Error)) return e;
    if (/^XmlTokenizer/.test(e.message)) return new XlsxParseError(`${this.part}: ${e.message}`);
    return e;
  }

  private fail(msg: string): never {
    throw new XlsxParseError(`Invalid XML in ${this.part}: ${msg}`);
  }
}

/** Parse a complete string with the strict guard. */
export function parseStrict(xml: string, h: XmlHandler, part: string): void {
  const s = new StrictXml(h, part);
  s.write(xml);
  s.end();
}

/** OOXML boolean attribute: "1" / "true" → true; absent → `dflt`. */
export function attrBool(v: string | undefined, dflt = false): boolean {
  if (v === undefined) return dflt;
  return v === '1' || v === 'true' || v === 'on';
}

/** Find an attribute by local name (e.g. `r:id` written with any prefix). */
export function attrLocal(a: XmlAttrs, local: string): string | undefined {
  for (const k in a) {
    const i = k.indexOf(':');
    if ((i < 0 ? k : k.slice(i + 1)) === local) return a[k];
  }
  return undefined;
}

export { decodeXString } from '../common/rich-text.js';
