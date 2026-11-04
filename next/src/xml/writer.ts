import { escapeAttr, escapeText } from './escape.js';

/** Attribute map for XmlWriter. `undefined` values are skipped, booleans write "1"/"0". */
export type XmlWriteAttrs = Record<string, string | number | boolean | undefined>;

const DEFAULT_FLUSH_AT = 64 * 1024;

/**
 * Minimal, allocation-light XML builder.
 *
 * Output is accumulated into a small string buffer; once it reaches `flushAt` UTF-16 units it is
 * handed to `onFlush` (streaming) or moved into an internal chunk array (in-memory).
 * Tag/attribute names are written verbatim (callers pass trusted names); values are escaped.
 */
export class XmlWriter {
  private buf: string = '';
  private readonly parts: string[] = [];
  private readonly onFlush: ((chunk: string) => void) | undefined;
  private readonly flushAt: number;

  constructor(onFlush?: (chunk: string) => void, flushAt: number = DEFAULT_FLUSH_AT) {
    this.onFlush = onFlush;
    this.flushAt = flushAt;
  }

  private push(s: string): void {
    const b = this.buf + s;
    if (b.length >= this.flushAt) {
      this.buf = '';
      if (this.onFlush) this.onFlush(b);
      else this.parts.push(b);
    } else {
      this.buf = b;
    }
  }

  /** Append pre-serialized XML verbatim (no escaping). */
  raw(s: string): this {
    this.push(s);
    return this;
  }

  /** `<tag a="..">` */
  open(tag: string, attrs?: XmlWriteAttrs): this {
    this.push(attrs ? `<${tag}${serializeAttrs(attrs)}>` : `<${tag}>`);
    return this;
  }

  /** `<tag a=".."/>` when `text` is undefined, else `<tag a="..">text</tag>`. */
  leaf(tag: string, attrs?: XmlWriteAttrs, text?: string): this {
    const a = attrs ? serializeAttrs(attrs) : '';
    this.push(text === undefined ? `<${tag}${a}/>` : `<${tag}${a}>${escapeText(text)}</${tag}>`);
    return this;
  }

  /** `</tag>` */
  close(tag: string): this {
    this.push(`</${tag}>`);
    return this;
  }

  /** Escaped text node. */
  text(s: string): this {
    this.push(escapeText(s));
    return this;
  }

  /** Content still buffered (not yet passed to onFlush). Does not clear the buffer. */
  toString(): string {
    return this.parts.length === 0 ? this.buf : this.parts.join('') + this.buf;
  }

  /** Emit all buffered content via onFlush (no-op without a callback or when empty). */
  flush(): void {
    if (!this.onFlush) return;
    const b = this.parts.length === 0 ? this.buf : this.parts.join('') + this.buf;
    this.parts.length = 0;
    this.buf = '';
    if (b !== '') this.onFlush(b);
  }
}

function serializeAttrs(attrs: XmlWriteAttrs): string {
  let out = '';
  for (const k in attrs) {
    const v = attrs[k];
    if (v === undefined) continue;
    if (typeof v === 'string') out += ` ${k}="${escapeAttr(v)}"`;
    else if (typeof v === 'number') out += ` ${k}="${v}"`;
    else out += v ? ` ${k}="1"` : ` ${k}="0"`;
  }
  return out;
}
