// Streaming XML push-tokenizer tuned for large OOXML parts.
//
// - Input arrives as string chunks that may be split at any position.
// - Only the unconsumed tail of a chunk is carried over; text runs are accumulated
//   separately so very long text never forces re-scanning.
// - `<?...?>` and comments are skipped; `<!DOCTYPE ...>` (incl. internal subset) is ignored
//   entirely, so custom/external entities are never expanded.
// - Only the 5 predefined entities and numeric char refs are decoded; unknown entities stay literal.
// - Line endings are normalized (CRLF / CR -> LF) and attribute whitespace (TAB/LF/CR -> space)
//   as required by the XML spec; char refs such as `&#10;` survive.
// - Adjacent character data (text, CDATA, text around comments/PIs) is merged into a single
//   `text` event, delivered right before the next open/close tag (or at end()).
// - Self-closing tags emit `open(name, attrs, true)` immediately followed by `close(name)`.
// - The tokenizer is not a validator: nesting is not checked.

export type XmlAttrs = Record<string, string>;

export interface XmlHandler {
  open?(name: string, attrs: XmlAttrs, selfClosing: boolean): void;
  close?(name: string): void;
  /** Entities decoded; CDATA delivered as text. */
  text?(text: string): void;
}

/**
 * Shared attrs object for tags without attributes (frozen, null-prototype).
 * Consumers must treat attrs as read-only.
 */
const EMPTY_ATTRS: XmlAttrs = Object.freeze(Object.create(null) as XmlAttrs);

const LT = 60; // <
const GT = 62; // >
const SLASH = 47; // /
const BANG = 33; // !
const QMARK = 63; // ?
const EQ = 61; // =
const DQ = 34; // "
const SQ = 39; // '
const AMP = 38; // &
const LBRACK = 91; // [
const RBRACK = 93; // ]

/** Strip a namespace prefix: `x:c` -> `c`. */
export function localName(name: string): string {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
}

export class XmlTokenizer {
  private readonly h: XmlHandler;
  /** Unconsumed tail (always starts at a `<` of an incomplete markup token). */
  private tail: string = '';
  /** Raw (undecoded) text that ran up to the end of the previous chunk. */
  private rawText: string = '';
  /** Decoded character data waiting to be emitted. */
  private text: string = '';
  /** Offset (relative to tail start) where an incomplete comment/CDATA/PI search resumes. */
  private resume: number = 0;
  private ended: boolean = false;

  constructor(h: XmlHandler) {
    this.h = h;
  }

  write(chunk: string): void {
    if (this.ended) throw new Error('XmlTokenizer: write after end');
    if (chunk.length === 0) return;
    const s = this.tail === '' ? chunk : this.tail + chunk;
    this.tail = '';
    this.run(s);
  }

  end(): void {
    if (this.ended) return;
    this.ended = true;
    if (this.tail !== '') throw new Error('XmlTokenizer: unexpected end of input inside markup');
    if (this.rawText !== '') {
      this.appendText(decodeText(this.rawText));
      this.rawText = '';
    }
    this.flushText();
  }

  private appendText(t: string): void {
    if (t !== '') this.text = this.text === '' ? t : this.text + t;
  }

  private flushText(): void {
    const t = this.text;
    if (t !== '') {
      this.text = '';
      this.h.text?.(t);
    }
  }

  private run(s: string): void {
    const n = s.length;
    let p = 0;
    while (p < n) {
      if (s.charCodeAt(p) !== LT) {
        const lt = s.indexOf('<', p);
        if (lt < 0) {
          // text continues into the next chunk; keep it raw (entities/CRLF may be split)
          this.rawText = this.rawText === '' ? s.slice(p) : this.rawText + s.slice(p);
          return;
        }
        if (this.rawText === '') {
          this.appendText(decodeText(s.slice(p, lt)));
        } else {
          this.appendText(decodeText(this.rawText + s.slice(p, lt)));
          this.rawText = '';
        }
        p = lt;
      } else if (this.rawText !== '') {
        this.appendText(decodeText(this.rawText));
        this.rawText = '';
      }
      const r = this.markup(s, p, n);
      if (r < 0) {
        this.tail = s.slice(p);
        return;
      }
      p = r;
    }
  }

  /** Parse one markup token starting at s[p] === '<'. Returns end index, or -1 if incomplete. */
  private markup(s: string, p: number, n: number): number {
    if (p + 1 >= n) return -1;
    const c = s.charCodeAt(p + 1);
    if (c === SLASH) return this.closeTag(s, p);
    if (c === QMARK) return this.skipTo(s, p, n, '?>', 2);
    if (c === BANG) {
      if (n - p < 4) return -1;
      if (s.charCodeAt(p + 2) === 45 && s.charCodeAt(p + 3) === 45)
        return this.skipTo(s, p, n, '-->', 4);
      if (n - p < 9 && '<![CDATA['.startsWith(s.slice(p, n))) return -1;
      if (s.startsWith('<![CDATA[', p)) return this.cdata(s, p, n);
      return skipDecl(s, p, n);
    }
    if (c <= 32) {
      // '<' followed by whitespace is not markup; be lenient and keep it as text
      this.appendText('<');
      return p + 1;
    }
    return this.openTag(s, p, n);
  }

  private skipTo(s: string, p: number, n: number, term: string, minOff: number): number {
    const from = p + Math.max(minOff, this.resume);
    const e = s.indexOf(term, from);
    if (e < 0) {
      this.resume = Math.max(minOff, n - p - term.length + 1);
      return -1;
    }
    this.resume = 0;
    return e + term.length;
  }

  private cdata(s: string, p: number, n: number): number {
    const from = p + Math.max(9, this.resume);
    const e = s.indexOf(']]>', from);
    if (e < 0) {
      this.resume = Math.max(9, n - p - 2);
      return -1;
    }
    this.resume = 0;
    const raw = s.slice(p + 9, e);
    this.appendText(raw.indexOf('\r') < 0 ? raw : normalizeEol(raw));
    return e + 3;
  }

  private closeTag(s: string, p: number): number {
    const e = s.indexOf('>', p + 2);
    if (e < 0) return -1;
    let ne = e;
    while (ne > p + 2 && s.charCodeAt(ne - 1) <= 32) ne--;
    this.flushText();
    this.h.close?.(s.slice(p + 2, ne));
    return e + 1;
  }

  private openTag(s: string, p: number, n: number): number {
    let i = p + 1;
    let c = 0;
    while (i < n) {
      c = s.charCodeAt(i);
      if (c === GT || c === SLASH || c <= 32) break;
      i++;
    }
    if (i >= n) return -1;
    const name = s.slice(p + 1, i);
    let attrs: XmlAttrs | null = null;
    for (;;) {
      while (i < n && s.charCodeAt(i) <= 32) i++;
      if (i >= n) return -1;
      c = s.charCodeAt(i);
      if (c === GT) {
        this.emitOpen(name, attrs, false);
        return i + 1;
      }
      if (c === SLASH) {
        if (i + 1 >= n) return -1;
        if (s.charCodeAt(i + 1) === GT) {
          this.emitOpen(name, attrs, true);
          return i + 2;
        }
        i++; // stray '/', ignore
        continue;
      }
      // attribute name
      const ns = i;
      while (i < n) {
        c = s.charCodeAt(i);
        if (c === EQ || c <= 32 || c === GT || c === SLASH) break;
        i++;
      }
      if (i >= n) return -1;
      const an = s.slice(ns, i);
      while (i < n && s.charCodeAt(i) <= 32) i++;
      if (i >= n) return -1;
      if (s.charCodeAt(i) !== EQ) {
        // valueless attribute (not well-formed XML): record as empty
        if (attrs === null) attrs = Object.create(null) as XmlAttrs;
        attrs[an] = '';
        continue;
      }
      i++;
      while (i < n && s.charCodeAt(i) <= 32) i++;
      if (i >= n) return -1;
      const q = s.charCodeAt(i);
      let value: string;
      if (q === DQ || q === SQ) {
        const e = s.indexOf(q === DQ ? '"' : "'", i + 1);
        if (e < 0) return -1;
        value = decodeAttr(s.slice(i + 1, e));
        i = e + 1;
      } else {
        // unquoted value (lenient)
        const vs = i;
        while (i < n) {
          c = s.charCodeAt(i);
          if (c <= 32 || c === GT) break;
          i++;
        }
        if (i >= n) return -1;
        value = decodeAttr(s.slice(vs, i));
      }
      if (attrs === null) attrs = Object.create(null) as XmlAttrs;
      attrs[an] = value;
    }
  }

  private emitOpen(name: string, attrs: XmlAttrs | null, selfClosing: boolean): void {
    this.flushText();
    const h = this.h;
    h.open?.(name, attrs ?? EMPTY_ATTRS, selfClosing);
    if (selfClosing) h.close?.(name);
  }
}

/** Skip `<!DOCTYPE ...>` / other `<!...>` declarations, including an internal subset `[...]`. */
function skipDecl(s: string, p: number, n: number): number {
  let i = p + 2;
  let depth = 0;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === DQ || c === SQ) {
      const e = s.indexOf(c === DQ ? '"' : "'", i + 1);
      if (e < 0) return -1;
      i = e + 1;
      continue;
    }
    if (c === LT && depth > 0) {
      if (s.startsWith('<!--', i)) {
        const e = s.indexOf('-->', i + 4);
        if (e < 0) return -1;
        i = e + 3;
        continue;
      }
      if (i + 1 < n && s.charCodeAt(i + 1) === QMARK) {
        const e = s.indexOf('?>', i + 2);
        if (e < 0) return -1;
        i = e + 2;
        continue;
      }
    } else if (c === LBRACK) depth++;
    else if (c === RBRACK) depth--;
    else if (c === GT && depth <= 0) return i + 1;
    i++;
  }
  return -1;
}

function normalizeEol(s: string): string {
  let out = '';
  let last = 0;
  let i = s.indexOf('\r');
  while (i >= 0) {
    out += `${s.slice(last, i)}\n`;
    last = s.charCodeAt(i + 1) === 10 ? i + 2 : i + 1;
    i = s.indexOf('\r', last);
  }
  return out + s.slice(last);
}

function decodeText(s: string): string {
  if (s.indexOf('&') < 0) return s.indexOf('\r') < 0 ? s : normalizeEol(s);
  return decodeSlow(s, 0, false);
}

function decodeAttr(s: string): string {
  const n = s.length;
  for (let i = 0; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c === AMP || c === 9 || c === 10 || c === 13) return decodeSlow(s, i, true);
  }
  return s;
}

function decodeSlow(s: string, start: number, attr: boolean): string {
  const n = s.length;
  let out = '';
  let last = 0;
  for (let i = start; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c === AMP) {
      const semi = s.indexOf(';', i + 1);
      if (semi < 0 || semi - i > 12) continue;
      const r = resolveEntity(s, i + 1, semi);
      if (r === undefined) continue;
      out += s.slice(last, i) + r;
      last = semi + 1;
      i = semi;
    } else if (c === 13) {
      out += s.slice(last, i) + (attr ? ' ' : '\n');
      if (i + 1 < n && s.charCodeAt(i + 1) === 10) i++;
      last = i + 1;
    } else if (attr && (c === 10 || c === 9)) {
      out += `${s.slice(last, i)} `;
      last = i + 1;
    }
  }
  return last === 0 ? s : out + s.slice(last);
}

function resolveEntity(s: string, a: number, b: number): string | undefined {
  const len = b - a;
  const c0 = s.charCodeAt(a);
  if (c0 === 35 /* # */) {
    let cp: number;
    if (len > 2 && (s.charCodeAt(a + 1) | 0x20) === 120 /* x */) {
      const hex = s.slice(a + 2, b);
      if (!/^[0-9a-fA-F]+$/.test(hex)) return undefined;
      cp = parseInt(hex, 16);
    } else {
      const dec = s.slice(a + 1, b);
      if (dec.length === 0 || !/^[0-9]+$/.test(dec)) return undefined;
      cp = parseInt(dec, 10);
    }
    if (cp === 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return undefined;
    return String.fromCodePoint(cp);
  }
  switch (len) {
    case 2:
      if (s.charCodeAt(a + 1) === 116 /* t */) {
        if (c0 === 108 /* l */) return '<';
        if (c0 === 103 /* g */) return '>';
      }
      return undefined;
    case 3:
      return s.startsWith('amp', a) ? '&' : undefined;
    case 4:
      if (s.startsWith('quot', a)) return '"';
      if (s.startsWith('apos', a)) return "'";
      return undefined;
    default:
      return undefined;
  }
}

/** Parse a complete XML string. */
export function parseXml(xml: string, h: XmlHandler): void {
  const t = new XmlTokenizer(h);
  t.write(xml);
  t.end();
}
