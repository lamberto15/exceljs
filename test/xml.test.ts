import { describe, expect, it } from 'vitest';
import {
  escapeAttr,
  escapeText,
  localName,
  parseXml,
  type XmlAttrs,
  type XmlHandler,
  XmlTokenizer,
  XmlWriter,
} from '../src/xml/index.js';

type Ev = ['open', string, Record<string, string>, boolean] | ['close', string] | ['text', string];

function recorder(): { events: Ev[]; h: XmlHandler } {
  const events: Ev[] = [];
  return {
    events,
    h: {
      open(name, attrs, sc) {
        events.push(['open', name, { ...attrs }, sc]);
      },
      close(name) {
        events.push(['close', name]);
      },
      text(t) {
        events.push(['text', t]);
      },
    },
  };
}

function parse(xml: string): Ev[] {
  const r = recorder();
  parseXml(xml, r.h);
  return r.events;
}

function parseChunks(chunks: string[]): Ev[] {
  const r = recorder();
  const t = new XmlTokenizer(r.h);
  for (const c of chunks) t.write(c);
  t.end();
  return r.events;
}

describe('escapeText / escapeAttr', () => {
  it('returns the same string when nothing to escape', () => {
    const s = 'Hello world 123 \t\n\r é 😀';
    expect(escapeText(s)).toBe(s);
    const a = 'plain value é 😀';
    expect(escapeAttr(a)).toBe(a);
  });

  it('escapes & < > in text but not quotes', () => {
    expect(escapeText('a & b < c > d "q" \'s\'')).toBe('a &amp; b &lt; c &gt; d "q" \'s\'');
  });

  it('escapes quotes and whitespace controls in attributes', () => {
    expect(escapeAttr('a&b<c>d"e\'f')).toBe("a&amp;b&lt;c&gt;d&quot;e'f");
    expect(escapeAttr('l1\nl2\r\tx')).toBe('l1&#10;l2&#13;&#9;x');
  });

  it('strips characters invalid in XML 1.0', () => {
    expect(escapeText('a\u0000b\u0001c\u0008d\u000Be\u000Cf\u001Fg')).toBe('abcdefg');
    expect(escapeText('x￾y￿z')).toBe('xyz');
    expect(escapeText('lone\uD800high')).toBe('lonehigh');
    expect(escapeText('lone\uDC00low')).toBe('lonelow');
    expect(escapeText('end\uD83D')).toBe('end');
    expect(escapeText('pair😀ok&')).toBe('pair😀ok&amp;');
    expect(escapeAttr('a\u0002b"')).toBe('ab&quot;');
  });

  it('round-trips through the tokenizer', () => {
    const nasty = 'a & b < c > "d" \'e\'\n\tline\r\n😀 &amp; ]]>';
    const xml = `<r a="${escapeAttr(nasty)}">${escapeText(nasty)}</r>`;
    const ev = parse(xml);
    expect(ev[0]).toEqual(['open', 'r', { a: nasty }, false]);
    // text: CRLF is normalized to LF by XML parsing rules
    expect(ev[1]).toEqual(['text', nasty.replace('\r\n', '\n')]);
  });
});

describe('XmlWriter', () => {
  it('builds elements with attributes, leaves and text', () => {
    const w = new XmlWriter();
    w.raw('<?xml version="1.0"?>')
      .open('row', { r: 1, spans: '1:2', hidden: undefined, customHeight: true, collapsed: false })
      .leaf('c', { r: 'A1', t: 's' }, '0')
      .leaf('c', { r: 'B1' })
      .open('c')
      .text('a<b & "c"')
      .close('c')
      .close('row');
    expect(w.toString()).toBe(
      '<?xml version="1.0"?><row r="1" spans="1:2" customHeight="1" collapsed="0">' +
        '<c r="A1" t="s">0</c><c r="B1"/><c>a&lt;b &amp; "c"</c></row>',
    );
  });

  it('escapes attribute values and leaf text', () => {
    const w = new XmlWriter();
    w.leaf('t', { v: 'x"y&\n' }, '<&>');
    expect(w.toString()).toBe('<t v="x&quot;y&amp;&#10;">&lt;&amp;&gt;</t>');
  });

  it('leaf with empty-string text writes open+close', () => {
    expect(new XmlWriter().leaf('t', undefined, '').toString()).toBe('<t></t>');
  });

  it('calls onFlush when buffer exceeds flushAt and on flush()', () => {
    const out: string[] = [];
    const w = new XmlWriter((c) => out.push(c), 32);
    for (let i = 0; i < 20; i++) w.leaf('v', undefined, String(i));
    expect(out.length).toBeGreaterThan(1);
    for (const c of out) expect(c.length).toBeGreaterThanOrEqual(32);
    w.flush();
    expect(w.toString()).toBe('');
    let expected = '';
    for (let i = 0; i < 20; i++) expected += `<v>${i}</v>`;
    expect(out.join('')).toBe(expected);
    const n = out.length;
    w.flush(); // nothing buffered -> no empty chunk
    expect(out.length).toBe(n);
  });

  it('keeps everything in toString() without onFlush', () => {
    const w = new XmlWriter(undefined, 16);
    let expected = '';
    for (let i = 0; i < 100; i++) {
      w.leaf('v', undefined, String(i));
      expected += `<v>${i}</v>`;
    }
    w.flush(); // no-op
    expect(w.toString()).toBe(expected);
  });
});

describe('XmlTokenizer', () => {
  it('parses namespaced OOXML cell markup', () => {
    expect(parse('<x:c r="A1" t="s"><x:v>1</x:v></x:c>')).toEqual([
      ['open', 'x:c', { r: 'A1', t: 's' }, false],
      ['open', 'x:v', {}, false],
      ['text', '1'],
      ['close', 'x:v'],
      ['close', 'x:c'],
    ]);
    expect(localName('x:c')).toBe('c');
    expect(localName('c')).toBe('c');
  });

  it('handles self-closing tags (open selfClosing=true, then close)', () => {
    expect(parse('<a><b x="1"/><c/><d /></a>')).toEqual([
      ['open', 'a', {}, false],
      ['open', 'b', { x: '1' }, true],
      ['close', 'b'],
      ['open', 'c', {}, true],
      ['close', 'c'],
      ['open', 'd', {}, true],
      ['close', 'd'],
      ['close', 'a'],
    ]);
  });

  it('attrs are null-prototype objects', () => {
    let seen: XmlAttrs[] = [];
    parseXml('<a x="1"><b/></a>', {
      open(_n, attrs) {
        seen.push(attrs);
      },
    });
    for (const a of seen) expect(Object.getPrototypeOf(a)).toBe(null);
    seen = [];
    parseXml('<a __proto__="x" constructor="y"/>', {
      open(_n, attrs) {
        seen.push(attrs);
      },
    });
    // biome-ignore lint/suspicious/noProto: verifies attrs are null-prototype
    expect(seen[0]?.__proto__).toBe('x');
    expect(seen[0]?.constructor).toBe('y');
  });

  it('supports single/double quotes and whitespace around =', () => {
    expect(parse(`<a x='1"2' y = "3'4"\n\tz='5'></a>`)).toEqual([
      ['open', 'a', { x: '1"2', y: "3'4", z: '5' }, false],
      ['close', 'a'],
    ]);
  });

  it('decodes predefined entities and numeric refs', () => {
    expect(
      parse('<t a="&lt;&amp;&#10;&#x1F600;">&lt;&gt;&amp;&quot;&apos;&#65;&#x42;&#x1f600;</t>'),
    ).toEqual([
      ['open', 't', { a: '<&\n😀' }, false],
      ['text', `<>&"'AB😀`],
      ['close', 't'],
    ]);
  });

  it('leaves unknown / malformed entities as literal text', () => {
    expect(parse('<t>&foo; &amp &#; &#xZZ; &#0; &#xD800; a&b</t>')[1]).toEqual([
      'text',
      '&foo; &amp &#; &#xZZ; &#0; &#xD800; a&b',
    ]);
  });

  it('normalizes line endings in text and whitespace in attributes', () => {
    expect(parse('<t a="x\r\ny\tz">l1\r\nl2\rl3&#13;</t>')).toEqual([
      ['open', 't', { a: 'x y z' }, false],
      ['text', 'l1\nl2\nl3\r'],
      ['close', 't'],
    ]);
  });

  it('delivers CDATA as (undecoded) text, merged with adjacent text', () => {
    expect(parse('<t>a<![CDATA[<b> &amp; ]] ]]>c</t>')).toEqual([
      ['open', 't', {}, false],
      ['text', 'a<b> &amp; ]] c'],
      ['close', 't'],
    ]);
  });

  it('skips comments and processing instructions', () => {
    expect(
      parse(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!-- c <a> -- x --><r>a<!-- in -->b<?pi x?></r>',
      ),
    ).toEqual([
      ['text', '\n'],
      ['open', 'r', {}, false],
      ['text', 'ab'],
      ['close', 'r'],
    ]);
  });

  it('delivers whitespace-only text but never empty text', () => {
    expect(parse('<a>\n  <b></b>\n</a>')).toEqual([
      ['open', 'a', {}, false],
      ['text', '\n  '],
      ['open', 'b', {}, false],
      ['close', 'b'],
      ['text', '\n'],
      ['close', 'a'],
    ]);
    expect(parse('<a></a>').some((e) => e[0] === 'text')).toBe(false);
  });

  it('ignores DOCTYPE and never expands custom entities (billion laughs)', () => {
    const xml = `<?xml version="1.0"?>
<!DOCTYPE lolz [
 <!ENTITY lol "lol">
 <!ELEMENT lolz (#PCDATA)>
 <!-- a comment with > and ' inside -->
 <?pi with > inside?>
 <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
 <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
 <!ENTITY lol9 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
 <!ENTITY xxe SYSTEM "file:///etc/passwd">
 <!ENTITY gt2 '>>'>
]>
<lolz>&lol9;&xxe;</lolz>`;
    expect(parse(xml)).toEqual([
      ['text', '\n\n'],
      ['open', 'lolz', {}, false],
      ['text', '&lol9;&xxe;'],
      ['close', 'lolz'],
    ]);
    expect(parse('<!DOCTYPE html><r/>')).toEqual([
      ['open', 'r', {}, true],
      ['close', 'r'],
    ]);
  });

  it('handles surrogate pairs and non-BMP text', () => {
    expect(parse('<t>😀 𝄞</t>')[1]).toEqual(['text', '😀 𝄞']);
  });

  it('trims whitespace in close tags', () => {
    expect(parse('<a></a >')).toEqual([
      ['open', 'a', {}, false],
      ['close', 'a'],
    ]);
  });

  it('throws on truncated markup at end()', () => {
    const t = new XmlTokenizer({});
    t.write('<a><b x="1');
    expect(() => t.end()).toThrow();
  });

  it('delivers trailing text at end()', () => {
    expect(parseChunks(['<a/>tail &am', 'p; x'])).toEqual([
      ['open', 'a', {}, true],
      ['close', 'a'],
      ['text', 'tail & x'],
    ]);
  });

  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!DOCTYPE worksheet [ <!ENTITY e "boom"> <!-- ]> --> ]>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- comment with <tags> & stuff -->
  <dimension ref="A1:C3"/>
  <sheetData>
    <row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="B1" t='inlineStr'><is><t xml:space="preserve"> a &amp; b &lt;c&gt; &#x1F600; &#65; </t></is></c></row>
    <row r="2"><c r="A2"><f>SUM(A1:B1)&gt;0</f><v>1.5E+3</v></c><c r="B2" t="str"><v><![CDATA[raw <cdata> &amp; ]]]]><![CDATA[>]]></v></c></row>
    <x:row r = '3' x:attr="&quot;q&quot;&#10;&apos;"><x:c r="A3" t="b"><x:v>1</x:v></x:c>😀𝄞 mixed&#13;&#10;text&unknown;</x:row>
  </sheetData>
  <?some-pi data?>
  <legacyDrawing r:id="rId1" />
</worksheet>
`;

  it('emits identical events regardless of chunk boundaries (every split point)', () => {
    const expected = parse(doc);
    expect(expected.length).toBeGreaterThan(30);
    for (let i = 0; i <= doc.length; i++) {
      expect(parseChunks([doc.slice(0, i), doc.slice(i)])).toEqual(expected);
    }
  });

  it('emits identical events when split into every pair of boundaries (sampled) and single chars', () => {
    const expected = parse(doc);
    expect(parseChunks(Array.from(doc))).toEqual(expected); // per code point
    expect(parseChunks(doc.split(''))).toEqual(expected); // per UTF-16 unit (splits surrogate pairs)
    for (let i = 1; i < doc.length; i += 7) {
      for (let j = i; j < doc.length; j += 13) {
        expect(parseChunks([doc.slice(0, i), doc.slice(i, j), doc.slice(j)])).toEqual(expected);
      }
    }
  });

  it('long text / CDATA / comments spanning many chunks', () => {
    const big = 'x'.repeat(100_000);
    const xml = `<a>${big}&amp;<!--${big}--><![CDATA[${big}]]></a>`;
    const chunks: string[] = [];
    for (let i = 0; i < xml.length; i += 997) chunks.push(xml.slice(i, i + 997));
    expect(parseChunks(chunks)).toEqual([
      ['open', 'a', {}, false],
      ['text', `${big}&${big}`],
      ['close', 'a'],
    ]);
  });
});

// Manual throughput check: RUN_BENCH=1 npx vitest run test/xml.test.ts
const env = (globalThis as { process?: { env: Record<string, string | undefined> } }).process?.env;
describe.skipIf(!env?.RUN_BENCH)('throughput (manual)', () => {
  it('tokenizes a ~20MB sheet-like document', () => {
    const parts: string[] = ['<?xml version="1.0"?><worksheet><sheetData>'];
    let size = 0;
    let r = 1;
    while (size < 20 * 1024 * 1024) {
      let row = `<row r="${r}" spans="1:10">`;
      for (let c = 0; c < 10; c++) {
        const ref = String.fromCharCode(65 + c) + r;
        row +=
          c % 3 === 0
            ? `<c r="${ref}" t="s" s="1"><v>${r * 10 + c}</v></c>`
            : c % 3 === 1
              ? `<c r="${ref}"><v>${(r * 1.234).toFixed(3)}</v></c>`
              : `<c r="${ref}" t="inlineStr"><is><t>Text &amp; ${r}</t></is></c>`;
      }
      row += '</row>';
      parts.push(row);
      size += row.length;
      r++;
    }
    parts.push('</sheetData></worksheet>');
    const xml = parts.join('');
    const chunkSize = 64 * 1024;
    let opens = 0;
    let texts = 0;
    const h: XmlHandler = {
      open() {
        opens++;
      },
      close() {},
      text() {
        texts++;
      },
    };
    for (let k = 0; k < 2; k++) {
      opens = 0;
      texts = 0;
      const t0 = performance.now();
      const t = new XmlTokenizer(h);
      for (let i = 0; i < xml.length; i += chunkSize) t.write(xml.slice(i, i + chunkSize));
      t.end();
      const ms = performance.now() - t0;
      const mb = xml.length / 1024 / 1024;
      console.log(
        `tokenizer: ${mb.toFixed(1)} MB in ${ms.toFixed(0)} ms = ${(mb / (ms / 1000)).toFixed(0)} MB/s (${opens} opens, ${texts} texts)`,
      );
    }

    const t0 = performance.now();
    const w = new XmlWriter(() => {});
    for (let i = 0; i < r; i++) {
      w.open('row', { r: i });
      for (let c = 0; c < 10; c++) w.leaf('c', { r: `A${i}`, t: 's' }, 'Text & more');
      w.close('row');
    }
    w.flush();
    console.log(`writer: ${r} rows x 10 cells in ${(performance.now() - t0).toFixed(0)} ms`);
    expect(opens).toBeGreaterThan(0);
  });
});
