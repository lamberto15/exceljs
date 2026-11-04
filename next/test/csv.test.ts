import { describe, expect, it } from 'vitest';
import type { CellValue } from '../src/core/types.js';
import {
  CsvError,
  CsvParser,
  csvRows,
  csvStream,
  formatCsvRow,
  parseCsv,
  writeCsv,
} from '../src/csv/index.js';

function chunked(text: string, sizes: number[]): string[][] {
  const out: string[][] = [];
  const p = new CsvParser((r) => out.push(r));
  let i = 0;
  for (const s of sizes) {
    p.write(text.slice(i, i + s));
    i += s;
  }
  p.write(text.slice(i));
  p.end();
  return out;
}

describe('parse basics', () => {
  it('simple rows', () => {
    expect(parseCsv('a,b,c\n1,2,3\n')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ]);
  });
  it('no trailing newline', () => {
    expect(parseCsv('a,b\n1,2')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });
  it('CRLF, LF, lone CR', () => {
    const exp = [
      ['a', 'b'],
      ['c', 'd'],
      ['e', 'f'],
    ];
    expect(parseCsv('a,b\r\nc,d\r\ne,f\r\n')).toEqual(exp);
    expect(parseCsv('a,b\nc,d\ne,f\n')).toEqual(exp);
    expect(parseCsv('a,b\rc,d\re,f\r')).toEqual(exp);
    expect(parseCsv('a,b\rc,d\ne,f\r\n')).toEqual(exp);
  });
  it('quoted fields with delimiter, newlines, escaped quotes', () => {
    expect(parseCsv('"a,b","line1\nline2","say ""hi"""\r\n"x\r\ny",""\n')).toEqual([
      ['a,b', 'line1\nline2', 'say "hi"'],
      ['x\r\ny', ''],
    ]);
  });
  it('empty fields, trailing delimiter, empty lines', () => {
    expect(parseCsv('a,,c\n,,\nx,\n')).toEqual([
      ['a', '', 'c'],
      ['', '', ''],
      ['x', ''],
    ]);
    expect(parseCsv('a,')).toEqual([['a', '']]);
    expect(parseCsv('a\n\n\nb\n')).toEqual([['a'], ['b']]);
    expect(parseCsv('a\r\n\r\nb')).toEqual([['a'], ['b']]);
    expect(parseCsv('a\n\nb', { skipEmptyLines: false })).toEqual([['a'], [''], ['b']]);
    expect(parseCsv('"",\n')).toEqual([['', '']]);
    expect(parseCsv('"" \n')).toEqual([[' ']]);
    expect(parseCsv('')).toEqual([]);
  });
  it('quote kept literal inside unquoted field; stray text after closing quote', () => {
    expect(parseCsv('a"b,c')).toEqual([['a"b', 'c']]);
    expect(parseCsv('"ab"cd,e')).toEqual([['abcd', 'e']]);
  });
  it('unterminated quote is lenient at EOF', () => {
    expect(parseCsv('a,"bc')).toEqual([['a', 'bc']]);
    expect(parseCsv('"a""')).toEqual([['a"']]);
  });
  it('BOM stripped by default, kept when bom:false', () => {
    expect(parseCsv('﻿a,b\n')).toEqual([['a', 'b']]);
    expect(parseCsv('﻿a,b\n', { bom: false })).toEqual([['﻿a', 'b']]);
  });
  it('custom delimiter and quote', () => {
    expect(parseCsv("a;'b;c';d\n", { delimiter: ';', quote: "'" })).toEqual([['a', 'b;c', 'd']]);
    expect(parseCsv('a\tb\n', { delimiter: '\t' })).toEqual([['a', 'b']]);
    expect(() => parseCsv('x', { delimiter: '::' })).toThrow(CsvError);
  });
  it('typed values', () => {
    const r = parseCsv(
      '1,-2.5,1e3,007,abc,"5",2024-01-02,2024-01-02T03:04:05Z,12345678901234567890\n',
      {
        parseNumbers: true,
        parseDates: true,
      },
    );
    expect(r[0]).toEqual([
      1,
      -2.5,
      1000,
      '007',
      'abc',
      '5',
      new Date(Date.UTC(2024, 0, 2)),
      new Date(Date.UTC(2024, 0, 2, 3, 4, 5)),
      '12345678901234567890',
    ]);
    expect(parseCsv('2024-01-02,3', { parseNumbers: true })).toEqual([['2024-01-02', 3]]);
  });
});

describe('chunk splitting', () => {
  const samples = [
    'a,b,c\r\n"x,1","he said ""hi""",\r\n"multi\nline\r\nfield",z\n\n"",e\rlast,"q"',
    '"""",""""""\n"a""b"\r\n',
    'a,b\r\n\r\nc,d\r',
    '﻿"é€😀",x\n😀,"ü"',
    '"ab"cd,e\n"f"\r\ng',
  ];
  for (const [n, text] of samples.entries()) {
    it(`every split position yields the same result (sample ${n})`, () => {
      const expected = parseCsv(text);
      for (let i = 0; i <= text.length; i++) {
        expect(chunked(text, [i]), `split at ${i}`).toEqual(expected);
      }
      for (let i = 0; i < text.length; i++) {
        for (let j = i; j <= text.length; j++) {
          expect(chunked(text, [i, j - i])).toEqual(expected);
        }
      }
      expect(chunked(text, Array(text.length).fill(1))).toEqual(expected);
    });
  }
  it('random chunking of a generated document', () => {
    let seed = 12345;
    const rnd = (): number => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const alphabet = ['a', 'b', ',', '"', '\n', '\r', ' ', '""', 'é'];
    for (let t = 0; t < 300; t++) {
      let text = '';
      const len = Math.floor(rnd() * 40);
      for (let i = 0; i < len; i++) text += alphabet[Math.floor(rnd() * alphabet.length)];
      const expected = parseCsv(text);
      const sizes: number[] = [];
      for (let i = 0; i < 6; i++) sizes.push(Math.floor(rnd() * 5));
      expect(chunked(text, sizes), JSON.stringify(text)).toEqual(expected);
    }
  });
  it('write after end throws; end is idempotent', () => {
    const p = new CsvParser(() => {});
    p.end();
    p.end();
    expect(() => p.write('a')).toThrow();
  });
});

describe('maxFieldBytes', () => {
  it('throws on oversized unquoted/quoted fields', () => {
    expect(() => parseCsv('a'.repeat(100), { maxFieldBytes: 50 })).toThrow(CsvError);
    expect(() => parseCsv('"' + 'a'.repeat(100) + '"', { maxFieldBytes: 50 })).toThrow(CsvError);
  });
  it('throws while streaming an unterminated quote', () => {
    const p = new CsvParser(() => {}, { maxFieldBytes: 100 });
    p.write('"');
    expect(() => {
      for (let i = 0; i < 100; i++) p.write('x'.repeat(10));
    }).toThrow(/maxFieldBytes/);
  });
  it('allows fields within limit', () => {
    expect(parseCsv('abc,def', { maxFieldBytes: 3 })).toEqual([['abc', 'def']]);
  });
});

describe('csvRows', () => {
  const enc = new TextEncoder();
  it('decodes multi-byte UTF-8 split across Uint8Array chunks', async () => {
    const text = 'é,€,😀\n"日本語",x\n';
    const bytes = enc.encode(text);
    for (let i = 1; i < bytes.length; i++) {
      const src = [bytes.slice(0, i), bytes.slice(i)];
      const rows: string[][] = [];
      for await (const r of csvRows(src as unknown as AsyncIterable<Uint8Array>)) rows.push(r);
      expect(rows).toEqual([
        ['é', '€', '😀'],
        ['日本語', 'x'],
      ]);
    }
  });
  it('byte-at-a-time with BOM from a ReadableStream', async () => {
    const bytes = enc.encode('﻿a,"b\nc"\r\n€,2');
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        for (const b of bytes) c.enqueue(new Uint8Array([b]));
        c.close();
      },
    });
    const rows: string[][] = [];
    for await (const r of csvRows(stream)) rows.push(r);
    expect(rows).toEqual([
      ['a', 'b\nc'],
      ['€', '2'],
    ]);
  });
  it('accepts async iterables of strings and typed option', async () => {
    async function* gen(): AsyncGenerator<string> {
      yield '1,a';
      yield '\n2,b\n';
    }
    const rows: (string | number | Date)[][] = [];
    for await (const r of csvRows(gen(), { parseNumbers: true })) rows.push(r);
    expect(rows).toEqual([
      [1, 'a'],
      [2, 'b'],
    ]);
  });
});

describe('writer', () => {
  it('quotes only when needed', () => {
    expect(formatCsvRow(['a', 'b c', 'x,y', 'q"r', 'l\nm', 'c\rd', ' lead', 'trail ', ''])).toBe(
      'a,b c,"x,y","q""r","l\nm","c\rd"," lead","trail ",',
    );
  });
  it('formats values', () => {
    const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
    const row: CellValue[] = [
      1.5,
      -3,
      true,
      false,
      d,
      null,
      undefined,
      { error: '#N/A' },
      { formula: 'A1', result: 7 },
      { formula: 'A1' },
      { formula: 'A1', result: { error: '#DIV/0!' } },
      { richText: [{ text: 'a' }, { text: 'b' }] },
      { text: 'link', hyperlink: 'http://x' },
    ];
    expect(formatCsvRow(row)).toBe(
      '1.5,-3,TRUE,FALSE,2024-01-02T03:04:05.000Z,,,#N/A,7,,#DIV/0!,ab,link',
    );
    expect(formatCsvRow([d], { dateFormat: (x) => String(x.getUTCFullYear()) })).toBe('2024');
  });
  it('options: delimiter, quote, eol, bom', () => {
    expect(formatCsvRow(['a;b', 'c,d'], { delimiter: ';' })).toBe('"a;b";c,d');
    expect(formatCsvRow(["it's", 'x'], { quote: "'" })).toBe("'it''s',x");
    expect(writeCsv([['a'], ['b']], { eol: '\n', bom: true })).toBe('﻿a\nb\n');
    expect(writeCsv([['a', 'b']])).toBe('a,b\r\n');
  });
  it('escapeFormulas prefixes strings but not numbers', () => {
    const o = { escapeFormulas: true };
    expect(formatCsvRow(['=1+1', '+x', '-x', '@s', '\tz', '\rz', 'ok', -5, 0], o)).toBe(
      "'=1+1,'+x,'-x,'@s,'\tz,\"'\rz\",ok,-5,0",
    );
    expect(formatCsvRow(['=1+1'])).toBe('=1+1');
    expect(formatCsvRow([{ text: '=HYPERLINK()', hyperlink: 'x' }], o)).toBe("'=HYPERLINK()");
  });
  it('single empty field survives round trip', () => {
    expect(formatCsvRow([''])).toBe('""');
    expect(parseCsv(writeCsv([[''], ['a']]))).toEqual([[''], ['a']]);
  });
  it('csvStream is pull-based and encodes UTF-8', async () => {
    let pulled = 0;
    function* gen(): Generator<string[]> {
      for (let i = 0; i < 100000; i++) {
        pulled++;
        yield ['é' + i, 'x'];
      }
    }
    const stream = csvStream(gen(), { bom: true });
    const reader = stream.getReader();
    const first = await reader.read();
    expect(first.value![0]).toBe(0xef);
    expect(pulled).toBeLessThan(100000);
    await reader.cancel();
    const rows: string[][] = [];
    for await (const r of csvRows(
      csvStream([
        ['a', 'b,c'],
        ['€', '"'],
      ]),
    ))
      rows.push(r);
    expect(rows).toEqual([
      ['a', 'b,c'],
      ['€', '"'],
    ]);
  });
  it('csvStream accepts async iterables', async () => {
    async function* g(): AsyncGenerator<string[]> {
      yield ['1'];
      yield ['2'];
    }
    const text = await new Response(csvStream(g(), { eol: '\n' })).text();
    expect(text).toBe('1\n2\n');
  });
});

describe('round trip', () => {
  it('parse(write(rows)) == rows', () => {
    const rows = [
      ['plain', 'with,comma', 'with "quote"', 'multi\nline', ' spaced ', ''],
      ['', '', ''],
      ['é€😀', '\r\n', ',', '"', '""'],
      ['=SUM(A1)', '-1', 'x'],
      ['only'],
    ];
    for (const eol of ['\r\n', '\n']) {
      expect(parseCsv(writeCsv(rows, { eol }))).toEqual(rows);
    }
    expect(
      parseCsv(writeCsv(rows, { delimiter: ';', quote: "'", bom: true }), {
        delimiter: ';',
        quote: "'",
      }),
    ).toEqual(rows);
  });
});
