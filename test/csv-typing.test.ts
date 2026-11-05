// CSV read typing (parseBooleans, parseErrors, dateFormats, localDates) and byte decoding
// (encoding, BOM sniffing, 'auto') for parseCsv / csvRows. Issue #81.

import { describe, expect, it } from 'vitest';
import {
  CsvError,
  type CsvRowsOptions,
  type CsvValue,
  csvRows,
  parseCsv,
} from '../src/csv/index.js';
import { CellKind, Workbook } from '../src/index.js';
import { arraysToSheet } from '../src/json/index.js';

const ERRORS = ['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!'];
const utc = (y: number, m: number, d: number, h = 0, mi = 0, s = 0, ms = 0) =>
  new Date(Date.UTC(y, m - 1, d, h, mi, s, ms));

/** Split bytes into chunks of `size` (0 = one chunk). */
function chunks(b: Uint8Array, size: number): Uint8Array[] {
  if (!size) return [b];
  const out: Uint8Array[] = [];
  for (let i = 0; i < b.length; i += size) out.push(b.slice(i, i + size));
  return out;
}

async function rows(
  src: Iterable<string | Uint8Array>,
  o: CsvRowsOptions = {},
): Promise<CsvValue[][]> {
  const out: CsvValue[][] = [];
  for await (const r of csvRows(src, o)) out.push(r);
  return out;
}

/** Every chunking that matters: whole, 1-, 2-, 3-byte chunks and every two-way split. */
async function allChunkings(b: Uint8Array, o: CsvRowsOptions = {}): Promise<CsvValue[][][]> {
  const res = [];
  for (const size of [0, 1, 2, 3, 7]) res.push(await rows(chunks(b, size), o));
  for (let i = 1; i < b.length; i++) res.push(await rows([b.slice(0, i), b.slice(i)], o));
  return res;
}

function win1252(s: string): Uint8Array {
  return Uint8Array.from(s, (c) => (c === '€' ? 0x80 : c.charCodeAt(0)));
}
function utf16(s: string, be = false, bom = true): Uint8Array {
  const t = bom ? `﻿${s}` : s;
  const b = new Uint8Array(t.length * 2);
  for (let i = 0; i < t.length; i++) {
    const c = t.charCodeAt(i);
    b[2 * i + (be ? 1 : 0)] = c & 0xff;
    b[2 * i + (be ? 0 : 1)] = c >> 8;
  }
  return b;
}
const utf8 = (s: string) => new TextEncoder().encode(s);

describe('parseBooleans', () => {
  it('unquoted true/false in any case become booleans', () => {
    expect(parseCsv('true,FALSE,True,fAlSe\n', { parseBooleans: true })).toEqual([
      [true, false, true, false],
    ]);
  });
  it('quoted values and look-alikes stay strings; off by default', () => {
    expect(parseCsv('"true",truex,yes,1,t, true\n', { parseBooleans: true })).toEqual([
      ['true', 'truex', 'yes', '1', 't', ' true'],
    ]);
    expect(parseCsv('true,false\n')).toEqual([['true', 'false']]);
    expect(parseCsv('true,1\n', { parseNumbers: true })).toEqual([['true', 1]]);
  });
});

describe('parseErrors', () => {
  it('unquoted Excel error literals become CellError', () => {
    expect(parseCsv(`${ERRORS.join(',')}\n`, { parseErrors: true })).toEqual([
      ERRORS.map((error) => ({ error })),
    ]);
  });
  it('quoted, lower-case, unknown and partial literals stay strings; off by default', () => {
    expect(parseCsv('"#N/A",#n/a,#GETTING_DATA,#N/A!,#REF,#\n', { parseErrors: true })).toEqual([
      ['#N/A', '#n/a', '#GETTING_DATA', '#N/A!', '#REF', '#'],
    ]);
    expect(parseCsv('#N/A\n')).toEqual([['#N/A']]);
  });
  it('arraysToSheet writes real error and boolean cells', () => {
    const data = parseCsv('a,b,c\n#N/A,TRUE,#DIV/0!\n', { parseErrors: true, parseBooleans: true });
    const ws = arraysToSheet(new Workbook().addWorksheet('s'), data);
    expect(ws.getKind('A2')).toBe(CellKind.Error);
    expect(ws.getValue('A2')).toEqual({ error: '#N/A' });
    expect(ws.getKind('B2')).toBe(CellKind.Boolean);
    expect(ws.getValue('B2')).toBe(true);
    expect(ws.getValue('C2')).toEqual({ error: '#DIV/0!' });
  });
  it('arraysToSheet still writes other objects as JSON text', () => {
    const ws = arraysToSheet(new Workbook().addWorksheet('s'), [
      [{ error: '#N/A', x: 1 }, { error: '#BOGUS' }, { error: '#N/A' }],
    ]);
    expect(ws.getValue('A1')).toBe('{"error":"#N/A","x":1}');
    expect(ws.getValue('B1')).toBe('{"error":"#BOGUS"}');
    expect(ws.getKind('C1')).toBe(CellKind.Error);
  });
});

describe('dateFormats', () => {
  it('day/month order is explicit, never guessed', () => {
    expect(parseCsv('04/11/2019', { dateFormats: ['DD/MM/YYYY'] })[0]![0]).toEqual(
      utc(2019, 11, 4),
    );
    expect(parseCsv('04/11/2019', { dateFormats: ['MM/DD/YYYY'] })[0]![0]).toEqual(
      utc(2019, 4, 11),
    );
    expect(parseCsv('04/11/2019', { parseDates: true })[0]![0]).toBe('04/11/2019');
  });
  it('times, one-digit tokens, literals and separators', () => {
    const p = (s: string, f: string) => parseCsv(s, { dateFormats: [f] })[0]![0];
    expect(p('04/11/2019 13:05:09', 'DD/MM/YYYY HH:mm:ss')).toEqual(utc(2019, 11, 4, 13, 5, 9));
    expect(p('4.1.2019 7:5', 'D.M.YYYY H:m')).toEqual(utc(2019, 1, 4, 7, 5));
    expect(p('14.11.2019 17:05:3', 'D.M.YYYY H:m:s')).toEqual(utc(2019, 11, 14, 17, 5, 3));
    expect(p('2019-11-04T10:17', 'YYYY-MM-DD[T]HH:mm')).toEqual(utc(2019, 11, 4, 10, 17));
    expect(p('20191104', 'YYYYMMDD')).toEqual(utc(2019, 11, 4));
    expect(p('2019/11/04 (x)', 'YYYY/MM/DD (x)')).toEqual(utc(2019, 11, 4));
    // '.' is literal, not "any char"; two-digit tokens need two digits
    expect(p('04x11x2019', 'DD.MM.YYYY')).toBe('04x11x2019');
    expect(p('4/11/2019', 'DD/MM/YYYY')).toBe('4/11/2019');
    expect(p('04/11/2019 extra', 'DD/MM/YYYY')).toBe('04/11/2019 extra');
  });
  it('rejects impossible dates and times', () => {
    const f = { dateFormats: ['DD/MM/YYYY HH:mm:ss', 'DD/MM/YYYY'] };
    const bad = [
      '31/02/2024',
      '29/02/2023',
      '00/01/2024',
      '01/13/2024',
      '01/00/2024',
      '31/04/2024',
    ];
    bad.push('01/01/2024 24:00:00', '01/01/2024 12:60:00', '01/01/2024 12:00:60');
    for (const s of bad) expect(parseCsv(s, f)[0]![0]).toBe(s);
    expect(parseCsv('29/02/2024', f)[0]![0]).toEqual(utc(2024, 2, 29));
    expect(parseCsv('31/12/2024 23:59:59', f)[0]![0]).toEqual(utc(2024, 12, 31, 23, 59, 59));
  });
  it('formats are tried in order after ISO 8601; an invalid match falls through', () => {
    const o = { dateFormats: ['MM/DD/YYYY', 'DD/MM/YYYY'] };
    expect(parseCsv('04/11/2019,31/12/2019,2019-11-04', o)[0]).toEqual([
      utc(2019, 4, 11),
      utc(2019, 12, 31),
      utc(2019, 11, 4),
    ]);
  });
  it('implies parseDates; numbers win when parseNumbers is set; quoted stays text', () => {
    expect(
      parseCsv('20191104,"04/11/2019"', { dateFormats: ['YYYYMMDD', 'DD/MM/YYYY'] })[0],
    ).toEqual([utc(2019, 11, 4), '04/11/2019']);
    expect(parseCsv('20191104', { dateFormats: ['YYYYMMDD'], parseNumbers: true })[0]![0]).toBe(
      20191104,
    );
  });
  it('bad formats throw BAD_OPTION', () => {
    for (const f of ['MM/DD', 'YY-MM-DD', 'DD/MM/YYYY/DD', 'HH:mm', 'YYYY-MM']) {
      expect(() => parseCsv('x', { dateFormats: [f] })).toThrow(CsvError);
    }
    try {
      parseCsv('x', { dateFormats: ['MM/DD'] });
    } catch (e) {
      expect((e as CsvError).code).toBe('BAD_OPTION');
    }
  });
  it('localDates uses the local wall clock for dates without an offset', () => {
    const o = { dateFormats: ['DD/MM/YYYY HH:mm'], localDates: true };
    expect(parseCsv('04/11/2019 10:30', o)[0]![0]).toEqual(new Date(2019, 10, 4, 10, 30));
    expect(
      parseCsv('2019-11-04,2019-11-04 10:30:15.25', { parseDates: true, localDates: true })[0],
    ).toEqual([new Date(2019, 10, 4), new Date(2019, 10, 4, 10, 30, 15, 250)]);
    // an explicit offset is absolute either way
    expect(
      parseCsv('2019-11-04T10:30:00+01:00', { parseDates: true, localDates: true })[0]![0],
    ).toEqual(utc(2019, 11, 4, 9, 30));
    // local dates land on the same Excel day as the text
    const data = parseCsv('04/11/2019', { dateFormats: ['DD/MM/YYYY'], localDates: true });
    const ws = arraysToSheet(new Workbook({ localDates: true }).addWorksheet('s'), data);
    expect(ws.getValue('A1')).toEqual(new Date(2019, 10, 4));
    const wsUtc = arraysToSheet(
      new Workbook().addWorksheet('s'),
      parseCsv('04/11/2019', { dateFormats: ['DD/MM/YYYY'] }),
    );
    expect(wsUtc.getValue('A1')).toEqual(utc(2019, 11, 4));
  });
  it('ISO 8601: fractions, offsets, and calendar validation', () => {
    expect(
      parseCsv('2024-01-02T03:04:05.1234,2024-01-02 03:04Z,2024-01-02T03:04:05+0130', {
        parseDates: true,
      })[0],
    ).toEqual([utc(2024, 1, 2, 3, 4, 5, 123), utc(2024, 1, 2, 3, 4), utc(2024, 1, 2, 1, 34, 5)]);
    // 2024-02-31 used to roll over to March 2; it now stays text
    expect(parseCsv('2024-02-31,2024-13-01,2024-01-01T25:00', { parseDates: true })[0]).toEqual([
      '2024-02-31',
      '2024-13-01',
      '2024-01-01T25:00',
    ]);
  });
});

describe('typing combinations', () => {
  const text = 'id,ok,err,when,n\n1,TRUE,#N/A,04/11/2019,"7"\n2,false,"#REF!",2019-11-05,x\n';
  const o = {
    parseNumbers: true,
    parseBooleans: true,
    parseErrors: true,
    dateFormats: ['DD/MM/YYYY'],
  };
  const expected = [
    ['id', 'ok', 'err', 'when', 'n'],
    [1, true, { error: '#N/A' }, utc(2019, 11, 4), '7'],
    [2, false, '#REF!', utc(2019, 11, 5), 'x'],
  ];
  it('parseCsv', () => {
    expect(parseCsv(text, o)).toEqual(expected);
    expect(parseCsv(text)[1]).toEqual(['1', 'TRUE', '#N/A', '04/11/2019', '7']);
  });
  it('csvRows applies the same options, with any chunking', async () => {
    for (const r of await allChunkings(utf8(text), o)) expect(r).toEqual(expected);
    expect(await rows([...text], o)).toEqual(expected);
  });
});

describe('encoding', () => {
  const text = 'name,city\ncafé,Zürich\n"€5",naïve\n';
  const expected = [
    ['name', 'city'],
    ['café', 'Zürich'],
    ['€5', 'naïve'],
  ];

  it('UTF-8 with and without BOM is unchanged (BOM stripped, or kept with bom: false)', async () => {
    for (const r of await allChunkings(utf8(text))) expect(r).toEqual(expected);
    for (const r of await allChunkings(utf8(`﻿${text}`))) expect(r).toEqual(expected);
    const kept = await rows(chunks(utf8(`﻿${text}`), 1), { bom: false });
    expect(kept[0]![0]).toBe('﻿name');
  });

  it('explicit encoding: windows-1252, iso-8859-1, utf-16le without BOM', async () => {
    for (const r of await allChunkings(win1252(text), { encoding: 'windows-1252' }))
      expect(r).toEqual(expected);
    expect(await rows([win1252('café\n')], { encoding: 'iso-8859-1' })).toEqual([['café']]);
    for (const r of await allChunkings(utf16(text, false, false), { encoding: 'utf-16le' })) {
      expect(r).toEqual(expected);
    }
    // the default still reads windows-1252 bytes as UTF-8 (replacement characters)
    expect(await rows([win1252('café\n')])).toEqual([['caf�']]);
  });

  it('BOM sniffing by default: UTF-16LE and UTF-16BE, incl. split BOMs and surrogate pairs', async () => {
    const t = `${text}😀,日本\n`;
    const exp = [...expected, ['😀', '日本']];
    for (const r of await allChunkings(utf16(t))) expect(r).toEqual(exp);
    for (const r of await allChunkings(utf16(t, true))) expect(r).toEqual(exp);
    expect((await rows(chunks(utf16(t), 1), { bom: false }))[0]![0]).toBe('﻿name');
  });

  it('a BOM wins over the encoding option', async () => {
    expect(await rows([utf8('﻿café\n')], { encoding: 'windows-1252' })).toEqual([['café']]);
    expect(await rows(chunks(utf16('café\n'), 1), { encoding: 'windows-1252' })).toEqual([
      ['café'],
    ]);
  });

  it("'auto': no BOM and invalid UTF-8 → windows-1252, valid UTF-8 → UTF-8", async () => {
    for (const r of await allChunkings(win1252(text), { encoding: 'auto' }))
      expect(r).toEqual(expected);
    for (const r of await allChunkings(utf8(text), { encoding: 'auto' }))
      expect(r).toEqual(expected);
    for (const r of await allChunkings(utf8(`﻿${text}`), { encoding: 'auto' }))
      expect(r).toEqual(expected);
    for (const r of await allChunkings(utf16(text), { encoding: 'auto' }))
      expect(r).toEqual(expected);
    for (const r of await allChunkings(utf8('a,b\n1,2\n'), { encoding: 'auto' })) {
      expect(r).toEqual([
        ['a', 'b'],
        ['1', '2'],
      ]);
    }
  });

  it("'auto' decides on the first chunk holding non-ASCII bytes", async () => {
    // a long ASCII prefix streams before the decision; the decision chunk sees the whole chunk
    const prefix = `${'x,y\n'.repeat(1000)}`;
    const win = win1252(`${prefix}é,ü\n`);
    expect((await rows(chunks(win, 64), { encoding: 'auto' })).at(-1)).toEqual(['é', 'ü']);
    // C3 A9 is valid UTF-8 ("é"): a lone sequence like that commits to UTF-8...
    const mixed = new Uint8Array([...utf8('é\n'), ...win1252('ü\n')]);
    expect(await rows(chunks(mixed, 3), { encoding: 'auto' })).toEqual([['é'], ['�']]);
    // ...but in one chunk the whole chunk is checked, so the invalid byte selects windows-1252
    expect(await rows([mixed], { encoding: 'auto' })).toEqual([['Ã©'], ['ü']]);
    // an incomplete sequence at the end of the input is not valid UTF-8
    expect(await rows([utf8('ab\n'), new Uint8Array([0xc3])], { encoding: 'auto' })).toEqual([
      ['ab'],
      ['Ã'],
    ]);
  });

  it('tiny inputs and BOM prefixes', async () => {
    expect(await rows([new Uint8Array([0x61])])).toEqual([['a']]);
    expect(await rows([new Uint8Array([0xef, 0xbb])])).toEqual([['�']]);
    expect(await rows([new Uint8Array([0xef, 0xbb])], { encoding: 'auto' })).toEqual([['ï»']]);
    expect(await rows([new Uint8Array([0xff])], { encoding: 'windows-1252' })).toEqual([['ÿ']]);
    expect(await rows([new Uint8Array([0xef, 0xbb, 0xbf])])).toEqual([]);
    expect(await rows([new Uint8Array([0xff, 0xfe])])).toEqual([]);
    expect(await rows([])).toEqual([]);
    expect(await rows([new Uint8Array(0), utf8('a')], { encoding: 'auto' })).toEqual([['a']]);
  });

  it('unknown encodings throw a RangeError when csvRows is called', () => {
    expect(() => csvRows([], { encoding: 'nope' })).toThrow(RangeError);
  });

  it('string chunks bypass decoding', async () => {
    expect(await rows(['café,', 'x\n'], { encoding: 'windows-1252' })).toEqual([['café', 'x']]);
  });
});
