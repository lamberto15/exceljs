import { describe, expect, it } from 'vitest';
import {
  type CellError,
  type Formula,
  type Hyperlink,
  type RichText,
  Workbook,
  type Worksheet,
} from '../src/index.js';
import {
  arraysToSheet,
  objectsToSheet,
  sheetToArrays,
  sheetToNdjson,
  sheetToObjects,
  toJsonValue,
} from '../src/json/index.js';

function newSheet(): Worksheet {
  return new Workbook().addWorksheet('S');
}

async function collect(gen: AsyncGenerator<string>): Promise<string[]> {
  const out: string[] = [];
  for await (const line of gen) out.push(line);
  return out;
}

describe('toJsonValue', () => {
  it('converts every cell kind to a JSON-safe value', () => {
    const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
    expect(toJsonValue(null)).toBeNull();
    expect(toJsonValue(undefined)).toBeNull();
    expect(toJsonValue(1.5)).toBe(1.5);
    expect(toJsonValue(Number.NaN)).toBeNull();
    expect(toJsonValue(Number.POSITIVE_INFINITY)).toBeNull();
    expect(toJsonValue('x')).toBe('x');
    expect(toJsonValue(false)).toBe(false);
    expect(toJsonValue(d)).toBe('2024-01-02T03:04:05.000Z');
    expect(toJsonValue(new Date(Number.NaN))).toBeNull();
    expect(toJsonValue({ error: '#DIV/0!' } as CellError)).toBe('#DIV/0!');
    expect(toJsonValue({ formula: 'A1+1', result: 3 } as Formula)).toBe(3);
    expect(toJsonValue({ formula: 'A1+1', result: { error: '#N/A' } } as Formula)).toBe('#N/A');
    expect(toJsonValue({ formula: 'A1+1' } as Formula)).toBeNull();
    expect(toJsonValue({ richText: [{ text: 'Hello ' }, { text: 'world' }] } as RichText)).toBe(
      'Hello world',
    );
    expect(toJsonValue({ text: 'link', hyperlink: 'https://x.y' } as Hyperlink)).toBe('link');
  });
});

describe('sheetToArrays', () => {
  it('reads the used range with empty cells as null', () => {
    const ws = newSheet();
    ws.addRow(['a', 1]);
    ws.setCell('C2', true);
    expect(sheetToArrays(ws)).toEqual([
      ['a', 1, null],
      [null, null, true],
    ]);
  });

  it('honours range', () => {
    const ws = newSheet();
    ws.addRow(['a', 'b', 'c']);
    ws.addRow([1, 2, 3]);
    ws.addRow([4, 5, 6]);
    expect(sheetToArrays(ws, { range: 'B2:C3' })).toEqual([
      [2, 3],
      [5, 6],
    ]);
    expect(sheetToArrays(ws, { range: 'A1' })).toEqual([['a']]);
  });

  it('raw=false converts values; raw=true keeps Dates and formulas', () => {
    const ws = newSheet();
    const d = new Date(Date.UTC(2020, 5, 1));
    ws.addRow([d, { formula: 'A1', result: 7 } as Formula]);
    expect(sheetToArrays(ws)[0]).toEqual([d, { formula: 'A1', result: 7 }]);
    expect(sheetToArrays(ws, { raw: false })[0]).toEqual(['2020-06-01T00:00:00.000Z', 7]);
  });

  it('returns [] for an empty sheet', () => {
    expect(sheetToArrays(newSheet())).toEqual([]);
  });
});

describe('sheetToObjects', () => {
  it('uses row 1 as header by default and skips blank rows', () => {
    const ws = newSheet();
    ws.addRow(['name', 'qty']);
    ws.addRow(['a', 1]);
    ws.addRow([null, null]);
    ws.addRow(['b', 2]);
    expect(sheetToObjects(ws)).toEqual([
      { name: 'a', qty: 1 },
      { name: 'b', qty: 2 },
    ]);
  });

  it('reads the header from another row number', () => {
    const ws = newSheet();
    ws.addRow(['report title']);
    ws.addRow(['id', 'label']);
    ws.addRow([1, 'x']);
    expect(sheetToObjects(ws, { header: 2 })).toEqual([{ id: 1, label: 'x' }]);
  });

  it('accepts explicit keys (no header row is consumed)', () => {
    const ws = newSheet();
    ws.addRow(['a', 'b']);
    ws.addRow(['c', 'd']);
    expect(sheetToObjects(ws, { header: ['x', 'y'] })).toEqual([
      { x: 'a', y: 'b' },
      { x: 'c', y: 'd' },
    ]);
  });

  it('honours range with header row outside the range', () => {
    const ws = newSheet();
    ws.addRow(['id', 'name', 'note']);
    ws.addRow([1, 'a', 'n1']);
    ws.addRow([2, 'b', 'n2']);
    ws.addRow([3, 'c', 'n3']);
    expect(sheetToObjects(ws, { range: 'B2:C3' })).toEqual([
      { name: 'a', note: 'n1' },
      { name: 'b', note: 'n2' },
    ]);
  });

  it('applies defval to empty cells, and omits them by default', () => {
    const ws = newSheet();
    ws.addRow(['a', 'b', 'c']);
    ws.addRow([1, null, 3]);
    expect(sheetToObjects(ws)).toEqual([{ a: 1, c: 3 }]);
    expect(sheetToObjects(ws, { defval: 0 })).toEqual([{ a: 1, b: 0, c: 3 }]);
    expect(Object.keys(sheetToObjects(ws, { defval: '' })[0]!)).toEqual(['a', 'b', 'c']);
  });

  it('defval is converted when raw=false', () => {
    const ws = newSheet();
    ws.addRow(['a', 'b']);
    ws.addRow([1, null]);
    const d = new Date(Date.UTC(2021, 0, 1));
    expect(sheetToObjects(ws, { defval: d, raw: false })).toEqual([
      { a: 1, b: '2021-01-01T00:00:00.000Z' },
    ]);
  });

  it('raw=true keeps Date and number; raw=false produces JSON-safe values', () => {
    const ws = newSheet();
    const d = new Date(Date.UTC(2022, 2, 3));
    ws.addRow(['when', 'v', 'f', 'rt', 'link', 'err']);
    ws.addRow([
      d,
      2.5,
      { formula: 'B2', result: 'ok' },
      { richText: [{ text: 'a' }, { text: 'b' }] },
      { text: 'site', hyperlink: 'https://e.x' },
      { error: '#N/A' },
    ]);
    const [rawRow] = sheetToObjects(ws);
    expect(rawRow).toEqual({
      when: d,
      v: 2.5,
      f: { formula: 'B2', result: 'ok' },
      rt: { richText: [{ text: 'a' }, { text: 'b' }] },
      link: { text: 'site', hyperlink: 'https://e.x' },
      err: { error: '#N/A' },
    });
    expect(sheetToObjects(ws, { raw: false })).toEqual([
      {
        when: '2022-03-03T00:00:00.000Z',
        v: 2.5,
        f: 'ok',
        rt: 'ab',
        link: 'site',
        err: '#N/A',
      },
    ]);
  });

  it('dedupes duplicate headers and names empty headers by column', () => {
    const ws = newSheet();
    ws.addRow(['id', 'id', '', 'id_2', 'id']);
    ws.addRow([1, 2, 3, 4, 5]);
    expect(sheetToObjects(ws)).toEqual([{ id: 1, id_2: 2, __EMPTY_3: 3, id_2_2: 4, id_3: 5 }]);
  });

  it('dedupes explicit keys as well', () => {
    const ws = newSheet();
    ws.addRow([1, 2, 3]);
    expect(sheetToObjects(ws, { header: ['k', 'k', ''] })).toEqual([
      { k: 1, k_2: 2, __EMPTY_3: 3 },
    ]);
  });

  it('guards prototype keys by renaming them with a leading underscore', () => {
    const ws = newSheet();
    ws.addRow(['__proto__', 'constructor', 'prototype', 'ok']);
    ws.addRow(['x', 'y', 'z', 'w']);
    const [row] = sheetToObjects(ws);
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(Object.keys(row!)).toEqual(['___proto__', '_constructor', '_prototype', 'ok']);
    expect(row!['___proto__']).toBe('x');
    expect(({} as Record<string, unknown>)['polluted']).toBeUndefined();
    expect(Object.hasOwn(row!, '__proto__')).toBe(false);
  });

  it('rejects an invalid header row number', () => {
    expect(() => sheetToObjects(newSheet(), { header: 0 })).toThrow(RangeError);
    expect(() => sheetToObjects(newSheet(), { header: 1.5 })).toThrow(RangeError);
  });

  it('returns [] for an empty sheet', () => {
    expect(sheetToObjects(newSheet())).toEqual([]);
  });
});

describe('sheetToNdjson', () => {
  it('yields one JSON line per row with raw=false semantics', async () => {
    const ws = newSheet();
    ws.addRow(['a', 'when']);
    ws.addRow([1, new Date(Date.UTC(2023, 0, 1))]);
    ws.addRow([2, null]);
    const lines = await collect(sheetToNdjson(ws, { defval: null }));
    expect(lines).toEqual(['{"a":1,"when":"2023-01-01T00:00:00.000Z"}\n', '{"a":2,"when":null}\n']);
    const parsed = lines.map((l) => JSON.parse(l) as unknown);
    expect(parsed).toEqual([
      { a: 1, when: '2023-01-01T00:00:00.000Z' },
      { a: 2, when: null },
    ]);
  });

  it('yields nothing for an empty sheet', async () => {
    expect(await collect(sheetToNdjson(newSheet()))).toEqual([]);
  });
});

describe('objectsToSheet', () => {
  it('uses the union of keys in first-seen order as header', () => {
    const ws = newSheet();
    objectsToSheet(ws, [
      { b: 1, a: 'x' },
      { c: true, a: 'y' },
    ]);
    expect(sheetToArrays(ws)).toEqual([
      ['b', 'a', 'c'],
      [1, 'x', null],
      [null, 'y', true],
    ]);
  });

  it('honours an explicit header and missing keys become empty cells', () => {
    const ws = newSheet();
    objectsToSheet(ws, [{ a: 1, b: 2 }, { a: 3 }], { header: ['b', 'a'] });
    expect(sheetToArrays(ws)).toEqual([
      ['b', 'a'],
      [2, 1],
      [null, 3],
    ]);
  });

  it('writeHeader=false writes only data rows', () => {
    const ws = newSheet();
    objectsToSheet(ws, [{ a: 1 }], { writeHeader: false });
    expect(sheetToArrays(ws)).toEqual([[1]]);
  });

  it('writes nested objects/arrays as JSON strings and keeps Dates', () => {
    const ws = newSheet();
    const d = new Date(Date.UTC(2019, 4, 5));
    objectsToSheet(ws, [{ o: { x: 1 }, arr: [1, 'two'], d, n: null, u: undefined }], {
      writeHeader: false,
    });
    // Trailing empty cells lie outside the used range, so the row is 3 wide.
    expect(sheetToArrays(ws)).toEqual([['{"x":1}', '[1,"two"]', d]]);
  });

  it('round-trips objectsToSheet -> sheetToObjects', () => {
    const ws = newSheet();
    const d = new Date(Date.UTC(2021, 11, 31, 12, 0, 0));
    const input = [
      { name: 'ann', age: 30, active: true, joined: d },
      { name: 'bob', age: 41.5, active: false, joined: null },
    ];
    objectsToSheet(ws, input, { header: ['name', 'age', 'active', 'joined'] });
    expect(sheetToObjects(ws)).toEqual([
      { name: 'ann', age: 30, active: true, joined: d },
      { name: 'bob', age: 41.5, active: false },
    ]);
    expect(sheetToObjects(ws, { defval: null })[1]).toEqual({
      name: 'bob',
      age: 41.5,
      active: false,
      joined: null,
    });
  });

  it('returns the same worksheet and handles empty input', () => {
    const ws = newSheet();
    expect(objectsToSheet(ws, [])).toBe(ws);
    expect(ws.rowCount).toBe(0);
  });

  it('accepts any iterable', () => {
    const ws = newSheet();
    function* gen(): Generator<Record<string, unknown>> {
      yield { k: 1 };
      yield { k: 2 };
    }
    objectsToSheet(ws, gen());
    expect(sheetToObjects(ws)).toEqual([{ k: 1 }, { k: 2 }]);
  });
});

describe('arraysToSheet', () => {
  it('appends arrays as rows and converts unknown values', () => {
    const ws = newSheet();
    const d = new Date(Date.UTC(2010, 0, 1));
    arraysToSheet(ws, [
      [1, 'a', true],
      [{ k: 'v' }, [1, 2], d, undefined],
    ]);
    expect(sheetToArrays(ws)).toEqual([
      [1, 'a', true],
      ['{"k":"v"}', '[1,2]', d],
    ]);
  });

  it('round-trips arraysToSheet -> sheetToArrays', () => {
    const ws = newSheet();
    const rows = [
      ['x', 1],
      ['y', 2],
    ];
    arraysToSheet(ws, rows);
    expect(sheetToArrays(ws)).toEqual(rows);
  });
});
