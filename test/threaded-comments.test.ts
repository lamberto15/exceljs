// Threaded comments (#111): authoring, round trip, the legacy notes Excel expects, persons.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addNote,
  addThreadedComment,
  createXlsxWriter,
  readXlsx,
  Workbook,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const text = (z: Map<string, Uint8Array>, p: string) => dec.decode(z.get(p)!);

function book(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['a', 'b']);
  addThreadedComment(ws, 'b1', {
    text: 'Is this right?',
    author: 'Ann',
    date: new Date('2024-05-01T10:30:00Z'),
    replies: [
      { text: 'Yes & <checked>', author: 'Bob', date: new Date('2024-05-02T08:00:00Z') },
      { text: 'Thanks', author: 'Ann' },
    ],
  });
  addThreadedComment(ws, 'A1', { text: 'Done here', author: 'Bob', done: true });
  addNote(ws, 'C3', 'a plain note');
  return wb;
}

describe('threaded comments', () => {
  it('writes the thread part, persons, and a legacy note per thread', async () => {
    const out = writeXlsx(book());
    const z = unzipSync(out);
    const t = text(z, 'xl/threadedComments/threadedComment1.xml');
    expect(t.match(/<threadedComment /g)).toHaveLength(4);
    expect(t).toMatch(
      /<threadedComment ref="B1" dT="2024-05-01T10:30:00.00" personId="\{[0-9A-F-]{36}\}" id="(\{[0-9A-F-]{36}\})"><text>Is this right\?<\/text>/,
    );
    const root = /ref="B1"[^>]* id="(\{[^"]+\})"/.exec(t)![1]!;
    expect(t).toContain(`parentId="${root}"><text>Yes &amp; &lt;checked&gt;</text>`);
    expect(t).toMatch(/ref="A1"[^>]* done="1"/);
    const persons = text(z, 'xl/persons/person.xml');
    expect(persons.match(/<person /g)).toHaveLength(2);
    expect(persons).toContain('displayName="Ann"');
    const comments = text(z, 'xl/comments1.xml');
    expect(comments).toContain(`<author>tc=${root}</author>`);
    expect(comments).toContain(
      'Comment:\n    Is this right?\nReply:\n    Yes &amp; &lt;checked&gt;\nReply:\n    Thanks',
    );
    expect(comments).toContain('a plain note');
    const ct = text(z, '[Content_Types].xml');
    expect(ct).toContain('application/vnd.ms-excel.threadedcomments+xml');
    expect(ct).toContain('application/vnd.ms-excel.person+xml');
    // exceljs sees the legacy notes
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(out as unknown as ArrayBuffer);
    // exceljs sees the legacy note of the thread (it only attaches notes to existing rows)
    expect(String(x.worksheets[0]!.getCell('B1').note)).toContain('Comment:\n    Is this right?');
  });

  it('round-trips: threads, replies, dates, done, persons; output is stable', () => {
    const wb = book();
    const out = writeXlsx(wb);
    const back = readXlsx(out);
    const ws = back.worksheets[0]!;
    expect(ws.comments).toEqual([{ ref: 'C3', text: 'a plain note' }]);
    expect(
      ws.threadedComments!.map(({ id: _, replies, ...t }) => ({
        ...t,
        replies: replies?.map(({ id: __, ...r }) => r),
      })),
    ).toEqual([
      {
        ref: 'B1',
        text: 'Is this right?',
        author: 'Ann',
        date: new Date('2024-05-01T10:30:00Z'),
        replies: [
          { text: 'Yes & <checked>', author: 'Bob', date: new Date('2024-05-02T08:00:00Z') },
          { text: 'Thanks', author: 'Ann' },
        ],
      },
      { ref: 'A1', text: 'Done here', author: 'Bob', done: true, replies: undefined },
    ]);
    expect(back.persons!.map((p) => p.displayName)).toEqual(['Ann', 'Bob']);
    const again = unzipSync(writeXlsx(back));
    const first = unzipSync(out);
    for (const p of [
      'xl/threadedComments/threadedComment1.xml',
      'xl/persons/person.xml',
      'xl/comments1.xml',
    ]) {
      expect(text(again, p), p).toBe(text(first, p));
    }
  });

  it('kept persons keep their account ids; new authors are added', () => {
    const wb = book();
    wb.persons = [
      {
        displayName: 'Ann',
        id: '{11111111-1111-4111-8111-111111111111}',
        userId: 'ann@x.com',
        providerId: 'AD',
      },
    ];
    const z = unzipSync(writeXlsx(wb));
    const persons = text(z, 'xl/persons/person.xml');
    expect(persons).toContain(
      'displayName="Ann" id="{11111111-1111-4111-8111-111111111111}" userId="ann@x.com" providerId="AD"',
    );
    expect(persons).toContain('displayName="Bob"');
    expect(text(z, 'xl/threadedComments/threadedComment1.xml')).toContain(
      'personId="{11111111-1111-4111-8111-111111111111}"',
    );
  });

  it('a note and a thread on the same cell is an error', () => {
    const wb = book();
    addNote(wb.worksheets[0]!, 'B1', 'clash');
    expect(() => writeXlsx(wb)).toThrow('has both a note and a threaded comment');
  });

  it('streaming writer writes threads too', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    s.addRow(['x']);
    addThreadedComment(s.worksheet, 'A1', { text: 'hi', author: 'Ann' });
    await w.end();
    const buf = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of chunks) {
      buf.set(c, o);
      o += c.length;
    }
    expect(readXlsx(buf).worksheets[0]!.threadedComments).toMatchObject([
      { ref: 'A1', text: 'hi', author: 'Ann' },
    ]);
  });
});
