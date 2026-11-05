// Notes (legacy comments): spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js and
// spec/unit/doc/cell.spec.js. Whitespace / newline notes (pr 896) are in prs.test.ts.

import { describe, expect, it } from 'vitest';
import { addNote, readXlsx, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad, streamWrite } from './helpers.js';

const texts = [
  {
    font: { size: 12, color: { argb: 'FFFF6600' }, name: 'Calibri', scheme: 'minor' as const },
    text: 'seven',
  },
];

async function notesBook(): Promise<Uint8Array> {
  return streamWrite((w) => {
    const s = w.addWorksheet('Hello');
    s.addRow([]);
    s.addRow([null, 5, null, 7]);
    addNote(s.worksheet, 'B2', 'five');
    addNote(s.worksheet, 'D2', { richText: texts });
  });
}

describe('exceljs parity: notes', () => {
  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "writes notes"
  it('streaming writer: plain and rich-text notes', async () => {
    const bytes = await notesBook();
    const ws = readXlsx(bytes).getWorksheet('Hello')!;
    expect(ws.getValue('B2')).toBe(5);
    expect(ws.getValue('D2')).toBe(7);
    const note = (ref: string) => ws.comments?.find((n) => n.ref === ref);
    expect(note('B2')?.text).toBe('five');
    expect(note('D2')?.text).toEqual({ richText: texts });
    const x = (await exceljsLoad(bytes)).getWorksheet('Hello')!;
    expect(x.getCell('B2').note).toBe('five');
    expect((x.getCell('D2').note as { texts: unknown[] }).texts).toEqual(texts);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "Cell annotation supports setting margins and protection properties"
  it('streaming writer: note margins, protection and editAs', async () => {
    const note = {
      text: { richText: texts },
      margins: { insetmode: 'custom' as const, inset: [0.25, 0.25, 0.35, 0.35] },
      // ➖ booleans (exceljs: 'True' / 'False' strings)
      protection: { locked: false, lockText: false },
      editAs: 'oneCells' as const,
    };
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('Hello');
      s.addRow([]);
      s.addRow([null, 5, null, 7]);
      addNote(s.worksheet, 'B2', 'five');
      addNote(s.worksheet, 'D2', note);
    });
    const ws = readXlsx(bytes).getWorksheet('Hello')!;
    expect(ws.getValue('B2')).toBe(5);
    expect(ws.comments?.find((n) => n.ref === 'B2')?.text).toBe('five');
    expect(ws.getValue('D2')).toBe(7);
    expect(ws.comments?.find((n) => n.ref === 'D2')).toEqual({ ref: 'D2', ...note });
    const d2 = (await exceljsLoad(bytes)).getWorksheet('Hello')!.getCell('D2').note as unknown as {
      texts: unknown;
      margins: unknown;
      protection: unknown;
      editAs: unknown;
    };
    expect(d2.texts).toEqual(texts);
    expect(d2.margins).toEqual(note.margins);
    expect(d2.protection).toEqual({ locked: 'False', lockText: 'False' });
    expect(d2.editAs).toBe('oneCells');
  });

  // exceljs: spec/unit/doc/cell.spec.js — "can set comment"
  it('a note keeps margins, protection and editAs', () => {
    const ws = new Workbook().addWorksheet('S');
    const note = {
      text: {
        richText: [{ font: { size: 12, color: { theme: 0 }, name: 'Calibri' }, text: 'This is ' }],
      },
      margins: { insetmode: 'auto' as const, inset: [0.13, 0.13, 0.25, 0.25] },
      protection: { locked: true, lockText: true }, // ➖ booleans
      editAs: 'twoCells' as const,
    };
    addNote(ws, 'A1', note);
    ws.getCell('A1').value = 'test set value';
    expect(ws.comments![0]).toEqual({ ref: 'A1', ...note });
    expect(readXlsx(writeXlsx(ws.workbook)).worksheets[0]!.comments![0]).toEqual({
      ref: 'A1',
      ...note,
    });
  });

  // exceljs: spec/unit/doc/cell.spec.js — "Cell comments supports setting margins, protection, and position properties"
  it('unset margins / editAs are the defaults (auto, absolute)', async () => {
    const ws = new Workbook().addWorksheet('S');
    ws.getCell('A1').value = 'test set value';
    addNote(ws, 'A1', {
      text: { richText: [{ font: { size: 12, name: 'Calibri' }, text: 'This is ' }] },
      protection: { locked: false, lockText: true },
    });
    const bytes = writeXlsx(ws.workbook);
    // ➖ defaults are not materialised in the model: undefined = auto margins, absolute
    expect(readXlsx(bytes).worksheets[0]!.comments![0]).toMatchObject({
      protection: { locked: false, lockText: true },
    });
    const n = (await exceljsLoad(bytes)).worksheets[0]!.getCell('A1').note as unknown as {
      margins: { insetmode: string };
      editAs: string;
      protection: unknown;
    };
    expect(n.margins.insetmode).toBe('auto');
    expect(n.editAs).toBe('absolute');
    expect(n.protection).toEqual({ locked: 'False', lockText: 'True' });
  });

  it('note author, visibility and size round-trip', async () => {
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('S');
      s.addRow(['x']);
      addNote(s.worksheet, 'A1', {
        text: 'hi',
        author: 'Ann',
        visible: true,
        width: 200,
        height: 60,
      });
    });
    const n = readXlsx(bytes).getWorksheet('S')!.comments![0]!;
    expect(n).toMatchObject({ ref: 'A1', text: 'hi', author: 'Ann', visible: true });
  });
});
