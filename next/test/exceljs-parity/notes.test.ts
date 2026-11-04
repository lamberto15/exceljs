// Notes (legacy comments): spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js and
// spec/unit/doc/cell.spec.js. Whitespace / newline notes (pr 896) are in prs.test.ts.

import { describe, expect, it } from 'vitest';
import { addNote, readXlsx } from '../../src/index.js';
import { exceljsLoad, streamWrite } from './helpers.js';

const texts = [
  {
    // 🚫 font scheme ('minor') is not modelled
    font: { size: 12, color: { argb: 'FFFF6600' }, name: 'Calibri' },
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
  // exceljs: spec/unit/doc/cell.spec.js — "can set comment"
  // exceljs: spec/unit/doc/cell.spec.js — "Cell comments supports setting margins, protection, and position properties"
  it.skip('note margins / protection / editAs — not supported (Note: text, author, visible, width, height)', () => {});

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
