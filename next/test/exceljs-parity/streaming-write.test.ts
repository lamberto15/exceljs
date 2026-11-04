// Streaming write (createXlsxWriter): spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js
// and the value tests of spec/integration/worksheet-xlsx-writer.spec.js. Feature-specific
// streaming tests live in their area files (styles, notes, images, data validation, ...).

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx } from '../../src/index.js';
import { exceljsLoad, streamWrite } from './helpers.js';

describe('exceljs parity: streaming write', () => {
  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "creates sheets with correct names"
  it('creates sheets with correct names', async () => {
    await streamWrite((w) => {
      expect(w.addWorksheet('Hello, World!').name).toBe('Hello, World!');
      // ➖ default names are "SheetN" (exceljs: "sheetN")
      expect(w.addWorksheet().name).toMatch(/^Sheet\d+$/);
    });
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "auto filter"
  it('auto filter', async () => {
    const bytes = await streamWrite(async (w) => {
      const s = w.addWorksheet('Hello', { autoFilter: 'A1:B1' });
      await s.addRows([
        [1, 1],
        [2, 2],
        [3, 3],
      ]);
    });
    expect(readXlsx(bytes).getWorksheet('Hello')!.autoFilter).toBe('A1:B1');
    expect(readXlsx(bytes).getWorksheet('Hello')!.getValue('B3')).toBe(3);
    expect((await exceljsLoad(bytes)).getWorksheet('Hello')!.autoFilter).toBe('A1:B1');
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "A lot of sheets"
  it('a lot of sheets (90)', async () => {
    const bytes = await streamWrite((w) => {
      for (let i = 1; i <= 90; i++) w.addWorksheet(`sheet${i}`).addRow([i]);
    });
    const wb2 = readXlsx(bytes);
    for (let i = 1; i <= 90; i++) expect(wb2.getWorksheet(`sheet${i}`)!.getValue('A1')).toBe(i);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "addRow"
  it('addRow to a sink (no row.commit needed)', async () => {
    const chunks: Uint8Array[] = [];
    const { createXlsxWriter } = await import('../../src/index.js');
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c), strings: 'shared' });
    const s = w.addWorksheet('test');
    expect(s.addRow(['hello'])).toBe(1);
    await s.commit();
    await w.end();
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of chunks) {
      bytes.set(c, o);
      o += c.length;
    }
    expect(readXlsx(bytes).getWorksheet('test')!.getValue('A1')).toBe('hello');
  });

  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "stores values properly"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "stores shared string values properly"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "assigns cell types properly"
  it('stores values and kinds (checked in the written file: streamed rows cannot be read back)', async () => {
    const now = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
    const bytes = await streamWrite(
      (w) => {
        const s = w.addWorksheet('blort');
        s.addRow([
          7,
          'Hello, World!',
          3.14,
          now,
          `${['Hello', 'World'].join(', ')}!`,
          { text: 'www.google.com', hyperlink: 'http://www.google.com' },
        ]);
        s.addRow([
          { formula: 'A1', result: 7 },
          { formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!' },
          { formula: 'D1' },
        ]);
      },
      { strings: 'shared' },
    );
    const ws = readXlsx(bytes).getWorksheet('blort')!;
    expect(ws.getValue('A1')).toBe(7);
    expect(ws.getValue('B1')).toBe('Hello, World!');
    expect(ws.getValue('C1')).toBe(3.14);
    expect(ws.getValue('D1')).toEqual(now);
    expect(ws.getValue('E1')).toBe(ws.getValue('B1'));
    expect(ws.getValue('F1')).toEqual({
      text: 'www.google.com',
      hyperlink: 'http://www.google.com',
    });
    expect(ws.getValue('A2')).toEqual({ formula: 'A1', result: 7 });
    expect(ws.getValue('B2')).toEqual({
      formula: 'CONCATENATE("Hello", ", ", "World!")',
      result: 'Hello, World!',
    });
    expect(['A1', 'B1', 'C1', 'D1', 'F1', 'A2', 'B2', 'C2'].map((r) => ws.getKind(r))).toEqual([
      CellKind.Number,
      CellKind.String,
      CellKind.Number,
      CellKind.Date,
      CellKind.Hyperlink,
      CellKind.Formula,
      CellKind.Formula,
      CellKind.Formula,
    ]);
  });

  // exceljs: spec/unit/doc/worksheet-writer.spec.js — "generates valid xml even when there is no data"
  it('an empty streamed sheet is valid', async () => {
    const bytes = await streamWrite((w) => {
      w.addWorksheet('empty');
    });
    expect(readXlsx(bytes).getWorksheet('empty')!.rowCount).toBe(0);
    await exceljsLoad(bytes);
  });

  // exceljs: spec/unit/doc/workbook-writer.spec.js — "returns undefined for non-existant sheet"
  it('the writer workbook returns undefined for a missing sheet', async () => {
    await streamWrite((w) => {
      w.addWorksheet('first');
      expect(w.workbook.getWorksheet('w00t')).toBeUndefined();
    });
  });
});
