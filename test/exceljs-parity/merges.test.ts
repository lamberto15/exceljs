// Merged cells: spec/integration/worksheet.spec.js, worksheet-xlsx-writer.spec.js,
// workbook/workbook.spec.js "Merge Cells" and spec/unit/doc/worksheet.merge.spec.js.

import { describe, expect, it } from 'vitest';
import { CellKind, insertRow, Workbook, type Worksheet } from '../../src/index.js';
import { roundtrip, streamWrite } from './helpers.js';
import { alignment, font, fontBack, numFmt1, styles } from './test-book.js';

function overlapping(merge: (r: string) => void): void {
  merge('B2:C3');
  for (const r of ['A1:B2', 'C1:D2', 'C3:D4', 'A3:B4', 'A1:D4']) {
    expect(() => merge(r), r).toThrow(/overlaps/);
  }
}

describe('exceljs parity: merges', () => {
  // exceljs: spec/integration/worksheet.spec.js — "references the same top-left value"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "references the same top-left value"
  // exceljs: spec/unit/doc/worksheet.merge.spec.js — "references the same top-left value"
  it('merged range: the top-left cell holds the value', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getCell('A1').value = 'A1';
    ws.getCell('B1').value = 'B1';
    ws.getCell('A2').value = 'A2';
    ws.getCell('B2').value = 'B2';
    ws.mergeCells('A1:B2');
    expect(ws.getCell('A1').value).toBe('A1');
    expect(ws.getCell('A1').kind).toBe(CellKind.String);
    expect(ws.merges).toEqual(['A1:B2']);
    // ➖ no Merge value type / master proxy: other cells keep what they held (Excel shows
    // only the top-left one)
    expect(ws.getCell('B2').value).toBe('B2');
  });

  // exceljs: spec/integration/worksheet.spec.js — "does not allow overlapping merges" (×2)
  // exceljs: spec/unit/doc/worksheet.merge.spec.js — "does not allow overlapping merges" (×2)
  it('does not allow overlapping merges', () => {
    const ws = new Workbook().addWorksheet('blort');
    overlapping((r) => ws.mergeCells(r));
  });

  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "does not allow overlapping merges"
  it('streaming writer: does not allow overlapping merges', async () => {
    await streamWrite((w) => {
      const s = w.addWorksheet('blort');
      overlapping((r) => s.mergeCells(r));
    });
  });

  // exceljs: spec/integration/worksheet.spec.js — "merges and unmerges"
  // exceljs: spec/unit/doc/worksheet.merge.spec.js — "merges and unmerges"
  it('merges and unmerges', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.mergeCells('A1:B2');
    expect(ws.merges).toEqual(['A1:B2']);
    // ➖ unmergeCells (exceljs: unMergeCells)
    ws.unmergeCells('A1:B2');
    expect(ws.merges).toEqual([]);
    ws.mergeCells('A1:B2');
    ws.unmergeCells('A1'); // any cell of the merge
    expect(ws.merges).toEqual([]);
    ws.mergeCells('A1:B2');
    ws.unmergeCells('B2');
    expect(ws.merges).toEqual([]);
    for (const r of ['A1:B2', 'D1:E2', 'A4:B5', 'D4:E5']) ws.mergeCells(r);
    expect(ws.merges).toHaveLength(4);
    ws.unmergeCells('B2:D4');
    expect(ws.merges).toEqual([]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "merges styles"
  // exceljs: spec/unit/doc/worksheet.merge.spec.js — "merges styles"
  it('merging does not copy the top-left style into the range (➖)', () => {
    const ws = new Workbook().addWorksheet('blort');
    const style = {
      font: font(styles.fonts.broadwayRedOutline20),
      border: styles.borders.doubleRed,
      alignment: alignment(styles.namedAlignments.middleCentre),
      numFmt: numFmt1,
    };
    ws.setCell('B2', 5, style);
    ws.mergeCells('B2:C3');
    expect(ws.getStyle('B2').font).toEqual(fontBack(styles.fonts.broadwayRedOutline20));
    // ➖ exceljs copies the master style to B3, C2, C3; turbosheet leaves them alone. To
    // style the whole merge (e.g. borders), set the style on each cell.
    expect(ws.getStyle('C3')).toEqual({});
  });

  // exceljs: spec/unit/doc/worksheet.merge.spec.js — "preserves merges after row inserts"
  it('merges shift with insertRow', () => {
    const ws = new Workbook().addWorksheet('testMergeAfterInsert');
    ws.addRow([1, 2]);
    ws.addRow([3, 4]);
    ws.mergeCells('A1:B2');
    insertRow(ws, 1, ['Inserted Row Text']);
    // ➖ exceljs inspects its internal NumberValue / MergeValue cells (1 number + 3 merge
    // slaves of A2); turbosheet has no Merge value type: the range moves, cells keep their values
    expect(ws.merges).toEqual(['A2:B3']);
    expect(ws.getRow(1).values).toEqual([undefined, 'Inserted Row Text']);
    expect(ws.getRow(2).values).toEqual([undefined, 1, 2]);
    expect(ws.getRow(3).values).toEqual([undefined, 3, 4]);
    expect(ws.getKind('A2')).toBe(CellKind.Number);
    expect(roundtrip(ws.workbook).getWorksheet('testMergeAfterInsert')!.merges).toEqual(['A2:B3']);
  });

  describe('Workbook Merge Cells', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "serialises and deserialises properly"
    it('serialises and deserialises properly', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('blort');
      ws.getCell('B2').value = 'B2';
      ws.mergeCells('B2:C3');
      const ws2 = roundtrip(wb).getWorksheet('blort')!;
      expect(ws2.getValue('B2')).toBe('B2');
      expect(ws2.getKind('B2')).toBe(CellKind.String);
      expect(ws2.merges).toEqual(['B2:C3']);
      for (const r of ['B3', 'C2', 'C3']) expect(ws2.getValue(r)).toBeNull(); // ➖ no master proxy
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "styles"
    it('styles set on every cell of a merge round-trip', () => {
      const wb = new Workbook();
      const ws: Worksheet = wb.addWorksheet('blort');
      const style = {
        font: font(styles.fonts.broadwayRedOutline20),
        border: styles.borders.doubleRed,
        fill: styles.fills.redDarkVertical, // 🚫 exceljs uses a gradient fill here
        alignment: alignment(styles.namedAlignments.middleCentre),
        numFmt: numFmt1,
      };
      ws.setCell('B2', 5, style);
      for (const r of ['B3', 'C2', 'C3']) ws.setStyle(r, style); // ➖ explicit
      ws.mergeCells('B2:C3');
      const ws2 = roundtrip(wb).getWorksheet('blort')!;
      for (const r of ['B2', 'B3', 'C2', 'C3']) {
        const s = ws2.getStyle(r);
        expect(s.font, r).toEqual(fontBack(styles.fonts.broadwayRedOutline20));
        expect(s.border, r).toEqual(styles.borders.doubleRed);
        expect(s.fill, r).toEqual(styles.fills.redDarkVertical);
        expect(s.alignment, r).toEqual(alignment(styles.namedAlignments.middleCentre));
        expect(s.numFmt, r).toBe(numFmt1);
      }
    });
  });
});
