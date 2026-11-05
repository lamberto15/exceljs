// Rows, columns and sheet names: spec/integration/worksheet.spec.js, worksheet-xlsx-writer.spec.js,
// workbook.spec.js (duplicate rows, splice) and the behaviour-level unit specs
// (spec/unit/doc/worksheet.values.spec.js, worksheet.views.spec.js, worksheet.pagebreak.spec.js).

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  CellKind,
  type CellValue,
  duplicateRow,
  getDefinedName,
  getRows,
  insertRow,
  insertRows,
  readXlsx,
  type Style,
  spliceColumns,
  spliceRows,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../../src/index.js';
import { sheetToArrays } from '../../src/json/index.js';
import { unzipSync } from '../../src/zip/index.js';
import { fixture, roundtrip, streamWrite } from './helpers.js';
import { alignment, alignmentBack, font, fontBack, numFmt1, numFmt2, styles } from './test-book.js';

const d1 = new Date(Date.UTC(1970, 1, 1));
const d2 = new Date(Date.UTC(1965, 1, 7));
const d3 = new Date(Date.UTC(1965, 1, 10));

// ---- spec/utils/test-spliced-sheet.js: the sheets the splice specs build and check ----------
// Each scenario has `add` (build + splice) and `check`, like createTestBook / checkTestBook.
const U = undefined;
const vals = (ws: Worksheet, r: number) => ws.getRow(r).values;
const kinds = (ws: Worksheet, addrs: Record<string, CellKind>) => {
  for (const [a, k] of Object.entries(addrs)) expect(ws.getKind(a), a).toBe(k);
};
const sheetOf = (wb: Workbook, name: string): Worksheet => {
  const ws = wb.getWorksheet(name);
  expect(ws).toBeDefined();
  return ws!;
};
const nameRef = (wb: Workbook, name: string) => getDefinedName(wb, name)?.value;
const q = (sheet: string, ref: string) => `'${sheet}'!${ref}`;
const fillRed: Style['fill'] = {
  type: 'pattern',
  pattern: 'darkVertical',
  fgColor: { argb: 'FFFF0000' },
};
const thin = { style: 'thin' } as const;
const box = { top: thin, left: thin, bottom: thin, right: thin };
// the alignment as turbosheet stores it ('middle' -> 'center', the default 'bottom' dropped)
const align = (horizontal: string, vertical: string) =>
  alignmentBack({ horizontal, vertical }) as NonNullable<Style['alignment']>;
/** ExcelJS `row.<prop> = x`: the row default and every existing cell of the row. */
function styleRow(ws: Worksheet, r: number, cols: number, patch: Style): void {
  ws.setRowProps(r, { style: { ...ws.getRowProps(r).style, ...patch } });
  for (let c = 1; c <= cols; c++) ws.setStyle(r, c, { ...ws.getStyle(r, c), ...patch });
}
/** ExcelJS `column.<prop> = x`: the column default and every existing cell of the column. */
function styleCol(ws: Worksheet, c: number, rows: number, patch: Style): void {
  ws.getColumn(c).style = { ...ws.getColumn(c).style, ...patch };
  for (let r = 1; r <= rows; r++) ws.setStyle(r, c, { ...ws.getStyle(r, c), ...patch });
}
/** Fill an A1:D4 grid of "r,c" strings. */
function grid4(ws: Worksheet): void {
  for (let r = 1; r <= 4; r++) ws.addRow([1, 2, 3, 4].map((c) => `${r},${c}`));
}
const grid5 = (ws: Worksheet) => {
  for (let r = 1; r <= 2; r++) ws.addRow([1, 2, 3, 4, 5].map((c) => `${r},${c}`));
  ws.setCell('A4', 4.1);
  ws.setCell('C4', 4.3);
  ws.setCell('E4', 4.5);
  ws.addRow([1, 2, 3, 4, 5].map((c) => `5,${c}`));
};

interface Scenario {
  add(wb: Workbook): void;
  check(wb: Workbook): void;
}

const spliced: { rows: Record<string, Scenario>; columns: Record<string, Scenario> } = {
  rows: {
    removeOnly: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-remove-only');
        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.setCell('A4', 4.1);
        ws.setCell('C4', 4.3);
        ws.addRow(['5,1', '5,2', '5,3']);
        spliceRows(ws, 2, 1);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-remove-only');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3']);
        kinds(ws, { A1: CellKind.String, B1: CellKind.String, C1: CellKind.String });
        kinds(ws, { A2: CellKind.Empty, B2: CellKind.Empty, C2: CellKind.Empty });
        expect(ws.getValue('A3')).toBe(4.1);
        expect(ws.getValue('C3')).toBe(4.3);
        kinds(ws, { A3: CellKind.Number, B3: CellKind.Empty, C3: CellKind.Number });
        expect(vals(ws, 4)).toEqual([U, '5,1', '5,2', '5,3']);
        ws.addRow(['5,1b', '5,2b', '5,3b']);
        expect(vals(ws, 5)).toEqual([U, '5,1b', '5,2b', '5,3b']);
      },
    },
    insertFewer: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-insert-fewer');
        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.setCell('A4', 4.1);
        ws.setCell('C4', 4.3);
        ws.addRow(['5,1', '5,2', '5,3']);
        spliceRows(ws, 2, 2, ['one', 'two', 'three']);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-insert-fewer');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3']);
        expect(vals(ws, 2)).toEqual([U, 'one', 'two', 'three']);
        expect(vals(ws, 3)).toEqual([U, 4.1, U, 4.3]);
        expect(vals(ws, 4)).toEqual([U, '5,1', '5,2', '5,3']);
      },
    },
    insertSame: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-insert-same');
        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.setCell('A4', 4.1);
        ws.setCell('C4', 4.3);
        ws.addRow(['5,1', '5,2', '5,3']);
        spliceRows(ws, 2, 2, ['one', 'two', 'three'], ['une', 'deux', 'trois']);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-insert-same');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3']);
        expect(vals(ws, 2)).toEqual([U, 'one', 'two', 'three']);
        expect(vals(ws, 3)).toEqual([U, 'une', 'deux', 'trois']);
        expect(vals(ws, 4)).toEqual([U, 4.1, U, 4.3]);
        expect(vals(ws, 5)).toEqual([U, '5,1', '5,2', '5,3']);
      },
    },
    insertMore: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-insert-more');
        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.setCell('A4', 4.1);
        ws.setCell('C4', 4.3);
        ws.addRow(['5,1', '5,2', '5,3']);
        spliceRows(
          ws,
          2,
          2,
          ['one', 'two', 'three'],
          ['une', 'deux', 'trois'],
          ['uno', 'due', 'tre'],
        );
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-insert-more');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3']);
        expect(vals(ws, 2)).toEqual([U, 'one', 'two', 'three']);
        expect(vals(ws, 3)).toEqual([U, 'une', 'deux', 'trois']);
        expect(vals(ws, 4)).toEqual([U, 'uno', 'due', 'tre']);
        expect(vals(ws, 5)).toEqual([U, 4.1, U, 4.3]);
        expect(vals(ws, 6)).toEqual([U, '5,1', '5,2', '5,3']);
      },
    },
    removeStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-remove-style');
        grid4(ws);
        ws.setStyle('A1', { numFmt: '# ?/?' });
        ws.setStyle('B2', { fill: fillRed });
        styleRow(ws, 3, 4, { border: box });
        styleRow(ws, 4, 4, { alignment: align('left', 'middle') });
        spliceRows(ws, 2, 2);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-remove-style');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3', '1,4']);
        expect(vals(ws, 2)).toEqual([U, '4,1', '4,2', '4,3', '4,4']);
        expect(ws.getStyle('A1')).toEqual({ numFmt: '# ?/?' });
        expect(ws.getRowProps(2).style).toEqual({ alignment: align('left', 'middle') });
        expect(ws.getStyle('A2').alignment).toEqual(align('left', 'middle'));
      },
    },
    insertStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-insert-style');
        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.setStyle('A2', { fill: fillRed });
        styleRow(ws, 2, 3, { alignment: align('left', 'middle') });
        spliceRows(ws, 2, 0, ['one', 'two', 'three']);
        ws.setStyle('A2', { ...ws.getStyle('A2'), border: box });
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-insert-style');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3']);
        expect(vals(ws, 2)).toEqual([U, 'one', 'two', 'three']);
        expect(vals(ws, 3)).toEqual([U, '2,1', '2,2', '2,3']);
        expect(ws.getRowProps(3).style?.alignment).toEqual(align('left', 'middle'));
        expect(ws.getStyle('A2').border).toEqual(box);
        expect(ws.getStyle('A3').alignment).toEqual(align('left', 'middle'));
        expect(ws.getStyle('A3').fill).toEqual(fillRed);
      },
    },
    replaceStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-row-replace-style');
        for (let r = 1; r <= 3; r++) ws.addRow([1, 2, 3, 4].map((c) => `${r},${c}`));
        ws.setStyle('B1', { numFmt: 'top' });
        ws.setStyle('B2', { numFmt: 'middle' });
        ws.setStyle('B3', { numFmt: 'bottom' });
        styleRow(ws, 1, 4, { alignment: align('left', 'top') });
        styleRow(ws, 2, 4, { alignment: align('center', 'middle') });
        styleRow(ws, 3, 4, { alignment: align('right', 'bottom') });
        spliceRows(ws, 2, 1, ['two-one', 'two-two', 'two-three', 'two-four']);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-row-replace-style');
        expect(vals(ws, 1)).toEqual([U, '1,1', '1,2', '1,3', '1,4']);
        expect(vals(ws, 2)).toEqual([U, 'two-one', 'two-two', 'two-three', 'two-four']);
        expect(vals(ws, 3)).toEqual([U, '3,1', '3,2', '3,3', '3,4']);
        expect(ws.getStyle('B1')).toEqual({ numFmt: 'top', alignment: align('left', 'top') });
        expect(ws.getStyle('B2')).toEqual({});
        expect(ws.getStyle('B3')).toEqual({
          numFmt: 'bottom',
          alignment: align('right', 'bottom'),
        });
        expect(ws.getRowProps(1).style).toEqual({ alignment: align('left', 'top') });
        expect(ws.getRowProps(2).style).toBeUndefined();
        expect(ws.getRowProps(3).style).toEqual({ alignment: align('right', 'bottom') });
      },
    },
    removeDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-row-remove-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceRows(sq, 2, 2);

        const si = wb.addWorksheet('splice-row-remove-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceRows(si, 2, 2);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-row-remove-name-square');
        expect(vals(sq, 1)).toEqual([U, '1,1', '1,2', '1,3', '1,4']);
        expect(vals(sq, 2)).toEqual([U, '4,1', '4,2', '4,3', '4,4']);
        // rows 1-2 stay named, nothing below
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$D$2'));

        const si = sheetOf(wb, 'splice-row-remove-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', U, U, '1,4']);
        expect(vals(si, 2)).toEqual([U, '4,1', U, U, '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$2'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$D$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$D$2'));
      },
    },
    insertDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-row-insert-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceRows(sq, 3, 0, ['foo', 'bar', 'baz', 'qux']);

        const si = wb.addWorksheet('splice-row-insert-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceRows(si, 3, 0, ['foo', 'bar', 'baz', 'qux']);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-row-insert-name-square');
        expect(vals(sq, 1)).toEqual([U, '1,1', '1,2', '1,3', '1,4']);
        expect(vals(sq, 2)).toEqual([U, '2,1', '2,2', '2,3', '2,4']);
        expect(vals(sq, 3)).toEqual([U, 'foo', 'bar', 'baz', 'qux']);
        expect(vals(sq, 4)).toEqual([U, '3,1', '3,2', '3,3', '3,4']);
        expect(vals(sq, 5)).toEqual([U, '4,1', '4,2', '4,3', '4,4']);
        // ➖ exceljs leaves the inserted row 3 outside the name (a hole in its cell list);
        // like Excel, a range spanning the insertion point grows
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$D$5'));

        const si = sheetOf(wb, 'splice-row-insert-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', U, U, '1,4']);
        expect(vals(si, 3)).toEqual([U, 'foo', 'bar', 'baz', 'qux']);
        expect(vals(si, 5)).toEqual([U, '4,1', U, U, '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$5'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$D$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$D$5'));
      },
    },
    replaceDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-row-replace-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceRows(sq, 2, 1, ['foo', 'bar', 'baz', 'qux']);

        const si = wb.addWorksheet('splice-row-replace-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceRows(si, 2, 1, ['foo', 'bar', 'baz', 'qux']);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-row-replace-name-square');
        expect(vals(sq, 1)).toEqual([U, '1,1', '1,2', '1,3', '1,4']);
        expect(vals(sq, 2)).toEqual([U, 'foo', 'bar', 'baz', 'qux']);
        expect(vals(sq, 3)).toEqual([U, '3,1', '3,2', '3,3', '3,4']);
        expect(vals(sq, 4)).toEqual([U, '4,1', '4,2', '4,3', '4,4']);
        // ➖ exceljs drops the replaced row 2 from the name; the row is overwritten in place,
        // so the reference to it stays
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$D$4'));

        const si = sheetOf(wb, 'splice-row-replace-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', U, U, '1,4']);
        expect(vals(si, 2)).toEqual([U, 'foo', 'bar', 'baz', 'qux']);
        expect(vals(si, 4)).toEqual([U, '4,1', U, U, '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$4'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$D$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$D$4'));
      },
    },
  },
  columns: {
    removeOnly: {
      add(wb) {
        const ws = wb.addWorksheet('splice-column-remove-only');
        ws.columns = [
          { key: 'id', width: 10 },
          { key: 'name', width: 32 },
          { key: 'dob', width: 10 },
        ];
        ws.addRow({ id: 'id1', name: 'name1', dob: 'dob1' });
        ws.addRow({ id: 2, dob: 'dob2' });
        ws.addRow({ name: 'name3', dob: 3 });
        spliceColumns(ws, 2, 1);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-column-remove-only');
        expect(ws.getValue('A1')).toBe('id1');
        expect(ws.getValue('B1')).toBe('dob1');
        kinds(ws, { A1: CellKind.String, B1: CellKind.String, C1: CellKind.Empty });
        expect(ws.getValue('A2')).toBe(2);
        expect(ws.getValue('B2')).toBe('dob2');
        kinds(ws, { A2: CellKind.Number, B2: CellKind.String, C2: CellKind.Empty });
        expect(ws.getValue('B3')).toBe(3);
        kinds(ws, { A3: CellKind.Empty, B3: CellKind.Number, C3: CellKind.Empty });
      },
    },
    insertFewer: {
      add(wb) {
        const ws = wb.addWorksheet('splice-column-insert-fewer');
        grid5(ws);
        spliceColumns(ws, 2, 2, ['one', 'two', 'three', 'four', 'five']);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-column-insert-fewer');
        expect(vals(ws, 1)).toEqual([U, '1,1', 'one', '1,4', '1,5']);
        expect(vals(ws, 2)).toEqual([U, '2,1', 'two', '2,4', '2,5']);
        expect(vals(ws, 3)).toEqual([U, U, 'three']);
        expect(vals(ws, 4)).toEqual([U, 4.1, 'four', U, 4.5]);
        expect(vals(ws, 5)).toEqual([U, '5,1', 'five', '5,4', '5,5']);
      },
    },
    insertSame: {
      add(wb) {
        const ws = wb.addWorksheet('splice-column-insert-same');
        grid5(ws);
        spliceColumns(
          ws,
          2,
          2,
          ['one', 'two', 'three', 'four', 'five'],
          ['une', 'deux', 'trois', 'quatre', 'cinq'],
        );
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-column-insert-same');
        expect(vals(ws, 1)).toEqual([U, '1,1', 'one', 'une', '1,4', '1,5']);
        expect(vals(ws, 2)).toEqual([U, '2,1', 'two', 'deux', '2,4', '2,5']);
        expect(vals(ws, 3)).toEqual([U, U, 'three', 'trois']);
        expect(vals(ws, 4)).toEqual([U, 4.1, 'four', 'quatre', U, 4.5]);
        expect(vals(ws, 5)).toEqual([U, '5,1', 'five', 'cinq', '5,4', '5,5']);
      },
    },
    insertMore: {
      add(wb) {
        const ws = wb.addWorksheet('splice-column-insert-more');
        grid5(ws);
        spliceColumns(
          ws,
          2,
          2,
          ['one', 'two', 'three', 'four', 'five'],
          ['une', 'deux', 'trois', 'quatre', 'cinq'],
          ['uno', 'due', 'tre', 'quatro', 'cinque'],
        );
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-column-insert-more');
        expect(vals(ws, 1)).toEqual([U, '1,1', 'one', 'une', 'uno', '1,4', '1,5']);
        expect(vals(ws, 2)).toEqual([U, '2,1', 'two', 'deux', 'due', '2,4', '2,5']);
        expect(vals(ws, 3)).toEqual([U, U, 'three', 'trois', 'tre']);
        expect(vals(ws, 4)).toEqual([U, 4.1, 'four', 'quatre', 'quatro', U, 4.5]);
        expect(vals(ws, 5)).toEqual([U, '5,1', 'five', 'cinq', 'cinque', '5,4', '5,5']);
      },
    },
    removeStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-col-remove-style');
        grid4(ws);
        ws.setStyle('A1', { numFmt: '# ?/?' });
        ws.setStyle('B2', { fill: fillRed });
        styleCol(ws, 3, 4, { border: box });
        styleCol(ws, 4, 4, { alignment: align('left', 'middle') });
        spliceColumns(ws, 2, 2);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-col-remove-style');
        for (let r = 1; r <= 4; r++) expect(vals(ws, r)).toEqual([U, `${r},1`, `${r},4`]);
        expect(ws.getStyle('A1')).toEqual({ numFmt: '# ?/?' });
        expect(ws.getColumn(2).style).toEqual({ alignment: align('left', 'middle') });
        expect(ws.getStyle('B4')).toEqual({ alignment: align('left', 'middle') });
      },
    },
    insertStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-col-insert-style');
        for (let r = 1; r <= 3; r++) ws.addRow([1, 2, 3].map((c) => `${r},${c}`));
        ws.setStyle('B2', { fill: fillRed });
        styleCol(ws, 2, 3, { alignment: align('left', 'middle') });
        spliceColumns(ws, 2, 0, ['one', 'two', 'three']);
        ws.setStyle('B2', { border: box });
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-col-insert-style');
        expect(vals(ws, 1)).toEqual([U, '1,1', 'one', '1,2', '1,3']);
        expect(vals(ws, 2)).toEqual([U, '2,1', 'two', '2,2', '2,3']);
        expect(vals(ws, 3)).toEqual([U, '3,1', 'three', '3,2', '3,3']);
        expect(ws.getColumn(3).style).toEqual({ alignment: align('left', 'middle') });
        expect(ws.getStyle('B2')).toEqual({ border: box });
        expect(ws.getStyle('C2')).toEqual({ alignment: align('left', 'middle'), fill: fillRed });
      },
    },
    replaceStyle: {
      add(wb) {
        const ws = wb.addWorksheet('splice-col-replace-style');
        for (let r = 1; r <= 3; r++) ws.addRow([1, 2, 3, 4].map((c) => `${r},${c}`));
        ws.setStyle('A2', { numFmt: 'left' });
        ws.setStyle('B2', { numFmt: 'center' });
        ws.setStyle('C2', { numFmt: 'right' });
        styleCol(ws, 1, 3, { alignment: align('left', 'top') });
        styleCol(ws, 2, 3, { alignment: align('center', 'middle') });
        styleCol(ws, 3, 3, { alignment: align('right', 'bottom') });
        spliceColumns(ws, 2, 1, ['one-two', 'two-two', 'three-two']);
      },
      check(wb) {
        const ws = sheetOf(wb, 'splice-col-replace-style');
        expect(vals(ws, 1)).toEqual([U, '1,1', 'one-two', '1,3', '1,4']);
        expect(vals(ws, 2)).toEqual([U, '2,1', 'two-two', '2,3', '2,4']);
        expect(vals(ws, 3)).toEqual([U, '3,1', 'three-two', '3,3', '3,4']);
        expect(ws.getStyle('A2')).toEqual({ numFmt: 'left', alignment: align('left', 'top') });
        expect(ws.getStyle('B2')).toEqual({});
        expect(ws.getStyle('C2')).toEqual({
          numFmt: 'right',
          alignment: align('right', 'bottom'),
        });
        expect(ws.getColumn(1).style).toEqual({ alignment: align('left', 'top') });
        expect(ws.getColumn(2).style).toEqual({});
        expect(ws.getColumn(3).style).toEqual({ alignment: align('right', 'bottom') });
      },
    },
    removeDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-col-remove-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceColumns(sq, 2, 2);

        const si = wb.addWorksheet('splice-col-remove-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceColumns(si, 2, 2);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-col-remove-name-square');
        for (let r = 1; r <= 4; r++) expect(vals(sq, r)).toEqual([U, `${r},1`, `${r},4`]);
        // columns A-B stay named, nothing right of them
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$B$4'));

        const si = sheetOf(wb, 'splice-col-remove-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', '1,4']);
        expect(vals(si, 4)).toEqual([U, '4,1', '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$4'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$B$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$B$4'));
      },
    },
    insertDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-col-insert-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceColumns(sq, 3, 0, ['foo', 'bar', 'baz', 'qux']);

        const si = wb.addWorksheet('splice-col-insert-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceColumns(si, 3, 0, ['foo', 'bar', 'baz', 'qux']);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-col-insert-name-square');
        expect(vals(sq, 1)).toEqual([U, '1,1', '1,2', 'foo', '1,3', '1,4']);
        expect(vals(sq, 2)).toEqual([U, '2,1', '2,2', 'bar', '2,3', '2,4']);
        expect(vals(sq, 3)).toEqual([U, '3,1', '3,2', 'baz', '3,3', '3,4']);
        expect(vals(sq, 4)).toEqual([U, '4,1', '4,2', 'qux', '4,3', '4,4']);
        // ➖ exceljs leaves the inserted column C outside the name; like Excel, a range
        // spanning the insertion point grows
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$E$4'));

        const si = sheetOf(wb, 'splice-col-insert-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', U, 'foo', U, '1,4']);
        expect(vals(si, 4)).toEqual([U, '4,1', U, 'qux', U, '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$4'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$E$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$E$4'));
      },
    },
    replaceDefinedNames: {
      add(wb) {
        const sq = wb.addWorksheet('splice-col-replace-name-square');
        grid4(sq);
        addDefinedName(wb, 'square', q(sq.name, 'A1:D4'));
        spliceColumns(sq, 2, 1, ['foo', 'bar', 'baz', 'qux']);

        const si = wb.addWorksheet('splice-col-replace-name-singles');
        si.setCell('A1', '1,1');
        si.setCell('A4', '4,1');
        si.setCell('D1', '1,4');
        si.setCell('D4', '4,4');
        for (const a of ['A1', 'A4', 'D1', 'D4']) addDefinedName(wb, `single${a}`, q(si.name, a));
        spliceColumns(si, 2, 1, ['foo', 'bar', 'baz', 'qux']);
      },
      check(wb) {
        const sq = sheetOf(wb, 'splice-col-replace-name-square');
        expect(vals(sq, 1)).toEqual([U, '1,1', 'foo', '1,3', '1,4']);
        expect(vals(sq, 2)).toEqual([U, '2,1', 'bar', '2,3', '2,4']);
        expect(vals(sq, 3)).toEqual([U, '3,1', 'baz', '3,3', '3,4']);
        expect(vals(sq, 4)).toEqual([U, '4,1', 'qux', '4,3', '4,4']);
        // ➖ exceljs drops the replaced column B from the name; the column is overwritten in
        // place, so the reference to it stays
        expect(nameRef(wb, 'square')).toBe(q(sq.name, '$A$1:$D$4'));

        const si = sheetOf(wb, 'splice-col-replace-name-singles');
        expect(vals(si, 1)).toEqual([U, '1,1', 'foo', U, '1,4']);
        expect(vals(si, 4)).toEqual([U, '4,1', 'qux', U, '4,4']);
        expect(nameRef(wb, 'singleA1')).toBe(q(si.name, '$A$1'));
        expect(nameRef(wb, 'singleA4')).toBe(q(si.name, '$A$4'));
        expect(nameRef(wb, 'singleD1')).toBe(q(si.name, '$D$1'));
        expect(nameRef(wb, 'singleD4')).toBe(q(si.name, '$D$4'));
      },
    },
  },
};

/** A scenario built and checked in memory, like the worksheet.values.spec "Splice" tests. */
function runSplice(s: Scenario): void {
  const wb = new Workbook();
  s.add(wb);
  s.check(wb);
}

describe('exceljs parity: rows and columns', () => {
  // exceljs: spec/integration/worksheet.spec.js — "adds columns"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds columns"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds columns"
  it('adds columns (by key, letter and number: same Column object)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { key: 'id', width: 10 },
      { key: 'name', width: 32 },
      { key: 'dob', width: 10 },
    ];
    for (const [key, n, letter, width] of [
      ['id', 1, 'A', 10],
      ['name', 2, 'B', 32],
      ['dob', 3, 'C', 10],
    ] as const) {
      expect(ws.getColumn(key).number).toBe(n);
      expect(ws.getColumn(key).width).toBe(width);
      expect(ws.getColumn(letter)).toBe(ws.getColumn(key));
      expect(ws.getColumn(n)).toBe(ws.getColumn(key));
    }
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds column headers"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers"
  it('adds column headers', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { header: 'Id', width: 10 },
      { header: 'Name', width: 32 },
      { header: 'D.O.B.', width: 10 },
    ];
    expect(ws.getCell('A1').value).toBe('Id');
    expect(ws.getCell('B1').value).toBe('Name');
    expect(ws.getCell('C1').value).toBe('D.O.B.');
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds column headers by number"
  // exceljs: spec/integration/worksheet.spec.js — "adds column headers by letter"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers by number"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds column headers by letter"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers by number"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds column headers by letter"
  it('column key / header / width by property (header cells come from ws.columns only)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getColumn(2).key = 'name';
    ws.getColumn('B').header = 'Name';
    ws.getColumn(2).width = 32;
    expect(ws.getColumn('name')).toBe(ws.getColumn(2));
    expect(ws.getColumn('B').key).toBe('name');
    expect(ws.getColumn(2).header).toBe('Name');
    expect(ws.getColumn(2).width).toBe(32);
    // ➖ no `column.defn` / `headers`, and setting `header` does not write row 1
    expect(ws.getCell('B1').value).toBeNull();
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by object"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by object"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by object"
  it('adds rows by object', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'dob', width: 10 },
    ];
    expect(ws.addRow({ id: 1, name: 'John Doe', dob: d1 })).toBe(2);
    ws.addRow({ id: 2, name: 'Jane Doe', dob: d2 });
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('B2').value).toBe('John Doe');
    expect(ws.getCell('C2').value).toEqual(d1);
    expect(ws.getCell('C3').value).toEqual(d2);
    const values: CellValue[][] = [
      [],
      [undefined, 'Id', 'Name', 'D.O.B.'],
      [undefined, 1, 'John Doe', d1],
      [undefined, 2, 'Jane Doe', d2],
    ];
    expect(ws.getRow(2).values).toEqual(values[2]);
    expect(ws.getRow(3).values).toEqual(values[3]);
    let n = 0;
    ws.eachRow((row, rowNumber) => {
      n++;
      expect(row.values).toEqual(values[rowNumber]);
      // ➖ eachCell receives the value, not a Cell
      row.eachCell((v, col) => expect(v).toEqual(values[rowNumber]![col]));
    });
    expect(n).toBe(3);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by contiguous array"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by contiguous array"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by contiguous array"
  it('adds rows by contiguous array', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.addRow([1, 'John Doe', d1]);
    ws.addRow([2, 'Jane Doe', d2]);
    expect(ws.getCell('A1').value).toBe(1);
    expect(ws.getCell('C2').value).toEqual(d2);
    expect(ws.getRow(1).values).toEqual([undefined, 1, 'John Doe', d1]);
    expect(ws.getRow(2).values).toEqual([undefined, 2, 'Jane Doe', d2]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows by sparse array"
  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds rows by sparse array"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "adds rows by sparse array"
  it('adds rows by sparse array (0-based arrays: index 0 is column A)', () => {
    const ws = new Workbook().addWorksheet('blort');
    // ➖ exceljs treats a sparse array as 1-based ([, 1, 'John Doe']); turbosheet arrays are 0-based
    const rows: CellValue[][] = [
      [1, 'John Doe', undefined, d1],
      [2, 'Jane Doe', undefined, d2],
    ];
    const row3: CellValue[] = [];
    row3[0] = 3;
    row3[2] = 'Sam';
    row3[4] = d1;
    rows.push(row3);
    for (const r of rows) ws.addRow(r);
    expect(ws.getCell('D1').value).toEqual(d1);
    expect(ws.getCell('C3').value).toBe('Sam');
    expect(ws.getCell('E3').value).toEqual(d1);
    ws.eachRow((row, rowNumber) => {
      expect(row.values).toEqual([undefined, ...rows[rowNumber - 1]!]);
    });
  });

  // regression (found porting workbook-xlsx-writer "shared formula"): a rejected value used to
  // leave a half-written row behind, shifting every later addRow down by one
  it('addRow is atomic: an invalid value adds no row', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.addRow([1]);
    expect(() => ws.addRow(['x', { sharedFormula: 'A1' } as never])).toThrow(TypeError);
    expect(ws.rowCount).toBe(1);
    expect(ws.addRow([2])).toBe(2);
    expect(ws.getValue('A2')).toBe(2);
  });

  // exceljs: spec/integration/worksheet.spec.js — "adds rows with style option"
  it('adds rows with style option', () => {
    const ws = new Workbook().addWorksheet('blort');
    const bold = { font: { bold: true } };
    const red = {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } },
    } as const;
    ws.addRow([1, 'a']);
    ws.getCell('A1').style = bold;
    ws.getCell('C1').style = red; // styled but empty
    ws.getRow(1).height = 30;
    expect(ws.addRow([2, 'b'], 'i+')).toBe(2);
    expect(ws.getCell('A2').font).toMatchObject({ bold: true });
    expect(ws.getCell('C2').fill).toEqual(red.fill); // 'i+' copies empty cells' styles too
    expect(ws.getRowProps(2).height).toBe(30);
    ws.addRow([3], 'i');
    expect(ws.getCell('A3').font).toMatchObject({ bold: true });
    expect(ws.getCell('C3').fill).toBeUndefined(); // 'i': only non-empty cells
    ws.addRows([[4], [5]], 'n');
    expect(ws.getCell('A4').font?.bold).toBeUndefined();
    ws.addRow([6], 'o'); // nothing below: no-op, as in exceljs
    expect(ws.getCell('A6').font?.bold).toBeUndefined();
  });

  // exceljs: lib/doc/worksheet.js insertRow / insertRows style option
  it('inserts rows with style option ("i", "o", "o+")', () => {
    const ws = new Workbook().addWorksheet('s');
    ws.addRows([['top'], ['r2'], ['r3']]);
    ws.getCell('A1').style = { font: { italic: true } };
    ws.getCell('A2').style = { font: { bold: true } };
    ws.getCell('B2').style = { numFmt: '0.00' }; // empty
    ws.getCell('A3').style = { font: { strike: true } };
    insertRow(ws, 2, ['new'], 'i');
    expect(ws.getCell('A2').font).toMatchObject({ italic: true });
    insertRows(ws, 3, [['x'], ['y']], 'o+'); // row k <- the row that was at 3 + k
    expect(ws.getCell('A3').font).toMatchObject({ bold: true });
    expect(ws.getCell('B3').numFmt).toBe('0.00');
    expect(ws.getCell('A4').font).toMatchObject({ strike: true });
    insertRow(ws, 1, ['n'], 'o');
    expect(ws.getCell('A1').font).toMatchObject({ italic: true });
    expect(ws.getCell('A1').value).toBe('n');
  });

  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by object"
  it('inserts rows by object', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'dob', width: 10 },
    ];
    ws.addRow({ id: 1, name: 'John Doe', dob: d1 });
    ws.addRow({ id: 2, name: 'Jane Doe', dob: d2 });
    // insert in 3, shifting the earlier row down (➖ a function, not a method; returns void)
    insertRow(ws, 3, { id: 3, name: 'Other Doe', dob: d3 });
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('B2').value).toBe('John Doe');
    expect(ws.getCell('C2').value).toEqual(d1);
    expect(ws.getCell('A3').value).toBe(3);
    expect(ws.getCell('B3').value).toBe('Other Doe');
    expect(ws.getCell('C3').value).toEqual(d3);
    expect(ws.getCell('A4').value).toBe(2);
    expect(ws.getCell('B4').value).toBe('Jane Doe');
    expect(ws.getCell('C4').value).toEqual(d2);
    const values: CellValue[][] = [
      [],
      [undefined, 'Id', 'Name', 'D.O.B.'],
      [undefined, 1, 'John Doe', d1],
      [undefined, 3, 'Other Doe', d3],
      [undefined, 2, 'Jane Doe', d2],
    ];
    let n = 0;
    ws.eachRow((row, rowNumber) => {
      n++;
      expect(row.values).toEqual(values[rowNumber]);
      // ➖ eachCell receives the value, not a Cell; no getRows
      row.eachCell((v, col) => expect(v).toEqual(values[rowNumber]![col]));
    });
    expect(n).toBe(4);
  });

  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by contiguous array"
  it('inserts rows by contiguous array', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.addRow([1, 'John Doe', d1]);
    ws.addRow([2, 'Jane Doe', d2]);
    insertRow(ws, 2, [3, 'Other Doe', d3]);
    expect(ws.getCell('A1').value).toBe(1);
    expect(ws.getCell('C1').value).toEqual(d1);
    expect(ws.getCell('A2').value).toBe(3);
    expect(ws.getCell('B2').value).toBe('Other Doe');
    expect(ws.getCell('C2').value).toEqual(d3);
    expect(ws.getCell('A3').value).toBe(2);
    expect(ws.getCell('B3').value).toBe('Jane Doe');
    expect(ws.getCell('C3').value).toEqual(d2);
    // ➖ no getRows
    expect(ws.getRow(1).values).toEqual([undefined, 1, 'John Doe', d1]);
    expect(ws.getRow(2).values).toEqual([undefined, 3, 'Other Doe', d3]);
    expect(ws.getRow(3).values).toEqual([undefined, 2, 'Jane Doe', d2]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "inserts rows by sparse array"
  it('inserts rows by sparse array (0-based arrays: index 0 is column A)', () => {
    const ws = new Workbook().addWorksheet('blort');
    // ➖ exceljs treats a sparse array as 1-based ([, 1, 'John Doe']); turbosheet arrays are 0-based
    const rows: CellValue[][] = [
      [1, 'John Doe', undefined, d1],
      [2, 'Jane Doe', undefined, d2],
    ];
    const row3: CellValue[] = [];
    row3[0] = 3;
    row3[2] = 'Other Doe';
    row3[4] = d3;
    rows.push(row3);
    // insert on row 1 every time, so the final order is reversed
    for (const r of rows) insertRow(ws, 1, r);
    expect(ws.getCell('A1').value).toBe(3);
    expect(ws.getCell('C1').value).toBe('Other Doe');
    expect(ws.getCell('E1').value).toEqual(d3);
    expect(ws.getCell('A2').value).toBe(2);
    expect(ws.getCell('B2').value).toBe('Jane Doe');
    expect(ws.getCell('D2').value).toEqual(d2);
    expect(ws.getCell('A3').value).toBe(1);
    expect(ws.getCell('B3').value).toBe('John Doe');
    expect(ws.getCell('D3').value).toEqual(d1);
    ws.eachRow((row, rowNumber) => {
      expect(row.values).toEqual([undefined, ...rows[rows.length - rowNumber]!]);
    });
  });

  // exceljs: spec/integration/worksheet.spec.js — "inserts rows with style option"
  it("inserts rows with style option ('above' / 'below' instead of 'i' / 'o' / 'n')", () => {
    const ws = new Workbook().addWorksheet('blort');
    const comic = font(styles.fonts.comicSansUdB16);
    const isComic = (r: number) => ws.getStyle(`A${r}`).font?.name === comic.name;
    const arr = [
      [5, 'Jane Doe', d2],
      [5, 'Jane Doe', d2],
      [5, 'Jane Doe', d2],
    ];
    const styleRow = (r: number) => {
      ws.setRowProps(r, { style: { font: comic } });
      for (const c of ['A', 'B', 'C']) ws.setStyle(`${c}${r}`, { font: comic });
    };
    ws.addRow([5, 'John Doe', d1]);
    styleRow(1);

    // ➖ ExcelJS 'o' ("original": the row that was at pos) = 'below'; 'i' ("inserted": the row
    // above pos) = 'above'; 'n' = no option
    insertRow(ws, 1, [5, 'Jane Doe', d2], { inheritStyle: 'below' });
    insertRow(ws, 1, [4, 'Jane Doe', d2], { inheritStyle: 'above' }); // no row above: unstyled
    insertRow(ws, 1, [3, 'Jane Doe', d2]);
    insertRow(ws, 1, [2, 'Jane Doe', d2], { inheritStyle: 'below' }); // row below is unstyled

    ws.addRow([6, 'Jane Doe', d2]);
    styleRow(6);
    insertRows(ws, 6, arr, { inheritStyle: 'below' });
    insertRows(ws, 10, arr, { inheritStyle: 'above' });
    insertRows(ws, 13, arr);

    expect([1, 2, 3].map(isComic)).toEqual([false, false, false]);
    expect(isComic(4)).toBe(true);
    expect(isComic(5)).toBe(true);
    expect(isComic(6)).toBe(true);
    expect(isComic(9)).toBe(true);
    // ➖ exceljs copies the style to the first inserted row only (A7, A8 stay unstyled);
    // turbosheet formats every inserted row like the row below / above
    expect([7, 8].map(isComic)).toEqual([true, true]);
    expect([10, 11, 12].map(isComic)).toEqual([true, true, true]);
    expect([13, 14, 15].map(isComic)).toEqual([false, false, false]);
    // the row style comes along (height / row xf), not only the cells
    expect(ws.getRowProps(7).style?.font?.name).toBe(comic.name);
    expect(ws.getValue('B13')).toBe('Jane Doe');
  });

  // exceljs: spec/integration/worksheet.spec.js — "should style of the inserted row with inherited style be mutable"
  it('style of an inserted row with inherited style can be changed without touching its source', () => {
    const ws = new Workbook().addWorksheet('blort');
    const comic = font(styles.fonts.comicSansUdB16);
    ws.addRow([1, 'John Doe', d1]);
    ws.setRowProps(1, { style: { font: comic } });
    ws.setStyle('A1', { font: comic });

    insertRow(ws, 2, [3, 'Jane Doe', d2], { inheritStyle: 'above' });
    insertRow(ws, 2, [2, 'Jane Doe', d2], { inheritStyle: 'below' });
    expect(ws.getStyle('A2').font?.name).toBe(comic.name);

    // ➖ styles are interned and immutable: "mutating" means assigning a new style to the row
    // or cell, which cannot leak into the rows it was copied from
    ws.setRowProps(2, { style: { font: font(styles.fonts.broadwayRedOutline20) } });
    ws.setRowProps(3, { style: { font: font(styles.fonts.broadwayRedOutline20) } });
    ws.setStyle('A2', { font: font(styles.fonts.arialBlackUI14) });
    ws.setStyle('A3', { font: font(styles.fonts.arialBlackUI14) });

    expect(ws.getRowProps(2).style?.font?.name).not.toBe(comic.name);
    expect(ws.getRowProps(3).style?.font?.name).not.toBe(comic.name);
    expect(ws.getStyle('A2').font?.name).not.toBe(comic.name);
    expect(ws.getStyle('A3').font?.name).not.toBe(comic.name);
    expect(ws.getStyle('A1').font?.name).toBe(comic.name);
    expect(ws.getRowProps(1).style?.font?.name).toBe(comic.name);
  });

  // exceljs: spec/integration/worksheet.spec.js — "iterates over rows"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "iterates over rows"
  it('iterates over rows (empty rows skipped; includeEmpty visits all)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getCell('A1').value = 1;
    ws.getCell('B2').value = 2;
    ws.getCell('D4').value = 4;
    ws.getCell('F6').value = 6;
    const seen: number[] = [];
    ws.eachRow((_row, n) => seen.push(n));
    expect(seen).toEqual([1, 2, 4, 6]);
    const all: number[] = [];
    ws.eachRow({ includeEmpty: true }, (row, n) => {
      expect(row.number).toBe(n);
      all.push(n);
    });
    expect(all).toEqual([1, 2, 3, 4, 5, 6]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "iterates over collumn cells"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "iterates over collumn cells"
  it('iterates over column cells (empty cells skipped; includeEmpty visits all)', () => {
    const ws = new Workbook().addWorksheet('blort');
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    ws.getCell('A4').value = 4;
    ws.getCell('A6').value = 6;
    const colA = ws.getColumn('A');
    const seen: number[] = [];
    // ➖ the callback gets the value, not a Cell
    colA.eachCell((v, rowNumber) => {
      expect(v).toBe(rowNumber);
      seen.push(rowNumber);
    });
    expect(seen).toEqual([1, 2, 4, 6]);
    let count = 1;
    colA.eachCell({ includeEmpty: true }, (_v, rowNumber) => {
      expect(rowNumber).toBe(count++);
    });
    expect(count).toBe(7);
  });

  // exceljs: spec/integration/worksheet.spec.js — "returns undefined when row range is less than 1"
  it('getRows with length < 1 (➖ an empty array, exceljs: undefined)', () => {
    const ws = new Workbook().addWorksheet('blort');
    expect(getRows(ws, 1, 0)).toEqual([]);
    ws.setCell('A2', 'x');
    expect(getRows(ws, 1, 3).map((r) => r.number)).toEqual([1, 2, 3]);
    expect(getRows(ws, 1, 3)[1]!.getValue(1)).toBe('x');
  });

  describe('worksheet names', () => {
    // exceljs: spec/integration/worksheet.spec.js — "save the original name"
    it('keeps names up to 31 characters', () => {
      const wb = new Workbook();
      let ws = wb.addWorksheet();
      ws.name = 'ThisIsAWorksheetName';
      expect(ws.name).toBe('ThisIsAWorksheetName');
      ws = wb.addWorksheet();
      ws.name = 'ThisIsAWorksheetNameWith31Chars';
      expect(ws.name).toBe('ThisIsAWorksheetNameWith31Chars');
    });

    // exceljs: spec/integration/worksheet.spec.js — "when empty should thrown an error"
    it('throws for an empty name', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = '';
      }).toThrow(/1-31 characters/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "when isn't string should thrown an error"
    it('throws for a non-string name', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = 0 as unknown as string;
      }).toThrow();
    });

    // exceljs: spec/integration/worksheet.spec.js — "thrown an error"
    it('rejects the reserved name History', () => {
      const ws = new Workbook().addWorksheet();
      expect(() => {
        ws.name = 'History';
      }).toThrow(/reserved/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "keep first 31 characters"
    it('names longer than 31 characters throw (exceljs truncates)', () => {
      const ws = new Workbook().addWorksheet();
      // ➖ silently truncating could create duplicates; turbosheet throws
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLongerThan31';
      }).toThrow(/1-31 characters/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (illegal characters)
    it('throws for illegal characters', () => {
      const wb = new Workbook();
      for (const ch of ['*', '?', ':', '/', '\\', '[', ']']) {
        const ws = wb.addWorksheet();
        expect(() => {
          ws.name = ch;
        }).toThrow(/invalid character/);
      }
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (leading / trailing quote)
    it('throws for a leading or trailing apostrophe', () => {
      const wb = new Workbook();
      for (const name of ["'sheetName", "sheetName'"]) {
        const ws = wb.addWorksheet();
        expect(() => {
          ws.name = name;
        }).toThrow(/apostrophe/);
      }
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (duplicate, case-insensitive)
    it('throws for a duplicate name (case-insensitive)', () => {
      const wb = new Workbook();
      wb.addWorksheet().name = 'thisisaworksheetnameinuppercase';
      const ws = wb.addWorksheet();
      expect(() => {
        ws.name = 'THISISAWORKSHEETNAMEINUPPERCASE';
      }).toThrow(/already exists/);
    });

    // exceljs: spec/integration/worksheet.spec.js — "throws an error" (duplicate after truncation)
    it('throws for a duplicate 31-character name', () => {
      const wb = new Workbook();
      wb.addWorksheet().name = 'ThisIsAWorksheetNameThatIsLonge';
      const ws = wb.addWorksheet();
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLonge';
      }).toThrow(/already exists/);
      // ➖ the >31 variant throws for its length rather than as a duplicate
      expect(() => {
        ws.name = 'ThisIsAWorksheetNameThatIsLongerThan31';
      }).toThrow();
    });
  });

  // exceljs: spec/integration/worksheet.spec.js — "returns sheet values"
  // exceljs: spec/unit/doc/worksheet.values.spec.js — "returns sheet values"
  it('sheet values (sheetToArrays instead of getSheetValues)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 11;
    ws.getCell('C1').value = 'C1';
    ws.getCell('A2').value = 21;
    ws.getCell('B2').value = 'B2';
    ws.getCell('A4').value = 'end';
    // ➖ dense, 0-based rows from turbosheet/json; exceljs returns sparse 1-based arrays
    expect(sheetToArrays(ws)).toEqual([
      [11, null, 'C1'],
      [21, 'B2', null],
      [null, null, null],
      ['end', null, null],
    ]);
  });

  // exceljs: spec/integration/worksheet.spec.js — "puts the lotion in the basket"
  it('puts the lotion in the basket', () => {
    const ws = new Workbook().addWorksheet('basket');
    ws.getCell('A1').value = 'lotion';
    expect(ws.getValue('A1')).toBe('lotion');
  });

  describe('Hidden', () => {
    for (const file of [
      'google-sheets',
      'libre-calc-as-excel-2007-365',
      'libre-calc-as-office-open-xml-spreadsheet',
    ]) {
      // exceljs: spec/integration/worksheet.spec.js — "Should set hidden attribute correctly (${file})"
      it(`reads hidden rows and columns (${file})`, () => {
        const ws = readXlsx(fixture(`hidden-test/${file}.xlsx`)).getWorksheet(1)!;
        expect(ws.getRow(1).hidden).toBe(false);
        expect(ws.getRow(2).hidden).toBe(true);
        expect(ws.getRow(3).hidden).toBe(false);
        expect(ws.getColumn(1).hidden).toBe(false);
        expect(ws.getColumn(2).hidden).toBe(true);
        expect(ws.getColumn(3).hidden).toBe(false);
      });
    }
  });

  // exceljs: spec/unit/doc/worksheet.values.spec.js — "calculates rowCount and actualRowCount"
  it('rowCount (actualRowCount is not supported)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 'A1';
    ws.getCell('C1').value = 'C1';
    ws.getCell('A3').value = 'A3';
    ws.getCell('D3').value = 'D3';
    ws.getCell('A4').value = null;
    ws.getCell('B5').value = 'B5';
    expect(ws.rowCount).toBe(5);
    let actual = 0;
    ws.eachRow(() => actual++);
    expect(actual).toBe(3);
  });

  // exceljs: spec/unit/doc/worksheet.values.spec.js — "calculates columnCount and actualColumnCount"
  it('columnCount (actualColumnCount is not supported)', () => {
    const ws = new Workbook().addWorksheet();
    ws.getCell('A1').value = 'A1';
    ws.getCell('C1').value = 'C1';
    ws.getCell('A3').value = 'A3';
    ws.getCell('D3').value = 'D3';
    ws.getCell('E4').value = null;
    ws.getCell('F5').value = 'F5';
    expect(ws.columnCount).toBe(6);
    expect(ws.dimensions?.ref).toBe('A1:F5');
  });

  // exceljs: spec/unit/doc/worksheet.views.spec.js — "adjusts collapsed property of columns"
  // exceljs: spec/unit/doc/worksheet.views.spec.js — "adjusts collapsed property of row"
  // ➖ exceljs derives `collapsed` (outlineLevel > ws.properties.outlineLevel*); here it is a
  // plain flag you set on the summary row / column, and ws.properties.outlineLevelRow/Col only
  // raise the written sheet-level outline level (the default is derived from the rows/columns)
  it('collapsed is a settable flag; sheet outline levels are written as the max', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('sheet1');
    ws.getColumn(2).outlineLevel = 1;
    ws.getColumn(3).outlineLevel = 2;
    ws.setRowProps(2, { outlineLevel: 1 });
    ws.setRowProps(3, { outlineLevel: 2 });
    expect(ws.getColumn(2).collapsed).toBe(false);
    expect(ws.getRowProps(2).collapsed).toBeFalsy();
    ws.getColumn(4).collapsed = true;
    ws.setRowProps(4, { collapsed: true });
    ws.properties.outlineLevelRow = 3;
    ws.properties.outlineLevelCol = 3;
    const ws2 = roundtrip(wb).getWorksheet('sheet1')!;
    expect(ws2.getColumn(4).collapsed).toBe(true);
    expect(ws2.getRowProps(4).collapsed).toBe(true);
    const xml = new TextDecoder().decode(unzipSync(writeXlsx(wb)).get('xl/worksheets/sheet1.xml')!);
    expect(xml).toMatch(/<sheetFormatPr [^>]*outlineLevelRow="3" outlineLevelCol="3"/);
  });

  // exceljs: spec/unit/doc/worksheet.views.spec.js — "sets outline levels via column headers"
  it('sets outline levels via column definitions (and round-trips them)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('sheet1');
    ws.columns = [
      { key: 'id', width: 10, outlineLevel: 1 },
      { key: 'name', width: 32, outlineLevel: 2 },
      { key: 'dob', width: 10, outlineLevel: 3 },
    ];
    ws.setCell('A1', 1);
    expect(ws.getColumn(1).outlineLevel).toBe(1);
    expect(ws.getColumn(2).outlineLevel).toBe(2);
    expect(ws.getColumn(3).outlineLevel).toBe(3);
    const ws2 = roundtrip(wb).getWorksheet('sheet1')!;
    expect([1, 2, 3].map((c) => ws2.getColumn(c).outlineLevel)).toEqual([1, 2, 3]);
  });

  // exceljs: spec/unit/doc/worksheet.pagebreak.spec.js — "adds multiple row breaks"
  it('row page breaks (➖ ws.rowBreaks = row numbers, no row.addPageBreak)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    for (const r of [1, 2, 3]) ws.addRow([`A${r}`, `B${r}`]);
    ws.rowBreaks = [1, 2];
    expect(ws.rowBreaks.length).toBe(2);
    expect(roundtrip(wb).getWorksheet('blort')!.rowBreaks).toEqual([1, 2]);
  });

  // exceljs: spec/integration/worksheet-xlsx-writer.spec.js — "adds multiple row breaks"
  it('row page breaks from the streaming writer', async () => {
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('blort');
      for (const r of [1, 2, 3]) s.addRow([`A${r}`, `B${r}`]);
      s.worksheet.rowBreaks = [1, 2];
    });
    expect(readXlsx(bytes).getWorksheet('blort')!.rowBreaks).toEqual([1, 2]);
  });

  describe('Duplicate Rows', () => {
    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows with styles properly"
    it('duplicates rows with styles properly (read from a file, written and read back)', () => {
      const wb = readXlsx(fixture('duplicateRowTest.xlsx'));
      const ws = wb.getWorksheet('duplicateTest')!;
      ws.setCell('A1', 'OneInfo', ws.getStyle('A1'));
      ws.setCell('A2', 'TwoInfo', ws.getStyle('A2'));
      // ➖ a function; overwrites the two rows below (insert = false is the default)
      duplicateRow(ws, 1, 2);

      const ws2 = roundtrip(wb).getWorksheet('duplicateTest')!;
      expect(ws2.getValue('A2')).toBe('OneInfo');
      expect(ws2.getStyle('A2')).toEqual(ws2.getStyle('A1'));
      expect(ws2.getValue('A3')).toBe('OneInfo');
      expect(ws2.getStyle('A3')).toEqual(ws2.getStyle('A1'));
      expect(ws2.getValue('A4')).toBeNull();
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows replacing properly"
    it('duplicates rows replacing the ones below', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('duplicateTest');
      ws.setCell('A1', 'OneInfo');
      ws.setCell('A2', 'TwoInfo');
      ws.setCell('A3', 'ThreeInfo');
      ws.setCell('A4', 'FourInfo');
      duplicateRow(ws, 1, 2, false);

      const ws2 = roundtrip(wb).getWorksheet('duplicateTest')!;
      expect(ws2.getValue('A1')).toBe('OneInfo');
      expect(ws2.getValue('A2')).toBe('OneInfo');
      expect(ws2.getValue('A3')).toBe('OneInfo');
      expect(ws2.getValue('A4')).toBe('FourInfo');
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows shifting properly"
    it('duplicates rows shifting the ones below', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('duplicateTest');
      ws.setCell('A1', 'OneInfo');
      ws.setCell('A2', 'TwoInfo');
      ws.setCell('A3', 'ThreeInfo');
      ws.setCell('A4', 'FourInfo');
      duplicateRow(ws, 1, 2, true);

      const ws2 = roundtrip(wb).getWorksheet('duplicateTest')!;
      expect(ws2.getValue('A1')).toBe('OneInfo');
      expect(ws2.getValue('A2')).toBe('OneInfo');
      expect(ws2.getValue('A3')).toBe('OneInfo');
      expect(ws2.getValue('A4')).toBe('TwoInfo');
      expect(ws2.getValue('A6')).toBe('FourInfo');
    });

    // exceljs: spec/integration/workbook/workbook.spec.js — "Duplicate rows with height properly"
    it('duplicates the row height', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('duplicateTest');
      ws.setCell('A1', 'OneInfo');
      ws.setCell('A2', 'TwoInfo');
      ws.setRowProps(1, { height: 25 });
      ws.setRowProps(2, { height: 15 });
      duplicateRow(ws, 1, 1, true);

      const ws2 = roundtrip(wb).getWorksheet('duplicateTest')!;
      expect(ws2.getValue('A1')).toBe('OneInfo');
      expect(ws2.getValue('A2')).toBe('OneInfo');
      expect(ws2.getRowProps(1).height).toBe(25);
      expect(ws2.getRowProps(2).height).toBe(ws2.getRowProps(1).height);
      expect(ws2.getRowProps(3).height).not.toBe(ws2.getRowProps(1).height);
    });

    // two rows of cells with fonts / fills / alignments and a row number format each
    function duplicateBook(): Worksheet {
      const ws = new Workbook().addWorksheet('blort');
      const arial = font(styles.fonts.arialBlackUI14);
      const comic = font(styles.fonts.comicSansUdB16);
      ws.setCell('A1', '1.1', { font: arial, numFmt: numFmt1 });
      ws.setCell('B1', '1.2', { font: comic, numFmt: numFmt1 });
      ws.setCell('C1', '1.3', { fill: styles.fills.redDarkVertical, numFmt: numFmt1 });
      ws.setRowProps(1, { style: { numFmt: numFmt1 } });
      const al = styles.namedAlignments;
      ws.setCell('A2', '2.1', { alignment: alignment(al.topLeft), numFmt: numFmt2 });
      ws.setCell('B2', '2.2', { alignment: alignment(al.middleCentre), numFmt: numFmt2 });
      ws.setCell('C2', '2.3', { alignment: alignment(al.bottomRight), numFmt: numFmt2 });
      ws.setRowProps(2, { style: { numFmt: numFmt2 } });
      return ws;
    }

    // exceljs: spec/unit/doc/workbook.spec.js — "inserts duplicates"
    it('inserts duplicates', () => {
      const ws = duplicateBook();
      duplicateRow(ws, 1, 2, true);
      expect(ws.getRow(1).values).toEqual([U, '1.1', '1.2', '1.3']);
      expect(ws.getRow(2).values).toEqual([U, '1.1', '1.2', '1.3']);
      expect(ws.getRow(3).values).toEqual([U, '1.1', '1.2', '1.3']);
      expect(ws.getRow(4).values).toEqual([U, '2.1', '2.2', '2.3']);

      const al = styles.namedAlignments;
      for (let i = 1; i <= 3; i++) {
        expect(ws.getStyle(`A${i}`).font).toEqual(fontBack(styles.fonts.arialBlackUI14));
        expect(ws.getStyle(`B${i}`).font).toEqual(fontBack(styles.fonts.comicSansUdB16));
        expect(ws.getStyle(`C${i}`).fill).toEqual(styles.fills.redDarkVertical);
      }
      expect(ws.getStyle('A4').alignment).toEqual(alignmentBack(al.topLeft));
      expect(ws.getStyle('B4').alignment).toEqual(alignmentBack(al.middleCentre));
      expect(ws.getStyle('C4').alignment).toEqual(alignmentBack(al.bottomRight));
      expect(ws.getRowProps(1).style?.numFmt).toBe(numFmt1);
      expect(ws.getRowProps(2).style?.numFmt).toBe(numFmt1);
      expect(ws.getRowProps(3).style?.numFmt).toBe(numFmt1);
      expect(ws.getRowProps(4).style?.numFmt).toBe(numFmt2);
    });

    // exceljs: spec/unit/doc/workbook.spec.js — "overwrites with duplicates"
    it('overwrites with duplicates', () => {
      const ws = duplicateBook();
      ws.setCell('A3', '3.1', { fill: styles.fills.redGreenDarkTrellis });
      ws.setCell('B3', '3.2', { fill: styles.fills.blueWhiteHGrad });
      ws.setCell('C3', '3.3', { fill: styles.fills.rgbPathGrad });
      ws.setRowProps(3, { style: { font: font(styles.fonts.broadwayRedOutline20) } });

      duplicateRow(ws, 1, 1, false);
      expect(ws.getRow(1).values).toEqual([U, '1.1', '1.2', '1.3']);
      expect(ws.getRow(2).values).toEqual([U, '1.1', '1.2', '1.3']);
      expect(ws.getRow(3).values).toEqual([U, '3.1', '3.2', '3.3']);

      for (let i = 1; i <= 2; i++) {
        expect(ws.getStyle(`A${i}`).font).toEqual(fontBack(styles.fonts.arialBlackUI14));
        expect(ws.getStyle(`A${i}`).alignment).toBeUndefined();
        expect(ws.getStyle(`B${i}`).font).toEqual(fontBack(styles.fonts.comicSansUdB16));
        expect(ws.getStyle(`B${i}`).alignment).toBeUndefined();
        expect(ws.getStyle(`C${i}`).fill).toEqual(styles.fills.redDarkVertical);
        expect(ws.getStyle(`C${i}`).alignment).toBeUndefined();
      }
      expect(ws.getRowProps(1).style?.numFmt).toBe(numFmt1);
      expect(ws.getRowProps(2).style?.numFmt).toBe(numFmt1);
      expect(ws.getRowProps(3).style?.numFmt).toBeUndefined();
      expect(ws.getRowProps(3).style?.font).toEqual(fontBack(styles.fonts.broadwayRedOutline20));
    });
  });

  describe('Splice', () => {
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove only" (Rows)
    it('spliceRows: remove only', () => runSplice(spliced.rows.removeOnly!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert fewer" (Rows)
    it('spliceRows: remove and insert fewer', () => runSplice(spliced.rows.insertFewer!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert same" (Rows)
    it('spliceRows: remove and insert same', () => runSplice(spliced.rows.insertSame!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert more" (Rows)
    it('spliceRows: remove and insert more', () => runSplice(spliced.rows.insertMore!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove style" (Rows)
    it('spliceRows: remove style', () => runSplice(spliced.rows.removeStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Insert style" (Rows)
    it('spliceRows: insert style', () => runSplice(spliced.rows.insertStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Replace style" (Rows)
    it('spliceRows: replace style', () => runSplice(spliced.rows.replaceStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove defined names" (Rows)
    it('spliceRows: remove defined names (the name range shrinks)', () =>
      runSplice(spliced.rows.removeDefinedNames!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Insert defined names" (Rows)
    it('spliceRows: insert defined names (➖ the name range grows)', () =>
      runSplice(spliced.rows.insertDefinedNames!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Replace defined names" (Rows)
    it('spliceRows: replace defined names (➖ the name keeps the replaced row)', () =>
      runSplice(spliced.rows.replaceDefinedNames!));

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "splices columns" (Columns)
    it('spliceColumns: remove only', () => runSplice(spliced.columns.removeOnly!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert fewer" (Columns)
    it('spliceColumns: remove and insert fewer', () => runSplice(spliced.columns.insertFewer!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert same" (Columns)
    it('spliceColumns: remove and insert same', () => runSplice(spliced.columns.insertSame!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove and insert more" (Columns)
    it('spliceColumns: remove and insert more', () => runSplice(spliced.columns.insertMore!));

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "handles column keys"
    it('spliceColumns: handles column keys (and widths)', () => {
      const ws = new Workbook().addWorksheet('splice-column-insert-fewer');
      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'dob', width: 20 },
        { key: 'name', width: 30 },
        { key: 'age', width: 40 },
      ];
      const values = [
        { id: '123', name: 'Jack', dob: new Date(), age: 0 },
        { id: '124', name: 'Jill', dob: new Date(), age: 0 },
      ];
      for (const v of values) ws.addRow(v);

      spliceColumns(ws, 2, 1, ['B1', 'B2'], ['C1', 'C2']);

      values.forEach((rowValues, index) => {
        for (const key of ['id', 'name', 'age'] as const) {
          expect(ws.getValue(index + 1, ws.getColumn(key).number), key).toBe(rowValues[key]);
        }
      });
      expect(ws.getColumn(1).width).toBe(10);
      expect(ws.getColumn(2).width).toBeUndefined();
      expect(ws.getColumn(3).width).toBeUndefined();
      expect(ws.getColumn(4).width).toBe(30);
      expect(ws.getColumn(5).width).toBe(40);
    });

    // six columns "Col-1".."Col-6" with headers in row 1, plus two rows of 1..6
    function sixColumns(): Worksheet {
      const ws = new Workbook().addWorksheet('splice-to-end');
      ws.columns = [1, 2, 3, 4, 5, 6].map((n) => ({ header: `Col-${n}`, width: 10 }));
      ws.addRow([1, 2, 3, 4, 5, 6]);
      ws.addRow([1, 2, 3, 4, 5, 6]);
      return ws;
    }
    function expectSixColumns(ws: Worksheet, kept: number[]): void {
      // rows 1 (headers), 2, 3: the kept columns in order, nothing after
      const header = kept.map((n) => `Col-${n}`);
      for (let c = 1; c <= 7; c++) {
        const k = kept[c - 1];
        expect(ws.getValue(1, c), `1,${c}`).toBe(k === undefined ? null : header[c - 1]);
        expect(ws.getValue(2, c), `2,${c}`).toBe(k === undefined ? null : k);
        expect(ws.getValue(3, c), `3,${c}`).toBe(k === undefined ? null : k);
      }
      for (let c = 1; c <= 6; c++) {
        const k = kept[c - 1];
        expect(ws.getColumn(c).header, `header ${c}`).toBe(
          k === undefined ? undefined : `Col-${k}`,
        );
      }
    }

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Splices to end"
    it('spliceColumns: splices to end', () => {
      const ws = sixColumns();
      spliceColumns(ws, 4, 3); // the last 3 columns
      expectSixColumns(ws, [1, 2, 3]);
    });

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Splices past end"
    it('spliceColumns: splices past end', () => {
      const ws = sixColumns();
      spliceColumns(ws, 4, 4);
      expectSixColumns(ws, [1, 2, 3]);
    });

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Splices almost to end"
    it('spliceColumns: splices almost to end', () => {
      const ws = sixColumns();
      spliceColumns(ws, 4, 2);
      expectSixColumns(ws, [1, 2, 3, 6]);
    });

    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove style" (Columns)
    it('spliceColumns: remove style', () => runSplice(spliced.columns.removeStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Insert style" (Columns)
    it('spliceColumns: insert style', () => runSplice(spliced.columns.insertStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Replace style" (Columns)
    it('spliceColumns: replace style', () => runSplice(spliced.columns.replaceStyle!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Remove defined names" (Columns)
    it('spliceColumns: remove defined names (the name range shrinks)', () =>
      runSplice(spliced.columns.removeDefinedNames!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Insert defined names" (Columns)
    it('spliceColumns: insert defined names (➖ the name range grows)', () =>
      runSplice(spliced.columns.insertDefinedNames!));
    // exceljs: spec/unit/doc/worksheet.values.spec.js — "Replace defined names" (Columns)
    it('spliceColumns: replace defined names (➖ the name keeps the replaced column)', () =>
      runSplice(spliced.columns.replaceDefinedNames!));

    // exceljs: spec/integration/workbook/workbook.spec.js — "spliced meat and ham"
    it('spliced sheets survive a write / read round trip', () => {
      const wb = new Workbook();
      const picked = [
        spliced.rows.removeOnly!,
        spliced.rows.insertFewer!,
        spliced.rows.insertSame!,
        spliced.rows.insertMore!,
        spliced.rows.insertStyle!,
        spliced.columns.removeOnly!,
        spliced.columns.insertFewer!,
        spliced.columns.insertSame!,
        spliced.columns.insertMore!,
      ];
      for (const s of picked) s.add(wb);
      const wb2 = roundtrip(wb);
      for (const s of picked) s.check(wb2);
    });

    // spec/unit/doc/row.spec.js "Splice" works on one Row (cells of that row only). turbosheet has
    // no Row.splice; on a one-row sheet spliceColumns(ws, start, count, [v], [v], ...) is the
    // same edit: one single-value column per inserted value.
    const rowSheet = (values: CellValue[]): Worksheet => {
      const ws = new Workbook().addWorksheet('s');
      ws.addRow(values);
      return ws;
    };
    const rowOf = (...cells: [number, CellValue][]): Worksheet => {
      const ws = new Workbook().addWorksheet('s');
      for (const [c, v] of cells) ws.setCell(1, c, v);
      return ws;
    };

    // exceljs: spec/unit/doc/row.spec.js — "remove only"
    it('row splice (via spliceColumns on a one-row sheet): remove only', () => {
      const ws = rowOf([2, 2], [3, 3], [5, 5], [7, 7], [11, 'eleven'], [13, 13]);
      spliceColumns(ws, 3, 3);
      expect(ws.getValue(1, 2)).toBe(2);
      expect(ws.getKind(1, 2)).toBe(CellKind.Number);
      expect(ws.getCell(1, 2).address).toBe('B1');
      expect(ws.getValue(1, 4)).toBe(7);
      expect(ws.getKind(1, 4)).toBe(CellKind.Number);
      expect(ws.getCell(1, 4).address).toBe('D1');
      expect(ws.getValue(1, 8)).toBe('eleven');
      expect(ws.getKind(1, 8)).toBe(CellKind.String);
      expect(ws.getCell(1, 8).address).toBe('H1');
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove to end"
    it('row splice (via spliceColumns on a one-row sheet): remove to end', () => {
      const ws = rowSheet([1, 2, 3, 4, 5]);
      spliceColumns(ws, 4, 2);
      expect([1, 2, 3].map((c) => ws.getValue(1, c))).toEqual([1, 2, 3]);
      expect([4, 5, 6].map((c) => ws.getValue(1, c))).toEqual([null, null, null]);
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove almost to end"
    it('row splice (via spliceColumns on a one-row sheet): remove almost to end', () => {
      const ws = rowSheet([1, 2, 3, 4, 5, 6]);
      spliceColumns(ws, 4, 2);
      expect([1, 2, 3, 4].map((c) => ws.getValue(1, c))).toEqual([1, 2, 3, 6]);
      expect([5, 6].map((c) => ws.getValue(1, c))).toEqual([null, null]);
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove past end"
    it('row splice (via spliceColumns on a one-row sheet): remove past end', () => {
      const ws = rowSheet([1, 2, 3, 4, 5, 6]);
      spliceColumns(ws, 4, 4);
      expect([1, 2, 3].map((c) => ws.getValue(1, c))).toEqual([1, 2, 3]);
      expect([4, 5, 6, 7, 8].map((c) => ws.getValue(1, c))).toEqual([null, null, null, null, null]);
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove and insert fewer"
    it('row splice (via spliceColumns on a one-row sheet): remove and insert fewer', () => {
      const ws = rowSheet([1, 2, 3, 4, 5, 6, 7, 8]);
      spliceColumns(ws, 4, 3, ['four'], ['five']);
      expect(ws.getRow(1).values).toEqual([U, 1, 2, 3, 'four', 'five', 7, 8]);
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove and insert replacements"
    it('row splice (via spliceColumns on a one-row sheet): remove and insert replacements', () => {
      const ws = rowSheet([1, 2, 3, 4, 5, 6, 7, 8]);
      spliceColumns(ws, 4, 3, ['four'], ['five'], ['six']);
      expect(ws.getRow(1).values).toEqual([U, 1, 2, 3, 'four', 'five', 'six', 7, 8]);
    });

    // exceljs: spec/unit/doc/row.spec.js — "remove and insert more"
    it('row splice (via spliceColumns on a one-row sheet): remove and insert more', () => {
      const ws = rowSheet([1, 2, 3, 4, 5, 6, 7, 8]);
      spliceColumns(ws, 4, 3, ['four'], ['five'], ['six'], ['six and a half']);
      expect(ws.getRow(1).values).toEqual([
        U,
        1,
        2,
        3,
        'four',
        'five',
        'six',
        'six and a half',
        7,
        8,
      ]);
    });
  });
});
