// Defined names: spec/integration/workbook/workbook.spec.js "defined names",
// workbook-xlsx-writer.spec.js "defined names", spec/unit/doc/defined-names.spec.js and
// bogus-defined-name.xlsx (workbook-xlsx-reader.spec.js). pr 567 is in prs.test.ts.

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  getDefinedName,
  getPrintArea,
  readXlsx,
  removeDefinedName,
  setPrintArea,
  spliceColumns,
  spliceRows,
  Workbook,
  type Worksheet,
} from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip, streamWrite } from './helpers.js';

describe('exceljs parity: defined names', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "defined names"
  it('names over cells, ranges and several sheets round-trip', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('blort');
    const b = wb.addWorksheet('foo');
    a.setCell('A1', 5);
    for (const r of ['A3', 'B3', 'C3', 'E1', 'E2', 'E3']) a.setCell(r, 3);
    // ➖ exceljs builds names from `cell.name` / `cell.names` / addName / removeName and merges
    // adjacent cells into ranges; turbosheet takes the range directly
    addDefinedName(wb, 'five', 'blort!A1');
    addDefinedName(wb, 'amigos', 'blort!A3:C3');
    addDefinedName(wb, 'verts', 'blort!E1:E3');
    addDefinedName(wb, 'squares', 'blort!B5:C6');
    addDefinedName(wb, 'sheets', 'blort!$B$7,foo!$B$7');
    addDefinedName(wb, 'thing1', 'blort!G1');
    addDefinedName(wb, 'thing2', 'blort!G1');
    addDefinedName(wb, 'once', 'blort!G2');
    addDefinedName(wb, 'twice', 'blort!G2');
    expect(removeDefinedName(wb, 'once')).toBe(true);
    b.setCell('B7', 2);

    const bytes = (await import('../../src/index.js')).writeXlsx(wb);
    const wb2 = readXlsx(bytes);
    // ➖ sheet names in references are always quoted
    const v = (n: string) => getDefinedName(wb2, n)?.value;
    expect(v('five')).toBe("'blort'!$A$1");
    expect(v('amigos')).toBe("'blort'!$A$3:$C$3");
    expect(v('verts')).toBe("'blort'!$E$1:$E$3");
    expect(v('squares')).toBe("'blort'!$B$5:$C$6");
    expect(v('sheets')).toBe('blort!$B$7,foo!$B$7');
    expect(v('thing1')).toBe("'blort'!$G$1");
    expect(v('thing2')).toBe("'blort'!$G$1");
    expect(v('once')).toBeUndefined();
    expect(v('twice')).toBe("'blort'!$G$2");

    const x = await exceljsLoad(bytes);
    expect(x.getWorksheet('blort')!.getCell('A1').name).toBe('five');
    expect(x.definedNames.getRanges('sheets').ranges).toEqual(
      expect.arrayContaining(['blort!$B$7', 'foo!$B$7']),
    );
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "defined names"
  it('streaming writer: defined names', async () => {
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('blort');
      s.addRow([5]);
      s.addRow([]);
      s.addRow(['drei', 'san']);
      addDefinedName(w.workbook, 'five', 'blort!A1');
      addDefinedName(w.workbook, 'threes', 'blort!A3:B3');
      addDefinedName(w.workbook, 'greens', 'blort!E1:E3');
    });
    const wb2 = readXlsx(bytes);
    expect(getDefinedName(wb2, 'five')?.value).toBe("'blort'!$A$1");
    expect(getDefinedName(wb2, 'threes')?.value).toBe("'blort'!$A$3:$B$3");
    expect(getDefinedName(wb2, 'greens')?.value).toBe("'blort'!$E$1:$E$3");
    const x = (await exceljsLoad(bytes)).getWorksheet('blort')!;
    expect(x.getCell('B3').name).toBe('threes');
  });

  // exceljs: spec/unit/doc/defined-names.spec.js — "adds names for cells"
  // exceljs: spec/unit/doc/defined-names.spec.js — "removes names for cells"
  // exceljs: spec/unit/doc/defined-names.spec.js — "gets the right ranges for a name"
  it('quoted sheet names, removal and lookup', () => {
    const wb = new Workbook();
    wb.addWorksheet('blo rt');
    wb.addWorksheet("it's");
    addDefinedName(wb, 'bar', "'blo rt'!B4");
    addDefinedName(wb, 'q', "'it''s'!A1:A3");
    addDefinedName(wb, 'Foo', 'blo rt!A1'.replace('blo rt', "'blo rt'"));
    expect(getDefinedName(wb, 'bar')?.value).toBe("'blo rt'!$B$4");
    expect(getDefinedName(wb, 'foo')?.value).toBe("'blo rt'!$A$1"); // case-insensitive
    expect(removeDefinedName(wb, 'FOO')).toBe(true);
    const wb2 = roundtrip(wb);
    expect(getDefinedName(wb2, 'q')?.value).toBe("'it''s'!$A$1:$A$3");
    expect(getDefinedName(wb2, 'foo')).toBeUndefined();
  });

  // exceljs: spec/unit/doc/defined-names.spec.js — "splices"
  it('defined names follow spliced rows and columns (➖ like Excel, not cell by cell)', () => {
    const wb = new Workbook();
    const sheets = Object.fromEntries(
      ['vertical', 'horizontal', 'square', 'single'].map((n) => [n, wb.addWorksheet(n)]),
    ) as Record<string, Worksheet>;
    // exceljs lists the named cells; turbosheet names a range
    addDefinedName(wb, 'vertical', 'vertical!A1:A4');
    addDefinedName(wb, 'horizontal', 'horizontal!A1:D1');
    addDefinedName(wb, 'square', 'square!A1:D4');
    for (const a of ['A1', 'D1', 'A4', 'D4']) addDefinedName(wb, `single${a}`, `single!${a}`);

    // exceljs: dn.spliceRows(name, 2, 2, 1) = delete 2 rows at 2, insert 1
    spliceRows(sheets.vertical!, 2, 2, ['new']);
    spliceColumns(sheets.horizontal!, 2, 2, ['new']);
    spliceRows(sheets.square!, 2, 2, ['new']);
    spliceColumns(sheets.square!, 2, 2, ['new']);
    spliceRows(sheets.single!, 2, 2, ['new']);
    spliceColumns(sheets.single!, 2, 2, ['new']);

    // ➖ exceljs: vertical = A1 and A3 (the inserted row 2 is not part of the name). The
    // replaced row is overwritten in place, so the reference to it stays: A1:A4 loses row 3
    const v = (n: string) => getDefinedName(wb, n)?.value;
    expect(v('vertical')).toBe("'vertical'!$A$1:$A$3");
    expect(v('horizontal')).toBe("'horizontal'!$A$1:$C$1");
    expect(v('square')).toBe("'square'!$A$1:$C$3");
    // single cells move exactly as in exceljs
    expect(v('singleA1')).toBe("'single'!$A$1");
    expect(v('singleD1')).toBe("'single'!$C$1");
    expect(v('singleA4')).toBe("'single'!$A$3");
    expect(v('singleD4')).toBe("'single'!$C$3");
  });

  // exceljs: spec/unit/doc/defined-names.spec.js — "creates matrix from model"
  it.skip('DefinedNames cell matrix — not applicable (exceljs internals)', () => {});

  // exceljs: spec/unit/doc/defined-names.spec.js — "skips values with invalid range"
  // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should not crash" (bogus-defined-name.xlsx)
  it('a defined name that only looks like a range is kept as-is', () => {
    const wb = readXlsx(fixture('bogus-defined-name.xlsx'));
    expect(wb.definedNames.length).toBeGreaterThan(0);
    const back = roundtrip(wb);
    expect(back.definedNames).toEqual(wb.definedNames);
  });

  // regression (found by this port): removeWorksheet left sheet-scoped names (print area,
  // print titles, local names) pointing at the wrong tab index
  it('sheet-scoped names stay on their sheet when an earlier sheet is removed', () => {
    const wb = new Workbook();
    wb.addWorksheet('A');
    wb.addWorksheet('B');
    const c = wb.addWorksheet('C');
    setPrintArea(c, 'A1:D20');
    addDefinedName(wb, 'local', 'C!A1', { localSheet: c });
    wb.removeWorksheet('A');
    expect(getPrintArea(c)).toEqual(['A1:D20']);
    const wb2 = roundtrip(wb);
    expect(getPrintArea(wb2.getWorksheet('C')!)).toEqual(['A1:D20']);
    expect(getPrintArea(wb2.getWorksheet('B')!)).toBeUndefined();
    expect(getDefinedName(wb2, 'local', 'C')?.value).toBe("'C'!$A$1");
  });
});
