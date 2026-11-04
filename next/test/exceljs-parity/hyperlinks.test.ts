// Hyperlinks: the hyperlinked formula of formulas.xlsx (spec/integration/workbook-xlsx-reader.spec.js)
// and the cell hyperlink of spec/integration/workbook/images.spec.js. Issues 275 and 877 are in
// issues.test.ts.

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad, fixture, roundtrip } from './helpers.js';

describe('exceljs parity: hyperlinks', () => {
  describe('with a cell that contains a hyperlinked formula (formulas.xlsx A1)', () => {
    const a1 = () => readXlsx(fixture('formulas.xlsx')).worksheets[0]!.getValue('A1');

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should be classified as a formula cell" (hyperlinked)
    it('keeps the formula and its cached text', () => {
      // ➖ a cell has one kind: the formula wins (exceljs reports a Hyperlink cell with a formula)
      expect(readXlsx(fixture('formulas.xlsx')).worksheets[0]!.getKind('A1')).toBe(
        CellKind.Formula,
      );
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should have text corresponding to the evaluated formula result" (hyperlinked)
    it('has the evaluated text', () => {
      expect((a1() as { result: string }).result).toBe('someone@example.com');
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should have the formula source" (hyperlinked)
    it('has the formula source', () => {
      expect((a1() as { formula: string }).formula).toBe('_xlfn.CONCAT("someone","@example.com")');
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should contain the linked url"
    // 🐛 (open, needs a design decision): the hyperlink of a formula cell is not modelled and
    // not preserved, so it is dropped on read and lost on a read → write round trip.
    it.skip('contains the linked url (mailto:) — hyperlink on a formula cell is dropped', () => {
      const wb = readXlsx(fixture('formulas.xlsx'));
      const out = writeXlsx(wb);
      expect(new TextDecoder().decode(out)).toContain('mailto:someone@example.com');
    });
  });

  // exceljs: spec/integration/workbook/images.spec.js — "stores embedded image and hyperlink" (cell part)
  it('cell hyperlinks (external, internal, tooltip) round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('blort');
    wb.addWorksheet('Other');
    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = { hyperlink: 'http://www.somewhere.com', text: 'www.somewhere.com' };
    ws.getCell('A3').value = { hyperlink: '#Other!B2', text: 'jump', tooltip: 'go there' };
    const bytes = writeXlsx(wb);
    const ws2 = readXlsx(bytes).getWorksheet('blort')!;
    expect(ws2.getValue('A1')).toBe('Hello, World!');
    expect(ws2.getValue('A2')).toEqual({
      hyperlink: 'http://www.somewhere.com',
      text: 'www.somewhere.com',
    });
    expect(ws2.getValue('A3')).toEqual({
      hyperlink: '#Other!B2',
      text: 'jump',
      tooltip: 'go there',
    });
    expect(roundtrip(readXlsx(bytes)).getWorksheet('blort')!.getValue('A3')).toEqual(
      ws2.getValue('A3'),
    );
    const x = (await exceljsLoad(bytes)).getWorksheet('blort')!;
    expect(x.getCell('A2').value).toEqual({
      hyperlink: 'http://www.somewhere.com',
      text: 'www.somewhere.com',
    });
  });

  // exceljs: spec/unit/doc/cell.spec.js — "upgrades from string to hyperlink"
  // exceljs: spec/unit/doc/cell.spec.js — "doesn't upgrade from non-string to hyperlink"
  it.skip('cell.hyperlink setter upgrading a string cell — not supported (assign a Hyperlink value)', () => {});
});
