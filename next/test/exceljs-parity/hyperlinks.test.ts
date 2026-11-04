// Hyperlinks: the hyperlinked formula of formulas.xlsx (spec/integration/workbook-xlsx-reader.spec.js)
// and the cell hyperlink of spec/integration/workbook/images.spec.js. Issues 275 and 877 are in
// issues.test.ts.

import { describe, expect, it } from 'vitest';
import { CellKind, readXlsx, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad, fixture, partText, roundtrip } from './helpers.js';

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
      expect((a1() as { formula: string }).formula).toBe('CONCAT("someone","@example.com")'); // _xlfn. stripped on read
    });

    // exceljs: spec/integration/workbook-xlsx-reader.spec.js — "should contain the linked url"
    // 🐛 fixed: the hyperlink of a formula cell used to be dropped (lost on round trip). It is
    // not modelled (a cell holds one kind) but kept raw in ws.preserved and written back while
    // the cell still holds the formula it was read with.
    it('keeps the linked url (mailto:) of a formula cell on round trip', async () => {
      const wb = readXlsx(fixture('formulas.xlsx'));
      wb.worksheets[0]!.setCell('C5', 'unrelated edit');
      const out = writeXlsx(wb);
      const sheet = partText(out, 'xl/worksheets/sheet1.xml')!;
      const id = /<hyperlink ref="A1" r:id="(rId\d+)"[^>]*\/>/.exec(sheet)?.[1];
      expect(id).toBeDefined();
      expect(sheet).toContain('display="mailto:someone@example.com"');
      const rels = partText(out, 'xl/worksheets/_rels/sheet1.xml.rels')!;
      expect(rels).toMatch(
        new RegExp(`Id="${id}"[^>]*Target="mailto:someone@example.com"[^>]*TargetMode="External"`),
      );
      // still a formula for turbosheet; exceljs sees the hyperlinked formula again
      expect(readXlsx(out).worksheets[0]!.getValue('A1')).toEqual(a1());
      const x = (await exceljsLoad(out)).worksheets[0]!;
      expect(x.getCell('A1').hyperlink).toBe('mailto:someone@example.com');
      // a second round trip keeps it too
      expect(partText(writeXlsx(readXlsx(out)), 'xl/worksheets/sheet1.xml')).toContain(
        '<hyperlink ref="A1"',
      );
    });

    it('drops the kept link when the formula cell is changed or cleared', () => {
      for (const edit of [
        (wb: ReturnType<typeof readXlsx>) => wb.worksheets[0]!.setCell('A1', 'plain'),
        (wb: ReturnType<typeof readXlsx>) =>
          wb.worksheets[0]!.setCell('A1', { formula: '_xlfn.CONCAT("a","b")', result: 'ab' }),
        (wb: ReturnType<typeof readXlsx>) => wb.worksheets[0]!.clearCell('A1'),
      ]) {
        const wb = readXlsx(fixture('formulas.xlsx'));
        edit(wb);
        const out = writeXlsx(wb);
        expect(partText(out, 'xl/worksheets/sheet1.xml')).not.toContain('<hyperlink');
        expect(partText(out, 'xl/worksheets/_rels/sheet1.xml.rels') ?? '').not.toContain('mailto:');
      }
    });

    it('kept formula links merge with modelled hyperlinks', () => {
      const wb = readXlsx(fixture('formulas.xlsx'));
      wb.worksheets[0]!.setCell('C1', { text: 'site', hyperlink: 'https://example.com' });
      const out = writeXlsx(wb);
      const sheet = partText(out, 'xl/worksheets/sheet1.xml')!;
      expect(sheet.match(/<hyperlinks>/g)).toHaveLength(1);
      expect(sheet).toMatch(/<hyperlink ref="C1"[^>]*\/><hyperlink ref="A1"/);
      const ws2 = readXlsx(out).worksheets[0]!;
      expect(ws2.getValue('C1')).toEqual({ text: 'site', hyperlink: 'https://example.com' });
      const rels = partText(out, 'xl/worksheets/_rels/sheet1.xml.rels')!;
      const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]);
      expect(new Set(ids).size).toBe(ids.length); // no rId collision
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
