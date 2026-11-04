// Print titles: spec/integration/workbook/workbook.spec.js "printTitlesRow", "printTitlesColumn",
// "printTitlesRowAndColumn". Page setup / header-footer round trips are part of the test book
// (values.test.ts).

import { describe, expect, it } from 'vitest';
import { getPrintTitles, readXlsx, setPrintTitles, Workbook, writeXlsx } from '../../src/index.js';
import { exceljsLoad } from './helpers.js';

async function check(set: { rows?: string; cols?: string }) {
  const wb = new Workbook();
  const ws = wb.addWorksheet('print');
  ws.getCell('A1').value = 'This is a header row repeated on every printed page';
  for (let i = 0; i < 100; i++) ws.addRow(['not header row']);
  // ➖ exceljs ws.pageSetup.printTitlesRow / printTitlesColumn
  setPrintTitles(ws, set);
  const bytes = writeXlsx(wb);
  const t = getPrintTitles(readXlsx(bytes).getWorksheet('print')!)!;
  expect(t.rows).toBe(set.rows);
  expect(t.cols).toBe(set.cols);
  const x = (await exceljsLoad(bytes)).getWorksheet('print')!;
  expect(x.pageSetup.printTitlesRow).toBe(set.rows);
  expect(x.pageSetup.printTitlesColumn).toBe(set.cols);
}

describe('exceljs parity: print', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "printTitlesRow"
  it('printTitlesRow', () => check({ rows: '1:2' }));

  // exceljs: spec/integration/workbook/workbook.spec.js — "printTitlesColumn"
  it('printTitlesColumn', () => check({ cols: 'A:B' }));

  // exceljs: spec/integration/workbook/workbook.spec.js — "printTitlesRowAndColumn"
  it('printTitlesRowAndColumn', () => check({ rows: '1:2', cols: 'A:B' }));
});
