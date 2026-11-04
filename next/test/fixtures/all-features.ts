// A deterministic workbook using every built-in feature (tables, notes, images, conditional
// formatting incl. an x14 data bar, data validation, print setup, sheet protection). Shared by
// test/lite.test.ts; takes the core API as an argument so the same builder works against any
// entry (root, lite, an older build).

import type { Workbook as WorkbookT } from '../../src/core/workbook.js';
import type * as Features from '../../src/xlsx/features/index.js';

const b64 = (s: string): Uint8Array => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
/** 1x1 PNG. */
export const PNG: Uint8Array = b64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
);

type Api = Pick<
  typeof Features,
  | 'addTable'
  | 'addNote'
  | 'addImage'
  | 'placeImage'
  | 'addConditionalFormatting'
  | 'addDataValidation'
>;

const FIXED = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));

export function allFeatures(WorkbookCtor: new () => WorkbookT, f: Api): WorkbookT {
  const wb = new WorkbookCtor();
  wb.created = FIXED;
  wb.modified = FIXED;
  wb.creator = 'fixture';
  const ws = wb.addWorksheet('Data');
  ws.getCell('A1').value = 'plain';
  f.addTable(ws, {
    name: 'Sales',
    topLeft: 'B2',
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Units', totalsRowFunction: 'sum' },
    ],
    rows: [
      ['North', 10],
      ['South', 20],
    ],
    totalsRow: true,
    headerStyle: { font: { bold: true, color: { argb: 'FF0000FF' } } },
  });
  f.addNote(ws, 'A1', { text: 'a note', author: 'Ann' });
  const img = f.addImage(wb, { data: PNG, extension: 'png' });
  f.placeImage(ws, img, 'F2:G4');
  f.addConditionalFormatting(ws, 'C3:C4', [
    {
      type: 'cellIs',
      operator: 'greaterThan',
      formulae: [15],
      style: { font: { bold: true, color: { argb: 'FF9C0006' } } },
    },
    { type: 'dataBar', color: { argb: 'FF638EC6' }, gradient: false },
  ]);
  f.addDataValidation(ws, { sqref: 'A2', type: 'list', values: ['x', 'y'] });
  ws.pageSetup = {
    orientation: 'landscape',
    paperSize: 9,
    margins: { left: 1, right: 1, top: 1.5, bottom: 1.5, header: 0.5, footer: 0.5 },
  };
  ws.headerFooter = { oddHeader: '&CReport' };
  ws.protection = { formatCells: true };
  const ws2 = wb.addWorksheet('Two');
  ws2.getCell('B2').value = 42;
  return wb;
}
