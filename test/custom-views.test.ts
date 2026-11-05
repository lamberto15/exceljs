// Custom views (#112) and threaded comments (#111) read from a file saved by Excel for Mac
// (test/fixtures/excel-saved-views-threads.xlsx: views "Small" and "Zoomed", two threads with a
// reply, one note), and custom views authored with addCustomView.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { addCustomView, readXlsx, Workbook, writeXlsx } from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const FIXTURE = readFileSync(
  fileURLToPath(new URL('./fixtures/excel-saved-views-threads.xlsx', import.meta.url)),
);
const dec = new TextDecoder();
const part = (buf: Uint8Array, p: string) => dec.decode(unzipSync(buf).get(p)!);

describe('custom views', () => {
  it('reads Excel-saved custom views and threads', () => {
    const wb = readXlsx(FIXTURE);
    expect(wb.customViews).toEqual([
      {
        name: 'Small',
        guid: '{0A922094-FD87-C147-94AD-7D30377714F5}',
        windowWidth: 1440,
        windowHeight: 615,
        activeSheetId: 1,
        attributes: { yWindow: '33' },
      },
      {
        name: 'Zoomed',
        guid: '{C175F84C-3796-4A8B-89C2-B4F2F5A1F2B1}',
        windowWidth: 16000,
        windowHeight: 9000,
        activeSheetId: 1,
      },
    ]);
    const ws = wb.worksheets[0]!;
    expect(ws.customSheetViews).toEqual([
      {
        guid: '{0A922094-FD87-C147-94AD-7D30377714F5}',
        scale: 150,
        children:
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
      },
      {
        guid: '{C175F84C-3796-4A8B-89C2-B4F2F5A1F2B1}',
        scale: 150,
        children:
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
      },
    ]);
    expect(
      ws.threadedComments!.map((t) => [t.ref, t.text, t.author, t.replies?.length ?? 0]),
    ).toEqual([
      ['B1', 'Is this right?', 'Ann', 1],
      ['A2', 'Resolved one', 'Bob', 0],
    ]);
    expect(ws.comments!.map((n) => n.ref)).toEqual(['B2']);
  });

  it('round-trips the Excel file: same views, same threads, no duplicates', () => {
    const wb = readXlsx(FIXTURE);
    const out = writeXlsx(wb);
    const book = part(out, 'xl/workbook.xml');
    expect(book.match(/<customWorkbookViews>/g)).toHaveLength(1);
    expect(book).toContain(
      '<customWorkbookView name="Small" guid="{0A922094-FD87-C147-94AD-7D30377714F5}" yWindow="33" windowWidth="1440" windowHeight="615" activeSheetId="1"/>',
    );
    const sheet = part(out, 'xl/worksheets/sheet1.xml');
    expect(sheet.match(/<customSheetViews>/g)).toHaveLength(1);
    expect(sheet).toMatch(/<\/sheetData>.*<customSheetViews>.*<\/customSheetViews>.*<pageMargins/);
    const back = readXlsx(out);
    expect(back.customViews).toEqual(wb.customViews);
    expect(back.worksheets[0]!.customSheetViews).toEqual(wb.worksheets[0]!.customSheetViews);
    expect(back.worksheets[0]!.threadedComments).toEqual(wb.worksheets[0]!.threadedComments);
    expect(back.persons).toEqual(wb.persons);
  });

  it('deleting the views removes them (the kept workbook element does not come back)', () => {
    const wb = readXlsx(FIXTURE);
    wb.customViews = [];
    delete wb.worksheets[0]!.customSheetViews;
    const out = writeXlsx(wb);
    expect(part(out, 'xl/workbook.xml')).not.toContain('customWorkbookView');
    expect(part(out, 'xl/worksheets/sheet1.xml')).not.toContain('customSheetView');
  });

  it('addCustomView captures each sheet; readers see it', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    a.getCell('A1').value = 1;
    a.views = [{ state: 'frozen', ySplit: 1, zoomScale: 120, showGridLines: false }];
    const b = wb.addWorksheet('B');
    b.state = 'hidden';
    const v = addCustomView(wb, 'Review', { includePrintSettings: false });
    expect(addCustomView(wb, 'review').guid).toBe(v.guid); // same name: replaced
    expect(wb.customViews).toHaveLength(1);
    const out = writeXlsx(wb);
    const sa = part(out, 'xl/worksheets/sheet1.xml');
    expect(sa).toContain(`<customSheetView guid="${v.guid}" scale="120" showGridLines="0">`);
    expect(sa).toContain('<pane ySplit="1" activePane="bottomLeft" state="frozen"/>');
    expect(part(out, 'xl/worksheets/sheet2.xml')).toContain(' state="hidden"');
    const back = readXlsx(out);
    expect(back.customViews![0]).toMatchObject({ name: 'review', guid: v.guid });
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(out as unknown as ArrayBuffer);
    expect(x.worksheets).toHaveLength(2);
  });

  it('lite without the feature keeps the elements raw', async () => {
    const { readXlsx: readLite, writeXlsx: writeLite } = await import('../src/lite.js');
    const wb = readLite(FIXTURE);
    expect(wb.customViews).toBeUndefined();
    const out = writeLite(wb);
    expect(part(out, 'xl/workbook.xml')).toContain('<customWorkbookView name="Small"');
    expect(part(out, 'xl/worksheets/sheet1.xml')).toContain('<customSheetViews>');
  });
});
