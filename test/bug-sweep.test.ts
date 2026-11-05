/// <reference types="node" />
// Reported ExcelJS bugs checked against turbosheet (issue #22), plus local-time dates (#23)
// and merged cells in streaming read (#24). Each `exceljs#N` names the upstream report.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addNote,
  addTable,
  createXlsxWriter,
  dateToSerial,
  groupRows,
  readXlsx,
  serialToDate,
  Workbook,
  writeXlsx,
  type XlsxRow,
  xlsxRows,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const sheetXml = (buf: Uint8Array): string =>
  dec.decode(unzipSync(buf).get('xl/worksheets/sheet1.xml')!);

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

describe('bug sweep (exceljs reports)', () => {
  it('exceljs#1325: every header/footer option is written and read back', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', 1);
    ws.headerFooter = {
      differentFirst: true,
      differentOddEven: true,
      oddHeader: '&Codd h',
      oddFooter: '&Codd f',
      evenHeader: '&Ceven h',
      evenFooter: '&Ceven f',
      firstHeader: '&Cfirst h',
      firstFooter: '&Cfirst f',
    };
    const buf = writeXlsx(wb);
    expect(sheetXml(buf)).toContain(
      '<headerFooter differentOddEven="1" differentFirst="1"><oddHeader>&amp;Codd h</oddHeader><oddFooter>&amp;Codd f</oddFooter><evenHeader>&amp;Ceven h</evenHeader><evenFooter>&amp;Ceven f</evenFooter><firstHeader>&amp;Cfirst h</firstHeader><firstFooter>&amp;Cfirst f</firstFooter></headerFooter>',
    );
    expect(readXlsx(buf).getWorksheet('S')!.headerFooter).toEqual(ws.headerFooter);
  });

  it('exceljs#1348: fitToPage + outline + tab color keep CT_SheetPr child order', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S', { tabColor: { argb: 'FFFF0000' } });
    for (let r = 1; r <= 5; r++) ws.addRow([r]);
    groupRows(ws, 2, 4, { collapsed: true });
    ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    const pr = sheetXml(writeXlsx(wb)).match(/<sheetPr>.*?<\/sheetPr>/)![0];
    expect(pr.indexOf('<tabColor')).toBeLessThan(pr.indexOf('<pageSetUpPr'));
    expect(pr).toBe('<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>');
  });

  it('exceljs#1369: a note on a cell inside a table', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addTable(ws, {
      name: 'T',
      columns: ['a', 'b'],
      rows: [
        [1, 2],
        [3, 4],
      ],
    });
    addNote(ws, 'B2', 'inside the table');
    addNote(ws, 'A1', 'on the header');
    const buf = writeXlsx(wb);
    const back = readXlsx(buf).getWorksheet('S')!;
    expect(back.comments!.map((n) => n.ref).sort()).toEqual(['A1', 'B2']);
    expect(back.tables![0]!.ref).toBe('A1:B3');
    const ex = await oracle(buf);
    expect(ex.getWorksheet('S')!.getCell('B2').note).toBeDefined();
  });

  it('exceljs#1430: streaming read returns Dates for date-formatted cells', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const d = new Date(Date.UTC(2024, 4, 17, 13, 30));
    ws.addRow([d, 45000]);
    ws.setStyle('B1', { numFmt: 'dd/mm/yyyy' });
    const rows: XlsxRow[] = [];
    for await (const r of xlsxRows(writeXlsx(wb))) rows.push(r);
    expect(rows[0]!.cells[0]).toEqual(d);
    expect(rows[0]!.cells[1]).toEqual(serialToDate(45000));
  });

  it('exceljs#1615: a table without a header row writes no autoFilter', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    addTable(ws, { name: 'NoHeader', columns: ['a', 'b'], rows: [[1, 2]], headerRow: false });
    const buf = writeXlsx(wb);
    const t = dec.decode(unzipSync(buf).get('xl/tables/table1.xml')!);
    expect(t).not.toContain('<autoFilter');
    expect(t).toContain('headerRowCount="0"');
    await oracle(buf);
  });

  it('exceljs#2267: the streaming writer keeps each rich-text cell its own', async () => {
    const chunks: Uint8Array[] = [];
    const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
    const s = w.addWorksheet('S');
    for (let i = 0; i < 5; i++) {
      s.addRow([
        { richText: [{ text: `bold ${i}`, font: { bold: true } }, { text: ` plain ${i}` }] },
      ]);
    }
    await w.end();
    const out = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    const back = readXlsx(out).getWorksheet('S')!;
    for (let i = 0; i < 5; i++) {
      expect(back.getValue(i + 1, 1)).toEqual({
        richText: [
          { text: `bold ${i}`, font: expect.objectContaining({ bold: true }) },
          { text: ` plain ${i}` },
        ],
      });
    }
  });

  it('exceljs#1988: a hyperlink cell whose text is rich text keeps text, link and formatting', async () => {
    const ex = new ExcelJS.Workbook();
    const ews = ex.addWorksheet('S');
    ews.getCell('A1').value = {
      text: { richText: [{ text: 'rich ', font: { bold: true } }, { text: 'link' }] },
      hyperlink: 'https://example.com',
    } as unknown as ExcelJS.CellValue;
    const buf = new Uint8Array(await ex.xlsx.writeBuffer());
    const v = readXlsx(buf).getWorksheet('S')!.getValue('A1') as {
      text: string;
      hyperlink: string;
    };
    expect(v.hyperlink).toBe('https://example.com');
    expect(v.text).toBe('rich link');
    // #34: the runs' formatting is kept (Hyperlink.runs) and written back
    const runs = [
      { text: 'rich ', font: expect.objectContaining({ bold: true }) },
      { text: 'link' },
    ];
    expect(v).toMatchObject({ runs });
    const again = readXlsx(writeXlsx(readXlsx(buf)))
      .getWorksheet('S')!
      .getValue('A1');
    expect(again).toMatchObject({ text: 'rich link', hyperlink: 'https://example.com', runs });
    // set by hand
    const wb = new Workbook();
    wb.addWorksheet('T').setCell('A1', {
      text: 'ab',
      hyperlink: 'https://example.com',
      runs: [{ text: 'a', font: { italic: true } }, { text: 'b' }],
    });
    expect(readXlsx(writeXlsx(wb)).getWorksheet('T')!.getValue('A1')).toMatchObject({
      text: 'ab',
      runs: [{ text: 'a', font: expect.objectContaining({ italic: true }) }, { text: 'b' }],
    });
  });
});

describe('local-time dates (localDates)', () => {
  it('converts by the local wall clock in setCell / getValue / readers', async () => {
    const local = new Date(2024, 0, 15, 9, 45);
    const serial = 45306 + (9 * 60 + 45) / 1440;
    expect(dateToSerial(local, false, true)).toBeCloseTo(serial, 9);
    expect(serialToDate(serial, false, true)).toEqual(local);
    expect(serialToDate(1, false, true)).toEqual(new Date(1900, 0, 1));

    const wb = new Workbook({ localDates: true });
    const ws = wb.addWorksheet('S');
    ws.setCell('A1', local);
    ws.addRow([new Date(2024, 1, 29)]);
    expect(ws.rowData[0]!.num[0]).toBeCloseTo(serial, 9);
    expect(ws.getValue('A1')).toEqual(local);
    const buf = writeXlsx(wb);
    expect(readXlsx(buf, { localDates: true }).getWorksheet('S')!.getValue('A2')).toEqual(
      new Date(2024, 1, 29),
    );
    const utc = readXlsx(buf).getWorksheet('S')!.getValue('A2');
    expect(utc).toEqual(new Date(Date.UTC(2024, 1, 29)));
    for await (const r of xlsxRows(buf, { localDates: true })) {
      if (r.row === 1) expect(r.cells[0]).toEqual(local);
    }
  });
});

describe('streaming read: merged cells (onSheetEnd)', () => {
  it('reports each sheet’s merges after its rows', async () => {
    const wb = new Workbook();
    const a = wb.addWorksheet('A');
    a.addRow([1, 2, 3]);
    a.mergeCells('A1:B1');
    a.mergeCells('C2:C4');
    wb.addWorksheet('B').addRow(['x']);
    const events: string[] = [];
    for await (const r of xlsxRows(writeXlsx(wb), {
      onSheetEnd: (s) => events.push(`end ${s.sheet} ${s.sheetIndex} ${s.merges.join(',')}`),
    })) {
      events.push(`row ${r.sheet} ${r.row}`);
    }
    expect(events).toEqual(['row A 1', 'end A 0 A1:B1,C2:C4', 'row B 1', 'end B 1 ']);
  });
});
