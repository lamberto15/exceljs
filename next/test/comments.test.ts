/// <reference types="node" />
// Cell notes (features/comments): comments part + legacy VML drawing, checked against exceljs.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/index.js';
import { addNote, type Note } from '../src/xlsx/features/index.js';
import { readXlsx } from '../src/xlsx/read/index.js';
import { createXlsxWriter, writeXlsx, xlsxStream } from '../src/xlsx/write/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
function parts(buf: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) out.set(k, dec.decode(v));
  return out;
}

async function oracle(buf: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // biome-ignore lint/suspicious/noExplicitAny: exceljs typings want a Node Buffer
  await wb.xlsx.load(buf as any);
  return wb;
}

const RICH = {
  richText: [
    { text: 'Bold ', font: { bold: true, size: 9, color: { argb: 'FFFF0000' }, name: 'Tahoma' } },
    { text: 'plain <&>' },
  ],
};

function build(wb: Workbook): void {
  const ws = wb.addWorksheet('Notes');
  // exceljs only attaches notes to cells that have a value
  for (const a of ['A1', 'B2', 'C3', 'D4']) ws.setCell(a, 'x');
  addNote(ws, 'B2', 'Hello note');
  addNote(ws, 'c3', { text: RICH, author: 'Ann', visible: true, width: 200, height: 100 });
  addNote(ws, 'D4', { text: 'second line\nhere', author: 'Bob' });
  const ws2 = wb.addWorksheet('Two');
  ws2.setCell('A1', 1);
  addNote(ws2, 'A1', { text: 'on sheet two', author: 'Ann' });
}

const EXPECT_NOTES: Note[] = [
  { ref: 'B2', text: 'Hello note' },
  { ref: 'C3', text: RICH, author: 'Ann', visible: true, width: 200, height: 100 },
  { ref: 'D4', text: 'second line\nhere', author: 'Bob' },
];

async function allWriters(): Promise<[string, Uint8Array][]> {
  const wb = new Workbook();
  build(wb);
  const chunks: Uint8Array[] = [];
  const w = createXlsxWriter({ sink: (c) => void chunks.push(c) });
  for (const src of wb.worksheets) {
    const s = w.addWorksheet(src.name);
    src.eachRow((row) => s.addRow(row.values.slice(1)));
    s.worksheet.comments = src.comments; // set before commit
    await s.commit();
  }
  await w.end();
  return [
    ['writeXlsx', writeXlsx(wb)],
    ['xlsxStream', new Uint8Array(await new Response(xlsxStream(wb)).arrayBuffer())],
    ['createXlsxWriter', new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())],
  ];
}

describe('comments', () => {
  it('addNote normalizes the address and replaces an existing note', () => {
    const ws = new Workbook().addWorksheet('S');
    addNote(ws, '$b$2', 'one');
    addNote(ws, 'B2', { text: 'two', author: 'Z' });
    addNote(ws, 'A1', RICH);
    expect(ws.comments).toEqual([
      { ref: 'B2', text: 'two', author: 'Z' },
      { ref: 'A1', text: RICH },
    ]);
  });

  it('writes comments, VML, rels, legacyDrawing and content types', () => {
    const wb = new Workbook();
    build(wb);
    const p = parts(writeXlsx(wb));
    const c1 = p.get('xl/comments1.xml')!;
    expect(c1).toContain('<authors><author></author><author>Ann</author><author>Bob</author>');
    expect(c1).toContain('<comment ref="B2" authorId="0"><text><r><t>Hello note</t></r></text>');
    expect(c1).toContain('<r><rPr><b/><sz val="9"/><color rgb="FFFF0000"/><rFont val="Tahoma"/>');
    expect(c1).toContain('plain &lt;&amp;&gt;');
    expect(p.get('xl/comments2.xml')).toContain('<comment ref="A1" authorId="0">');
    const vml = p.get('xl/drawings/vmlDrawing1.vml')!;
    expect(vml).toContain('<o:idmap v:ext="edit" data="1"/>');
    expect(vml).toContain('<x:ClientData ObjectType="Note">');
    expect(vml).toContain('<x:Row>2</x:Row><x:Column>2</x:Column><x:Visible/>');
    expect(vml).toContain('width:150pt;height:75pt');
    // shape ids don't collide across sheets
    expect(p.get('xl/drawings/vmlDrawing2.vml')).toContain('data="2"');
    const sheet = p.get('xl/worksheets/sheet1.xml')!;
    expect(sheet).toMatch(/<pageMargins[^>]*\/><legacyDrawing r:id="rId2"\/><\/worksheet>/);
    const rels = p.get('xl/worksheets/_rels/sheet1.xml.rels')!;
    expect(rels).toContain(
      'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"',
    );
    expect(rels).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"',
    );
    const ct = p.get('[Content_Types].xml')!;
    expect(ct).toContain(
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
    );
    expect(ct).toContain(
      '<Override PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
    );
  });

  it('derives the VML anchor from actual column widths and row heights', () => {
    const anchorOf = (setup: (ws: ReturnType<Workbook['addWorksheet']>) => void): string => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.setCell('B2', 'x');
      setup(ws);
      addNote(ws, 'B2', { text: 'n', width: 200, height: 100 });
      const vml = parts(writeXlsx(wb)).get('xl/drawings/vmlDrawing1.vml') as string;
      return /<x:Anchor>([^<]*)<\/x:Anchor>/.exec(vml)![1]!;
    };
    // defaults: 64 x 20 px cells (15 + 200 = 215 = 3 cols + 23; 10 + 100 = 110 = 5 rows + 10)
    expect(anchorOf(() => {})).toBe('2, 15, 0, 10, 5, 23, 5, 10');
    // column C = 20 chars = 140 px, row 1 = 30pt = 40 px
    expect(
      anchorOf((ws) => {
        ws.getColumn('C').width = 20;
        ws.setRowProps(1, { height: 30 });
      }),
    ).toBe('2, 15, 0, 10, 4, 11, 4, 10');
    // hidden columns / rows take no room
    expect(
      anchorOf((ws) => {
        ws.getColumn('C').hidden = true;
        ws.setRowProps(1, { hidden: true });
      }),
    ).toBe('2, 15, 0, 10, 6, 23, 6, 10');
  });

  it('positions the VML box with the actual widths (margin-left/top)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.setCell('C3', 'x');
    ws.getColumn('A').width = 20; // 140 px
    ws.setRowProps(1, { height: 30 }); // 40 px
    addNote(ws, 'C3', 'hi');
    const vml = parts(writeXlsx(wb)).get('xl/drawings/vmlDrawing1.vml') as string;
    // left = A(140) + B(64) + C(64) + 15 = 283 px; top = row 1 (40 px) + 10 = 50 px (box sits a row above)
    expect(vml).toContain('margin-left:212.25pt;margin-top:37.5pt');
  });

  it('round-trips through ours in every writer', async () => {
    for (const [name, buf] of await allWriters()) {
      const wb = readXlsx(buf);
      expect(wb.getWorksheet('Notes')!.comments, name).toEqual(EXPECT_NOTES);
      expect(wb.getWorksheet('Two')!.comments, name).toEqual([
        { ref: 'A1', text: 'on sheet two', author: 'Ann' },
      ]);
      // and written again unchanged
      const again = readXlsx(writeXlsx(wb));
      expect(again.getWorksheet('Notes')!.comments, name).toEqual(EXPECT_NOTES);
    }
  });

  it('exceljs reads our notes', async () => {
    for (const [name, buf] of await allWriters()) {
      const x = await oracle(buf);
      const xs = x.getWorksheet('Notes')!;
      expect(xs.getCell('B2').note, name).toBe('Hello note');
      const rich = xs.getCell('C3').note as ExcelJS.Comment;
      expect(rich.texts, name).toEqual([
        {
          text: 'Bold ',
          font: { bold: true, size: 9, color: { argb: 'FFFF0000' }, name: 'Tahoma' },
        },
        { text: 'plain <&>' },
      ]);
      expect(xs.getCell('D4').note, name).toBe('second line\nhere');
      expect(x.getWorksheet('Two')!.getCell('A1').note, name).toBe('on sheet two');
    }
  });

  it('reads exceljs notes', async () => {
    const x = new ExcelJS.Workbook();
    const xs = x.addWorksheet('S');
    xs.getCell('A1').value = 1;
    xs.getCell('A1').note = 'simple';
    xs.getCell('E5').note = {
      texts: [{ font: { bold: true, size: 12, name: 'Calibri' }, text: 'B' }, { text: 'rest' }],
    } as ExcelJS.Comment;
    const ws = readXlsx(new Uint8Array(await x.xlsx.writeBuffer())).getWorksheet('S')!;
    expect(ws.comments!.map((n) => [n.ref, n.text, n.visible])).toEqual([
      ['A1', 'simple', undefined],
      [
        'E5',
        {
          richText: [
            { text: 'B', font: { bold: true, size: 12, name: 'Calibri' } },
            { text: 'rest' },
          ],
        },
        undefined,
      ],
    ]);
  });

  it('tolerates a missing VML part and ignores threaded comments', () => {
    const wb = new Workbook();
    build(wb);
    const files = unzipSync(writeXlsx(wb));
    files.delete('xl/drawings/vmlDrawing1.vml');
    // a threaded-comments rel + part as Excel 365 writes them
    const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    files.set(
      relsPath,
      enc.encode(
        dec
          .decode(files.get(relsPath)!)
          .replace(
            '</Relationships>',
            '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
          ),
      ),
    );
    files.set(
      'xl/threadedComments/threadedComment1.xml',
      enc.encode(
        '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B2" id="{1}"><text>hi</text></threadedComment></ThreadedComments>',
      ),
    );
    const ws = readXlsx(
      zipSync(
        [...files].map(([path, data]) => ({ path, data })),
        6,
      ),
    ).getWorksheet('Notes')!;
    expect(ws.comments!.map((n) => n.ref)).toEqual(['B2', 'C3', 'D4']);
    expect(ws.comments![1]!.visible).toBeUndefined(); // came from VML
  });

  it('many notes get several VML id blocks', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 1100; r++) addNote(ws, `A${r}`, `n${r}`);
    wb.addWorksheet('T').comments = [{ ref: 'A1', text: 't' }];
    const p = parts(writeXlsx(wb));
    expect(p.get('xl/drawings/vmlDrawing1.vml')).toContain('data="1,2"');
    expect(p.get('xl/drawings/vmlDrawing1.vml')).toContain('id="_x0000_s2125"');
    expect(p.get('xl/drawings/vmlDrawing2.vml')).toContain('data="3"');
    expect(readXlsx(writeXlsx(wb)).getWorksheet('S')!.comments!.length).toBe(1100);
  });
});
