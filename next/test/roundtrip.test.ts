import { describe, expect, it } from 'vitest';
import {
  type CellValue,
  createXlsxWriter,
  readXlsx,
  type Style,
  Workbook,
  type WriteOptions,
  writeXlsx,
  xlsxRows,
  xlsxStream,
} from '../src/index.js';

const bold: Style = { font: { bold: true, color: { argb: 'FFFF0000' } }, numFmt: '0.00' };

function sample(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.columns = [
    { header: 'id', key: 'id', width: 8 },
    { header: 'name', key: 'name', width: 20 },
    { header: 'when', key: 'when', width: 14 },
    { header: 'amount', key: 'amount' },
  ];
  ws.addRow({ id: 1, name: 'alpha <&> "q"', when: new Date(Date.UTC(2024, 0, 31)), amount: 12.5 });
  ws.addRow({
    id: 2,
    name: '  padded  ',
    when: new Date(Date.UTC(2024, 1, 29, 13, 45)),
    amount: -3,
  });
  ws.addRow({ id: 3, name: 'emoji 🚀', when: null, amount: { formula: 'D2+D3', result: 9.5 } });
  ws.setCell('E2', true);
  ws.setCell('E3', { error: '#N/A' });
  ws.setCell('F2', { text: 'site', hyperlink: 'https://example.com' });
  ws.setCell('G2', 42, bold);
  ws.mergeCells('H2:I3');
  ws.freeze(1, 0);
  const hidden = wb.addWorksheet('Hidden', { state: 'hidden' });
  hidden.addRow(['x', 1]);
  return wb;
}

function snapshot(wb: Workbook): Record<string, CellValue[][]> {
  const out: Record<string, CellValue[][]> = {};
  for (const ws of wb.worksheets) {
    const rows: CellValue[][] = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row: CellValue[] = [];
      for (let c = 1; c <= ws.columnCount; c++) row.push(ws.getValue(r, c) ?? null);
      rows.push(row);
    }
    out[ws.name] = rows;
  }
  return out;
}

async function collect(rs: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  for await (const c of rs as unknown as AsyncIterable<Uint8Array>) parts.push(c);
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

const profiles: WriteOptions[] = [
  {},
  { profile: 'low-memory' },
  { profile: 'max-speed' },
  { compression: 0 },
  { strings: 'inline' },
];

describe('round-trip: our writer → our reader', () => {
  for (const opts of profiles) {
    it(`writeXlsx ${JSON.stringify(opts)}`, () => {
      const wb = sample();
      const back = readXlsx(writeXlsx(wb, opts));
      expect(snapshot(back)).toEqual(snapshot(wb));
      const ws = back.getWorksheet('Data')!;
      expect(back.getWorksheet('Hidden')!.state).toBe('hidden');
      expect(ws.merges).toEqual(['H2:I3']);
      expect(ws.getColumn(2).width).toBe(20);
      expect(ws.getStyle('G2').font?.bold).toBe(true);
      expect(ws.getStyle('G2').numFmt).toBe('0.00');
    });
  }

  it('xlsxStream output equals writeXlsx content', async () => {
    const wb = sample();
    const back = readXlsx(await collect(xlsxStream(wb)));
    expect(snapshot(back)).toEqual(snapshot(wb));
  });

  it('streaming writer → streaming reader, 50k rows', async () => {
    const w = createXlsxWriter({ profile: 'low-memory' });
    const bytesP = collect(w.readable!);
    const s = w.addWorksheet('Big');
    for (let i = 0; i < 50_000; i++) s.addRow([i, `row ${i}`, i % 2 === 0]);
    await s.commit();
    await w.end();
    const bytes = await bytesP;
    let n = 0;
    for await (const row of xlsxRows(bytes)) {
      if (n === 0 || n === 49_999) expect(row.values).toEqual([n, `row ${n}`, n % 2 === 0]);
      n++;
    }
    expect(n).toBe(50_000);
  });
});
