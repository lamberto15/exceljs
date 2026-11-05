// Reading charts into the model (#113): supported charts become ws.charts; unchanged ones are
// written back byte for byte, edited ones are regenerated; unsupported ones stay raw.

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { addChart, type Chart, readXlsx, Workbook, writeXlsx } from '../src/index.js';
import { unzipSync, zipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
const parts = (buf: Uint8Array) => {
  const m = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) m.set(k, dec.decode(v));
  return m;
};
// saved by Excel for Mac: a column + line combo on a secondary axis, a line chart with a
// trendline and data labels, a scatter with an exponential trendline
const EXCEL = new Uint8Array(
  readFileSync(new URL('./fixtures/excel-saved-charts.xlsx', import.meta.url)),
);
// written by Excel (ExcelJS test data): a chart on a chart sheet
const CHART_SHEET = new Uint8Array(
  readFileSync(new URL('./fixtures/exceljs/data/chart-sheet.xlsx', import.meta.url)),
);

function data(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['Month', 'Units', 'Share', 'X', 'Y']);
  for (const [i, m] of ['Jan', 'Feb', 'Mar', 'Apr'].entries())
    ws.addRow([m, 10 + i, 0.1 * i, i, i * i]);
  return wb;
}

describe('reading charts', () => {
  it('Excel-saved charts: modelled (combo, secondary axis, trendlines, labels)', () => {
    const ws = readXlsx(EXCEL).worksheets[0]!;
    const [combo, trend, exp] = ws.charts!;
    expect(combo).toMatchObject({
      type: 'column',
      title: 'Combo',
      name: 'Chart 1',
      anchor: { from: { col: 6, row: 1 }, to: { col: 14, row: 16 } },
      axes: { y: { title: 'Units' }, y2: { title: 'Share', numFmt: '0%' } },
    });
    expect(combo!.series[1]).toMatchObject({
      type: 'line',
      axis: 'secondary',
      values: 'Data!$C$2:$C$7',
    });
    expect(trend!.series[0]).toMatchObject({
      trendline: { type: 'linear', forward: 1, displayEquation: true, displayRSquared: true },
      dataLabels: { showValue: true, position: 't' },
    });
    expect(exp!.series[0]!.trendline).toEqual({ type: 'exponential', displayEquation: true });
  });

  it('unchanged read charts are written back byte for byte; one anchor each', () => {
    const src = parts(EXCEL);
    const out = parts(writeXlsx(readXlsx(EXCEL)));
    for (const n of [1, 2, 3]) {
      const name = `xl/charts/chart${n}.xml`;
      expect(out.get(name), name).toBe(src.get(name));
    }
    expect(out.get('xl/drawings/drawing1.xml')!.match(/<xdr:graphicFrame /g)).toHaveLength(3);
  });

  it('editing a read chart regenerates only that chart', () => {
    const wb = readXlsx(EXCEL);
    const src = parts(EXCEL);
    wb.worksheets[0]!.charts![1]!.title = 'Edited title';
    const out = parts(writeXlsx(wb));
    const all = [1, 2, 3].map((n) => out.get(`xl/charts/chart${n}.xml`)!);
    const edited = all.find((x) => x.includes('<a:t>Edited title</a:t>'))!;
    expect(edited).toContain('<c:trendline>');
    expect(
      all.filter(
        (x) => x === src.get('xl/charts/chart1.xml') || x === src.get('xl/charts/chart3.xml'),
      ),
    ).toHaveLength(2);
    const back = readXlsx(writeXlsx(wb)).worksheets[0]!.charts!;
    expect(back.map((c) => c.title)).toEqual(['Combo', 'Edited title', 'Exp']);
  });

  it('chart sheets stay raw (round-trip unchanged)', () => {
    const wb = readXlsx(CHART_SHEET);
    expect(wb.worksheets.every((w) => !w.charts)).toBe(true);
    expect(parts(writeXlsx(wb)).get('xl/charts/chart1.xml')).toBe(
      parts(CHART_SHEET).get('xl/charts/chart1.xml'),
    );
  });

  it('our charts read back to the same model (combo, secondary axis, trendlines, labels)', () => {
    const wb = data();
    const ws = wb.worksheets[0]!;
    const charts: Chart[] = [
      {
        type: 'column',
        title: 'Combo',
        anchor: { from: { col: 6, row: 1 }, to: { col: 13, row: 15 } },
        series: [
          {
            name: { ref: 'Data!$B$1' },
            categories: 'Data!$A$2:$A$5',
            values: 'Data!$B$2:$B$5',
            color: { argb: 'FF4472C4' },
          },
          {
            name: 'Share',
            categories: 'Data!$A$2:$A$5',
            values: 'Data!$C$2:$C$5',
            type: 'line',
            axis: 'secondary',
            marker: false,
            dataLabels: { showValue: true, position: 't', numFmt: '0%' },
          },
        ],
        axes: {
          x: { title: 'Month' },
          y: { title: 'Units', min: 0, max: 20 },
          y2: { title: 'Share', numFmt: '0%' },
        },
        legend: { position: 'b' },
        stacked: undefined,
        style: 3,
        name: 'Chart 1',
      },
      {
        type: 'scatter',
        anchor: { from: { col: 6, row: 17 }, to: { col: 13, row: 30 } },
        series: [
          {
            name: 'y',
            xValues: 'Data!$D$2:$D$5',
            values: 'Data!$E$2:$E$5',
            line: true,
            smooth: true,
            trendline: [
              { type: 'polynomial', order: 2, displayEquation: true },
              { type: 'linear', forward: 1 },
            ],
          },
        ],
        legend: false,
        name: 'Chart 2',
      },
    ];
    for (const c of charts) addChart(ws, c);
    const back = readXlsx(writeXlsx(wb)).worksheets[0]!.charts!;
    const clean = (c: Chart) => JSON.parse(JSON.stringify(c));
    expect(back.map(clean)).toEqual(charts.map(clean));
  });

  it('an unsupported chart (3-D) stays raw and round-trips unchanged', () => {
    const wb = data();
    addChart(wb.worksheets[0]!, {
      type: 'column',
      series: [{ values: 'B2:B5' }],
      anchor: 'G2:M15',
    });
    const z = unzipSync(writeXlsx(wb));
    const name = 'xl/charts/chart1.xml';
    const as3d = dec
      .decode(z.get(name)!)
      .replace(/<c:barChart>/, '<c:bar3DChart>')
      .replace(/<\/c:barChart>/, '<c:shape val="box"/></c:bar3DChart>');
    z.set(name, enc.encode(as3d));
    const src = zipSync(
      [...z].map(([path, d]) => ({ path, data: d })),
      6,
    );
    const read = readXlsx(src);
    expect(read.worksheets[0]!.charts).toBeUndefined();
    expect(parts(writeXlsx(read)).get(name)).toBe(as3d);
  });
});
