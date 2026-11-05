// Pivot table authoring (#115). The report grids below are what Excel for Mac shows after it
// opens (and refreshes) the written file — checked cell by cell; the files open without repair.

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  addPivotTable,
  type CellValue,
  readXlsx,
  Workbook,
  type Worksheet,
  writeXlsx,
} from '../src/index.js';
import { unzipSync } from '../src/zip/index.js';

const dec = new TextDecoder();
const parts = (buf: Uint8Array) => {
  const m = new Map<string, string>();
  for (const [k, v] of unzipSync(buf)) m.set(k, dec.decode(v));
  return m;
};

function book(): { wb: Workbook; pv: Worksheet } {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['Region', 'Product', 'Qtr', 'Units', 'Price']);
  let k = 0;
  for (const r of ['North', 'South', 'East'])
    for (const p of ['Apples', 'Pears'])
      for (const q of ['Q1', 'Q2']) {
        k++;
        ws.addRow([r, p, q, 10 + k, Math.round((1.5 + k / 10) * 100) / 100]);
      }
  return { wb, pv: wb.addWorksheet('Pivot') };
}

/** Rows of the sheet from `from`, `width` columns, null for empty cells. */
function grid(ws: Worksheet, from: number, to: number, width: number): CellValue[][] {
  const out: CellValue[][] = [];
  for (let r = from; r <= to; r++) {
    const row: CellValue[] = [];
    for (let c = 1; c <= width; c++) row.push(ws.getValue(r, c));
    out.push(row);
  }
  return out;
}

describe('pivot tables', () => {
  it('one row field, one value: report cells and parts', async () => {
    const { wb, pv } = book();
    addPivotTable(pv, {
      name: 'PT',
      source: 'Data!A1:E13',
      ref: 'A3',
      rows: ['Region'],
      values: [{ field: 'Units' }],
    });
    expect(grid(pv, 3, 7, 2)).toEqual([
      ['Row Labels', 'Sum of Units'],
      ['East', 82],
      ['North', 50],
      ['South', 66],
      ['Grand Total', 198],
    ]);
    const out = writeXlsx(wb);
    const z = parts(out);
    const def = z.get('xl/pivotCache/pivotCacheDefinition1.xml')!;
    expect(def).toContain('refreshOnLoad="1"');
    expect(def).toContain('<worksheetSource ref="A1:E13" sheet="Data"/>');
    expect(def).toContain(
      '<cacheField name="Region" numFmtId="0"><sharedItems count="3"><s v="North"/><s v="South"/><s v="East"/></sharedItems></cacheField>',
    );
    expect(def).toContain(
      '<cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="11" maxValue="22"/>',
    );
    expect(z.get('xl/pivotCache/pivotCacheRecords1.xml')).toContain(
      '<r><x v="0"/><x v="0"/><x v="0"/><n v="11"/><n v="1.6"/></r>',
    );
    const t = z.get('xl/pivotTables/pivotTable1.xml')!;
    expect(t).toContain(
      '<location ref="A3:B7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>',
    );
    // items in display order (East, North, South) refer to the shared items
    expect(t).toContain(
      '<pivotField axis="axisRow" showAll="0"><items count="4"><item x="2"/><item x="0"/><item x="1"/><item t="default"/></items></pivotField>',
    );
    expect(t).toContain(
      '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems>',
    );
    expect(t).toContain('<dataField name="Sum of Units" fld="3" baseField="0" baseItem="0"/>');
    expect(z.get('xl/workbook.xml')).toMatch(
      /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/,
    );
    expect(z.get('xl/worksheets/_rels/sheet2.xml.rels')).toContain(
      'Target="../pivotTables/pivotTable1.xml"',
    );
    expect(z.get('xl/pivotTables/_rels/pivotTable1.xml.rels')).toContain(
      'Target="../pivotCache/pivotCacheDefinition1.xml"',
    );
    const ct = z.get('[Content_Types].xml')!;
    for (const k of ['pivotTable', 'pivotCacheDefinition', 'pivotCacheRecords'])
      expect(ct).toContain(`spreadsheetml.${k}+xml`);
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(out as unknown as ArrayBuffer);
    expect(x.getWorksheet('Pivot')!.getCell('B7').value).toBe(198);
  });

  it('nested rows with a column field (Excel compact layout)', () => {
    const { wb, pv } = book();
    addPivotTable(pv, {
      name: 'PT',
      source: 'Data!A1:E13',
      ref: 'A3',
      rows: ['Region', 'Product'],
      columns: ['Qtr'],
      values: [{ field: 'Units' }],
    });
    expect(grid(pv, 3, 14, 4)).toEqual([
      ['Sum of Units', 'Column Labels', null, null],
      ['Row Labels', 'Q1', 'Q2', 'Grand Total'],
      ['East', 40, 42, 82],
      ['Apples', 19, 20, 39],
      ['Pears', 21, 22, 43],
      ['North', 24, 26, 50],
      ['Apples', 11, 12, 23],
      ['Pears', 13, 14, 27],
      ['South', 32, 34, 66],
      ['Apples', 15, 16, 31],
      ['Pears', 17, 18, 35],
      ['Grand Total', 96, 102, 198],
    ]);
    const t = parts(writeXlsx(wb)).get('xl/pivotTables/pivotTable1.xml')!;
    expect(t).toContain('firstHeaderRow="1" firstDataRow="2" firstDataCol="1"');
    expect(t).toContain(
      '<rowItems count="10"><i><x/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i><x v="1"/></i>',
    );
    expect(t).toContain(
      '<colFields count="1"><field x="2"/></colFields><colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>',
    );
  });

  it('several values and a report filter', () => {
    const { wb, pv } = book();
    addPivotTable(pv, {
      name: 'PT',
      source: 'Data!A1:E13',
      ref: 'A1',
      rows: ['Product'],
      values: [{ field: 'Units' }, { field: 'Price', fn: 'average' }],
      filters: ['Region'],
    });
    expect(grid(pv, 1, 7, 3)).toEqual([
      ['Region', '(All)', null],
      [null, null, null],
      [null, 'Values', null],
      ['Row Labels', 'Sum of Units', 'Average of Price'],
      ['Apples', 93, 2.05],
      ['Pears', 105, 2.25],
      ['Grand Total', 198, 2.15],
    ]);
    const t = parts(writeXlsx(wb)).get('xl/pivotTables/pivotTable1.xml')!;
    expect(t).toContain(
      '<location ref="A3:C7" firstHeaderRow="1" firstDataRow="2" firstDataCol="1" rowPageCount="1" colPageCount="1"/>',
    );
    expect(t).toContain('<colFields count="1"><field x="-2"/></colFields>');
    expect(t).toContain('<pageFields count="1"><pageField fld="0" hier="-1"/></pageFields>');
    expect(t).toContain(
      '<dataField name="Average of Price" fld="4" subtotal="average" baseField="0" baseItem="0"/>',
    );
  });

  it('blank items, numeric items and other functions', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['Size', 'Region', 'Units']);
    for (const r of [
      [2, 'N', 5],
      [1, null, 7],
      [2, 'S', 1],
      [1, 'N', 3],
    ])
      ws.addRow(r);
    const pv = wb.addWorksheet('Pivot');
    addPivotTable(pv, {
      name: 'PT',
      source: 'Data!A1:C5',
      ref: 'A1',
      rows: ['Region'],
      columns: ['Size'],
      values: [{ field: 'Units', fn: 'max', name: 'Most' }],
    });
    expect(grid(pv, 1, 6, 4)).toEqual([
      ['Most', 'Column Labels', null, null],
      ['Row Labels', 1, 2, 'Grand Total'],
      ['N', 3, 5, 5],
      ['S', null, 1, 1],
      ['(blank)', 7, null, 7],
      ['Grand Total', 7, 5, 7],
    ]);
    const def = parts(writeXlsx(wb)).get('xl/pivotCache/pivotCacheDefinition1.xml')!;
    expect(def).toContain('containsBlank="1" count="3"><s v="N"/><m/><s v="S"/>');
  });

  it('validation errors', () => {
    const { wb, pv } = book();
    const base = {
      name: 'PT',
      source: 'Data!A1:E13',
      ref: 'A3',
      rows: ['Region'],
      values: [{ field: 'Units' }],
    };
    expect(() => addPivotTable(pv, { ...base, rows: ['Nope'] })).toThrow('no source column "Nope"');
    expect(() => addPivotTable(pv, { ...base, rows: [] })).toThrow('row field');
    expect(() => addPivotTable(pv, { ...base, columns: ['Qtr', 'Product'] })).toThrow(
      'one column field',
    );
    expect(() =>
      addPivotTable(pv, {
        ...base,
        columns: ['Qtr'],
        values: [{ field: 'Units' }, { field: 'Price' }],
      }),
    ).toThrow('not supported yet');
    expect(() => addPivotTable(pv, { ...base, columns: ['Region'] })).toThrow('one axis only');
    pv.setCell('B5', 'x');
    expect(() => addPivotTable(pv, base)).toThrow('B5 is not empty');
    pv.clearCell('B5');
    addPivotTable(pv, base);
    expect(() => addPivotTable(pv, { ...base, ref: 'H3' })).toThrow('used more than once');
    expect(wb.worksheets[1]!.pivotTables).toHaveLength(1);
  });

  it('two pivots, two caches; adding one to a workbook read with pivots throws', () => {
    const { wb, pv } = book();
    addPivotTable(pv, {
      name: 'A',
      source: 'Data!A1:E13',
      ref: 'A3',
      rows: ['Region'],
      values: [{ field: 'Units' }],
    });
    addPivotTable(pv, {
      name: 'B',
      source: 'Data!A1:E13',
      ref: 'E3',
      rows: ['Product'],
      values: [{ field: 'Price' }],
    });
    const out = writeXlsx(wb);
    const z = parts(out);
    expect(z.get('xl/workbook.xml')!.match(/<pivotCache /g)).toHaveLength(2);
    expect(z.get('xl/pivotTables/pivotTable2.xml')).toContain('cacheId="2"');
    const back = readXlsx(out);
    expect(back.worksheets[1]!.pivotTables).toBeUndefined(); // read pivots are kept raw
    expect(parts(writeXlsx(back)).get('xl/pivotTables/pivotTable1.xml')).toBe(
      z.get('xl/pivotTables/pivotTable1.xml'),
    );
    addPivotTable(back.worksheets[1]!, {
      name: 'C',
      source: 'Data!A1:E13',
      ref: 'J3',
      rows: ['Qtr'],
      values: [{ field: 'Units' }],
    });
    expect(() => writeXlsx(back)).toThrow('not supported yet');
  });

  it('lite: needs pivotTablesWriter', async () => {
    const lite = await import('../src/lite.js');
    const { wb, pv } = book();
    lite.addPivotTable(pv, {
      name: 'PT',
      source: 'Data!A1:E13',
      ref: 'A3',
      rows: ['Region'],
      values: [{ field: 'Units' }],
    });
    expect(() => lite.writeXlsx(wb)).toThrow(/pivotTablesWriter/);
    expect(
      parts(lite.writeXlsx(wb, { features: [lite.pivotTablesWriter] })).has(
        'xl/pivotTables/pivotTable1.xml',
      ),
    ).toBe(true);
  });
});
