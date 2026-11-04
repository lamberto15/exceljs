// Conditional formatting: spec/utils/test-conditional-formatting-sheet.js written by
// spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js.

import { describe, expect, it } from 'vitest';
import {
  addConditionalFormatting,
  type CfRule,
  type ConditionalFormatting,
  readXlsx,
  Workbook,
  writeXlsx,
} from '../../src/index.js';
import { exceljsLoad, partText, streamWrite, utilData } from './helpers.js';

const json = utilData('conditional-formatting');
const blocks = (): ConditionalFormatting[] =>
  (json.types as string[]).filter((t) => json[t]).map((t) => json[t] as ConditionalFormatting);

function check(cf: ConditionalFormatting[] | undefined): void {
  expect(cf).toBeDefined();
  expect(cf!.length).toBe(blocks().length);
  for (const item of cf!) {
    const want = json[item.rules[0]!.type] as ConditionalFormatting;
    expect(item.ref).toBe(want.ref);
    expect(item.rules.length).toBe(want.rules.length);
    item.rules.forEach((r, i) => {
      expect(r).toMatchObject({ type: want.rules[i]!.type });
    });
  }
}

describe('exceljs parity: conditional formatting', () => {
  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "with conditional formatting"
  it('streaming writer: conditional formatting sheet round-trips', async () => {
    const bytes = await streamWrite(
      (w) => {
        const s = w.addWorksheet('conditional-formatting');
        for (const b of blocks()) addConditionalFormatting(s.worksheet, b.ref, b.rules);
      },
      { strings: 'shared' },
    );
    const ws = readXlsx(bytes).getWorksheet('conditional-formatting')!;
    check(ws.conditionalFormatting);
    // rule details survive (styles become dxfs)
    const expr = ws.conditionalFormatting![0]!.rules[0]!;
    expect(expr).toMatchObject({
      type: 'expression',
      formulae: ['MOD(ROW()+COLUMN(),2)=0'],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } },
    });
    const x = (await exceljsLoad(bytes)).getWorksheet('conditional-formatting')!;
    expect(
      (x as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings,
    ).toHaveLength(blocks().length);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "with conditional formatting that contains numFmt (#1814)"
  it('conditional formatting styles with a numFmt (#1814)', async () => {
    const cf = json.abbreviation as { ref: string; rules: CfRule[] };
    const wb = new Workbook();
    const ws = wb.addWorksheet('conditionalFormatting');
    addConditionalFormatting(ws, cf.ref, cf.rules);
    const bytes = writeXlsx(wb);
    const styles = partText(bytes, 'xl/styles.xml')!;
    // every dxf numFmt has a numeric id and a format code
    const dxfs = /<dxfs[^>]*>(.*)<\/dxfs>/s.exec(styles)![1]!;
    expect([...dxfs.matchAll(/<numFmt numFmtId="(\d+)" formatCode="[^"]+"\/>/g)]).toHaveLength(2);
    const [cf2] = readXlsx(bytes).getWorksheet('conditionalFormatting')!.conditionalFormatting!;
    expect(cf2!.ref).toBe('A:A');
    for (const r of cf2!.rules) {
      // ➖ numFmt is the format code (exceljs: { id, formatCode })
      expect(typeof (r as { style?: { numFmt?: string } }).style?.numFmt).toBe('string');
    }
    const xr = (
      (await exceljsLoad(bytes)).getWorksheet('conditionalFormatting') as unknown as {
        conditionalFormattings: { rules: { style: { numFmt: { id: number } } }[] }[];
      }
    ).conditionalFormattings[0]!.rules;
    for (const r of xr) expect(typeof r.style.numFmt.id).toBe('number');
  });
});
