// Data validation: the 'data-validations' test sheet (spec/utils/test-data-validation-sheet.js)
// written by spec/integration/workbook/workbook.spec.js "dataValidations" and the streaming writer
// specs. Issues 1027 / 771 / 1842 and pr 1204 are in issues.test.ts / prs.test.ts.

import { describe, expect, it } from 'vitest';
import {
  addDataValidation,
  addDefinedName,
  type DataValidation,
  dateToSerial,
  readXlsx,
  Workbook,
  type Worksheet,
  type WriteOptions,
  writeXlsx,
} from '../../src/index.js';
import { exceljsLoad, streamWrite, utilData } from './helpers.js';

const json = utilData('data-validations');

/** exceljs { formulae: [...] } validation → turbosheet DataValidation at `sqref`. */
function toDv(sqref: string, x: Record<string, unknown>): DataValidation {
  const { formulae, errorStyle, ...rest } = x as { formulae?: unknown[]; errorStyle?: string };
  const dv = { sqref, ...rest } as Record<string, unknown>;
  if (formulae?.[0] !== undefined) dv.formula1 = formulae[0];
  if (formulae?.[1] !== undefined) dv.formula2 = formulae[1];
  // ➖ exceljs 'error' is OOXML 'stop'
  if (errorStyle) dv.errorStyle = errorStyle === 'error' ? 'stop' : errorStyle;
  return dv as unknown as DataValidation;
}

/** What reads back: formulas as text (dates as serials), sqref as written. */
function back(dv: DataValidation): DataValidation {
  const f = (v: unknown) => (v instanceof Date ? String(dateToSerial(v)) : String(v));
  const o = { ...dv } as Record<string, unknown>;
  if (dv.formula1 !== undefined) o.formula1 = f(dv.formula1);
  if (dv.formula2 !== undefined) o.formula2 = f(dv.formula2);
  return o as unknown as DataValidation;
}

function matrix(): DataValidation[] {
  const out: DataValidation[] = [];
  json.types.forEach((type: string, ri: number) => {
    json.operators.forEach((operator: string, ci: number) => {
      const v = json.values[type];
      const formulae = [v.v1];
      if (operator === 'between' || operator === 'notBetween') formulae.push(v.v2);
      const ref = `${String.fromCharCode(67 + ci)}${8 + ri}`;
      out.push(
        toDv(ref, {
          type,
          operator,
          allowBlank: true,
          showInputMessage: true,
          showErrorMessage: true,
          formulae,
        }),
      );
    });
  });
  return out;
}

function expected(): DataValidation[] {
  return [
    toDv('B1', json.B1),
    toDv('B3', json.B3),
    toDv('B5', json.B5),
    ...matrix(),
    toDv('B13', json.B13),
    toDv('E13', json.E13),
    toDv('B15', json.B15),
    toDv('B17', json.B17),
    toDv('B19', json.B19),
    toDv('A20', {
      type: 'date',
      operator: 'greaterThan',
      showErrorMessage: true,
      allowBlank: true,
      formulae: [new Date(Date.UTC(2016, 0, 1))],
    }),
    // ➖ one validation over a range (exceljs: per cell, merged on write)
    toDv('B22:F23', json.shared),
  ];
}

function fill(wb: Workbook, ws: Worksheet): void {
  ws.getCell('D1').value = 'Hewie';
  ws.getCell('E1').value = 'Dewie';
  ws.getCell('F1').value = 'Louie';
  // ➖ exceljs `cell.name = 'Nephews'`; turbosheet: a defined name
  addDefinedName(wb, 'Nephews', `'${ws.name}'!$D$1:$F$1`);
  ws.getCell('D5').value = 'Tom';
  ws.getCell('E5').value = 'Dick';
  ws.getCell('F5').value = 'Harry';
  for (const dv of expected()) addDataValidation(ws, dv);
}

function check(wb: Workbook): void {
  const ws = wb.getWorksheet('data-validations')!;
  expect(ws.dataValidations).toEqual(expected().map(back));
}

describe('exceljs parity: data validation', () => {
  // exceljs: spec/integration/workbook/workbook.spec.js — "dataValidations"
  it('dataValidations test sheet round-trips', async () => {
    const wb = new Workbook();
    fill(wb, wb.addWorksheet('data-validations'));
    const bytes = writeXlsx(wb);
    check(readXlsx(bytes));
    const x = (await exceljsLoad(bytes)).getWorksheet('data-validations')!;
    expect(x.getCell('B3').dataValidation).toEqual(json.B3);
    expect(x.getCell('B13').dataValidation).toEqual(json.B13);
    expect(x.getCell('E23').dataValidation).toEqual(json.shared);
  });

  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "serializes and deserializes dataValidations"
  // exceljs: spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js — "with zip compression option"
  it.each<[string, WriteOptions]>([
    ['default', {}],
    ['compression 9', { compression: 9 }],
  ])('streaming writer: dataValidations (%s)', async (_n, opts) => {
    const bytes = await streamWrite((w) => {
      const s = w.addWorksheet('data-validations');
      s.addRow([null, null, null, 'Hewie', 'Dewie', 'Louie']);
      addDefinedName(w.workbook, 'Nephews', "'data-validations'!$D$1:$F$1");
      s.worksheet.dataValidations = expected();
    }, opts);
    check(readXlsx(bytes));
  });

  // exceljs: spec/unit/doc/workbook.spec.js — date validation of createSimpleWorkbook (D1)
  it('date validations follow the 1904 date system', () => {
    const wb = new Workbook({ date1904: true });
    const ws = wb.addWorksheet('s');
    const d = new Date(Date.UTC(2016, 0, 1));
    addDataValidation(ws, { sqref: 'D1', type: 'date', operator: 'greaterThan', formula1: d });
    const dv = readXlsx(writeXlsx(wb)).getWorksheet('s')!.dataValidations![0]!;
    expect(dv.formula1).toBe(String(dateToSerial(d, true)));
  });
});
