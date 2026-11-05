// Incremental recalculation (#106): track(wb).set / update recalculate only what depends on the
// changed cells, and always agree with a full evaluate.

import { describe, expect, it } from 'vitest';
import { evaluate, track } from '../src/formula/index.js';
import { addDefinedName, addTable, Workbook, type Worksheet } from '../src/index.js';

const res = (ws: Worksheet, a: string) => {
  const v = ws.getValue(a);
  return v && typeof v === 'object' && 'formula' in v ? v.result : v;
};

/** Every formula result of the workbook, by address. */
function snapshot(wb: Workbook): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const ws of wb.worksheets) {
    ws.eachRow((row, r) => {
      row.eachCell((v, c) => {
        if (v && typeof v === 'object' && 'formula' in v)
          out[`${ws.name}!${ws.getCell(r, c).address}`] = v.result;
      });
    });
  }
  return out;
}

/** Same edits on a tracked workbook and on a copy evaluated in full: results must match. */
function sameAsFull(build: () => Workbook, edits: [string, string, number | string][]): void {
  const a = build();
  const b = build();
  const calc = track(a);
  for (const [s, addr, v] of edits) {
    calc.set(s, addr, v);
    b.getWorksheet(s)!.setCell(addr, v);
  }
  evaluate(b);
  expect(snapshot(a)).toEqual(snapshot(b));
}

function chain(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 50; r++) {
    ws.setCell(r, 1, r); // A: inputs
    ws.setCell(r, 2, { formula: `A${r}*2` }); // B: per row
    ws.setCell(r, 3, { formula: r === 1 ? 'B1' : `C${r - 1}+B${r}` }); // C: running total
  }
  ws.setCell('E1', { formula: 'SUM(B1:B50)' });
  ws.setCell('E2', { formula: 'MAX(C:C)' });
  ws.setCell('E3', { formula: 'Rate*E1' });
  ws.setCell('F1', 0.5);
  addDefinedName(wb, 'Rate', 'S!$F$1');
  const t = wb.addWorksheet('T');
  t.setCell('A1', { formula: 'S!E1+1' });
  t.setCell('A2', { formula: 'IF(S!A10>100,"big","small")' });
  return wb;
}

describe('incremental recalculation', () => {
  it('recalculates only the dependents of a change', () => {
    const wb = chain();
    const ws = wb.getWorksheet('S')!;
    const calc = track(wb);
    expect(res(ws, 'E1')).toBe(2550);
    const changed = calc.set('S', 'A49', 100);
    expect(new Set(changed)).toEqual(
      new Set(['S!B49', 'S!C49', 'S!C50', 'S!E1', 'S!E2', 'S!E3', 'T!A1']),
    );
    expect(res(ws, 'B49')).toBe(200);
    expect(res(ws, 'E1')).toBe(2550 - 98 + 200);
    expect(res(wb.getWorksheet('T')!, 'A1')).toBe(2653);
    // through a defined name
    expect(calc.set('S', 'F1', 2)).toEqual(['S!E3']);
    expect(res(ws, 'E3')).toBe(2652 * 2);
    // an untouched cell recalculates nothing
    expect(calc.set('S', 'Z99', 1)).toEqual([]);
  });

  it('agrees with a full evaluate after many edits', () => {
    sameAsFull(chain, [
      ['S', 'A1', 7],
      ['S', 'A10', 500],
      ['S', 'F1', 3],
      ['S', 'A50', 'text'],
      ['S', 'A10', 1],
    ]);
  });

  it('tables, volatile functions and cross-sheet references', () => {
    const build = (): Workbook => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      addTable(ws, {
        name: 'T',
        topLeft: 'A1',
        totalsRow: true,
        columns: [
          { name: 'q', totalsRowFunction: 'sum' },
          { name: 'd', calculatedColumnFormula: 'T[[#This Row],[q]]*2' },
        ],
        rows: [[1], [2], [3]],
      });
      ws.setCell('E1', { formula: 'SUM(T[d])' });
      ws.setCell('E2', { formula: 'INDIRECT("A"&F1)' });
      ws.setCell('F1', 3);
      return wb;
    };
    const wb = build();
    const calc = track(wb);
    const ws = wb.getWorksheet('S')!;
    expect(new Set(calc.set('S', 'A3', 10))).toEqual(new Set(['S!B3', 'S!A5', 'S!E1', 'S!E2']));
    expect(res(ws, 'E1')).toBe(28);
    expect(res(ws, 'E2')).toBe(10);
    sameAsFull(build, [
      ['S', 'A2', 5],
      ['S', 'F1', 2],
    ]);
  });

  it('formula edits, cycles and spills fall back to a full recalculation', () => {
    const wb = chain();
    const ws = wb.getWorksheet('S')!;
    const calc = track(wb);
    const all = calc.set('S', 'B1', { formula: 'A1*3' } as never);
    expect(all.length).toBe(calc.report.evaluated);
    expect(res(ws, 'E1')).toBe(2551);
    // dynamic array: the spill area may change
    ws.setCell('H1', { formula: 'SEQUENCE(F2)' });
    ws.setCell('F2', 2);
    calc.recalc();
    calc.set('S', 'F2', 4);
    expect(res(ws, 'H1')).toBe(1);
    expect(ws.getValue('H4')).toBe(4);
    // a cycle
    ws.setCell('J1', { formula: 'J2+A1' });
    ws.setCell('J2', { formula: 'J1' });
    expect(calc.recalc().circular.sort()).toEqual(['S!J1', 'S!J2']);
    expect(calc.set('S', 'A1', 9).length).toBe(calc.report.evaluated);
  });

  it('update() after editing cells directly', () => {
    const wb = chain();
    const ws = wb.getWorksheet('S')!;
    const calc = track(wb);
    ws.setCell('A1', 10);
    ws.setCell('A2', 10);
    calc.update([
      { sheet: 'S', address: 'A1' },
      { sheet: ws, address: 'A2' },
    ]);
    expect(res(ws, 'C2')).toBe(40);
    expect(() => calc.update([{ sheet: 'Nope', address: 'A1' }])).toThrow(RangeError);
  });
});
