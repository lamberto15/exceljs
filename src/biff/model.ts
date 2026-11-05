// Cell-level model glue shared by the binary readers: cached values → cell kinds, formula cells
// (incl. shared / array formulas announced by a PtgExp token) and the read options.

import type { StyleRegistry } from '../core/styles.js';
import type { CellValue, ErrorCode, Formula, Hyperlink, Limits, RichText } from '../core/types.js';
import { CellKind, type Worksheet } from '../core/worksheet.js';
import { decodeFormula, type FormulaContext } from './formula.js';

/** Options of `readXlsb` / `readXls`. */
export interface BinaryReadOptions {
  /**
   * Size limits. `.xlsb`: the zip-bomb limits of readXlsx. `.xls`: `maxEntryBytes` caps each
   * stream of the compound file.
   */
  limits?: Partial<Limits> | undefined;
  /** Only load these sheet names (case-insensitive); others are not created. */
  sheets?: readonly string[] | undefined;
  /** Dates by local wall clock (sets `workbook.localDates`, see WorkbookOptions). */
  localDates?: boolean | undefined;
}

/** A shared or array formula definition, keyed by its master cell. */
export interface FormulaDef {
  readonly rgce: Uint8Array;
  readonly rgcb: Uint8Array;
  /** Array formula range ("A1:B2"); undefined for a shared formula. */
  readonly array: string | undefined;
}

/** A formula cell whose tokens are a PtgExp: resolved once the sheet is read. */
interface PendingExp {
  r: number;
  c: number;
  key: number;
  kind: number;
  num: number;
  obj: unknown;
  xf: number;
}

/** Writes decoded cells into a worksheet (the shared tail of both binary readers). */
export class CellWriter {
  readonly defs: Map<number, FormulaDef> = new Map();
  private pending: PendingExp[] = [];

  constructor(
    readonly ws: Worksheet,
    private readonly styles: StyleRegistry,
    private readonly ctx: FormulaContext,
  ) {}

  /** A number cell: Date kind when its format is a date format (readXlsx semantics). */
  number(r: number, c: number, v: number, xf: number): void {
    const date = xf !== 0 && this.styles.isDateXf(xf);
    this.ws.setCellRaw(r, c, date ? CellKind.Date : CellKind.Number, v, undefined, xf);
  }

  value(r: number, c: number, kind: number, num: number, obj: unknown, xf: number): void {
    if (kind === CellKind.Number) this.number(r, c, num, xf);
    else if (kind !== CellKind.Empty || xf !== 0) this.ws.setCellRaw(r, c, kind, num, obj, xf);
  }

  /**
   * A formula cell with its cached value (kind / num / obj as for `value`). `rgce` starting with
   * PtgExp (`exp` = master cell key) is resolved at `finish()`, when the shared / array formula
   * definitions that follow their first cell are known.
   */
  formula(
    r: number,
    c: number,
    kind: number,
    num: number,
    obj: unknown,
    xf: number,
    rgce: Uint8Array,
    rgcb: Uint8Array,
    exp: number | undefined,
  ): void {
    if (exp !== undefined) {
      this.pending.push({ r, c, key: exp, kind, num, obj, xf });
      this.value(r, c, kind, num, obj, xf);
      return;
    }
    this.put(r, c, kind, num, obj, xf, decodeFormula(this.ctx, rgce, rgcb, r, c), undefined);
  }

  private put(
    r: number,
    c: number,
    kind: number,
    num: number,
    obj: unknown,
    xf: number,
    text: string | undefined,
    array: string | undefined,
  ): void {
    if (text === undefined) {
      this.value(r, c, kind, num, obj, xf);
      return;
    }
    const f: { -readonly [K in keyof Formula]: Formula[K] } = { formula: text };
    if (kind === CellKind.Number) f.result = num;
    else if (kind === CellKind.Boolean) f.result = num !== 0;
    else if (kind === CellKind.Error) f.result = { error: obj as ErrorCode };
    else if (kind === CellKind.String) f.result = obj as string;
    if (array) f.array = array;
    this.ws.setCellRaw(r, c, CellKind.Formula, 0, f, xf);
  }

  /** Resolve PtgExp cells: array formulas on their master cell, shared formulas everywhere. */
  finish(): void {
    for (const e of this.pending) {
      const def = this.defs.get(e.key);
      if (!def) continue;
      if (def.array !== undefined) {
        if (e.key === key(e.r, e.c))
          this.put(e.r, e.c, e.kind, e.num, e.obj, e.xf, this.decodeDef(def, e.r, e.c), def.array);
      } else
        this.put(e.r, e.c, e.kind, e.num, e.obj, e.xf, this.decodeDef(def, e.r, e.c), undefined);
    }
    this.pending = [];
  }

  private decodeDef(def: FormulaDef, r: number, c: number): string | undefined {
    return decodeFormula(this.ctx, def.rgce, def.rgcb, r, c);
  }

  /** Turn a cell into a hyperlink (like readXlsx: formula cells keep their formula). */
  link(r: number, c: number, target: string, tooltip: string | undefined): void {
    const ws = this.ws;
    const row = ws.rowData[r];
    const i = row ? row.find(c) : -1;
    let text = '';
    let runs: RichText['richText'] | undefined;
    let xf = 0;
    if (i >= 0) {
      const k = row!.kind[i]!;
      if (k === CellKind.Formula || k === CellKind.Hyperlink) return;
      xf = row!.style[i]!;
      const v = ws.getValue(r + 1, c + 1);
      text = display(v);
      if (k === CellKind.RichText) runs = (v as RichText).richText;
    }
    const h: { -readonly [K in keyof Hyperlink]: Hyperlink[K] } = { text, hyperlink: target };
    if (runs) h.runs = runs;
    if (tooltip) h.tooltip = tooltip;
    ws.setCellRaw(r, c, CellKind.Hyperlink, 0, h, xf);
  }
}

/** Map key of a cell (0-based). */
export function key(r: number, c: number): number {
  return r * 16384 + c;
}

function display(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if ('richText' in v) return v.richText.map((x) => x.text).join('');
  if ('error' in v) return v.error;
  return '';
}
