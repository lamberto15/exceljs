// Pieces shared by the binary readers (`./xlsb` BIFF12 and `./xls` BIFF8): enumerations that
// both formats number the same way, RK numbers, string decoding, and the steps that turn
// decoded records into the model (styles, sheets, defined names).

import type { StyleRegistry } from '../core/styles.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  DefinedName,
  ErrorCode,
  Fill,
  Font,
  Protection,
} from '../core/types.js';
import type { Workbook } from '../core/workbook.js';
import type { SheetState, Worksheet } from '../core/worksheet.js';

/** Malformed or unsupported .xls / .xlsb data. */
export class BiffParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'BiffParseError';
  }
}

/** Cell error codes (BErr). */
export const ERRORS: Readonly<Record<number, ErrorCode>> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};

/** Fill patterns (FillPattern / fls) in file order. */
const PATTERNS = [
  'none',
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
];

/** Border line styles (dg). */
const LINES: (BorderStyle | undefined)[] = [
  undefined,
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];

const H_ALIGN = [
  undefined,
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
] as const;
const V_ALIGN = ['top', 'center', undefined, 'justify', 'distributed'] as const;

/** RK number: a 30-bit integer or the high 30 bits of a double, optionally divided by 100. */
export function rkValue(rk: number): number {
  let v: number;
  if (rk & 2) v = rk >> 2;
  else {
    const dv = new DataView(new ArrayBuffer(8));
    dv.setInt32(4, rk & ~3, true);
    v = dv.getFloat64(0, true);
  }
  return rk & 1 ? v / 100 : v;
}

/** `n` UTF-16LE code units at `p`. */
export function utf16(d: Uint8Array, p: number, n: number): string {
  let s = '';
  for (let i = 0; i < n; i += 4096) {
    const end = Math.min(n, i + 4096);
    const codes = new Array<number>(end - i);
    for (let k = i; k < end; k++) codes[k - i] = d[p + 2 * k]! | (d[p + 2 * k + 1]! << 8);
    s += String.fromCharCode(...codes);
  }
  return s;
}

/** `n` 8-bit characters at `p` (BIFF8 "compressed" UTF-16: the high bytes are zero). */
export function latin1(d: Uint8Array, p: number, n: number): string {
  let s = '';
  for (let i = 0; i < n; i += 4096) {
    s += String.fromCharCode(...d.subarray(p + i, p + Math.min(n, i + 4096)));
  }
  return s;
}

/** ARGB hex of color channels ("FF0000FF" = opaque blue). */
export function argb(a: number, r: number, g: number, b: number): string {
  return (((a << 24) | (r << 16) | (g << 8) | b) >>> 0).toString(16).padStart(8, '0').toUpperCase();
}

/** A border edge, or undefined for "no line". */
export function edge(dg: number, color: Font['color']): BorderEdge | undefined {
  const style = LINES[dg];
  if (!style) return undefined;
  return color ? { style, color } : { style };
}

export function fill(fls: number, fg: Font['color'], bg: Font['color']): Fill | undefined {
  const pattern = PATTERNS[fls];
  if (pattern === undefined || pattern === 'none') return undefined;
  const f: { type: 'pattern'; pattern: string; fgColor?: Font['color']; bgColor?: Font['color'] } =
    { type: 'pattern', pattern };
  if (fg) f.fgColor = fg;
  if (bg) f.bgColor = bg;
  return f as Fill;
}

export function alignment(
  alc: number,
  alcv: number,
  wrap: boolean,
  shrink: boolean,
  indent: number,
  trot: number,
): Alignment | undefined {
  const a: { -readonly [K in keyof Alignment]: Alignment[K] } = {};
  const h = H_ALIGN[alc];
  const v = V_ALIGN[alcv];
  if (h) a.horizontal = h;
  if (v) a.vertical = v;
  if (wrap) a.wrapText = true;
  if (shrink) a.shrinkToFit = true;
  if (indent) a.indent = indent;
  if (trot) a.textRotation = trot;
  for (const _ in a) return a;
  return undefined;
}

/** A cell format resolved to model pieces (before interning). */
export interface XfRecord {
  numFmtId: number;
  /** The cell font; undefined = the workbook default font. */
  font: Font | undefined;
  fill: Fill | undefined;
  border: Border;
  alignment: Alignment | undefined;
  protection: Protection;
}

/**
 * Interns decoded formats into `reg`: `defaultFont` becomes the workbook default font (font 0,
 * like readXlsx), number formats are registered under the ids the file uses.
 */
export class StyleSink {
  constructor(
    private readonly reg: StyleRegistry,
    defaultFont: Font,
    private readonly formats: ReadonlyMap<number, string>,
  ) {
    reg.setDefaultFont(defaultFont);
  }

  intern(x: XfRecord): number {
    const reg = this.reg;
    const code = this.formats.get(x.numFmtId);
    if (code !== undefined) reg.registerNumFmt(x.numFmtId, code);
    return reg.internXf({
      numFmtId: x.numFmtId,
      fontId: x.font ? reg.addFont(x.font) : 0,
      fillId: x.fill ? reg.addFill(x.fill) : 0,
      borderId: reg.addBorder(x.border),
      alignment: x.alignment,
      protection: x.protection,
    });
  }
}

/** Add a worksheet, renaming names Excel would reject (like readXlsx). */
export function addSheet(wb: Workbook, name: string, state: SheetState): Worksheet {
  const opts = state !== 'visible' ? { state } : {};
  try {
    return wb.addWorksheet(name, opts);
  } catch {
    let base =
      name
        .replace(/[\\/?*[\]:]/g, '_')
        .replace(/^'+|'+$/g, '')
        .slice(0, 28) || 'Sheet';
    if (base.toLowerCase() === 'history') base = 'History_';
    let n = base;
    for (let i = 2; wb.getWorksheet(n); i++) n = `${base}(${i})`;
    return wb.addWorksheet(n, opts);
  }
}

/** A sheet name as written in a formula: quoted when it is not a plain identifier. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_¡-￿][\w.¡-￿]*$/.test(name) &&
    !/^([A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|TRUE|FALSE)$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

/** A defined name as decoded from the file; `sheet` = 0-based file sheet index. */
export interface RawName {
  name: string;
  value: string;
  sheet: number | undefined;
  hidden: boolean;
}

/**
 * Defined names → `wb.definedNames`, with sheet scopes remapped to loaded sheets
 * (`sheetMap`: file index → wb index); `_xlnm._FilterDatabase` becomes `ws.autoFilter`, as in
 * readXlsx.
 */
export function applyNames(
  wb: Workbook,
  names: readonly RawName[],
  sheetMap: ReadonlyMap<number, number>,
): void {
  const out: DefinedName[] = [];
  for (const n of names) {
    let local: number | undefined;
    if (n.sheet !== undefined) {
      local = sheetMap.get(n.sheet);
      if (local === undefined) continue;
    }
    if (n.name === '_xlnm._FilterDatabase') {
      const ws = local !== undefined ? wb.worksheets[local] : undefined;
      const ref = n.value.slice(n.value.lastIndexOf('!') + 1).replace(/\$/g, '');
      if (ws && ws.autoFilter === undefined && /^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(ref)) {
        ws.autoFilter = ref;
      }
      continue;
    }
    const d: { -readonly [K in keyof DefinedName]: DefinedName[K] } = {
      name: n.name,
      value: n.value,
    };
    if (local !== undefined) d.localSheetId = local;
    if (n.hidden) d.hidden = true;
    out.push(d);
  }
  wb.definedNames = out;
}

/** Built-in defined names of .xls (one-character codes) without the `_xlnm.` prefix. */
export const BUILTIN_NAMES: readonly string[] = [
  'Consolidate_Area',
  'Auto_Open',
  'Auto_Close',
  'Extract',
  'Database',
  'Criteria',
  'Print_Area',
  'Print_Titles',
  'Recorder',
  'Data_Form',
  'Auto_Activate',
  'Auto_Deactivate',
  'Sheet_Title',
  '_FilterDatabase',
];
