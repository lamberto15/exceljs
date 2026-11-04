// Feature: defined-name helpers (workbook-level named ranges, print area, print titles).
//
//   addDefinedName(wb, 'Rate', 'Sheet1!A1');                       // workbook scope
//   addDefinedName(wb, 'Local', 'B2:C3', { localSheet: ws });      // qualified + absolute
//   setPrintArea(ws, 'A1:D20');   setPrintTitles(ws, { rows: '1:2', cols: 'A:B' });
//   getPrintArea(ws) → ['A1:D20'];  getPrintTitles(ws) → { rows: '1:2', cols: 'A:B' }
//
// The model is `workbook.definedNames` (already read and written by the core); these helpers
// only build correct values: sheet names are quoted ('My Sheet'!, `'` doubled), plain ranges
// become absolute ($A$1:$B$2), names are validated with Excel's rules. Nothing to register:
// there is no reader/writer of its own. Scope (`localSheet`) is stored as the sheet's index in
// `workbook.worksheets` (removing sheets later shifts it — set scoped names last).

import { decodeCol } from '../../core/address.js';
import type { DefinedName } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';

export interface DefinedNameOptions {
  /** Sheet scope: the sheet, its name or its index in `workbook.worksheets`. Default: workbook. */
  localSheet?: Worksheet | string | number | undefined;
  hidden?: boolean | undefined;
  /** Comment shown in Excel's Name Manager. */
  comment?: string | undefined;
}

export interface PrintTitles {
  /** Rows to repeat at the top of every page, e.g. "1:2". */
  rows?: string | undefined;
  /** Columns to repeat at the left of every page, e.g. "A:B". */
  cols?: string | undefined;
}

const PRINT_AREA = '_xlnm.Print_Area';
const PRINT_TITLES = '_xlnm.Print_Titles';

/**
 * Excel's rules for names (defined names and table names): starts with a letter, `_` or `\`;
 * then letters, digits, `.`, `_`, `\`, `?`; no spaces; at most 255 characters; not a cell
 * reference ("A1", "XFD1048576", "R1C1") and not "R" / "C".
 */
export function isValidName(name: string): boolean {
  if (typeof name !== 'string' || name.length === 0 || name.length > 255) return false;
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u.test(name)) return false;
  if (/^(?:[Rr](?:\d+)?(?:[Cc](?:\d+)?)?|[Cc]\d*)$/.test(name)) return false; // R, C, RC, R1, C1, R1C1
  const m = /^([A-Za-z]{1,3})\d+$/.exec(name);
  if (m) {
    try {
      decodeCol(m[1]!.toUpperCase());
      return false; // a real cell address
    } catch {
      /* beyond XFD: a legal name like "ABCD1" or "XFE1" */
    }
  }
  return true;
}

function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const COL = String.raw`\$?[A-Za-z]{1,3}`;
const ROW = String.raw`\$?\d+`;
const RANGE = new RegExp(
  `^(?:('(?:[^']|'')*'|[^!]+)!)?((?:${CELL}(?::${CELL})?)|(?:${COL}:${COL})|(?:${ROW}:${ROW}))$`,
);

function abs(part: string): string {
  const c = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(part);
  if (c) return `$${c[1]!.toUpperCase()}$${c[2]}`;
  return `$${part.replace(/\$/g, '').toUpperCase()}`;
}

/**
 * "A1:B2" / "$A$1" / "A:B" / "1:2", optionally "Sheet!"-prefixed, → "'Sheet'!$A$1:$B$2".
 * Anything else (constants, formulas) is returned as is, without a leading '='.
 */
function qualify(ref: string, sheet: string | undefined): string {
  const r = ref.startsWith('=') ? ref.slice(1) : ref;
  const m = RANGE.exec(r);
  if (!m) return r;
  let sh = m[1];
  if (sh !== undefined) {
    if (sh.startsWith("'")) sh = sh.slice(1, -1).replace(/''/g, "'");
  } else sh = sheet;
  if (sh === undefined) {
    throw new Error(
      `Reference "${ref}" has no sheet: prefix it ("Sheet1!${ref}") or use localSheet`,
    );
  }
  return `${quoteSheet(sh)}!${m[2]!.split(':').map(abs).join(':')}`;
}

function scopeOf(wb: Workbook, s: DefinedNameOptions['localSheet']): number | undefined {
  if (s === undefined) return undefined;
  const i =
    typeof s === 'number'
      ? s
      : typeof s === 'string'
        ? wb.worksheets.findIndex((w) => w.name === s)
        : wb.worksheets.indexOf(s);
  if (!(i >= 0 && i < wb.worksheets.length))
    throw new Error('localSheet is not a sheet of this workbook');
  return i;
}

function indexOfName(wb: Workbook, name: string, scope: number | undefined): number {
  const k = name.toLowerCase();
  return wb.definedNames.findIndex((d) => d.localSheetId === scope && d.name.toLowerCase() === k);
}

function put(wb: Workbook, d: DefinedName): DefinedName {
  const i = indexOfName(wb, d.name, d.localSheetId);
  if (i >= 0) wb.definedNames[i] = d;
  else wb.definedNames.push(d);
  return d;
}

function build(
  name: string,
  value: string,
  scope: number | undefined,
  o: DefinedNameOptions,
): DefinedName {
  const d: { -readonly [K in keyof DefinedName]: DefinedName[K] } = { name, value };
  if (scope !== undefined) d.localSheetId = scope;
  if (o.hidden) d.hidden = true;
  if (o.comment) d.comment = o.comment;
  return d;
}

/**
 * Define (or replace, same name and scope) a name. `ref` is a range ("B2:C3", needs
 * `localSheet`), a sheet-qualified range ("Sheet1!A1:B2" — quoted and made absolute), or any
 * formula / constant ("0.2", "SUM(Sheet1!$A:$A)"), a leading '=' is dropped.
 */
export function addDefinedName(
  wb: Workbook,
  name: string,
  ref: string,
  options: DefinedNameOptions = {},
): DefinedName {
  if (!isValidName(name)) throw new Error(`Invalid defined name "${name}"`);
  const scope = scopeOf(wb, options.localSheet);
  const dflt = scope === undefined ? undefined : wb.worksheets[scope]!.name;
  return put(wb, build(name, qualify(ref, dflt), scope, options));
}

/** The defined name with that scope (workbook-wide unless `localSheet`), case-insensitive. */
export function getDefinedName(
  wb: Workbook,
  name: string,
  localSheet?: DefinedNameOptions['localSheet'],
): DefinedName | undefined {
  return wb.definedNames[indexOfName(wb, name, scopeOf(wb, localSheet))];
}

/** Remove a defined name; returns whether it existed. */
export function removeDefinedName(
  wb: Workbook,
  name: string,
  localSheet?: DefinedNameOptions['localSheet'],
): boolean {
  const i = indexOfName(wb, name, scopeOf(wb, localSheet));
  if (i < 0) return false;
  wb.definedNames.splice(i, 1);
  return true;
}

function setBuiltin(ws: Worksheet, name: string, value: string | undefined): void {
  const wb = ws.workbook;
  const scope = wb.worksheets.indexOf(ws);
  if (scope < 0) throw new Error('Worksheet does not belong to its workbook');
  if (value === undefined) {
    const i = indexOfName(wb, name, scope);
    if (i >= 0) wb.definedNames.splice(i, 1);
  } else put(wb, { name, value, localSheetId: scope });
}

function getBuiltin(ws: Worksheet, name: string): string | undefined {
  const wb = ws.workbook;
  const i = wb.worksheets.indexOf(ws);
  return i < 0 ? undefined : wb.definedNames[indexOfName(wb, name, i)]?.value;
}

/** Split a defined-name value on commas outside quoted sheet names. */
function splitRefs(v: string): string[] {
  return v.match(/(?:'(?:[^']|'')*'|[^,'])+/g) ?? [];
}

function plain(ref: string): string {
  const bang = ref.lastIndexOf('!');
  return (bang >= 0 ? ref.slice(bang + 1) : ref).replace(/\$/g, '');
}

/** Set the print area: one range or several ("A1:D20"); undefined / [] removes it. */
export function setPrintArea(ws: Worksheet, range: string | readonly string[] | undefined): void {
  const list = range === undefined ? [] : typeof range === 'string' ? [range] : range;
  const parts: string[] = [];
  for (const r of list) {
    const m = RANGE.exec(r);
    if (!m || m[1] !== undefined) throw new Error(`Invalid print area range "${r}"`);
    parts.push(qualify(r, ws.name));
  }
  setBuiltin(ws, PRINT_AREA, parts.length ? parts.join(',') : undefined);
}

/** The sheet's print area ranges without sheet prefix or `$` ("A1:D20"), if any. */
export function getPrintArea(ws: Worksheet): string[] | undefined {
  const v = getBuiltin(ws, PRINT_AREA);
  return v === undefined ? undefined : splitRefs(v).map(plain);
}

/** Set rows / columns repeated on every printed page; undefined or `{}` removes them. */
export function setPrintTitles(ws: Worksheet, titles: PrintTitles | undefined): void {
  const parts: string[] = [];
  const { cols, rows } = titles ?? {};
  if (cols !== undefined) {
    if (!/^\$?[A-Za-z]{1,3}(:\$?[A-Za-z]{1,3})?$/.test(cols))
      throw new Error(`Invalid columns "${cols}"`);
    parts.push(qualify(cols.includes(':') ? cols : `${cols}:${cols}`, ws.name));
  }
  if (rows !== undefined) {
    if (!/^\$?\d+(:\$?\d+)?$/.test(rows)) throw new Error(`Invalid rows "${rows}"`);
    parts.push(qualify(rows.includes(':') ? rows : `${rows}:${rows}`, ws.name));
  }
  setBuiltin(ws, PRINT_TITLES, parts.length ? parts.join(',') : undefined);
}

/** The sheet's print titles ({ rows: "1:2", cols: "A:B" }), if any. */
export function getPrintTitles(ws: Worksheet): PrintTitles | undefined {
  const v = getBuiltin(ws, PRINT_TITLES);
  if (v === undefined) return undefined;
  const out: PrintTitles = {};
  for (const r of splitRefs(v)) {
    const p = plain(r);
    if (/^\d+:\d+$/.test(p)) out.rows = p;
    else if (/^[A-Za-z]+:[A-Za-z]+$/.test(p)) out.cols = p;
  }
  return out;
}
