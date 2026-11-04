// Workbook: ordered worksheets + one shared StyleRegistry.

import type { WorkbookFeatureFields } from './features.js';
import { StyleRegistry } from './styles.js';
import type { Color, DefinedName } from './types.js';
import { type SheetState, Worksheet } from './worksheet.js';

export interface WorkbookOptions {
  /** Use the 1904 date system. Set before storing dates (serials are stored, not Dates). */
  date1904?: boolean | undefined;
}

export interface WorksheetOptions {
  tabColor?: Color | undefined;
  state?: SheetState | undefined;
}

const INVALID_NAME = /[\\/?*[\]:]/;

/** Feature-owned optional fields, see core/features.ts. */
export interface Workbook extends WorkbookFeatureFields {}

// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: type-only feature fields
export class Workbook {
  /** 1904 date system flag; changing it after storing dates shifts them by 1462 days. */
  date1904: boolean;
  readonly styles: StyleRegistry = new StyleRegistry();
  creator: string | undefined = undefined;
  lastModifiedBy: string | undefined = undefined;
  created: Date = new Date();
  modified: Date = new Date();
  /**
   * Defined names, written to workbook.xml. `_xlnm._FilterDatabase` names are managed through
   * `worksheet.autoFilter` (readers move them there; writers generate them) — don't add them.
   */
  definedNames: DefinedName[] = [];

  private readonly _sheets: Worksheet[] = [];
  private _nextId = 1;

  constructor(options: WorkbookOptions = {}) {
    this.date1904 = options.date1904 === true;
  }

  /** Worksheets in tab order. */
  get worksheets(): readonly Worksheet[] {
    return this._sheets;
  }

  /** Add a worksheet (default name "SheetN"). Names are unique case-insensitively. */
  addWorksheet(name?: string, options: WorksheetOptions = {}): Worksheet {
    let n = name;
    if (n === undefined) {
      let i = this._sheets.length + 1;
      while (this._find(`Sheet${i}`)) i++;
      n = `Sheet${i}`;
    }
    this._checkName(n, undefined);
    const ws = new Worksheet(this, this._nextId++, n);
    if (options.tabColor) ws.tabColor = options.tabColor;
    if (options.state) ws.state = options.state;
    this._sheets.push(ws);
    return ws;
  }

  /**
   * Find a worksheet by name (case-insensitive) or by sheet id (number, 1-based; equals the
   * tab position until sheets are removed). No argument → first worksheet.
   */
  getWorksheet(nameOrId?: string | number): Worksheet | undefined {
    if (nameOrId === undefined) return this._sheets[0];
    if (typeof nameOrId === 'number') return this._sheets.find((s) => s.id === nameOrId);
    return this._find(nameOrId);
  }

  /** Remove a worksheet by name, id or instance. */
  removeWorksheet(ws: string | number | Worksheet): void {
    const target = ws instanceof Worksheet ? ws : this.getWorksheet(ws);
    const i = target ? this._sheets.indexOf(target) : -1;
    if (i < 0) return;
    this._sheets.splice(i, 1);
    // sheet-scoped names are keyed by tab index: drop the removed sheet's, shift later ones
    const names: DefinedName[] = [];
    for (const d of this.definedNames) {
      const s = d.localSheetId;
      if (s === undefined || s < i) names.push(d);
      else if (s > i) names.push({ ...d, localSheetId: s - 1 });
    }
    this.definedNames = names;
  }

  /** @internal Validate a sheet name (Excel rules) and uniqueness. */
  _checkName(name: string, self: Worksheet | undefined): void {
    if (name.length === 0 || name.length > 31) {
      throw new Error(`Worksheet name must be 1-31 characters: "${name}"`);
    }
    if (INVALID_NAME.test(name)) {
      throw new Error(`Worksheet name contains an invalid character (\\ / ? * [ ] :): "${name}"`);
    }
    if (name.charCodeAt(0) === 39 || name.charCodeAt(name.length - 1) === 39) {
      throw new Error(`Worksheet name cannot start or end with an apostrophe: "${name}"`);
    }
    if (name.toLowerCase() === 'history') throw new Error('"History" is a reserved sheet name');
    const other = this._find(name);
    if (other && other !== self) throw new Error(`Worksheet name already exists: "${name}"`);
  }

  private _find(name: string): Worksheet | undefined {
    const lower = name.toLowerCase();
    return this._sheets.find((s) => s.name.toLowerCase() === lower);
  }
}
