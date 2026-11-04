// workbook.xml reader: sheet list (name, order, state, rel id), date system, defined names.

import type { SheetState } from '../../core/worksheet.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { attrBool, attrLocal } from './xml-util.js';

export interface SheetEntry {
  readonly name: string;
  readonly sheetId: number;
  readonly state: SheetState;
  /** Relationship id into xl/_rels/workbook.xml.rels. */
  readonly rId: string | undefined;
}

export interface DefinedName {
  readonly name: string;
  /** Raw formula text, e.g. "Sheet1!$A$1:$D$10". */
  readonly value: string;
  /** 0-based index into the sheet list for sheet-scoped names. */
  readonly localSheetId?: number | undefined;
  readonly hidden?: boolean | undefined;
}

export class WorkbookHandler implements XmlHandler {
  readonly sheets: SheetEntry[] = [];
  readonly definedNames: DefinedName[] = [];
  date1904 = false;
  private dn: { name: string; localSheetId?: number; hidden?: boolean } | undefined;
  private dnText = '';

  open(name: string, a: XmlAttrs): void {
    switch (localName(name)) {
      case 'workbookPr':
        this.date1904 = attrBool(a.date1904);
        break;
      case 'sheet': {
        const state = a.state === 'hidden' || a.state === 'veryHidden' ? a.state : 'visible';
        this.sheets.push({
          name: a.name ?? `Sheet${this.sheets.length + 1}`,
          sheetId: a.sheetId !== undefined ? +a.sheetId : this.sheets.length + 1,
          state,
          rId: attrLocal(a, 'id'),
        });
        break;
      }
      case 'definedName':
        this.dn = { name: a.name ?? '' };
        if (a.localSheetId !== undefined) this.dn.localSheetId = +a.localSheetId;
        if (attrBool(a.hidden)) this.dn.hidden = true;
        this.dnText = '';
        break;
    }
  }

  close(name: string): void {
    if (this.dn && localName(name) === 'definedName') {
      this.definedNames.push({ ...this.dn, value: this.dnText });
      this.dn = undefined;
    }
  }

  text(t: string): void {
    if (this.dn) this.dnText += t;
  }
}
