// workbook.xml reader: sheet list (name, order, state, rel id), date system, defined names.

import type { DefinedName, WorkbookView } from '../../core/types.js';
import type { SheetState } from '../../core/worksheet.js';
import { localName, type XmlAttrs, type XmlHandler } from '../../xml/index.js';
import { stripXlfn } from '../common/xlfn.js';
import { attrBool, attrLocal } from './xml-util.js';

export interface SheetEntry {
  readonly name: string;
  readonly sheetId: number;
  readonly state: SheetState;
  /** Relationship id into xl/_rels/workbook.xml.rels. */
  readonly rId: string | undefined;
}

export class WorkbookHandler implements XmlHandler {
  readonly sheets: SheetEntry[] = [];
  readonly definedNames: DefinedName[] = [];
  date1904 = false;
  /** `<calcPr fullCalcOnLoad>`. */
  fullCalcOnLoad = false;
  /** `<workbookView>`s (Excel windows). */
  readonly views: WorkbookView[] = [];
  private dn:
    | { name: string; localSheetId?: number; hidden?: boolean; comment?: string }
    | undefined;
  private dnText = '';

  open(name: string, a: XmlAttrs): void {
    switch (localName(name)) {
      case 'calcPr':
        this.fullCalcOnLoad = attrBool(a.fullCalcOnLoad);
        break;
      case 'workbookView':
        {
          const n = (k: string): number | undefined => (a[k] !== undefined ? +a[k]! : undefined);
          const v: WorkbookView = {};
          // the first window's activeTab is not kept: the sheets' tabSelected carries it
          if (a.activeTab !== undefined && this.views.length) v.activeTab = n('activeTab');
          if (a.firstSheet !== undefined) v.firstSheet = n('firstSheet');
          if (a.xWindow !== undefined) v.x = n('xWindow');
          if (a.yWindow !== undefined) v.y = n('yWindow');
          if (a.windowWidth !== undefined) v.width = n('windowWidth');
          if (a.windowHeight !== undefined) v.height = n('windowHeight');
          if (a.visibility === 'hidden' || a.visibility === 'veryHidden')
            v.visibility = a.visibility;
          this.views.push(v);
        }
        break;
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
        if (a.comment) this.dn.comment = a.comment;
        this.dnText = '';
        break;
    }
  }

  close(name: string): void {
    if (this.dn && localName(name) === 'definedName') {
      this.definedNames.push({ ...this.dn, value: stripXlfn(this.dnText) });
      this.dn = undefined;
    }
  }

  text(t: string): void {
    if (this.dn) this.dnText += t;
  }
}
