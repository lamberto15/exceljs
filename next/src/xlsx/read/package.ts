// Package model: which zip entry plays which role (workbook, styles, shared strings, sheets),
// resolved through [Content_Types].xml and the relationship parts. Works incrementally (the
// streaming reader feeds parts as they appear in the zip) and with fallbacks for incomplete
// packages (no _rels/.rels, no workbook.xml, ...), finalized via `finalize()`.

import {
  type ContentTypes,
  parseContentTypes,
  parseRels,
  partDir,
  REL_OFFICE_DOCUMENT,
  REL_SHARED_STRINGS,
  REL_STYLES,
  REL_WORKSHEET,
  type Relationship,
  relTypeIs,
} from '../common/r-opc.js';
import type { SheetEntry } from './workbook-part.js';

export const CT_PATH = '[Content_Types].xml';

const CT_MAIN = /\.sheet\.main\+xml$|\.template\.main\+xml$|sheet\.macroEnabled\.main\+xml$/;
const CT_WORKSHEET = /\.worksheet\+xml$/;
const CT_STYLES = /\.styles\+xml$/;
const CT_SST = /\.sharedStrings\+xml$/;

export type Role = 'workbook' | 'styles' | 'sst' | 'sheet' | 'ignore' | 'unknown';

export interface ResolvedSheet extends SheetEntry {
  /** Zip entry path of the sheet part. */
  readonly path: string;
  /** 0-based index in tab order (among all sheets of the workbook, incl. skipped ones). */
  readonly index: number;
}

export class XlsxPackage {
  ct: ContentTypes | undefined;
  readonly rels: Map<string, Map<string, Relationship>> = new Map();
  /** Lower-case → actual entry name, for case-insensitive lookups (filled by the caller). */
  readonly names: Map<string, string> = new Map();
  final = false;
  /** Sheets from workbook.xml (set by the caller once parsed). */
  sheetEntries: SheetEntry[] | undefined;

  /** Register an entry name (for case-insensitive part matching). */
  addName(path: string): void {
    this.names.set(path.toLowerCase(), path);
  }

  /** Feed a metadata part ([Content_Types].xml or any .rels). Returns true if consumed. */
  addMeta(path: string, xml: string): boolean {
    if (path === CT_PATH) {
      this.ct = parseContentTypes(xml);
      return true;
    }
    if (path.endsWith('.rels')) {
      // "xl/_rels/workbook.xml.rels" → source "xl/workbook.xml"; "_rels/.rels" → ""
      const i = path.lastIndexOf('_rels/');
      const source = path.slice(0, i) + path.slice(i + 6, -5);
      this.rels.set(source.toLowerCase(), parseRels(xml, source));
      return true;
    }
    return false;
  }

  relsOf(part: string): Map<string, Relationship> | undefined {
    return this.rels.get(part.toLowerCase());
  }

  /** Actual entry name for a resolved path (case-insensitive), or the path itself. */
  entry(path: string): string {
    return this.names.get(path.toLowerCase()) ?? path;
  }

  private has(path: string): boolean {
    return this.names.has(path.toLowerCase());
  }

  private ctFind(re: RegExp): string[] {
    const out: string[] = [];
    if (this.ct) for (const [p, t] of this.ct.overrides) if (re.test(t)) out.push(p);
    return out;
  }

  /** Workbook part path, or undefined while not determinable yet. */
  workbookPath(): string | undefined {
    const root = this.relsOf('');
    if (root) {
      for (const r of root.values()) {
        if (!r.external && relTypeIs(r.type, REL_OFFICE_DOCUMENT)) return this.entry(r.path);
      }
    }
    if (this.ct && !root) {
      const m = this.ctFind(CT_MAIN)[0];
      if (m) return this.entry(m);
    }
    if (this.final || root) {
      for (const p of ['xl/workbook.xml', 'xl/workbook.bin']) if (this.has(p)) return this.entry(p);
      return 'xl/workbook.xml';
    }
    return undefined;
  }

  private wbRel(kind: string): string | null | undefined {
    const wb = this.workbookPath();
    if (wb === undefined) return undefined;
    const rels = this.relsOf(wb);
    if (rels) {
      for (const r of rels.values())
        if (!r.external && relTypeIs(r.type, kind)) return this.entry(r.path);
    }
    if (!this.final && !rels) return undefined;
    // fallbacks: content types, then the conventional name next to the workbook
    const re = kind === REL_STYLES ? CT_STYLES : CT_SST;
    const c = this.ctFind(re)[0];
    if (c && this.has(c)) return this.entry(c);
    const conv = partDir(wb) + (kind === REL_STYLES ? 'styles.xml' : 'sharedStrings.xml');
    return this.has(conv) ? this.entry(conv) : null;
  }

  /** Styles part path; null = none; undefined = not known yet. */
  stylesPath(): string | null | undefined {
    return this.wbRel(REL_STYLES);
  }

  /** Shared strings part path; null = none; undefined = not known yet. */
  sstPath(): string | null | undefined {
    return this.wbRel(REL_SHARED_STRINGS);
  }

  /**
   * Worksheets in tab order with their part paths (chartsheets / dialog sheets and sheets whose
   * part is missing are dropped). Undefined while not determinable yet.
   */
  sheets(): ResolvedSheet[] | undefined {
    const wb = this.workbookPath();
    if (wb === undefined) return undefined;
    const rels = this.relsOf(wb);
    const entries = this.sheetEntries;
    if (entries && rels) {
      const out: ResolvedSheet[] = [];
      entries.forEach((s, index) => {
        const r = s.rId !== undefined ? rels.get(s.rId) : undefined;
        if (!r || r.external || !relTypeIs(r.type, REL_WORKSHEET)) return;
        if (this.final && !this.has(r.path)) return;
        out.push({ ...s, path: this.entry(r.path), index });
      });
      return out;
    }
    if (!this.final) return undefined;
    // Incomplete package: synthesize from rels / content types / entry names.
    let paths: string[] = [];
    if (rels) {
      for (const r of rels.values()) {
        if (!r.external && relTypeIs(r.type, REL_WORKSHEET)) paths.push(r.path);
      }
    }
    if (paths.length === 0) paths = this.ctFind(CT_WORKSHEET);
    if (paths.length === 0) {
      for (const n of this.names.values())
        if (/(^|\/)worksheets\/[^/]+\.xml$/i.test(n)) paths.push(n);
      paths.sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
    }
    const out: ResolvedSheet[] = [];
    let index = 0;
    for (const p of paths) {
      if (!this.has(p)) continue;
      const e = entries?.[index];
      out.push({
        name: e?.name ?? `Sheet${index + 1}`,
        sheetId: index + 1,
        state: e?.state ?? 'visible',
        rId: undefined,
        path: this.entry(p),
        index,
      });
      index++;
    }
    return out;
  }

  /** Role of an entry given what is known so far. */
  roleOf(path: string): Role {
    const wb = this.workbookPath();
    if (wb === undefined) return 'unknown';
    const lp = path.toLowerCase();
    if (lp === wb.toLowerCase()) return 'workbook';
    const st = this.stylesPath();
    const ss = this.sstPath();
    if (st && lp === st.toLowerCase()) return 'styles';
    if (ss && lp === ss.toLowerCase()) return 'sst';
    const rels = this.relsOf(wb);
    if (!rels && !this.final) return 'unknown';
    if (rels) {
      for (const r of rels.values()) {
        if (!r.external && relTypeIs(r.type, REL_WORKSHEET) && r.path.toLowerCase() === lp) {
          return 'sheet';
        }
      }
      return 'ignore';
    }
    const sh = this.sheets();
    return sh?.some((s) => s.path.toLowerCase() === lp) ? 'sheet' : 'ignore';
  }
}

/** Whether an entry might matter before roles are known (buffered by the streaming reader). */
export function maybeRelevant(path: string): boolean {
  return (
    /\.xml$/i.test(path) &&
    /workbook|styles|sharedstrings|sheet/i.test(path) &&
    !/^docProps\//.test(path)
  );
}
