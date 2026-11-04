// Package bookkeeping shared by all writers: extra parts queued by features, content types
// (Default by extension / Override by part), relationship lists, part-family counters.

import { escapeAttr } from '../../xml/index.js';
import type { PackageWriteCtx } from '../features/types.js';

const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export interface PartRec {
  readonly path: string;
  readonly data: string | Uint8Array;
}

export interface RelRec {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

/** Relationships of one source part; ids are "rId<n>" from `first`. */
export class RelList {
  readonly recs: RelRec[] = [];
  private next: number;

  constructor(first = 1) {
    this.next = first;
  }

  add(type: string, target: string, external = false): string {
    const id = `rId${this.next++}`;
    this.recs.push({ id, type: type.includes(':') ? type : REL + type, target, external });
    return id;
  }
}

/** Relationship elements (without the surrounding part), e.g. to append to built-in ones. */
export function relElems(recs: readonly RelRec[]): string {
  let s = '';
  for (const r of recs) {
    s += `<Relationship Id="${r.id}" Type="${escapeAttr(r.type)}" Target="${escapeAttr(r.target)}"${
      r.external ? ' TargetMode="External"' : ''
    }/>`;
  }
  return s;
}

export function relsXml(recs: readonly RelRec[]): string {
  return `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">${relElems(recs)}</Relationships>`;
}

export class PackageParts implements PackageWriteCtx {
  /** Set when a dynamic-array formula was written (xl/metadata.xml is needed). */
  dynamicArrays = false;
  /** [part path, content type] registered by features. */
  readonly overrides: [string, string][] = [];
  /** Extension (lower-case) → content type registered by features. */
  readonly defaults: Map<string, string> = new Map();
  private pending: PartRec[] = [];
  private readonly counters = new Map<string, number>();
  private readonly reserved = new Set<string>();

  /** Make `nextIndex(family)` skip the number in `key` ("drawing1" → drawing 1 is taken). */
  reserve(key: string): void {
    this.reserved.add(key);
  }

  addPart(path: string, data: string | Uint8Array, contentType?: string): void {
    this.pending.push({ path, data });
    if (contentType) this.overrides.push([path, contentType]);
  }

  addDefault(ext: string, contentType: string): void {
    this.defaults.set(ext.toLowerCase(), contentType);
  }

  nextIndex(family: string): number {
    let n = (this.counters.get(family) ?? 0) + 1;
    while (this.reserved.has(family + n)) n++;
    this.counters.set(family, n);
    return n;
  }

  /** Parts queued since the last call. */
  take(): PartRec[] {
    const p = this.pending;
    this.pending = [];
    return p;
  }
}
