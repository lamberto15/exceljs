// OPC (Open Packaging Conventions) helpers for readers: part-name resolution, relationships,
// content types, and byte → string decoding of XML parts.

import { localName, parseXml } from '../../xml/index.js';

export interface Relationship {
  readonly id: string;
  /** Relationship type URI (compare with {@link relTypeIs}). */
  readonly type: string;
  /** Raw target as written in the file. */
  readonly target: string;
  /** Resolved zip entry path (no leading slash) for internal targets; raw target when external. */
  readonly path: string;
  readonly external: boolean;
}

export const REL_OFFICE_DOCUMENT = 'officeDocument';
export const REL_WORKSHEET = 'worksheet';
export const REL_STYLES = 'styles';
export const REL_SHARED_STRINGS = 'sharedStrings';

/**
 * Match a relationship type by its last path segment, so both the transitional
 * (`http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`) and strict
 * (`http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet`) URIs match.
 */
export function relTypeIs(type: string, kind: string): boolean {
  const i = type.lastIndexOf('/');
  return (i < 0 ? type : type.slice(i + 1)) === kind;
}

/** Directory of a part path ("xl/workbook.xml" → "xl/"; "workbook.xml" → ""). */
export function partDir(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i + 1);
}

/** Relationship part for a part: "xl/workbook.xml" → "xl/_rels/workbook.xml.rels"; "" → "_rels/.rels". */
export function relsPathFor(path: string): string {
  const i = path.lastIndexOf('/');
  return `${path.slice(0, i + 1)}_rels/${path.slice(i + 1)}.rels`;
}

function safeDecodeURI(s: string): string {
  if (s.indexOf('%') < 0) return s;
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/**
 * Resolve a relationship target against the source part's directory into a zip entry path
 * (no leading slash, `.`/`..` collapsed, backslashes normalized, percent-decoding applied).
 * Absolute targets ("/xl/worksheets/sheet1.xml") resolve from the package root.
 */
export function resolveTarget(baseDir: string, target: string): string {
  let t = safeDecodeURI(target.replace(/\\/g, '/'));
  const hash = t.indexOf('#');
  if (hash >= 0) t = t.slice(0, hash);
  const full = t.charCodeAt(0) === 47 ? t.slice(1) : baseDir + t;
  const out: string[] = [];
  for (const seg of full.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

/** Parse a .rels part. `sourcePath` is the part the rels belong to ("" for the package). */
export function parseRels(xml: string, sourcePath: string): Map<string, Relationship> {
  const map = new Map<string, Relationship>();
  const base = partDir(sourcePath);
  parseXml(xml, {
    open(name, a) {
      if (localName(name) !== 'Relationship') return;
      const id = a.Id ?? a.id;
      const target = a.Target ?? a.target;
      if (id === undefined || target === undefined) return;
      const external = (a.TargetMode ?? a.targetMode) === 'External';
      map.set(id, {
        id,
        type: a.Type ?? a.type ?? '',
        target,
        path: external ? target : resolveTarget(base, target),
        external,
      });
    },
  });
  return map;
}

export interface ContentTypes {
  /** Extension (lower-case, no dot) → content type. */
  readonly defaults: Map<string, string>;
  /** Part path (no leading slash) → content type. */
  readonly overrides: Map<string, string>;
}

export function parseContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  parseXml(xml, {
    open(name, a) {
      const ln = localName(name);
      if (ln === 'Default') {
        if (a.Extension && a.ContentType) defaults.set(a.Extension.toLowerCase(), a.ContentType);
      } else if (ln === 'Override') {
        if (a.PartName && a.ContentType)
          overrides.set(resolveTarget('', a.PartName), a.ContentType);
      }
    },
  });
  return { defaults, overrides };
}

/** Content type of a part, or undefined. Override lookup is case-insensitive per OPC. */
export function contentTypeOf(ct: ContentTypes, path: string): string | undefined {
  const o = ct.overrides.get(path);
  if (o !== undefined) return o;
  const lower = path.toLowerCase();
  for (const [k, v] of ct.overrides) if (k.toLowerCase() === lower) return v;
  const dot = path.lastIndexOf('.');
  return dot < 0 ? undefined : ct.defaults.get(path.slice(dot + 1).toLowerCase());
}

const CHUNK = 1 << 20;

function decoderFor(bytes: Uint8Array): TextDecoder {
  if (bytes.length >= 2) {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le');
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be');
  }
  return new TextDecoder('utf-8');
}

/** Decode a whole part to a string (BOM stripped, UTF-16 detected by BOM). */
export function decodePart(bytes: Uint8Array): string {
  return decoderFor(bytes).decode(bytes);
}

/**
 * Decode a part in ~1 MB string chunks (keeps peak memory low for huge sheets: the whole
 * decoded XML never exists as one string).
 */
export function decodePartChunked(bytes: Uint8Array, onChunk: (s: string) => void): void {
  if (bytes.length <= CHUNK) {
    onChunk(decodePart(bytes));
    return;
  }
  const dec = decoderFor(bytes);
  for (let i = 0; i < bytes.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, bytes.length);
    const s = dec.decode(bytes.subarray(i, end), { stream: end < bytes.length });
    if (s.length > 0) onChunk(s);
  }
}
