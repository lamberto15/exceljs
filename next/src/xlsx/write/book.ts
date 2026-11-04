// In-memory workbook → XLSX: writeXlsx (buffered) and xlsxStream (pull-based ReadableStream).
// Both drive the same generator, which yields sheet XML in ~CHUNK-sized pieces so the stream
// variant only materializes one piece at a time.

import { SharedStrings } from '../../core/sharedstrings.js';
import type { WriteOptions } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { nativeCompressor } from '../../zip/index.js';
import type { PPart } from '../common/preserved.js';
import { WRITERS } from '../features/registry.js';
import type { BookWriteCtx } from '../features/types.js';
import { type ResolvedWriteOptions, resolveWriteOptions } from './options.js';
import { PackageParts, type PartRec, RelList, relElems } from './package.js';
import {
  appPropsXml,
  type BookInfo,
  contentTypesXml,
  corePropsXml,
  rootRelsXml,
  type SheetEntry,
  sharedStringsXml,
  stylesXml,
  THEME_CT,
  THEME_PATH,
  themeXml,
  workbookRelsXml,
  workbookXml,
} from './parts.js';
import { planPreserved, presBook, sheetIds } from './preserve.js';
import { afterSheetParts, encodeRow, newSheetCtx, sheetHead, sheetTail } from './sheet.js';
import { concatChunks, ZipOut } from './zipout.js';

/** Sheet XML is handed to the zip in pieces of about this many UTF-16 units. */
export const CHUNK: number = 128 * 1024;

interface Piece {
  path: string;
  data: string | Uint8Array;
  last: boolean;
}

/**
 * Workbook-level parts, written after all sheets: feature book hooks/parts, workbook.xml,
 * styles, shared strings, workbook rels, content types, root rels, docProps. Shared by every
 * writer.
 */
export function bookParts(
  wb: Workbook,
  book: BookInfo,
  pkg: PackageParts,
  sst: SharedStrings | undefined,
  kept?: readonly PPart[],
): PartRec[] {
  const n = book.sheets.length;
  // ids after the built-in worksheet / styles / sharedStrings relationships
  const rels = new RelList(n + 3);
  const ctx: BookWriteCtx = {
    workbook: wb,
    styles: wb.styles,
    addRel: (type, target, external) => rels.add(type, target, external),
    addPart: (path, data, ct) => pkg.addPart(path, data, ct),
    addDefault: (ext, ct) => pkg.addDefault(ext, ct),
    nextIndex: (f) => pkg.nextIndex(f),
  };
  for (const f of WRITERS) f.book?.write?.(ctx);
  const pb = wb.preserved;
  // a theme kept from the source file replaces the default one (never two themes)
  const keptTheme = kept && pb?.implicit.some((k) => /\/theme$/.test(pb.rels[k]!.type));
  if (book.theme !== false && !keptTheme) {
    pkg.addPart(THEME_PATH, themeXml(), THEME_CT);
    rels.add('theme', 'theme/theme1.xml');
  }
  const extra: [number, string, number][] = [];
  for (const f of WRITERS) {
    for (const e of f.book?.emitters ?? []) {
      const x = e.emit(ctx);
      if (x) extra.push([e.slot, x, e.order ?? 0]);
    }
  }
  let info = book;
  let rootExtra = '';
  if (kept) {
    const pres = presBook(
      wb,
      pkg,
      kept,
      ctx.addRel,
      book.sheets.map((s) => s.sheetId),
    );
    extra.push(...pres.frags);
    info = { ...book, prAttrs: pb?.prAttrs, calcAttrs: pb?.calcAttrs, extraSheets: pres.sheets };
    const root = new RelList(4);
    for (const r of pres.root) {
      const t = r.target;
      if (typeof t === 'string') root.add(r.type, t, true);
      else root.add(r.type, t.out!);
    }
    rootExtra = relElems(root.recs);
  }
  extra.sort((a, b) => a[0] - b[0] || a[2] - b[2]);
  const out = pkg.take();
  const shared = sst !== undefined;
  out.push(
    {
      path: 'xl/workbook.xml',
      data: workbookXml(
        info,
        extra.map((e) => [e[0], e[1]]),
      ),
    },
    { path: 'xl/styles.xml', data: stylesXml(wb.styles, pb?.tableStyles, pb?.stylesTail) },
  );
  if (sst) out.push({ path: 'xl/sharedStrings.xml', data: sharedStringsXml(sst) });
  out.push(
    { path: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml(n, shared, relElems(rels.recs)) },
    {
      path: '[Content_Types].xml',
      data: contentTypesXml(n, shared, pkg.defaults, pkg.overrides),
    },
    { path: '_rels/.rels', data: rootRelsXml(rootExtra) },
    { path: 'docProps/core.xml', data: corePropsXml(book) },
    { path: 'docProps/app.xml', data: appPropsXml(book) },
  );
  return out;
}

/** Index of the active (selected) tab: first visible sheet with tabSelected, else first visible. */
export function activeTabOf(
  sheets: readonly { state: string; views: readonly { tabSelected?: boolean | undefined }[] }[],
): number {
  let first = -1;
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i]!;
    if (s.state !== 'visible') continue;
    if (s.views[0]?.tabSelected) return i;
    if (first < 0) first = i;
  }
  if (first < 0) throw new Error('At least one worksheet must be visible');
  return first;
}

export function sheetEntry(ws: Worksheet): SheetEntry {
  return { name: ws.name, sheetId: ws.id, state: ws.state, autoFilter: ws.autoFilter };
}

function* pieces(wb: Workbook, o: ResolvedWriteOptions): Generator<Piece, void, undefined> {
  const sheets = wb.worksheets;
  if (sheets.length === 0) throw new Error('Workbook has no worksheets');
  const active = activeTabOf(sheets);
  const sst = o.strings === 'shared' ? new SharedStrings() : undefined;
  const pkg = new PackageParts();
  const kept = planPreserved(wb, pkg);
  const ids = sheetIds(wb);
  let formulas = false;

  for (let si = 0; si < sheets.length; si++) {
    const ws = sheets[si]!;
    const ctx = newSheetCtx({ sst, pkg, ws, index: si, streaming: false });
    const path = ctx.path;
    const rows = ws.rowData;
    let outline = 0;
    for (let r = 0; r < rows.length; r++) {
      const ol = rows[r]?.outlineLevel;
      if (ol !== undefined && ol > outline) outline = ol;
    }
    ctx.dimension = ws.dimensions?.ref;
    ctx.tabSelected = si === active;
    ctx.outlineLevelRow = outline;
    let buf = sheetHead(ctx);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (row === undefined) continue;
      buf += encodeRow(row, r, ctx);
      if (buf.length >= CHUNK) {
        yield { path, data: buf, last: false };
        buf = '';
      }
    }
    yield { path, data: buf + sheetTail(ctx), last: true };
    for (const p of afterSheetParts(ctx, pkg)) yield { path: p.path, data: p.data, last: true };
    if (ctx.formulas) formulas = true;
  }

  const book: BookInfo = {
    sheets: sheets.map((ws, i) => ({ ...sheetEntry(ws), sheetId: ids[i]! })),
    date1904: wb.date1904,
    sharedStrings: sst !== undefined,
    formulas,
    activeTab: active,
    creator: wb.creator,
    lastModifiedBy: wb.lastModifiedBy,
    created: wb.created,
    modified: wb.modified,
    definedNames: wb.definedNames,
    theme: o.theme,
  };
  for (const p of bookParts(wb, book, pkg, sst, kept))
    yield { path: p.path, data: p.data, last: true };
}

/**
 * Serialize a workbook to XLSX bytes. The date system is the workbook's own (`wb.date1904`);
 * `opts.date1904` is ignored here because cell serials were computed when values were stored.
 */
export function writeXlsx(wb: Workbook, opts?: WriteOptions): Uint8Array {
  const o = resolveWriteOptions(opts);
  const chunks: Uint8Array[] = [];
  const zip = new ZipOut((c) => void chunks.push(c), o.compression);
  for (const p of pieces(wb, o)) zip.write(p.path, p.data, p.last);
  zip.end();
  return concatChunks(chunks);
}

/**
 * Stream a workbook as XLSX. Pull-based: sheet XML is generated one piece (~128K chars of rows)
 * per pull, so a slow consumer stalls generation instead of buffering the whole file.
 * The workbook must not be modified until the stream has finished.
 */
export function xlsxStream(wb: Workbook, opts?: WriteOptions): ReadableStream<Uint8Array> {
  const o = resolveWriteOptions(opts);
  let gen: Generator<Piece, void, undefined> | undefined;
  let zip: ZipOut | undefined;
  return new ReadableStream<Uint8Array>({
    start(controller) {
      gen = pieces(wb, o);
      // Native CompressionStream (when available) deflates off the main thread; chunks are
      // enqueued as the compressor produces them.
      zip = new ZipOut(
        (c) => c.length > 0 && controller.enqueue(c),
        o.compression,
        nativeCompressor(),
      );
    },
    async pull(controller) {
      for (;;) {
        const r = gen!.next();
        if (r.done) {
          zip!.end();
          await zip!.idle();
          controller.close();
          return;
        }
        zip!.write(r.value.path, r.value.data, r.value.last);
        const w = zip!.ready();
        if (w) await w;
        if ((controller.desiredSize ?? 0) <= 0) return;
      }
    },
    cancel() {
      gen?.return();
    },
  });
}
