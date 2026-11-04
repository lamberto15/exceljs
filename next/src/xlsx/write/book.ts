// In-memory workbook → XLSX: writeXlsx (buffered) and xlsxStream (pull-based ReadableStream).
// Both drive the same generator, which yields sheet XML in ~CHUNK-sized pieces so the stream
// variant only materializes one piece at a time.

import { SharedStrings } from '../../core/sharedstrings.js';
import type { WriteOptions } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import type { Worksheet } from '../../core/worksheet.js';
import { type ResolvedWriteOptions, resolveWriteOptions } from './options.js';
import {
  appPropsXml,
  type BookInfo,
  contentTypesXml,
  corePropsXml,
  rootRelsXml,
  type SheetEntry,
  sharedStringsXml,
  sheetPath,
  sheetRelsPath,
  sheetRelsXml,
  stylesXml,
  workbookRelsXml,
  workbookXml,
} from './parts.js';
import { encodeRow, hasExternalLinks, newSheetCtx, sheetHead, sheetTail } from './sheet.js';
import { concatChunks, ZipOut } from './zipout.js';

/** Sheet XML is handed to the zip in pieces of about this many UTF-16 units. */
export const CHUNK: number = 128 * 1024;

interface Piece {
  path: string;
  data: string;
  last: boolean;
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
  const shared = sst !== undefined;
  let formulas = false;

  yield { path: '[Content_Types].xml', data: contentTypesXml(sheets.length, shared), last: true };
  yield { path: '_rels/.rels', data: rootRelsXml(), last: true };
  yield {
    path: 'xl/_rels/workbook.xml.rels',
    data: workbookRelsXml(sheets.length, shared),
    last: true,
  };

  for (let si = 0; si < sheets.length; si++) {
    const ws = sheets[si]!;
    const path = sheetPath(si);
    const ctx = newSheetCtx(sst);
    const rows = ws.rowData;
    let outline = 0;
    for (let r = 0; r < rows.length; r++) {
      const ol = rows[r]?.outlineLevel;
      if (ol !== undefined && ol > outline) outline = ol;
    }
    let buf = sheetHead(ws, {
      dimension: ws.dimensions?.ref,
      tabSelected: si === active,
      outlineLevelRow: outline,
    });
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (row === undefined) continue;
      buf += encodeRow(row, r, ctx);
      if (buf.length >= CHUNK) {
        yield { path, data: buf, last: false };
        buf = '';
      }
    }
    yield { path, data: buf + sheetTail(ws, ctx), last: true };
    if (hasExternalLinks(ctx)) {
      yield { path: sheetRelsPath(si), data: sheetRelsXml(ctx.links), last: true };
    }
    if (ctx.formulas) formulas = true;
  }

  const book: BookInfo = {
    sheets: sheets.map(sheetEntry),
    date1904: wb.date1904,
    sharedStrings: shared,
    formulas,
    activeTab: active,
    creator: wb.creator,
    lastModifiedBy: wb.lastModifiedBy,
    created: wb.created,
    modified: wb.modified,
  };
  yield { path: 'xl/workbook.xml', data: workbookXml(book), last: true };
  yield { path: 'xl/styles.xml', data: stylesXml(wb.styles), last: true };
  if (sst) yield { path: 'xl/sharedStrings.xml', data: sharedStringsXml(sst), last: true };
  yield { path: 'docProps/core.xml', data: corePropsXml(book), last: true };
  yield { path: 'docProps/app.xml', data: appPropsXml(book), last: true };
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
  let out: Uint8Array[] = [];
  return new ReadableStream<Uint8Array>({
    start() {
      gen = pieces(wb, o);
      zip = new ZipOut((c) => void out.push(c), o.compression);
    },
    pull(controller) {
      for (;;) {
        const r = gen!.next();
        if (r.done) zip!.end();
        else zip!.write(r.value.path, r.value.data, r.value.last);
        if (out.length > 0) {
          const batch = out;
          out = [];
          for (const c of batch) controller.enqueue(c);
        }
        if (r.done) {
          controller.close();
          return;
        }
        if ((controller.desiredSize ?? 0) <= 0) return;
      }
    },
    cancel() {
      gen?.return();
    },
  });
}
