// XLSX writer public API (package root): every writer runs all built-in features. The
// feature-agnostic implementations (book.ts, stream-writer.ts) only run `options.features`;
// the `./lite` entry exports those as is.

import type { WriteOptions } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import { WRITERS, withAll } from '../features/registry.js';
import { writeXlsx as writeCore, xlsxStream as xlsxStreamCore } from './book.js';
import { XlsxWriter as XlsxWriterCore, type XlsxWriterOptions } from './stream-writer.js';

export { listPreservedParts, type PreservedPart } from './preserve.js';
export {
  type ChunkSink,
  StreamSheet,
  type StreamSheetOptions,
  type XlsxWriterOptions,
} from './stream-writer.js';

function all<T extends WriteOptions>(opts: T | undefined): T {
  return { ...(opts as T), features: withAll(WRITERS, opts?.features) };
}

/**
 * Serialize a workbook to XLSX bytes, with every built-in feature (tables, notes, images,
 * conditional formatting, data validation, print setup, sheet protection, theme);
 * `opts.features` adds custom feature writers. The date system is the workbook's own
 * (`wb.date1904`); `opts.date1904` is ignored here because cell serials were computed when
 * values were stored.
 */
export function writeXlsx(wb: Workbook, opts?: WriteOptions): Uint8Array {
  return writeCore(wb, all(opts));
}

/**
 * Stream a workbook as XLSX (every built-in feature, see writeXlsx). Pull-based: sheet XML is
 * generated one piece (~128K chars of rows) per pull, so a slow consumer stalls generation
 * instead of buffering the whole file. The workbook must not be modified until the stream has
 * finished.
 */
export function xlsxStream(wb: Workbook, opts?: WriteOptions): ReadableStream<Uint8Array> {
  return xlsxStreamCore(wb, all(opts));
}

/** Streaming XLSX writer with every built-in feature (see createXlsxWriter). */
export class XlsxWriter extends XlsxWriterCore {
  constructor(opts?: XlsxWriterOptions) {
    super(all(opts));
  }
}

/**
 * Create a streaming XLSX writer (every built-in feature). Its workbook is new and private:
 * content preserved by `readXlsx` (preserveUnknown) is never carried over — use writeXlsx /
 * xlsxStream for that.
 * ```ts
 * const w = createXlsxWriter({ sink: (chunk) => file.write(chunk) });
 * const s = w.addWorksheet('Data', { columns: [{ header: 'Id', key: 'id', width: 10 }] });
 * for (const r of rows) s.addRow(r);
 * await s.commit();
 * await w.end();
 * ```
 */
export function createXlsxWriter(opts?: XlsxWriterOptions): XlsxWriter {
  return new XlsxWriter(opts);
}
