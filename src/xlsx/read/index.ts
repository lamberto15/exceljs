// XLSX reader public API (package root): `readXlsx` with every built-in feature reader. The
// feature-agnostic implementation lives in ./reader.ts (exported as is by the `./lite` entry).

import type { ReadOptions } from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import { READERS, withAll } from '../features/registry.js';
import { readXlsx as readCore } from './reader.js';

export { ZipLimitError, ZipPathError } from '../../zip/limits.js';

export {
  EncryptedFileError,
  type SheetEntry,
  translateFormula,
  XlsxParseError,
  type XlsxRow,
  xlsxRows,
} from './reader.js';

/**
 * Read a whole XLSX file into a {@link Workbook} (synchronous), modelling every built-in
 * feature (tables, notes, images, conditional formatting, data validation, print setup, sheet
 * protection). `opts.features` adds custom feature readers. See the `./lite` entry for an
 * opt-in variant with a smaller bundle.
 * - `opts.sheets`: only load these sheet names (case-insensitive); others are not created.
 * - `opts.limits`: zip-bomb limits (entry count, sizes, compression ratio).
 * - `opts.preserveUnknown` (default true): keep what the model does not represent (charts,
 *   shapes, pivot tables, threaded comments, VBA, custom XML, unknown worksheet elements, ...)
 *   so `writeXlsx` / `xlsxStream` write it back (see `listPreservedParts`). Kept parts stay
 *   compressed in the input buffer (inflated when written): do not mutate `data` afterwards.
 *   `false` drops all of it (and does not retain the input buffer).
 * `profile` is ignored.
 * Throws ZipLimitError / ZipPathError for hostile archives, XlsxParseError for broken XML and
 * EncryptedFileError for password-protected files (decrypt with `decryptXlsx` from `./crypto`).
 */
export function readXlsx(data: Uint8Array | ArrayBuffer, opts: ReadOptions = {}): Workbook {
  return readCore(data, { ...opts, features: withAll(READERS, opts.features) });
}
