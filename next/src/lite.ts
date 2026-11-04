// `./lite` entry: the same API and Workbook model as the package root, but no feature is
// enabled by default — pass the halves you use in `options.features`:
//
//   import { Workbook, writeXlsx, readXlsx, tablesWriter, tablesReader } from 'turbosheet/lite';
//   writeXlsx(wb, { features: [tablesWriter] });
//   readXlsx(buf, { features: [tablesReader] });
//
// Writing a model field of a feature that is not enabled throws (no silent data loss); reading
// without a feature keeps its elements raw and writes them back unchanged (preserveUnknown).
// Never import features/registry.ts from here (it would pull every feature in).

export * from './core/index.js';
export * from './xlsx/features/index.js';
export {
  EncryptedFileError,
  readXlsx,
  type SheetEntry,
  translateFormula,
  XlsxParseError,
  type XlsxRow,
  xlsxRows,
} from './xlsx/read/reader.js';
export { writeXlsx, xlsxStream } from './xlsx/write/book.js';
export { type ResolvedWriteOptions, resolveWriteOptions } from './xlsx/write/options.js';
export { listPreservedParts, type PreservedPart } from './xlsx/write/preserve.js';
export {
  type ChunkSink,
  createXlsxWriter,
  StreamSheet,
  type StreamSheetOptions,
  XlsxWriter,
  type XlsxWriterOptions,
} from './xlsx/write/stream-writer.js';
export { ZipLimitError, ZipPathError } from './zip/limits.js';
