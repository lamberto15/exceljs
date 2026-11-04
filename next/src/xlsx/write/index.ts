// XLSX writer public API.
export { writeXlsx, xlsxStream } from './book.js';
export { type ResolvedWriteOptions, resolveWriteOptions } from './options.js';
export {
  type ChunkSink,
  createXlsxWriter,
  StreamSheet,
  type StreamSheetOptions,
  XlsxWriter,
  type XlsxWriterOptions,
} from './stream-writer.js';
