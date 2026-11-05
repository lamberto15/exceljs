// `./xlsb` entry: read Excel binary workbooks (.xlsb) into the same Workbook model as
// readXlsx. A separate entry so the main bundles do not grow.
//
//   import { readXlsb } from 'turbosheet/xlsb';
//   const wb = readXlsb(new Uint8Array(await file.arrayBuffer()));

export { BiffParseError } from '../biff/common.js';
export type { BinaryReadOptions } from '../biff/model.js';
export { ZipLimitError, ZipPathError } from '../zip/limits.js';
export { readXlsb } from './reader.js';
