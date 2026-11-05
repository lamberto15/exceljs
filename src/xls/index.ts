// `./xls` entry: read legacy Excel 97-2003 workbooks (.xls, BIFF8) into the same Workbook
// model as readXlsx. A separate entry so the main bundles do not grow.
//
//   import { readXls } from 'turbosheet/xls';
//   const wb = readXls(new Uint8Array(await file.arrayBuffer()));

export { BiffParseError } from '../biff/common.js';
export type { BinaryReadOptions } from '../biff/model.js';
export { CfbError } from '../crypto/cfb.js';
export { EncryptedXlsError, readXls } from './reader.js';
