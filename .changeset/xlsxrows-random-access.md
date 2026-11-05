---
"turbosheet": minor
---

`xlsxRows` reads `Uint8Array` / `ArrayBuffer` and (new) `Blob` / `File` sources by the zip central directory: the content types, rels, workbook, styles and shared strings are loaded first and each sheet is inflated as it is parsed, so files that store worksheets before `styles.xml` (Excel 16.x, ExcelJS, `createXlsxWriter`) stream at constant memory instead of buffering the whole sheet. In Node, `xlsxRows(await fs.openAsBlob(path))` reads a 1M x 35 `createXlsxWriter` file (176 MB) with default limits at 156 MB peak RSS, rows arriving from the start; before, it failed with `ZipLimitError` or buffered 1.4 GB. Blob entries are inflated with `DecompressionStream` (fflate fallback). Byte inputs yield to the event loop about every 50 ms while inflating. Stream sources are unchanged. Closes #86.
