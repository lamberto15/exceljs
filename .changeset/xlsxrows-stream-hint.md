---
"turbosheet": patch
---

`xlsxRows`: when a stream source has to buffer a sheet stored before `styles.xml` / `sharedStrings.xml` (Excel, ExcelJS and `createXlsxWriter` files) and the sheet hits a byte limit, the `ZipLimitError` message now says to pass a Blob (`await fs.openAsBlob(path)`, a `File` in browsers) or the bytes, which stream it at constant memory. `docs/api.md` recommends saving uploads to disk and reading them with `openAsBlob`. Closes #89.
