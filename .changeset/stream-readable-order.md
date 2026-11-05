---
"turbosheet": minor
---

`writeXlsx` and `xlsxStream` now store `[Content_Types].xml`, `_rels/.rels`, `workbook.xml`, its rels, `styles.xml` and `sharedStrings.xml` before the worksheets, so `xlsxRows` streams turbosheet's own files at constant memory instead of buffering the whole sheet (rows arrive from the first chunks). `xlsxRows` also applies `maxEntryBytes` / `maxTotalBytes` only to parts it buffers: a sheet parsed as it inflates is bounded by `maxCompressionRatio` only, so large legitimate exports no longer hit the 1 GiB entry limit. Docs corrected: Excel 16.x and ExcelJS store worksheets before `styles.xml`, so their sheets are buffered by `xlsxRows`; `createXlsxWriter` output too. Closes #74.
