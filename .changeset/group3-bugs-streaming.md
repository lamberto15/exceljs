---
"turbosheet": minor
---

`localDates` option (`new Workbook({ localDates: true })`, `ReadOptions.localDates`, also for `xlsxRows`): JS Dates are converted by their local wall clock instead of UTC, so `new Date(2024, 0, 1)` is Jan 1 in Excel in every time zone; `dateToSerial` / `serialToDate` take a `local` flag. `xlsxRows` reports each sheet's merged ranges through `onSheetEnd`. Regression tests for seven reported ExcelJS bugs (header/footer options, sheetPr order with fitToPage, notes in tables, streaming dates, headerless tables, streamed rich text, rich-text hyperlinks).
