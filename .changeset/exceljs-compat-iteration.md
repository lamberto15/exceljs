---
"turbosheet": minor
---

ExcelJS migration compatibility. `ws.eachRow`, `row.eachCell`, the new `column.eachCell` and `eachColumnCell` accept an optional first-argument `{ includeEmpty: true }` with ExcelJS semantics: rows 1 to `ws.rowCount`, or for `row.eachCell` columns 1 to the row's last cell, with `null` for empty cells. A call without a callback now throws a clear `TypeError` instead of "cb is not a function". `ReadOptions.maxRows` / `maxCols` are enforced by `readXlsx` and `xlsxRows`: the first cell beyond either limit throws the new exported `SheetLimitError` (`sheet`, `limit`, `max`), before the rest of the sheet is parsed. Cell setters (`setCell`, `addRow`, `cell.value`, the streaming writer) accept ExcelJS `{ sharedFormula: 'B1', result? }` values (type `CellValueInput`): they are stored and written as the master cell's formula with relative references moved by the offset. Every bundle with a `Workbook` grows by about 1 KB. Closes #80.
