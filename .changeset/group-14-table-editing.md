---
"turbosheet": minor
---

Table editing (#103): `addTableColumn`, `removeTableColumns`, `renameTableColumn`, `removeTableRows`, `setTableHeaderRow`, `setTableTotalsRow`. These are free functions covering ExcelJS's `table.addColumn` / `removeColumns` / `removeRows` / column rename / `headerRow` / `totalsRow` + `commit()`. As in ExcelJS, the table keeps its top-left cell and only its own cells move. Renaming a column updates its structured references in every formula of the workbook. Removing a column turns references to it into `#REF!`, as Excel does.

Fix: the `[@Col]` structured-reference shorthand (also `T[@Col]`, `T[@[Col]]`, `T[@]`) is now written as `[[#This Row],[Col]]`. Excel drops a table whose calculated column uses `@` in the file, and shows `#REF!` in cells that do.
