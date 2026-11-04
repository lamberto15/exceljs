---
"turbosheet": minor
---

Row / column structure edits: `insertRows`, `insertRow`, `deleteRows`, `spliceRows`, `insertColumns`, `deleteColumns`, `spliceColumns` and `duplicateRow`. Like Excel's Insert / Delete Sheet Rows / Columns, every reference to the edited sheet follows its cells: formulas on all sheets, array ranges, defined names (print area / titles), chart series, merges, conditional formatting and data validation, tables (inserted rows get calculated columns, inserted columns become table columns), notes, picture and chart anchors and the auto-filter. Edits are validated first and throw without changing anything when they would push cells off the sheet or break a table.
