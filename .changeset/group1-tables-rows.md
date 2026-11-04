---
"turbosheet": minor
---

Tables: `tableRows(ws, name)` reads a table's data rows as objects keyed by column name, `addTableRows(ws, name, rows)` appends rows (above the totals row; cells below move down, calculated columns are filled), `getTable(ws, name)`, and `addTable` columns accept `width`. Outline groups: `groupRows` / `groupColumns` with `{ collapsed }`, and a `collapsed` flag on rows (`setRowProps`) and columns, written and read. Iterators: `ws.rows()` and `row.cells()`.
