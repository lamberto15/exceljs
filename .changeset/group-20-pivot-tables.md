---
"turbosheet": minor
---

Pivot table authoring (#115): `addPivotTable(ws, { name, source, ref, rows, columns, values, filters })`.

- Writes the pivot cache (definition + records, refreshed when Excel opens the file) and the pivot table part.
- Fills the report cells in Excel's compact layout, so readers that don't refresh pivots show the totals too. Checked cell by cell against what Excel for Mac displays.
- Supports nested row fields, one column field, several value fields (sum, count, average, max, min, product, countNums, stdDev(p), var(p)) and report filters.
- New feature half: `pivotTablesWriter`. Pivot tables in read files are still kept raw.
