---
"turbosheet": minor
---

Dynamic-array formulas (FILTER, SORT, UNIQUE, SEQUENCE, ...) are written with Excel's dynamic-array metadata (`cm="1"` + `xl/metadata.xml`), so Excel spills them instead of adding an implicit-intersection `@`; `Formula.dynamic` forces / disables it and is read back. `copyWorksheet(ws, target?, name?)` copies a sheet within or across workbooks; `copyRange(src, range, dst, to)` copies cells, styles and merges like copy-paste. Column keys may be dotted paths (`'address.city'`) in `addRow`.
