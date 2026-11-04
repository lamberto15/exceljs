---
"turbosheet": patch
---

Fixes found by porting the ExcelJS test suite (see docs/exceljs-parity.md):

- Styles: a negative `alignment.textRotation` (-90..-1, the ExcelJS convention) is stored in Excel's encoding (90 + |deg|, e.g. -45 → 135) instead of being written as an invalid negative attribute that Excel rejects; out-of-range rotations and negative `indent` values are dropped.
- Styles: 6-digit `argb` colors (`'FF0000'`) get an opaque alpha (`'FFFF0000'`) in the model and in every written `rgb` attribute, so they are valid and read back unchanged.
- `Worksheet.addRow` (and the streaming writer's `addRow`) no longer leaves a half-written row behind when a value is rejected, which shifted every later row down by one.
- `Workbook.removeWorksheet` keeps sheet-scoped defined names (print area, print titles, local names) on their sheet: later sheets' indexes are shifted and the removed sheet's names are dropped.
