---
"turbosheet": minor
---

ExcelJS compatibility, group 13 (#99-#102):

- `cell.hyperlink` / `cell.tooltip` on `CellRef`. Setting `hyperlink` on a string or rich-text cell turns it into a link. On a link cell it changes the target, and `undefined` turns the cell back into text. Other cells are left unchanged, as in ExcelJS.
- ExcelJS row style codes: `ws.addRow(values, 'i' | 'i+' | 'n')`, `ws.addRows(rows, code)` and `insertRow(s)(ws, pos, values, 'i' | 'o' | 'i+' | 'o+' | 'n')` copy the format of the row above or below. The `+` variants include empty cells.
- `fillFormula` ranges, `{ sharedFormula }` values and shared formulas read from a file are now written as Excel shared formulas (`<f t="shared">`). For a 100k-row filled column the file is about a third smaller. Cells edited after the fill are written as their own formulas. Reading still expands shared formulas to one formula per cell.
- New sheet view fields: `style` (`'pageLayout'` / `'pageBreakPreview'`), `showRuler`, `zoomScaleNormal` and `activePane` (read and written).
