---
"turbosheet": patch
---

ExcelJS parity additions (see docs/exceljs-parity.md):

- Fonts model `family` and `scheme` (`'major' | 'minor' | 'none'`): read from styles, rich text and notes, written back, and part of the style identity. The default font is Calibri 11 family 2 scheme minor, as Excel writes it. Editing and saving an Excel file no longer drops `<scheme val="minor"/>` from its fonts.
- Alignment accepts the ExcelJS aliases `vertical: 'middle'` (stored as `'center'`) and `textRotation: 'vertical'` (stored as 255).
- `arraysToSheet` / `objectsToSheet`: empty strings (such as empty CSV fields from `parseCsv`) leave the cell empty; pass `{ emptyAs: 'string' }` for the previous behaviour.
- `readXlsx`: hyperlinks on formula cells are kept and written back while the cell still holds that formula (they used to be lost on a round trip).
