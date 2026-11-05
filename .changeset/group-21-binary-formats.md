---
"turbosheet": minor
---

Binary workbooks, group 21 (#116, #117): two new read-only entries produce the same `Workbook` model as `readXlsx`.

- **`turbosheet/xlsb`: `readXlsb(bytes, opts?)`** reads Excel binary workbooks (BIFF12 records in the .xlsx zip package).
- **`turbosheet/xls`: `readXls(bytes, opts?)`** reads legacy Excel 97-2003 files (BIFF8 in an OLE compound file). The compound-file reader is shared with `turbosheet/crypto`.
- Both read sheets (order, names, hidden state), cell values, shared strings incl. rich text, dates (numbers with a date format, as in readXlsx), styles (number formats, fonts, fills, borders, alignment, protection; the full theme / RGB colors of `.xls` files saved by Excel 2007+), merges, column widths, row heights, hidden rows / columns, outline levels, defined names, hyperlinks and the 1904 date system.
- Formulas are decompiled to text with their cached results: references (incl. 3-D, whole rows / columns, shared formulas), array formulas, operators, constants, array constants, built-in and future functions, defined names. A formula using an unsupported token (structured references, external links, ...) keeps its cached value.
- Errors: `BiffParseError`, `EncryptedXlsError` (password-protected `.xls`). An encrypted `.xlsb` is read with `readXlsb(await decryptXlsx(bytes, password))`.
- Not read: charts, pictures, comments, data validation, conditional formatting, tables, pivot tables, sheet views, print setup, document properties. Excel 95 and older `.xls` files are rejected.
- The existing bundles do not change; `./xlsb` is 23.5 KB and `./xls` 20.2 KB (min + brotli).
