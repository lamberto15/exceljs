# turbosheet

## 0.1.0

Initial release. A clean rewrite inspired by and partly derived from ExcelJS (MIT).

- XLSX writing: `writeXlsx` (sync), `xlsxStream` (web `ReadableStream`), `createXlsxWriter` (row-by-row with backpressure); `low-memory` / `balanced` / `max-speed` profiles, deflate level and shared/inline string control.
- XLSX reading: `readXlsx` (whole workbook, sync) and `xlsxRows` (streaming rows from bytes, Node streams or web streams), with zip-bomb limits and no XML entity expansion.
- Workbook model with typed-array cell storage, interned immutable styles, rich text, hyperlinks, formulas (cached results, no evaluation), merges, frozen panes, auto-filter, column and row properties.
- Features: tables, notes, pictures, conditional formatting, data validation, defined names, print area and titles, print setup, headers and footers, sheet protection.
- Round-trip preservation of content the model does not represent (charts, pivot tables, shapes, VBA, threaded comments, themes, unknown XML).
- `turbosheet/lite`: the same API with features opt-in, 20 KB write-only (min + brotli).
- `turbosheet/csv`: streaming RFC 4180 parser and writer with optional formula-injection escaping.
- `turbosheet/json`: worksheet to and from objects, arrays and NDJSON.
- `turbosheet/crypto`: password-protected files (agile read/write, standard read).
- Not yet: chart writing, formula evaluation, gradient fills, row/column insert and splice, verification in Microsoft Excel (compatibility is checked with LibreOffice and ExcelJS in CI).
