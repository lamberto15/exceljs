# turbosheet

## 0.3.0

### Minor Changes

- 616bbd1: Add formula evaluation as a separate entry point, `turbosheet/formula`: `parseFormula`, `evaluate(wb)` (dependency-ordered, circular-reference detection, optional iterative calculation, results stored into `Formula.result`), `evaluateCell`, `evaluateFormula` and `createEvaluator({ functions })` for per-instance custom functions. 87 built-in functions (math and conditional aggregation, logical, lookup incl. XLOOKUP/XMATCH, text incl. TEXT formats, dates, info) with Excel's coercion, comparison and wildcard rules. The core bundles do not change.

### Patch Changes

- ddcd58b: ExcelJS parity additions (see docs/exceljs-parity.md):
  
  - Fonts model `family` and `scheme` (`'major' | 'minor' | 'none'`): read from styles, rich text and notes, written back, and part of the style identity. The default font is Calibri 11 family 2 scheme minor, as Excel writes it. Editing and saving an Excel file no longer drops `<scheme val="minor"/>` from its fonts.
  - Alignment accepts the ExcelJS aliases `vertical: 'middle'` (stored as `'center'`) and `textRotation: 'vertical'` (stored as 255).
  - `arraysToSheet` / `objectsToSheet`: empty strings (such as empty CSV fields from `parseCsv`) leave the cell empty; pass `{ emptyAs: 'string' }` for the previous behaviour.
  - `readXlsx`: hyperlinks on formula cells are kept and written back while the cell still holds that formula (they used to be lost on a round trip).
- 8d33b5e: Fixes found by porting the ExcelJS test suite (see docs/exceljs-parity.md):
  
  - Styles: a negative `alignment.textRotation` (-90..-1, the ExcelJS convention) is stored in Excel's encoding (90 + |deg|, e.g. -45 → 135) instead of being written as an invalid negative attribute that Excel rejects; out-of-range rotations and negative `indent` values are dropped.
  - Styles: 6-digit `argb` colors (`'FF0000'`) get an opaque alpha (`'FFFF0000'`) in the model and in every written `rgb` attribute, so they are valid and read back unchanged.
  - `Worksheet.addRow` (and the streaming writer's `addRow`) no longer leaves a half-written row behind when a value is rejected, which shifted every later row down by one.
  - `Workbook.removeWorksheet` keeps sheet-scoped defined names (print area, print titles, local names) on their sheet: later sheets' indexes are shifted and the removed sheet's names are dropped.
- 12e23ea: Write the `_xlfn.` prefix for Excel future functions so Excel no longer shows #NAME?; strip it on read.

## 0.2.1

### Patch Changes

- f870704: Round-trip fidelity: `readXlsx` → `writeXlsx` now keeps autoFilter criteria (filter columns, sort state), named cell styles (`cellStyleXfs` / `cellStyles` and each cell format's style link, exposed read-only as `Xf.xfId`), `sheetFormatPr` attributes (default column width, base width, zero height, ...), form controls on sheets that also have notes (VML shapes merged with the regenerated notes; Excel's `mc:AlternateContent`-wrapped `<controls>` and their ctrlProps parts), the tab position of chartsheets / dialogsheets (activeTab and sheet-scoped names follow), and the queryTable link of unchanged query tables. See "Round-trip preservation" in docs/features.md for what happens when you edit nearby.

## 0.2.0

### Minor Changes

- ab291c6: Chart writing: `addChart(ws, chart)` / `ws.charts` writes native bar, column, line, pie, doughnut,
  scatter and area charts (titles, legend, axis titles / bounds / number formats, stacking, data
  labels, series colors, markers, smoothing, built-in styles) with caches of the referenced cells,
  sharing the sheet's drawing with pictures and with charts kept from read files. New opt-in
  `chartsWriter` for `turbosheet/lite`. Reading keeps charts raw (round-trip preservation).
- 4e42e76: Add column auto-fit: `autoFitColumns(ws, opts?)` sizes columns from their displayed contents (number formats, dates, rich text, bold/size, wrap, merges), and `estimateColumnWidths(rows, opts?)` computes widths from a sample for the streaming writer. Available from both `turbosheet` and `turbosheet/lite`; not included in a bundle unless imported.

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
