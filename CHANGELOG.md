# turbosheet

## 0.6.1

### Patch Changes

- e83c470: `xlsxRows`: when a stream source has to buffer a sheet stored before `styles.xml` / `sharedStrings.xml` (Excel, ExcelJS and `createXlsxWriter` files) and the sheet hits a byte limit, the `ZipLimitError` message now says to pass a Blob (`await fs.openAsBlob(path)`, a `File` in browsers) or the bytes, which stream it at constant memory. `docs/api.md` recommends saving uploads to disk and reading them with `openAsBlob`. Closes #89.

## 0.6.0

### Minor Changes

- a6ff5c4: `xlsxRows` reads `Uint8Array` / `ArrayBuffer` and (new) `Blob` / `File` sources by the zip central directory: the content types, rels, workbook, styles and shared strings are loaded first and each sheet is inflated as it is parsed, so files that store worksheets before `styles.xml` (Excel 16.x, ExcelJS, `createXlsxWriter`) stream at constant memory instead of buffering the whole sheet. In Node, `xlsxRows(await fs.openAsBlob(path))` reads a 1M x 35 `createXlsxWriter` file (176 MB) with default limits at 156 MB peak RSS, rows arriving from the start; before, it failed with `ZipLimitError` or buffered 1.4 GB. Blob entries are inflated with `DecompressionStream` (fflate fallback). Byte inputs yield to the event loop about every 50 ms while inflating. Stream sources are unchanged. Closes #86.

## 0.5.0

### Minor Changes

- 9cededf: CSV read typing and encodings. `parseCsv` / `csvRows` / `CsvParser` take new opt-in options: `parseBooleans` (unquoted `true` / `false`, any case), `parseErrors` (unquoted `#N/A`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!` become `{ error }`, which `arraysToSheet` now writes as error cells), `dateFormats` (e.g. `['DD/MM/YYYY HH:mm']`, tried in order after ISO 8601, with an explicit day/month order and calendar validation) and `localDates` (dates without an offset by local wall clock instead of UTC). `CsvValue` now includes `boolean` and `CellError`. ISO dates are calendar-checked too: `2024-02-31` stays text instead of rolling over to March 2. `csvRows` takes `encoding` (any `TextDecoder` label, e.g. `'windows-1252'`, `'utf-16le'`), sniffs a UTF-8 / UTF-16LE / UTF-16BE BOM by default (UTF-16 files from Excel read correctly with no option), and `encoding: 'auto'` falls back to windows-1252 when the first chunk with non-ASCII bytes is not valid UTF-8. The `./csv` entry grows from 2.38 to 3.14 KB. Closes #81.
- fab308c: ExcelJS migration compatibility. `ws.eachRow`, `row.eachCell`, the new `column.eachCell` and `eachColumnCell` accept an optional first-argument `{ includeEmpty: true }` with ExcelJS semantics: rows 1 to `ws.rowCount`, or for `row.eachCell` columns 1 to the row's last cell, with `null` for empty cells. A call without a callback now throws a clear `TypeError` instead of "cb is not a function". `ReadOptions.maxRows` / `maxCols` are enforced by `readXlsx` and `xlsxRows`: the first cell beyond either limit throws the new exported `SheetLimitError` (`sheet`, `limit`, `max`), before the rest of the sheet is parsed. Cell setters (`setCell`, `addRow`, `cell.value`, the streaming writer) accept ExcelJS `{ sharedFormula: 'B1', result? }` values (type `CellValueInput`): they are stored and written as the master cell's formula with relative references moved by the offset. Every bundle with a `Workbook` grows by about 1 KB. Closes #80.
- bc338f5: ExcelJS metadata parity. Several sheet views per sheet: `ws.views = [{...}, { zoomScale: 50 }]` writes one `<sheetView>` per view, one per Excel window (`workbookViewId`, which defaults to the view's position in `ws.views`). Writers add `<workbookView>`s when needed, every `wb.views` entry is written, and all views and windows are read back. Sheet `properties` (`addWorksheet(name, { properties, views })`, `ws.properties`): `tabColor` (same value as `ws.tabColor`), `defaultRowHeight`, `defaultColWidth`, `outlineLevelRow`, `outlineLevelCol` and `dyDescent` are written to `<sheetPr>` / `<sheetFormatPr>` (the `x14ac` namespace is declared Ignorable when needed). Outline levels are written as the higher of the property and the levels in the rows / columns, and are not read back. Document properties `language` and `revision` are added. Notes take `margins` (`insetmode`, `inset` in cm), `protection` (`locked`, `lockText`) and `editAs`, written to the VML like ExcelJS and read back. The ExcelJS parity docs are refreshed. Closes #82.
- fde8cba: `writeXlsx` and `readXlsx` use native zlib in Node, Bun and Deno (looked up at runtime with `process.getBuiltinModule`; browsers keep fflate). `writeXlsx` deflates with `zlib.deflateRawSync` in ~1 MiB sync-flushed batches: 200k x 35 balanced 10.9 s to 7.3 s with ~1.5% smaller files; `max-speed` 3.8 s to 1.7 s. `readXlsx` inflates parts up to 64 MiB with `zlib.inflateRawSync` (5.8 s to 4.8 s on the same file), with zip-bomb limits enforced through `maxOutputLength`; larger parts still stream through fflate. Closes #75.
- 6dbeb00: Streaming writers (`createXlsxWriter`, `xlsxStream`) now honor `compression` / `profile` in Node, Bun and Deno: entries are deflated off the main thread by `node:zlib` at the requested level (looked up at runtime with `process.getBuiltinModule`, no static `node:*` import), instead of `CompressionStream`, which has no level. `max-speed` streaming is ~2.5x faster (300k rows: 1.98 s to 0.76 s). Zip CRC-32 uses `zlib.crc32` when available (~5x faster than the JS table), so the default level is ~15% faster too. Browsers keep `CompressionStream` (platform default level, now documented). Closes #71.
- b40478e: `writeXlsx` and `xlsxStream` now store `[Content_Types].xml`, `_rels/.rels`, `workbook.xml`, its rels, `styles.xml` and `sharedStrings.xml` before the worksheets, so `xlsxRows` streams turbosheet's own files at constant memory instead of buffering the whole sheet (rows arrive from the first chunks). `xlsxRows` also applies `maxEntryBytes` / `maxTotalBytes` only to parts it buffers: a sheet parsed as it inflates is bounded by `maxCompressionRatio` only, so large legitimate exports no longer hit the 1 GiB entry limit. Docs corrected: Excel 16.x and ExcelJS store worksheets before `styles.xml`, so their sheets are buffered by `xlsxRows`; `createXlsxWriter` output too. Closes #74.
- 1f2fe58: `createXlsxWriter` now defaults to `profile: 'low-memory'` (inline strings). With shared strings the table held every distinct string until `end()` and was serialized there in one blocking pass (1M x 35 rows: 527 MB peak RSS and a 304 ms stall, vs 207 MB / 42 ms inline, and 16% faster). Pass `profile: 'balanced'` or `strings: 'shared'` to keep sharedStrings.xml. `writeXlsx` and `xlsxStream` still default to `balanced`. Closes #72.

## 0.4.0

### Minor Changes

- 957af79: Tables: `tableRows(ws, name)` reads a table's data rows as objects keyed by column name, `addTableRows(ws, name, rows)` appends rows (above the totals row; cells below move down, calculated columns are filled), `getTable(ws, name)`, and `addTable` columns accept `width`. Outline groups: `groupRows` / `groupColumns` with `{ collapsed }`, and a `collapsed` flag on rows (`setRowProps`) and columns, written and read. Iterators: `ws.rows()` and `row.cells()`.
- ad522de: Gradient fills (`{ type: 'gradient', gradient: 'angle' | 'path', degree, center, stops }`, ExcelJS shape) are written, read and deduplicated; fonts gain `charset`, `outline`, `shadow`, `condense` and `extend`; `wb.views` sets Excel's window (active tab, first visible tab, position, size, visibility) and is read back.
- 5de95b8: `cellText(ws, addr)` / `formatValue(value, numFmt)` (in `turbosheet/formula`) return the text Excel displays; `richTextFromHtml(html)` converts simple HTML to rich text; `columnValues` / `eachColumnCell` (ExcelJS `column.values` / `eachCell`); documented stream-transform recipe.
- 4daaef9: Excel-like polish: `new Workbook({ defaultFont })` sets the workbook's Normal font; text with line breaks gets `wrapText` automatically and unstyled hyperlink cells get Excel's hyperlink look (both as Excel does when typing); `ws.ignoredErrors` hides error-check triangles such as "number stored as text" (new `ignoredErrorsWriter` / `ignoredErrorsReader` feature); manual page breaks via `ws.rowBreaks` / `ws.colBreaks` (print setup feature; they and ignored-error ranges follow row/column edits); `setSheetStyle(ws, style)` formats the whole sheet, e.g. to unlock all cells before protecting it.
- bbd1236: `localDates` option (`new Workbook({ localDates: true })`, `ReadOptions.localDates`, also for `xlsxRows`): JS Dates are converted by their local wall clock instead of UTC, so `new Date(2024, 0, 1)` is Jan 1 in Excel in every time zone; `dateToSerial` / `serialToDate` take a `local` flag. `xlsxRows` reports each sheet's merged ranges through `onSheetEnd`. Regression tests for seven reported ExcelJS bugs (header/footer options, sheetPr order with fitToPage, notes in tables, streaming dates, headerless tables, streamed rich text, rich-text hyperlinks).
- 5a01199: In-cell checkboxes: `setCheckbox(ws, range, checked?)` and `style.checkbox` (Excel 365 2024+; boolean cells with the checkbox format, written with the feature property bag part and read back). Pictures: `fitImage(ws, id, range)` scales a picture into a cell or range (aspect kept, centered), `imageSize(bytes)` reads PNG / JPEG / GIF sizes, anchors accept ExcelJS-style fractional positions (`{ col: 1.5, row: 2.25 }`) and `{ tl, br }` corners (pictures and charts).
- 92f0ae7: Dynamic-array formulas (FILTER, SORT, UNIQUE, SEQUENCE, ...) are written with Excel's dynamic-array metadata (`cm="1"` + `xl/metadata.xml`), so Excel spills them instead of adding an implicit-intersection `@`; `Formula.dynamic` forces / disables it and is read back. `copyWorksheet(ws, target?, name?)` copies a sheet within or across workbooks; `copyRange(src, range, dst, to)` copies cells, styles and merges like copy-paste. Column keys may be dotted paths (`'address.city'`) in `addRow`.
- 59badfc: `sortRange(ws, range, keys, { header })` sorts the rows of a range like Excel; `fillTemplate(wb, data)` fills `{{placeholders}}` and repeats `{{items[].field}}` rows (totals ranges grow over the copies); `addDataValidation` moves lists over Excel's 255-character limit (or with commas) to a hidden `_lists` sheet instead of throwing; ExcelJS migration helpers `actualRowCount`, `lastRow`, `getRows`, `getSheetValues`.
- 38429e2: Document properties: `wb.title`, `subject`, `keywords`, `description`, `category`, `contentStatus`, `company`, `manager`, written to docProps and read back. `protectSheet(ws, { password, editable, allow })` unlocks editable ranges and protects the sheet in one call. Pixel conversions for pictures and the default row height follow the workbook default font (`fontMetrics(font)`).
- 8ad654a: `fillFormula(ws, range, formula, results?)` fills a formula over a range with relative references adjusted (ExcelJS `fillFormula`); `wb.calcProperties.fullCalcOnLoad` controls Excel's recalculation on open (read back); `turbosheet/formula` gains `OFFSET`, `INDIRECT` and `LOOKUP` (90 functions).
- f6cd6f9: Hyperlink cells keep the formatting of rich-text link text: `Hyperlink.runs` (formatted runs; `text` stays the plain text) is read from files and written back.
- 0aa7bb7: Row / column structure edits: `insertRows`, `insertRow`, `deleteRows`, `spliceRows`, `insertColumns`, `deleteColumns`, `spliceColumns` and `duplicateRow`. Like Excel's Insert / Delete Sheet Rows / Columns, every reference to the edited sheet follows its cells: formulas on all sheets, array ranges, defined names (print area / titles), chart series, merges, conditional formatting and data validation, tables (inserted rows get calculated columns, inserted columns become table columns), notes, picture and chart anchors and the auto-filter. Edits are validated first and throw without changing anything when they would push cells off the sheet or break a table.

### Patch Changes

- 3d4b001: Streaming: documented and tested backpressure (`addRows` / `flush` wait for slow consumers); regression tests for reported ExcelJS streaming bugs (pictures and number formats in the streaming writer, multi-byte text split across read chunks, constant memory, merges over tables, background picture with notes).

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
