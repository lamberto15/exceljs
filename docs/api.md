# API reference

Generated from the 0.1.0 type declarations (`dist/*.d.ts`). `turbosheet` and `turbosheet/lite` export the same names; the only difference is which [features](./features.md) are on by default (see [bundle-size.md](./bundle-size.md)). `turbosheet/csv`, `turbosheet/json`, `turbosheet/crypto` and `turbosheet/formula` are separate entries.

Conventions: **rows and columns in the public API are 1-based** (like Excel); columns may be given as a number or a letter (`3` or `'C'`). Cell addresses are A1 strings (`'B3'`). The low-level helpers in `address` (`decodeAddr`, `decodeRange`, `encodeAddr`, `encodeRange`, `encodeCol`, `decodeCol`) are **0-based** and ranges are `{ top, left, bottom, right }`.

Contents: [Workbook](#workbook) - [Worksheet](#worksheet) - [Rows, cells, columns](#rows-cells-columns) - [CellValue](#cellvalue) - [Style](#style) - [Writing](#writing) - [Reading](#reading) - [Streaming writer](#streaming-writer) - [Features](#features) - [CSV](#csv) - [JSON](#json) - [Crypto](#crypto) - [Formula](#formula) - [Errors](#errors) - [Addresses and dates](#addresses-and-dates)

## Workbook

```ts
new Workbook(options?: { date1904?: boolean; defaultFont?: Font; localDates?: boolean })
```

| Member | Description |
| --- | --- |
| `addWorksheet(name?, { tabColor?, state?, properties?, views? })` | Add a sheet (default name `SheetN`). Names follow Excel rules and are unique case-insensitively. |
| `getWorksheet(nameOrId?)` | By name (case-insensitive) or 1-based sheet id; no argument = first sheet. `undefined` when absent. |
| `removeWorksheet(nameOrIdOrSheet)` | Remove a sheet. |
| `worksheets` | `readonly Worksheet[]` in tab order. |
| `styles` | The shared `StyleRegistry` (interns styles; see [Style](#style)). |
| `creator`, `lastModifiedBy`, `created`, `modified` | Document properties. |
| `date1904` | 1904 date system. Set it **before** storing dates; changing it afterwards shifts them by 1462 days. |
| `definedNames` | `DefinedName[]`; use [`addDefinedName`](#defined-names) rather than editing it. |
| `images` | Pictures stored once per workbook (see [Pictures](#pictures)). |
| `preserved` | Opaque round-trip data kept by `readXlsx`. Delete it to drop unmodelled content. |


Document properties (plain fields, written to docProps and read back): `title`, `subject`, `keywords`, `description`, `category`, `contentStatus`, `company`, `manager`, `language`, `revision` (number), plus `creator`, `lastModifiedBy`, `created`, `modified`. `fontMetrics(font)` → `{ mdw, rowPt }` (Excel's digit width in px and default row height for a default font).

`wb.views = [{ activeTab, firstSheet, x, y, width, height, visibility }]` (ExcelJS shape) sets Excel's windows, one `<workbookView>` per entry: the active tab (default: the first visible sheet with `tabSelected`), the first tab shown, window position / size in twips. Read back except the first window's `activeTab` (the sheets' `tabSelected` carries the selection). When a sheet has more views than there are windows, writers add windows (copies of the first entry).
## Worksheet

| Member | Description |
| --- | --- |
| `name` (get/set), `id`, `state` (`'visible' \| 'hidden' \| 'veryHidden'`), `tabColor` | Sheet properties. |
| `properties: WorksheetProperties` | ExcelJS `ws.properties`: `tabColor` (same value as `ws.tabColor`), `defaultRowHeight` (pt), `defaultColWidth` (characters), `outlineLevelRow`, `outlineLevelCol`, `dyDescent`. Unset = derived when writing: row height from the default font, outline levels from the rows / columns (an explicit level only raises the derived one, so set it for streamed outlines). Read back except the outline levels (derived again on write). |
| `views: WorksheetView[]` | One per Excel window (usually one): `state` (`normal`/`frozen`/`split`), `xSplit`, `ySplit`, `topLeftCell`, `activeCell`, `showGridLines`, `showRowColHeaders`, `zoomScale`, `rightToLeft`, `tabSelected`, `workbookViewId` (the window; default: the view's position, so `[{}, { zoomScale: 50 }]` sets windows 1 and 2). |
| `freeze(rows, cols?)` | Freeze the first rows/columns (`freeze(1)` = header row). |
| `autoFilter: string \| undefined` | e.g. `'A1:D1'`. |
| `mergeCells(range)` / `mergeCells(top, left, bottom, right)` | Throws on overlap. |
| `unmergeCells(range)`, `merges` (A1 strings), `mergeRanges` (0-based) | |
| `rowCount`, `columnCount`, `dimensions` | Last row / column with content; `dimensions` is `{ top, left, bottom, right, ref }` (1-based) or `undefined`. |
| `columns` (get/set) | Setting `ColumnDef[]` (`header`, `key`, `width`, `style`, `hidden`, `outlineLevel`) writes a header row when any def has a `header`. |
| `getColumn(numberOrLetterOrKey)` | `Column`: `number`, `letter`, `key`, `header`, `width`, `hidden`, `outlineLevel`, `style`, `eachCell([options], (value, rowNumber) => {})` (non-empty cells; `{ includeEmpty: true }`: rows 1 to `rowCount`, `null` for empty cells). |

Feature fields (plain properties, `undefined` until set): `pageSetup`, `headerFooter`, `comments`, `images`, `backgroundImage`, `conditionalFormatting`, `dataValidations`, `protection`, `tables`, `preserved`. See [Features](#features).

## Rows, cells, columns

All indices below are 1-based.

```ts
ws.addRow(values, style?)    // array (index 0 -> column A) or object keyed by ws.columns[].key; returns the row number
ws.addRows(iterable, style?)
ws.getRow(n)                 // RowView handle
ws.getRowProps(n) / ws.setRowProps(n, { height, hidden, outlineLevel, style, styleIndex })
ws.eachRow((row, rowNumber) => {})   // rows that contain cells, ascending
ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {})   // every row 1..rowCount (ExcelJS)
```

Cells (fast path, no allocation):

```ts
ws.setCell('B3', value, style?)          // or ws.setCell(row, col, value, style?)
ws.getValue('B3')                        // or ws.getValue(row, col); null when empty
ws.getKind('B3')                         // CellKind: Empty, Number, String, Boolean, Date, Error, Formula, RichText, Hyperlink
ws.getStyle('B3')  ws.setStyle('B3', style | xfIndex | undefined)
ws.getStyleIndex('B3')
ws.clearCell('B3')
```

ExcelJS-style handle (allocates per call; prefer the methods above in hot loops):

```ts
const cell = ws.getCell('B3');   // CellRef
cell.value = 42;  cell.address;  cell.row;  cell.col;  cell.kind;
cell.style = { ... };  cell.numFmt = '0.00';  cell.font = {...};  cell.fill;  cell.border;  cell.alignment;  cell.protection;
```

Insert, delete, splice and duplicate rows / columns (functions; every reference to the sheet follows, as in Excel — see [features.md](./features.md#inserting-and-deleting-rows-and-columns)):

```ts
insertRows(ws, pos, rows | count, { inheritStyle?: 'above' | 'below' })   // rows: arrays or keyed objects, like addRow
insertRow(ws, pos, values?, opts?)
deleteRows(ws, start, count = 1)
spliceRows(ws, start, count, ...rows)          // ExcelJS spliceRows: overwrite min(count, rows) in place, delete / insert the rest
insertColumns(ws, pos, cols | count)           // pos: number or letter; cols: arrays of values from row 1
deleteColumns(ws, start, count = 1)
spliceColumns(ws, start, count, ...cols)
duplicateRow(ws, rowNum, count, insert = false) // copies below rowNum: inserted, or overwriting
groupRows(ws, from, to, { collapsed? })         // outline one level deeper (max 7); summary row below
groupColumns(ws, from, to, { collapsed? })      // summary column to the right
```

Display and conversion: `cellText(ws, addr)` / `formatValue(value, numFmt, { date1904 })` from `turbosheet/formula` give the text Excel shows (number formats applied); `richTextFromHtml(html)` → RichText; `columnValues(ws, col)` / `eachColumnCell(ws, col, [options], cb)` (ExcelJS `column.values` / `eachCell`; `{ includeEmpty: true }` visits rows 1 to `rowCount` with `null` for empty cells).

Sort and templates: `sortRange(ws, range, keys: SortKey | SortKey[], { header? })` (`SortKey = { col, desc? }`); `fillTemplate(wb, data, { sheets?, missing? })` (`{{path}}`, repeated rows with `{{items[].field}}`). ExcelJS helpers: `actualRowCount(ws)`, `lastRow(ws)`, `getRows(ws, start, n)`, `getSheetValues(ws)`.

Formulas: `fillFormula(ws, range, formula, results?)`; `wb.calcProperties = { fullCalcOnLoad }`.

Copy: `copyWorksheet(src, target = src.workbook, name?) -> Worksheet` (adds the copy last; "Name (2)" by default) and `copyRange(src, 'A1:D10', dst, 'F2')` (values, styles, merges; relative refs adjusted). Column keys may be dotted paths (`key: 'address.city'` reads nested objects in `addRow`).

Iterators: `for (const row of ws.rows())` (rows that contain cells) and `for (const [col, value] of row.cells())` (non-empty cells). Row props also have `collapsed` (`setRowProps(n, { collapsed: true })`); columns have `column.collapsed`.

`RowView`: `number`, `cellCount`, `values` (sparse array indexed by 1-based column, index 0 unused), `getValue(col)`, `setValue(col, value, style?)`, `getCell(col)`, `eachCell([options], (value, colNumber, xf) => {})` (`{ includeEmpty: true }`: columns 1 to the row's last cell, styled empty cells included; empty cells get `null` and their effective style index), `height`, `hidden`, `outlineLevel`.

## CellValue

```ts
type CellValue = number | string | boolean | Date | CellError | Formula | RichText | Hyperlink | null | undefined;
```

| Type | Shape |
| --- | --- |
| number, string, boolean | as is |
| `Date` | stored as a serial using the workbook date system (UTC based, the Excel 1900 leap-year bug is honored). Number cells with a date format read back as `Date`. |
| `Formula` | `{ formula: 'SUM(A1:A2)', result?: number \| string \| boolean \| CellError, array?: 'A1:B3' }`. No leading `=`, no `_xlfn.` prefix: writers add it to functions newer than Excel 2007 (`XLOOKUP`, `IFS`, `STDEV.S`, ...; also in defined names, validations, conditional formats and table formulas) and readers remove it. `result` is the cached value written to the file; the core does not evaluate formulas (see [`turbosheet/formula`](#formula)). For date-formatted formula cells `result` is the serial (use `serialToDate(result, wb.date1904)`). |
| `RichText` | `{ richText: [{ text, font? }, ...] }` |
| `Hyperlink` | `{ text, hyperlink, tooltip? }` |
| `CellError` | `{ error: '#N/A' }` (codes: `#NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A #GETTING_DATA`) |
| `null` / `undefined` | empty |
| `{ sharedFormula: 'B1', result? }` | setters only (`CellValueInput`): ExcelJS shared-formula value, stored as the master cell's formula with relative references moved by the offset (`translateFormula`); the master must already hold a formula, otherwise a `TypeError`. In the streaming writer the master must be in a row not yet flushed. Never returned by getters. |


`Formula.dynamic`: dynamic-array (spilling) formula, written with Excel's metadata so no `@` is added; automatic for FILTER / SORT / UNIQUE / SEQUENCE / ... unless `array` is set (legacy array formula) or `dynamic: false`.

`Hyperlink.runs`: optional formatted runs of the link text (`text` stays the plain concatenation), read from files and written back as a rich string.
## Style

Styles are **plain, immutable-by-convention objects**. They are interned into the workbook's `StyleRegistry` (one `xf` index per distinct style); objects returned by `getStyle` or `StyleRegistry.get` are shared and must not be mutated. Change a style by assigning a new object.

```ts
interface Style {
  numFmt?: string;            // 'yyyy-mm-dd', '#,##0.00', ...
  font?: Font;                // name, size, bold, italic, underline, strike, color, vertAlign, family, scheme
  fill?: Fill;                // { type: 'pattern', pattern, fgColor?, bgColor? }
  border?: Border;            // top/left/bottom/right/diagonal: { style?, color? }
  alignment?: Alignment;      // horizontal, vertical, wrapText, shrinkToFit, indent, textRotation (Excel encoding; -45, 'vertical', vertical: 'middle' accepted)
  protection?: Protection;    // locked, hidden
}
type Color = { argb?: 'FFFF0000'; theme?: number; tint?: number; indexed?: number };
```

Gradient fills are not modelled in 0.1.0 (only `type: 'pattern'`).

`StyleRegistry` (via `wb.styles`): `intern(style) -> xf`, `get(xf) -> Style`, `internDxf(style) -> dxfId`, `numFmtId(code)`, `addFont/addFill/addBorder`, `isDateXf(xf)`, plus read-only tables `fonts`, `fills`, `borders`, `numFmts`, `xfs`, `dxfs`. Exported helpers: `BUILTIN_NUM_FMTS`, `builtinFormat(id)`, `builtinFormatId(code)`, `isDateFormat(code)`, `isDateFormatId(id)`.


`Fill` = pattern fill or `GradientFill` (`{ type: 'gradient', gradient: 'angle' | 'path', degree?, center?, stops: [{ position, color }] }`). `Font` also has `charset`, `outline`, `shadow`, `condense`, `extend`.
## Writing

```ts
writeXlsx(wb: Workbook, opts?: WriteOptions): Uint8Array            // synchronous
xlsxStream(wb: Workbook, opts?: WriteOptions): ReadableStream<Uint8Array>
createXlsxWriter(opts?: XlsxWriterOptions): XlsxWriter             // row-by-row, see below
```

```ts
interface WriteOptions {
  profile?: 'low-memory' | 'balanced' | 'max-speed';  // default 'balanced' ('low-memory' for createXlsxWriter)
  compression?: number;            // deflate level 0-9 (0 = store); overrides the profile
  strings?: 'shared' | 'inline';   // overrides the profile
  date1904?: boolean;              // ignored by writeXlsx: the workbook's own flag wins
  theme?: boolean;                 // write xl/theme/theme1.xml (default true)
  features?: readonly FeatureWriter[];  // root: added to the built-ins; /lite: the only features enabled
}
```

| Profile | compression | strings | Use |
| --- | --- | --- | --- |
| `balanced` (default for `writeXlsx`, `xlsxStream`) | 6 | shared | smallest files, normal speed |
| `low-memory` (default for `createXlsxWriter`) | 6 | inline | less RAM (no string table) |
| `max-speed` | 1 | inline | fastest, files ~30% larger in our benchmark |

`resolveWriteOptions(opts)` returns the resolved values. `xlsxStream` is pull-based: sheet XML is generated as the consumer reads, so a slow consumer does not make memory grow; do not modify the workbook until the stream ends. Content kept by `readXlsx` (charts, pivots, ...) is written back by `writeXlsx` and `xlsxStream`, not by `createXlsxWriter`.

## Reading

```ts
readXlsx(data: Uint8Array | ArrayBuffer, opts?: ReadOptions): Workbook   // synchronous
xlsxRows(source, opts?): AsyncGenerator<XlsxRow>                          // streaming
```

```ts
interface ReadOptions {
  limits?: Partial<Limits>;        // maxEntries 10_000, maxEntryBytes 1 GiB, maxTotalBytes 2 GiB, maxCompressionRatio 1000
  preserveUnknown?: boolean;       // default true: keep charts, pivots, shapes, VBA, ... and write them back
  sheets?: readonly string[];      // only load these sheets (case-insensitive)
  maxRows?: number;                // throw SheetLimitError at the first cell in a row after maxRows (also xlsxRows)
  maxCols?: number;                // same for columns after maxCols; unset / 0 = no limit
  localDates?: boolean;            // Dates by local wall clock instead of UTC (also xlsxRows)
  onSheetEnd?: (info: { sheet: string; sheetIndex: number; merges: string[] }) => void; // xlsxRows: after each sheet's last row
  profile?: 'low-memory' | 'balanced' | 'max-speed';  // used by xlsxRows (low-memory re-deflates buffered sheets)
  features?: readonly FeatureReader[];  // root: added to the built-ins; /lite: the only features modelled
}
interface XlsxRow { sheet: string; sheetIndex: number; row: number /* 1-based */; values: CellValue[] /* 0-based by column */ }
```

`xlsxRows` accepts `Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>`. Memory is constant for sheet data when `workbook.xml`, `styles.xml` and `sharedStrings.xml` come before the worksheets in the archive, as `writeXlsx` and `xlsxStream` write them. Otherwise such a sheet is buffered (see the JSDoc of `xlsxRows`). That includes files saved by Excel 16.x and ExcelJS, which store worksheets before `styles.xml`, and `createXlsxWriter` output (styles and shared strings are only final at `end()`). The byte limits (`maxEntryBytes`, `maxTotalBytes`) apply only to buffered parts. A streamed sheet keeps nothing, so only `maxCompressionRatio` applies to it, and a 1 GiB+ sheet XML is fine. The shared-string table is always held in memory. Rows without values are skipped; hyperlinks are not applied; `features` is ignored. Merged ranges come after the cell data in a sheet's XML, so `xlsxRows` reports them through `onSheetEnd` once the sheet's rows are done.

Other exports: `listPreservedParts(wb)` -> `{ path, contentType }[]` of the parts the next write will emit from the original file; `translateFormula`; `SheetEntry`.

## Streaming writer

```ts
const w = createXlsxWriter({ sink?, flushRows?, highWaterMark?, creator?, lastModifiedBy?, created?, modified?, ...WriteOptions });
const s = w.addWorksheet(name?, { columns?, views?, freeze?, autoFilter?, tabColor?, state?, properties? });
s.addRow(valuesOrObject, style?, rowProps?)   // returns the 1-based row number
await s.addRows(iterableOrAsyncIterable, style?)   // with backpressure
s.mergeCells('A1:C1')                          // written at commit
await s.flush();  await s.commit();
await w.end();                                 // commits the open sheet, finishes the zip
```

- `profile` defaults to `'low-memory'` here (inline strings: memory stays flat however many distinct strings you write). `profile: 'balanced'` / `strings: 'shared'` writes sharedStrings.xml instead: smaller files when strings repeat a lot, but the table holds every distinct string until `end()` and is serialized in one blocking pass there.
- Its output stores the worksheets before `styles.xml` / `sharedStrings.xml`, which are only final at `end()`. `xlsxRows` therefore buffers each sheet of such a file before emitting rows, and `readXlsx` is unaffected. `writeXlsx` and `xlsxStream` store the metadata first, so their output streams.
- Without `sink`, read the zip from `w.readable` (a `ReadableStream<Uint8Array>`). With `sink: (chunk) => void | Promise<void>` the promise applies backpressure.
- Sheets are written one after another: `addWorksheet` commits the previous sheet. Rows cannot be changed once added.
- `s.worksheet` is a metadata-only `Worksheet` (columns, views, merges, feature fields). Set **head** data (columns, views, freeze) before adding rows; feature fields set after the rows (tables, validations, ...) are applied at `commit()`.
- `w.workbook` holds the shared styles, so `w.workbook.styles.intern(style)` works.
- Async writers (`xlsxStream`, `createXlsxWriter`) deflate off the main thread: with `node:zlib` in Node, Bun and Deno (the `compression` level is honored), with `CompressionStream('deflate-raw')` elsewhere (browsers, workers: the platform's default level, ~6; `compression` only switches between store and deflate), and with fflate on the main thread when neither exists. Zip CRC-32 uses `zlib.crc32` when available.
- `writeXlsx` deflates with `zlib.deflateRawSync` (in ~1 MiB batches) and `readXlsx` inflates with `zlib.inflateRawSync` when `node:zlib` exists, which is 2-4x faster than fflate. Zip limits still apply (through `maxOutputLength`). Parts declared larger than 64 MiB are inflated as a stream with fflate to keep memory bounded. Browsers always use fflate.

## Features

Each is a plain model field plus an `add*` helper (where noted). Full examples in [features.md](./features.md).

| Feature | Model | Helpers | Feature halves (for `/lite`) |
| --- | --- | --- | --- |
| Tables | `ws.tables: Table[]` | `addTable(ws, AddTableOptions) -> Table` | `tablesWriter`, `tablesReader` |
| Notes | `ws.comments: Note[]` | `addNote(ws, ref, text \| RichText \| { text, author?, visible?, width?, height?, margins?, protection?, editAs? }) -> Note` | `commentsWriter`, `commentsReader` |
| Pictures | `wb.images`, `ws.images`, `ws.backgroundImage` | `addImage(wb, { data \| base64, extension }) -> id`, `placeImage(ws, id, anchor, opts?) -> ImagePlacement` | `imagesWriter`, `imagesReader` |
| Charts | `ws.charts: Chart[]` | `addChart(ws, Chart) -> Chart` | `chartsWriter` (write only; read charts are kept raw) |
| Conditional formatting | `ws.conditionalFormatting` | `addConditionalFormatting(ws, ref, rules[])` | `conditionalFormattingWriter`, `conditionalFormattingReader` |
| Data validation | `ws.dataValidations` | `addDataValidation(ws, dv)` | `dataValidationWriter`, `dataValidationReader` |
| Print setup | `ws.pageSetup`, `ws.headerFooter`, `ws.rowBreaks`, `ws.colBreaks` | (plain objects) | `printSetupWriter`, `printSetupReader` |
| Ignored errors | `ws.ignoredErrors` | (plain objects) | `ignoredErrorsWriter`, `ignoredErrorsReader` |
| Sheet protection | `ws.protection` | `hashPassword(password, saltB64, spinCount)` | `sheetProtectionWriter`, `sheetProtectionReader` |
| Default theme | | | `themeWriter` (write only) |

### Column auto-fit

```ts
autoFitColumns(ws, opts?): (number | undefined)[]      // measures, sets column widths, returns them (index 0 = column A)
estimateColumnWidths(rows: Iterable<readonly CellValue[]>, opts?): (number | undefined)[]   // no sheet needed (streaming writer)

interface AutoFitOptions { columns?: (number | string)[]; min?: number /*0*/; max?: number /*100*/; padding?: number /*1.5*/;
  includeHeader?: boolean /*true*/; rows?: { from?: number; to?: number }; sample?: number; apply?: boolean /*true*/ }
interface EstimateWidthOptions { min?; max?; padding?; includeHeader?; headerBold?: boolean; numFmts?: (string | undefined)[]; font?: Font; sample?: number }
```

Pure model helpers exported from `turbosheet` and `turbosheet/lite`; not part of any writer, so unused they cost nothing. Model and limits: [features.md](./features.md#column-auto-fit).

### Defined names

```ts
addDefinedName(wb, name, ref, { localSheet?, hidden?, comment? }) -> DefinedName
getDefinedName(wb, name, localSheet?)   removeDefinedName(wb, name, localSheet?) -> boolean
setPrintArea(ws, 'A1:D20' | string[] | undefined)   getPrintArea(ws)
setPrintTitles(ws, { rows?: '1:2', cols?: 'A:B' })   getPrintTitles(ws)
```

### Tables

`AddTableOptions`: `name`, `displayName?`, `ref?` or `topLeft?` (default `A1`), `columns: (string | AddTableColumn)[]` (`AddTableColumn` = `TableColumn` + `width?`, set on the sheet column), `rows?`, `headerRow?`, `totalsRow?`, `style?: TableStyle` (`theme` such as `TableStyleMedium2`, `showFirstColumn`, `showLastColumn`, `showRowStripes`, `showColumnStripes`), `autoFilter?`, `headerStyle?`, `dataStyle?`, `totalsStyle?`. `TableColumn`: `name`, `totalsRowFunction` (`none sum average count countNums max min stdDev var custom`), `totalsRowLabel`, `totalsRowFormula`, `calculatedColumnFormula`, `headerStyle`, `dataStyle`, `totalsStyle`. It throws on invalid names, overlaps, duplicate columns or rows that do not fit.

```ts
getTable(ws, name) -> Table                       // case-insensitive; throws when missing
tableRows(ws, name) -> Record<string, CellValue>[] // data rows keyed by column name (values as stored)
addTableRows(ws, name, rows)                      // append above the totals row; arrays (table column order) or objects by column name
```

### Pictures

`ImageExtension`: `png jpeg jpg gif bmp tiff svg emf wmf webp`. `placeImage` anchors: a range string (`'B2:D6'`), `{ tl: 'B2' | CellPosition, ext: { width, height } }` (pixels), `{ tl, br }` (corners), or an `ImageAnchor` (`{ from, to }`, `{ from, ext }`, `{ pos, ext }`). Options: `editAs`, `hyperlink`, `tooltip`, `description`, `name`. Identical pictures are stored once. Set `ws.backgroundImage = id` for a sheet background. Positions may be fractional (`{ col: 1.5, row: 2.25 }`, sized with the sheet's widths / heights). `fitImage(ws, id, range, { padding?, upscale?, ...PlaceImageOptions })` scales a picture into a cell / range (aspect kept, centered); `imageSize(bytes)` → `{ width, height }` for PNG / JPEG / GIF.

In-cell checkboxes: `setCheckbox(ws, range, checked = false)` (Excel 365 2024+; a boolean value with `style.checkbox`).

### Charts

`Chart`: `type` (`bar column line pie doughnut scatter area`), `series: ChartSeries[]`, `anchor` (same forms as `placeImage`), `title?`, `legend?: { position: 'r' | 'l' | 't' | 'b' | 'tr' } | false` (default `r`), `axes?: { x?: ChartAxis, y?: ChartAxis }` (x = category axis or scatter x values; y = value axis), `stacked?: 'stacked' | 'percentStacked'` (bar, column, line, area), `dataLabels?: { showValue?, showPercent?, showCategoryName?, showSeriesName? }`, `style?` (1-48), `varyColors?`, `name?`. `ChartSeries`: `values` (A1 reference), `categories?`, `xValues?` (scatter), `name?: string | { ref }` (or a string starting with `=`), `color?: Color` (`argb`, or `theme` 0-9), `marker?`, `smooth?`, `line?` (scatter), `cache?: { name?, categories?, values?, xValues? }` (written instead of reading cells; the streaming writer has no cells). `ChartAxis`: `title?`, `numFmt?`, `min?`, `max?`. References without a sheet name point at the chart's sheet.

### Conditional formatting rules

`cellIs`, `expression`, `top10`, `aboveAverage`, `duplicateValues`/`uniqueValues`, `containsText`/`notContainsText`/`beginsWith`/`endsWith`, `containsBlanks`/`notContainsBlanks`/`containsErrors`/`notContainsErrors`, `timePeriod`, `colorScale`, `dataBar`, `iconSet`. Rules with formatting take `style` (a partial `Style`, written as a dxf); every rule takes `priority?` and `stopIfTrue?`.

### Data validation

`DataValidation`: `sqref`, `type` (`list whole decimal date time textLength custom any`), `operator?`, `formula1?`, `formula2?`, `values?` (literal list items, no commas, 255 characters max joined), `allowBlank?`, `hideDropDown?`, `showInputMessage?`, `promptTitle?`, `prompt?`, `showErrorMessage?`, `errorStyle?` (`stop warning information`), `errorTitle?`, `error?`.

### Print setup and protection

`PageSetup`: `margins`, `orientation`, `paperSize`, `scale`, `fitToPage`, `fitToWidth`, `fitToHeight`, `horizontalCentered`, `verticalCentered`, `showGridLines`, `showRowColHeaders`, `pageOrder`, `blackAndWhite`, `draft`, `cellComments`, `errors`, `firstPageNumber`, `copies`, and DPI/printer flags. Page breaks: `ws.rowBreaks` / `ws.colBreaks` (manual breaks after these 1-based rows / columns). `ws.ignoredErrors: IgnoredError[]` (`sqref` + `evalError`, `twoDigitTextYear`, `numberStoredAsText`, `formula`, `formulaRange`, `unlockedFormula`, `emptyCellReference`, `listDataValidation`, `calculatedColumn`; feature `ignoredErrorsWriter` / `ignoredErrorsReader`). `setSheetStyle(ws, style)` formats the whole sheet (all cells, rows, columns). `protectSheet(ws, { password?, editable?: range | range[], allow?: { sort, autoFilter, formatCells, ... } })` unlocks editable ranges and protects the sheet. `HeaderFooter`: `oddHeader`, `oddFooter`, `evenHeader`, `evenFooter`, `firstHeader`, `firstFooter`, `differentOddEven`, `differentFirst`. `SheetProtection`: `password` (hashed at write time, `spinCount` default 100000) or a precomputed `algorithmName`/`hashValue`/`saltValue`, `sheet`, `objects`, `scenarios`, `selectLockedCells`, `selectUnlockedCells`, `formatCells`, `formatColumns`, `formatRows`, `insertColumns`, `insertRows`, `insertHyperlinks`, `deleteColumns`, `deleteRows`, `sort`, `autoFilter`, `pivotTables`.

### Custom features

`FeatureWriter` and `FeatureReader` (exported types) are the extension points the built-in features use; see [`src/xlsx/features/README.md`](../src/xlsx/features/README.md).

## CSV

`import { ... } from 'turbosheet/csv'`

```ts
parseCsv(text, opts?): string[][]                       // CsvValue[][] when a typing option is set
csvRows(source, opts?): AsyncIterable<string[]>         // source: AsyncIterable|Iterable<string|Uint8Array> | ReadableStream<Uint8Array>
class CsvParser { constructor(onRow, opts?); write(chunk: string); end() }
writeCsv(rows, opts?): string                           // each row followed by eol
formatCsvRow(row, opts?): string                        // one row, no eol
csvStream(rows, opts?): ReadableStream<Uint8Array>      // pull-based
type CsvValue = string | number | boolean | Date | CellError
```

Read options: `delimiter` (single char, default `,`), `quote` (default `"`), `bom` (strip a leading U+FEFF, default true), `maxFieldBytes` (default 16 MiB), `skipEmptyLines` (default true). Write options: `delimiter`, `quote`, `eol` (default `\r\n`), `bom` (default false), `escapeFormulas` (default false), `dateFormat` (default `toISOString()`). Errors: `CsvError` with `code` `FIELD_TOO_LARGE | BAD_OPTION`.

Typing (all opt-in, unquoted fields only: `"5"`, `"true"`, `"#N/A"` stay strings; rows are `string[]` without them):

| Option | Converts | Notes |
| --- | --- | --- |
| `parseNumbers` | `-1.5`, `1e3` → number | no leading zeros (`007` stays text); integers over 15 digits stay text |
| `parseDates` | ISO 8601 → `Date` | `2024-01-02`, `2024-01-02T03:04[:05[.123]]`, with `Z` / `±hh:mm` / `±hhmm` |
| `dateFormats` | e.g. `['DD/MM/YYYY', 'MM/DD/YYYY HH:mm']` → `Date` | tried in order after ISO 8601; implies `parseDates` |
| `localDates` | — | dates without an offset use the local wall clock instead of UTC |
| `parseBooleans` | `true` / `false`, any case → boolean | |
| `parseErrors` | `#N/A`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!` → `{ error }` | exact case; `arraysToSheet` writes error cells |

A field is tried as a number, then a date, then a boolean, then an error. `dateFormats` tokens: `YYYY` (4 digits), `MM` `DD` `HH` `mm` `ss` (2 digits), `M` `D` `H` `m` `s` (1 or 2 digits); `[text]` is literal and any other character must match itself. Each format needs `YYYY`, a month and a day; a bad format throws `CsvError` `BAD_OPTION`. Day/month order is always the format's, never guessed. Impossible values (`31/02/2024`, month 13, `24:00`) stay text and the next format is tried; ISO dates are checked the same way. Dates are UTC wall clock by default (`04/11/2019` with `DD/MM/YYYY` is `2019-11-04T00:00Z`), consistent with `Workbook` dates; use `localDates` together with `new Workbook({ localDates: true })`.

Decoding (`csvRows` byte chunks; string chunks are used as is): option `encoding` is any `TextDecoder` label (`'windows-1252'`, `'iso-8859-1'`, `'utf-16le'`, `'shift_jis'`, ...; default `'utf-8'`; an unknown label throws `RangeError` when `csvRows` is called). Invalid bytes become U+FFFD. A BOM is sniffed first and always wins over `encoding`: `EF BB BF` UTF-8, `FF FE` UTF-16LE, `FE FF` UTF-16BE (so Excel's "Unicode Text" export reads correctly with no option); the BOM itself is then stripped per `bom`. `encoding: 'auto'`: without a BOM, ASCII bytes stream through as they arrive (they read the same in UTF-8 and windows-1252); the first chunk that holds a non-ASCII byte decides: if that chunk (from its first non-ASCII byte to its end) is valid UTF-8 the input is read as UTF-8, otherwise as windows-1252. A UTF-8 sequence cut at the chunk end waits for the next chunk; one cut at the end of the input counts as invalid. Once decided the choice is final: invalid UTF-8 in a later chunk becomes U+FFFD. A whole-file input (`csvRows([bytes])`) is therefore checked as a whole. Decoding is chunk-safe for every encoding: a BOM or a multi-byte character split across chunks (even 1-byte chunks) decodes correctly.

## JSON

`import { ... } from 'turbosheet/json'`

```ts
sheetToObjects(ws, { header?: number | string[], range?, defval?, raw? }): Record<string, unknown>[]
sheetToArrays(ws, { range?, raw? }): unknown[][]
objectsToSheet(ws, rows, { header?: string[], writeHeader?: boolean, emptyAs?: 'empty' | 'string' }): Worksheet
arraysToSheet(ws, rows, { emptyAs?: 'empty' | 'string' }): Worksheet   // '' → no cell by default; { error } → error cell
sheetToNdjson(ws, opts?): AsyncGenerator<string>        // one JSON object + "\n" per row
toJsonValue(cellValue): unknown
```

`header` is a 1-based header row number (default 1) or explicit keys. `raw: false` returns JSON-safe values (dates become ISO strings, rich text and hyperlinks become text, formulas their cached result). Empty header cells become `__EMPTY_<col>`, duplicates get `_2`, `_3`, and `__proto__`/`constructor`/`prototype` are prefixed with `_`. Nested objects and arrays are written as JSON strings.

## Crypto

`import { ... } from 'turbosheet/crypto'` (async; uses WebCrypto)

```ts
isEncrypted(data: Uint8Array): boolean
decryptXlsx(data, password, { verifyIntegrity?, limits?, maxSpinCount? }): Promise<Uint8Array>   // inner zip bytes
encryptXlsx(zipBytes, password, { cipher?: 'AES-128'|'AES-192'|'AES-256', hash?: 'SHA1'|'SHA256'|'SHA384'|'SHA512', spinCount? }): Promise<Uint8Array>
```

Defaults match Excel 2010+: AES-256-CBC, SHA-512, 100000 spins, HMAC integrity. `decryptXlsx` reads agile and standard (Excel 2007) encryption; `encryptXlsx` writes agile. `verifyIntegrity` defaults to true; `maxSpinCount` defaults to 10,000,000.

## Formula

`import { ... } from 'turbosheet/formula'` (synchronous; details, semantics and the function list: [formula.md](./formula.md))

```ts
evaluate(wb, opts?): EvaluationReport             // computes every formula cell, stores Formula.result
evaluateCell(wb, sheet, 'D10', opts?): FormulaResult | null
evaluateFormula(wb, 'SUM(A1:A9)', { sheet?, address?, array?, now? }?): FormulaResult | FormulaResult[][]
parseFormula(text): FormulaNode                   // throws FormulaParseError { position, formula }
createEvaluator({ functions?, now? }): Evaluator  // per-instance functions: .registerFunction(name, impl)
supportedFunctions(): string[]
toNumber(v) toText(v) toBoolean(v) scalarValues(v) isError(v)   // helpers for custom functions
```

`EvaluateOptions`: `sheets`, `writeResults` (default `true` for `evaluate`, `false` for `evaluateCell`), `maxIterations` (default 0: circular references give `#REF!`), `maxChange` (0.001), `onUnsupported(name, address)`, `now`. `EvaluationReport`: `evaluated`, `circular` (addresses), `unsupported` (function names), `errors` (`{ address, message }`), `results` (`Map` when not writing). A formula never makes evaluation throw: errors become `CellError` results. `FormulaResult` = `number | string | boolean | CellError`; dates are serials in the workbook's date system.

## Errors

| Class | Thrown by | Meaning |
| --- | --- | --- |
| `XlsxParseError` | `readXlsx`, `xlsxRows` | Broken XML or package structure. |
| `EncryptedFileError` (extends `XlsxParseError`) | `readXlsx` | Input is an OLE container (password-protected, or a legacy `.xls`); use `decryptXlsx`. |
| `SheetLimitError` | `readXlsx`, `xlsxRows` | A sheet has a cell beyond `maxRows` / `maxCols`; thrown at the first such cell. Fields `sheet`, `limit` (`'maxRows' \| 'maxCols'`), `max`. |
| `ZipLimitError`, `ZipPathError` | `readXlsx`, `xlsxRows` | A `limits` bound was exceeded, or an unsafe path (`ZipPathError` extends `ZipLimitError`). Exported from `turbosheet` and `turbosheet/lite`. |
| `RangeError`, `TypeError` | model methods | Invalid addresses, overlapping merges, bad sheet names, bad options. |
| `WrongPasswordError`, `IntegrityError`, `UnsupportedEncryptionError`, `CfbError` | `turbosheet/crypto` | Wrong password; tampered or corrupt data; RC4 / extensible encryption; malformed container. |
| `CsvError` | `turbosheet/csv` | Option or size-limit problem. |
| `FormulaParseError` (extends `SyntaxError`) | `parseFormula`, `evaluateFormula` | Formula text that does not parse; `position` is the 0-based offset. `evaluate` reports such cells instead of throwing. |

In `turbosheet/lite`, writing a feature field whose writer is not enabled throws (for example ``Worksheet 'Data': ws.tables is set but the 'tables' feature is not enabled; pass features: [tablesWriter]``) instead of dropping data silently.

## Addresses and dates

```ts
encodeCol(0) // 'A'          decodeCol('XFD') // 16383
encodeAddr(2, 1) // 'B3'     decodeAddr('B3') // { row: 2, col: 1 }
decodeRange('A1:C3') // { top: 0, left: 0, bottom: 2, right: 2 }     encodeRange(range)
MAX_ROWS // 1048576          MAX_COLS // 16384
dateToSerial(date, date1904?)   serialToDate(serial, date1904?)
```

Also exported: `CellKind`, `DEFAULT_LIMITS`, `SharedStrings`, and the types `CellAddress`, `CellRange`, `Dimensions`, `WorksheetView`, `ColumnDef`, `RowProps`, `Limits`, `Profile`, `SheetState`, `ChunkSink`, `StreamSheetOptions`, `XlsxWriterOptions`, `ResolvedWriteOptions`.
