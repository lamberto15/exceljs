# Features

Every feature is available on `turbosheet`. On `turbosheet/lite` you opt in per call (see [Lite](#lite-opt-in-features)). Field-by-field options are in [api.md](./api.md#features). All samples assume `wb` is a `Workbook` and `ws` one of its sheets.

- [Cell values, rich text, hyperlinks, formulas](#cell-values-rich-text-hyperlinks-formulas)
- [Styles and number formats](#styles-and-number-formats)
- [Column auto-fit](#column-auto-fit)
- [Inserting and deleting rows and columns](#inserting-and-deleting-rows-and-columns)
- [Tables](#tables)
- [Notes (comments)](#notes-comments)
- [Pictures](#pictures)
- [Charts](#charts)
- [Conditional formatting](#conditional-formatting)
- [Data validation](#data-validation)
- [Defined names, print area and titles](#defined-names-print-area-and-titles)
- [Print setup and headers/footers](#print-setup-and-headersfooters)
- [Sheet protection](#sheet-protection)
- [Round-trip preservation](#round-trip-preservation)
- [Streaming with features](#streaming-with-features)
- [Encryption](#encryption)
- [Lite: opt-in features](#lite-opt-in-features)

## Cell values, rich text, hyperlinks, formulas

```ts
ws.setCell('A1', 'plain text');
ws.setCell('A2', { richText: [{ text: 'bold ', font: { bold: true } }, { text: 'and plain' }] });
ws.setCell('A3', { text: 'turbosheet', hyperlink: 'https://example.com', tooltip: 'Home page' });
ws.setCell('B1', 10);
ws.setCell('B2', 32);
ws.setCell('B3', { formula: 'SUM(B1:B2)', result: 42 }); // cached result is written as given
ws.setCell('C1', new Date(Date.UTC(2025, 5, 1)));        // gets a default date format
ws.setCell('C2', { error: '#N/A' });
```

turbosheet never evaluates formulas: provide `result` if readers that do not recalculate (previewers, parsers) should see a value. Excel recalculates on open.

## Styles and number formats

```ts
ws.setStyle('A1', {
  font: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } },
  alignment: { horizontal: 'center', wrapText: true },
  border: { bottom: { style: 'thin', color: { argb: 'FF000000' } } },
});
ws.getCell('B1').numFmt = '#,##0.00';
ws.getColumn('B').width = 14;
```

Styles are values: assign a new object to change one, do not mutate `getStyle()` results. Equal styles share one entry in `styles.xml`, so styling a million cells with the same object is cheap. For hot loops intern once and pass the index: `const xf = wb.styles.intern(style); ws.setCell(r, c, v, xf)`.

## Column auto-fit

```ts
import { autoFitColumns, estimateColumnWidths } from 'turbosheet'; // also from 'turbosheet/lite'

const widths = autoFitColumns(ws);                       // sets ws column widths, returns them (0-based index)
autoFitColumns(ws, { columns: ['A', 'C'], max: 60, includeHeader: false, rows: { from: 2 }, sample: 5000 });
```

Options: `columns` (numbers, letters or keys; default every used column), `min` (default 0: columns with nothing to measure are left alone; with `min > 0` requested empty columns get it), `max` (100), `padding` (1.5 characters, includes Excel's 5 px cell padding), `includeHeader` (true; false skips the first row of the range), `rows: { from, to }` (1-based), `sample` (scan only the first N non-empty rows), `apply` (true; false only returns the widths).

How it estimates (Excel's AutoFit, without fonts):

- Units are characters of the default font's digit width (Calibri 11: 7 px). Text width is the sum of Calibri advance widths (a 95-entry ASCII table; CJK, full-width forms and emoji count as 2 digits, other characters as 1), scaled by `font.size / defaultFont.size` and by 1.1 for bold.
- It measures the displayed text: General numbers (shortest form up to 11 characters, `E+` notation outside 1e-9..1e11), fixed decimals, thousands separators, percent, scientific, currency/literal text, negative sections like `(0.00)`; date formats by the width of the format (month and day names use the widest name, so they err wide); `TRUE`/`FALSE`; error codes; formula cached results; hyperlink text; rich text (runs concatenated, widest run font).
- Multi-line text (wrapped or containing `\n`) uses its widest line. Cells in a merge spanning several columns are skipped. Rotated text (90, 180, 255) gets a small fixed width. `indent` adds about 1.3 characters per level.
- Accuracy: within about 10 % of Excel for Latin text in Calibri (the tests derive references from Excel's width formula and hinted glyph widths). Other fonts are measured with Calibri metrics scaled by size, so wide fonts (Verdana) come out narrow and narrow ones (Arial Narrow) wide; kerning, hinting at other sizes, locale-specific formats, fractions (`# ?/?`) and conditional-format text are not modelled. Pass a larger `padding` if you prefer to err wide.
- Speed: one pass over `rowData` with per-style caching and no per-cell allocation for numbers, strings and dates; 1,000,000 mixed cells take about 0.2 s.

The streaming writer cannot fit columns after rows are flushed. Estimate from a sample first, then declare the widths:

```ts
const sample = rows.slice(0, 1000);                      // any Iterable<CellValue[]>; first row = header
const widths = estimateColumnWidths(sample, { headerBold: true, numFmts: [undefined, '#,##0.00'], max: 60 });
const w = createXlsxWriter(sink);
const sheet = w.addWorksheet('Data', { columns: widths.map((width) => ({ width })) });
```

`estimateColumnWidths(rows, { min, max, padding, includeHeader, headerBold, numFmts, font, sample })` uses the same measuring code; it has no styles, so give number formats per column with `numFmts` (0-based) and the body font with `font`. Dates use `yyyy-mm-dd` (with time when any value has one) unless the column's `numFmts` entry is a date format.

## Inserting and deleting rows and columns

```ts
import { deleteRows, duplicateRow, insertColumns, insertRows, spliceRows } from 'turbosheet'; // also from 'turbosheet/lite'

ws.setCell('D1', { formula: 'SUM(B2:B5)' });
insertRows(ws, 3, [['Inserted', 1]], { inheritStyle: 'above' }); // D1 is now SUM(B2:B6)
deleteRows(ws, 2);                                               // ... and now SUM(B2:B5)
insertColumns(ws, 'B', 1);                                       // ... and now SUM(C2:C5), in E1
spliceRows(ws, 2, 1, ['Replaced', 2]);                           // ExcelJS spliceRows
duplicateRow(ws, 2, 2, true);                                    // two copies of row 2 below it
```

The edits work like Excel's *Insert / Delete Sheet Rows / Columns*: cells move, and every reference to the edited sheet follows them:

- **Formulas** of every sheet (unqualified refs on the edited sheet, `Data!A1` / `'My Data'!A1` on others), array-formula ranges, **defined names** (print area / titles too) and **chart series**. `$` makes no difference (it only matters for copy and fill). A range spanning an insertion point grows; a range losing some of its rows shrinks; a reference whose cells are all deleted becomes `#REF!` (a print area / titles part that is gone is dropped).
- **Sheet data**: row and column properties (height, width, hidden, styles, keys), merges (dropped when deleted or reduced to one cell), conditional formatting and data validation ranges and formulas, notes, picture and chart anchors (two-cell anchors stretch, one-cell anchors move), `autoFilter`.
- **Tables**: rows inserted inside a table extend it and get its calculated-column formulas; columns inserted inside one become new columns (`Column4`, header cell written); deleted columns drop their definitions; a table whose cells are all deleted is removed. Deleting a table's header row, or all of its data rows, throws.

An edit is validated before anything changes: one that would push cells off the grid (row 1,048,576 / column XFD) throws and leaves the workbook as it was.

ExcelJS differences: these are functions, not methods; `spliceRows(ws, start, count, ...rows)` overwrites the first `min(count, rows.length)` rows in place (references to them stay) and deletes or inserts only the difference; `duplicateRow` adjusts the copied formulas' relative references, as a paste does; `inheritStyle` replaces the `'i'` / `'o'` style codes (and also formats empty cells, like `'i+'` / `'o+'`).

Not adjusted: content kept raw from a read file (charts read from a file, pivot tables, sparklines, threaded comments, ...), hyperlink targets (`#Sheet!A1`), `ws.views` (frozen panes, selection), 3D (`Sheet1:Sheet3!A1`) and external-workbook references.

## Tables

```ts
addTable(ws, {
  name: 'Sales',
  columns: ['Region', { name: 'Units', totalsRowFunction: 'sum' }],
  rows: [['North', 120], ['South', 80]],
  totalsRow: true,
  style: { theme: 'TableStyleMedium2', showRowStripes: true },
});
```

Writes the header, data and totals cells (`SUBTOTAL` formulas) and the table part. `ref`/`topLeft` position it; `calculatedColumnFormula`, `headerStyle`/`dataStyle`/`totalsStyle` (per table or per column) are supported. Existing tables are read into `ws.tables` and can be edited.

## Notes (comments)

```ts
addNote(ws, 'B2', 'Check this value');
addNote(ws, 'B3', { text: 'Approved', author: 'Ann', visible: true, width: 200, height: 60 });
addNote(ws, 'B4', { richText: [{ text: 'Important', font: { bold: true } }] });
```

These are classic notes (legacy comments with a VML box). Threaded comments in existing files are preserved, not modelled.

## Pictures

```ts
const id = addImage(wb, { data: pngBytes, extension: 'png' }); // or { base64: '...' }
placeImage(ws, id, 'D2:F6');                                    // stretch over a range
placeImage(ws, id, { tl: 'H2', ext: { width: 120, height: 80 } }, { description: 'Logo' });
ws.backgroundImage = id;                                        // sheet background
```

## Charts

Bar, column, line, pie, doughnut, scatter and area charts, written as native Excel charts (write only: charts in files you read are kept raw and written back).

```ts
ws.addRow(['Month', 'North', 'South']);
ws.addRow(['Jan', 10, 30]);
ws.addRow(['Feb', 15, 27]);
ws.addRow(['Mar', 20, 24]);
addChart(ws, {
  type: 'column', // 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area'
  title: 'Units per month',
  series: [
    { name: { ref: 'B1' }, categories: 'A2:A4', values: 'B2:B4' }, // the chart's own sheet
    { name: 'South', categories: "'Data'!$A$2:$A$4", values: "'Data'!$C$2:$C$4", color: { argb: 'FFC00000' } },
  ],
  anchor: 'E2:L16', // same forms as placeImage
  legend: { position: 'b' }, // or false
  axes: { y: { title: 'Units', min: 0 } },
  stacked: 'stacked', // bar / column / line / area
  dataLabels: { showValue: true },
});
```

- References are A1 ranges, optionally sheet-qualified (`'My Data'!$B$2:$B$10`); unqualified ranges mean the chart's own sheet. They are written absolute, with the sheet name quoted when needed.
- Each reference carries a cache of the cells' current values, so Excel and LibreOffice show the data immediately. The streaming writer holds no cells: it writes references only (Excel recalculates on open; LibreOffice shows an empty chart) unless you pass `series.cache: { name?, categories?, values?, xValues? }`.
- Series options: `color` (`argb` or `theme` 0-9), `marker: false` and `smooth` (line, scatter), `line: true` (scatter: connect the points; default markers only), `xValues` (scatter). Chart options: `style` (built-in style 1-48), `varyColors` (default on for pie/doughnut), `dataLabels: { showValue, showPercent, showCategoryName, showSeriesName }`, axis `title`, `numFmt`, `min`, `max`, shape `name`.
- Charts share the sheet's drawing with pictures; on a sheet read with charts, the ones you add are written next to the kept ones.

## Conditional formatting

```ts
addConditionalFormatting(ws, 'B2:B20', [
  { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } },
]);
addConditionalFormatting(ws, 'C2:C20', [
  { type: 'colorScale', colors: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }] },
]);
addConditionalFormatting(ws, 'D2:D20', [{ type: 'dataBar', color: { argb: 'FF638EC6' } }]);
```

## Data validation

```ts
addDataValidation(ws, { sqref: 'A2:A20', type: 'list', values: ['North', 'South', 'East'] });
addDataValidation(ws, {
  sqref: 'B2:B20',
  type: 'whole',
  operator: 'between',
  formula1: 0,
  formula2: 1000,
  showErrorMessage: true,
  errorTitle: 'Units',
  error: 'Enter a whole number from 0 to 1000',
});
```

## Defined names, print area and titles

```ts
addDefinedName(wb, 'Rate', '0.2');
addDefinedName(wb, 'Totals', 'Data!B2:B20');
setPrintArea(ws, 'A1:D50');
setPrintTitles(ws, { rows: '1:1' });
getPrintArea(ws); // ['A1:D50']
```

## Print setup and headers/footers

```ts
ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
ws.headerFooter = { oddHeader: '&CMonthly report', oddFooter: '&LPage &P of &N' };
```

## Sheet protection

```ts
ws.protection = { password: 's3cret', formatCells: false, insertRows: false };
```

`password` is hashed at write time (SHA-512, `spinCount` 100000 by default; lower it to speed up writes). Sheet protection is not encryption: anyone can remove it. Use [encryption](#encryption) to restrict opening the file.

## Round-trip preservation

`readXlsx` keeps what the model does not represent (charts, pivot tables, shapes, VBA, threaded comments, themes, custom XML, unknown worksheet elements, ...) and `writeXlsx`/`xlsxStream` write it back, so loading a file, editing cells and saving does not strip it.

```ts
const wb = readXlsx(bytes);                 // preserveUnknown defaults to true
listPreservedParts(wb);                     // [{ path: 'xl/charts/chart1.xml', contentType: '...' }, ...]
wb.getWorksheet('Data')!.setCell('A1', 'edited');
const out = writeXlsx(wb);                  // charts and pivots are still there
```

- `readXlsx(bytes, { preserveUnknown: false })` drops them (and does not retain the input buffer).
- Kept parts stay compressed inside the input buffer until written: do not mutate `bytes` after reading.
- Parts reachable only from removed or renamed sheets are dropped. Delete `wb.preserved` / `ws.preserved` to drop content explicitly.
- `createXlsxWriter` starts from a new private workbook and never carries preserved content.

What is kept next to the model, and what happens when you edit nearby:

| Kept | When you edit |
| --- | --- |
| AutoFilter criteria (`filterColumn`: values, custom, top 10, dynamic, color, icon filters) and its `sortState`; sheet-level `sortState` | `ws.autoFilter` is the range. Changing the range keeps the filter columns only when it still starts in the same column and covers their position; the autoFilter's sort state is dropped. Removing `ws.autoFilter` removes the criteria. |
| Named cell styles (`cellStyleXfs` / `cellStyles`: "Heading 1", "Good", custom styles) and each cell format's link to its style | Styles you set get no named style. Reusing a style object read from a cell (`ws.setStyle(ref, ws.getStyle(other))`) keeps the link. |
| `sheetFormatPr` (`defaultColWidth`, `baseColWidth`, non-default `defaultRowHeight`, `customHeight`, `zeroHeight`, `thickTop`/`thickBottom`) | Outline levels are recomputed from rows and columns. Extension attributes (`x14ac:dyDescent`) are dropped. |
| Form controls (checkboxes, buttons: `<controls>`, ctrlProps parts, legacy VML shapes), also on sheets with notes | Notes are regenerated from `ws.comments` and the control shapes are merged into the same VML, keeping their shape ids. Shapes that link other parts (pictures in VML) cannot be merged on a sheet with notes, so they are dropped there. |
| Chartsheet / dialogsheet tab positions | A kept sheet stays right after the worksheet it followed. If that worksheet was removed, it goes right before the worksheet it preceded, otherwise last. New worksheets go last. |
| Query tables (table part → queryTable part, `tableType="queryTable"`) | While the table in `ws.tables` is unchanged, its original part is written back. A modified table is regenerated from the model without the query link. |
| Hyperlinks on formula cells (a cell holds one value kind: the formula is modelled, the link kept raw) | Written back while the cell still holds the formula it was read with. Setting a new value (even an equal formula) or clearing the cell drops the link. |

Dropped on purpose: `calcChain.xml` (Excel rebuilds it), kept content that references a sheet renamed or removed since reading, and kept content of a feature whose model you replace (for example, setting `ws.images` on a sheet read without the images reader).

## Streaming with features

The streaming writer exposes metadata-only sheets, so sheet-level features work as long as they do not need the cells:

```ts
const w = createXlsxWriter({ sink });
const s = w.addWorksheet('Data', { columns: [{ header: 'Region', key: 'region' }], freeze: 1 });
s.worksheet.dataValidations = [{ sqref: 'A2:A1000', type: 'list', values: ['North', 'South'] }];
s.worksheet.charts = [{ type: 'line', series: [{ values: 'B2:B1000' }], anchor: 'D2:K16' }]; // no caches
for (const r of rows) s.addRow(r);
await s.commit();
await w.end();
```

`addTable` writes cells into a worksheet that holds them, so use it with the buffered `Workbook`/`writeXlsx` path.

## Encryption

See [`turbosheet/crypto` in api.md](./api.md#crypto):

```ts
const encrypted = await encryptXlsx(writeXlsx(wb), 'hunter2');
const wb2 = readXlsx(await decryptXlsx(encrypted, 'hunter2'));
```

## Lite: opt-in features

```ts
import { Workbook, addTable, readXlsx, tablesReader, tablesWriter, writeXlsx } from 'turbosheet/lite';

const wb = new Workbook();
addTable(wb.addWorksheet('Data'), { name: 'T', columns: ['Region', 'Units'], rows: [['N', 1]] });
const bytes = writeXlsx(wb, { features: [tablesWriter] });
const back = readXlsx(bytes, { features: [tablesReader] });
```

Without a feature enabled, writing its field throws, and reading keeps its XML raw and writes it back unchanged. Details and size numbers: [bundle-size.md](./bundle-size.md).
