# Features

Every feature is available on `turbosheet`. On `turbosheet/lite` you opt in per call (see [Lite](#lite-opt-in-features)). Field-by-field options are in [api.md](./api.md#features). All samples assume `wb` is a `Workbook` and `ws` one of its sheets.

- [Cell values, rich text, hyperlinks, formulas](#cell-values-rich-text-hyperlinks-formulas)
- [Styles and number formats](#styles-and-number-formats)
- [Document properties](#document-properties)
- [Column auto-fit](#column-auto-fit)
- [Inserting and deleting rows and columns](#inserting-and-deleting-rows-and-columns)
- [Sorting and templates](#sorting-and-templates)
- [Tables](#tables)
- [Notes (comments)](#notes-comments)
- [Pictures](#pictures)
- [In-cell checkboxes](#in-cell-checkboxes)
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


Dynamic-array formulas (FILTER, SORT, UNIQUE, SEQUENCE, ... that spill in Excel 365) are written with Excel's dynamic-array metadata, so Excel shows `=SORT(A1:A9)` and spills it instead of turning it into `=@SORT(A1:A9)`. It is automatic for formulas using a spilling function; `{ formula, dynamic: true }` forces it, `dynamic: false` turns it off, and `array: 'C1:C9'` (without `dynamic`) still writes a legacy Ctrl+Shift+Enter array formula. Files read keep the flag.

Filling: `fillFormula(ws, 'C2:C100', 'A2*B2', results?)` puts the formula in every cell with relative references adjusted (like Excel's fill handle); `results` (array or `(row, col) => value`) are the cached values. Recalculation: workbooks with formulas are written with `fullCalcOnLoad` so Excel recomputes on open; `wb.calcProperties = { fullCalcOnLoad: false }` turns that off when the cached results are current (e.g. after `evaluate`), `true` forces it.

Copying: `copyWorksheet(ws)` adds a copy to the same workbook ("Data (2)"), `copyWorksheet(ws, otherWb, 'Name')` to another one (styles re-registered, pictures re-added, table names made unique, dates converted between date systems). `copyRange(src, 'A1:D10', dst, 'F2')` copies cells, styles and merges like copy-paste, adjusting relative references.

Column keys may be dotted paths: with `ws.columns = [{ key: 'address.city' }]`, `ws.addRow({ address: { city: 'Oslo' } })` fills that column.
## Document properties

```ts
wb.title = 'Q3 report';
wb.subject = 'Sales';
wb.keywords = 'sales, q3';
wb.description = 'Quarterly numbers';
wb.category = 'Finance';
wb.contentStatus = 'Draft';
wb.company = 'ACME';
wb.manager = 'Jane';
// also wb.creator, wb.lastModifiedBy, wb.created, wb.modified
```

Written to `docProps/core.xml` / `docProps/app.xml` (File > Info in Excel) and read back.

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

Gradient fills use ExcelJS's shape: `{ type: 'gradient', gradient: 'angle', degree: 90, stops: [{ position: 0, color: { argb: 'FFFF0000' } }, { position: 1, color: { argb: 'FF0000FF' } }] }`, or `gradient: 'path'` with `center: { left: 0.5, top: 0.5 }`. Fonts also take `charset`, `outline`, `shadow`, `condense` and `extend`.

Excel-like defaults:

- **Default font**: `new Workbook({ defaultFont: { name: 'Arial', size: 10 } })` sets Excel's "Normal" font (default Calibri 11): unstyled cells use it and partial fonts (`{ bold: true }`) are completed with its name and size. Column widths are in units of that font's digit width, so pixel conversions (pictures, `fitImage`) and the default row height follow it (`fontMetrics(font)`; approximations for fonts other than Calibri). `wb.styles.setDefaultFont(font)` does the same later, but set it before styling cells.
- **Line breaks**: text containing `\n` gets `alignment.wrapText` merged into its style, as when typing Alt+Enter in Excel (line breaks only show in wrapped cells).
- **Hyperlinks**: a hyperlink cell stored without a style gets Excel's hyperlink look (underlined, theme hyperlink color); pass a style to choose your own.
- **Whole sheet**: `setSheetStyle(ws, style)` formats every cell, like selecting all and formatting: existing cells, row styles and all columns get the style's properties merged over theirs, and cells created later inherit it. See [sheet protection](#sheet-protection) for unlocking all cells.
- **Ignored errors**: hide Excel's green error triangles, e.g. "number stored as text" on a column of IDs: `ws.ignoredErrors = [{ sqref: 'A2:A1000', numberStoredAsText: true }]` (also `evalError`, `twoDigitTextYear`, `formula`, `formulaRange`, `unlockedFormula`, `emptyCellReference`, `listDataValidation`, `calculatedColumn`).

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

Outline groups: `groupRows(ws, 2, 5)` / `groupColumns(ws, 'B', 'D')` put rows / columns one outline level deeper (Excel's *Data > Group*, max 7 levels); `{ collapsed: true }` also hides them and flags the summary row below (column to the right). Iterate with `for (const row of ws.rows())` and `for (const [col, value] of row.cells())`.

Not adjusted: content kept raw from a read file (charts read from a file, pivot tables, sparklines, threaded comments, ...), hyperlink targets (`#Sheet!A1`), `ws.views` (frozen panes, selection), 3D (`Sheet1:Sheet3!A1`) and external-workbook references.

## Sorting and templates

```ts
import { fillTemplate, sortRange } from 'turbosheet';

sortRange(ws, 'A1:D100', [{ col: 'C', desc: true }, { col: 'A' }], { header: true });

// template: A1 "Invoice for {{customer.name}}", row 4 "{{items[].name}}" | "{{items[].qty}}", B6 =SUM(B4:B4)
fillTemplate(wb, { customer: { name: 'ACME' }, items: [{ name: 'bolt', qty: 3 }, { name: 'nut', qty: 5 }] });
```

`sortRange` orders the rows of a range like Excel (numbers and dates, then text case-insensitively, then booleans, then errors; blanks last in both directions; ties keep their order). Only the range's cells move; formulas in moved rows shift with them, references from outside keep their positions. Merged cells inside the range throw.

`fillTemplate` replaces `{{path}}` placeholders in text cells, keeping styles: a cell holding only a placeholder gets the typed value (number, date, ...), mixed text gets text. A row with `{{items[].field}}` placeholders is repeated for each element of `items` (rows below move down, formulas in the row adjust, and ranges ending at the row, like the `SUM(B4:B4)` total, grow over the copies); an empty array removes the row. Options: `sheets`, `missing` (text for unknown paths).

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

Writes the header, data and totals cells (`SUBTOTAL` formulas) and the table part. `ref`/`topLeft` position it; `calculatedColumnFormula`, `headerStyle`/`dataStyle`/`totalsStyle` (per table or per column) are supported. Existing tables are read into `ws.tables` and can be edited. A column's `width` sets the sheet column width.

```ts
import { addTableRows, tableRows } from 'turbosheet';

addTableRows(ws, 'Sales', [['East', 95], { Region: 'West', Units: 60 }]); // above the totals row
tableRows(ws, 'Sales'); // [{ Region: 'North', Units: 120 }, ..., { Region: 'West', Units: 60 }]
```

`addTableRows` inserts sheet rows (cells below the table move down, references follow as with [`insertRows`](#inserting-and-deleting-rows-and-columns)), grows the table and fills calculated columns. `tableRows` returns the data rows keyed by column name; formula cells give their `Formula` object (`.result` is the cached value).

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
fitImage(ws, id, 'B2:C4');                                      // scaled to fit (aspect kept), centered
placeImage(ws, id, { tl: { col: 1.5, row: 2.25 }, br: 'E8' });  // fractional positions, like ExcelJS
```

Fractional `col` / `row` (0-based, as in ExcelJS: `1.5` = halfway into column B) are converted to cell + pixel offset with the sheet's column widths and row heights, as `fitImage` does; set widths and heights first. `fitImage` reads PNG, JPEG and GIF sizes (`imageSize(bytes)`), shrinks pictures larger than the range (`upscale: true` also grows small ones) and keeps `padding` (2 px) free.

## In-cell checkboxes

```ts
import { setCheckbox } from 'turbosheet';

ws.addRow(['Task', 'Done']);
ws.addRow(['write tests', true]);
setCheckbox(ws, 'B2:B10'); // B2 stays checked; empty cells become unchecked (FALSE)
```

Excel 365's in-cell checkboxes (2024+) are boolean cells with a checkbox format (`style.checkbox`): formulas see TRUE / FALSE, and Excel versions without checkboxes show those values. Written and read back.

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

Lists longer than Excel's 255-character limit, or with commas in items, are written by `addDataValidation` to a hidden sheet `_lists` and referenced from there, so the dropdown still works.

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
ws.rowBreaks = [40, 80]; // manual page breaks after rows 40 and 80
ws.colBreaks = [6];      // and after column F
```

## Sheet protection

```ts
ws.protection = { password: 's3cret', formatCells: false, insertRows: false };
```

`password` is hashed at write time (SHA-512, `spinCount` 100000 by default; lower it to speed up writes). One call: `protectSheet(ws, { password: 's3cret', editable: ['B2:D20', 'F:F'], allow: { sort: true, autoFilter: true } })` unlocks the editable ranges (whole columns / rows through their default style), keeps everything else locked and sets the allowed actions (sorting and filtering only work on unlocked cells, filtering needs an `autoFilter`). By hand: to protect a sheet but keep most cells editable, unlock everything and lock the few cells that matter: `setSheetStyle(ws, { protection: { locked: false } })`, then `ws.setStyle('A1', { ...ws.getStyle('A1'), protection: { locked: true } })`. Sheet protection is not encryption: anyone can remove it. Use [encryption](#encryption) to restrict opening the file.

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

Backpressure: `addRow` is synchronous; `await s.addRows(iterable)` (sync or async iterable) and `await s.flush()` wait until the sink or the readable's consumer has caught up, so a slow consumer (a network upload, a slow disk) keeps memory flat. In a long synchronous loop call `await s.flush()` every few thousand rows. Pictures (`placeImage(s.worksheet, ...)`), number formats and the other style options work as in the buffered writer.


Transform recipe (constant memory both ways): read with `xlsxRows`, write with `createXlsxWriter`.

```ts
const w = createXlsxWriter({ sink });
let out = w.addWorksheet('Orders');
for await (const row of xlsxRows(createReadStream('in.xlsx'))) {
  const [item, qty, price] = row.values;
  if (row.row === 1) out.addRow([item, qty, price, 'total']);
  else if (qty) out.addRow([item, qty, price, qty * price]);
  if (row.row % 5000 === 0) await out.flush(); // backpressure
}
await w.end();
```

Rows come with `row.sheet`: start a new `w.addWorksheet(row.sheet)` when it changes. Styles are not carried over (`xlsxRows` returns values); use `readXlsx` + `writeXlsx` when formatting must survive.
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
