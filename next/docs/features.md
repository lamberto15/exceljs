# Features

Every feature is available on `turbosheet`. On `turbosheet/lite` you opt in per call (see [Lite](#lite-opt-in-features)). Field-by-field options are in [api.md](./api.md#features). All samples assume `wb` is a `Workbook` and `ws` one of its sheets.

- [Cell values, rich text, hyperlinks, formulas](#cell-values-rich-text-hyperlinks-formulas)
- [Styles and number formats](#styles-and-number-formats)
- [Column auto-fit](#column-auto-fit)
- [Tables](#tables)
- [Notes (comments)](#notes-comments)
- [Pictures](#pictures)
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

## Streaming with features

The streaming writer exposes metadata-only sheets, so sheet-level features work as long as they do not need the cells:

```ts
const w = createXlsxWriter({ sink });
const s = w.addWorksheet('Data', { columns: [{ header: 'Region', key: 'region' }], freeze: 1 });
s.worksheet.dataValidations = [{ sqref: 'A2:A1000', type: 'list', values: ['North', 'South'] }];
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
