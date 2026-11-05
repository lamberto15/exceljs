# Migrating from ExcelJS

turbosheet is a clean rewrite inspired by (and partly derived from) ExcelJS 4.x. The cell/row/column vocabulary and 1-based addressing are familiar, but the I/O is synchronous where it can be, features are functions instead of methods, and several rarely used APIs are not there yet. This page maps ExcelJS 4.4 to turbosheet 0.6. Every turbosheet sample here runs in CI (`test/docs-examples.test.ts`). For a test-by-test account of which ExcelJS behaviours are supported, different or missing, see [exceljs-parity.md](./exceljs-parity.md).

## The big differences

| Topic | ExcelJS | turbosheet |
| --- | --- | --- |
| In-memory write | `await wb.xlsx.writeBuffer()` (async, returns `Buffer`) | `writeXlsx(wb)` is **synchronous** and returns a `Uint8Array`. |
| In-memory read | `await wb.xlsx.load(buffer)` into an existing workbook | `readXlsx(bytes)` is **synchronous** and **returns** a new `Workbook`. |
| Files | `wb.xlsx.writeFile(path)` / `readFile(path)` | Use your runtime's file API: `writeFileSync(path, writeXlsx(wb))`, `readXlsx(readFileSync(path))`; for big files `xlsxRows(await openAsBlob(path))`. The library has no static `node:*` imports. |
| Streaming write | `new ExcelJS.stream.xlsx.WorkbookWriter({ stream })`, `row.commit()` | `createXlsxWriter({ sink })`; rows are batched for you, no `row.commit()`; inline strings by default (`profile: 'low-memory'`). `xlsxStream(wb)` gives a web `ReadableStream`. |
| Streaming read | `WorkbookReader` with events or `for await` | `for await (const row of xlsxRows(source))`; bytes and Blobs are read by random access, so files written by Excel or ExcelJS stream at constant memory too. |
| Features | Methods and properties on worksheet objects (`ws.addTable`, `ws.addConditionalFormatting`, `cell.note`) | Functions that take the worksheet: `addTable(ws, ...)`, `addConditionalFormatting(ws, ...)`, `addNote(ws, ...)`. They tree-shake. |
| Styles | Mutable objects on `cell.style`, deep-cloned | **Plain values, interned** into `wb.styles` (one `xf` per distinct style). Assign a new object to change a style; never mutate what `getStyle()` returns. |
| Cells | Cell objects per cell | No per-cell objects. Use `ws.setCell` / `ws.getValue` in hot loops; `ws.getCell()` returns a lightweight handle for ExcelJS-style code. |
| Runtimes | Node-oriented (`Buffer`, streams) | Web standards (`Uint8Array`, Web Streams); Node streams work where an `AsyncIterable<Uint8Array>` is accepted. |
| Formulas | Not evaluated | Not evaluated by reading or writing; provide `result` for the cached value. `evaluate(wb)` from [`turbosheet/formula`](./formula.md) computes them (90 functions). |
| Unknown content | Dropped on load/save (charts, pivot tables, ...) | **Preserved** on round trip (`preserveUnknown`, default true). |

Indexing: rows and columns are **1-based** everywhere in the public API, like ExcelJS. Two spots are 0-based: `XlsxRow.values` from `xlsxRows` (`values[0]` is column A) and the low-level address helpers (`decodeAddr('B3')` is `{ row: 2, col: 1 }`). `RowView.values` is 1-based like ExcelJS (index 0 unused).

## API map

### Workbook and worksheets

| ExcelJS | turbosheet |
| --- | --- |
| `new ExcelJS.Workbook()` | `new Workbook()` |
| `wb.creator`, `lastModifiedBy`, `created`, `modified` | same |
| `wb.addWorksheet('S', { properties: { tabColor }, state, views })` | same (`{ tabColor }` also works) |
| `wb.getWorksheet('S' \| id)`, `wb.worksheets` | same (`getWorksheet` returns `undefined` when absent) |
| `wb.eachSheet(cb)` | `for (const ws of wb.worksheets)` |
| `wb.removeWorksheet(id)` | `wb.removeWorksheet(nameOrIdOrSheet)` |
| `wb.definedNames.add(ref, name)` | `addDefinedName(wb, name, ref)` |
| `ws.pageSetup.printArea` | `setPrintArea(ws, 'A1:D20')` |
| `ws.pageSetup.printTitlesRow` | `setPrintTitles(ws, { rows: '1:1' })` |
| `wb.calcProperties`, `wb.views` | supported (`calcProperties.fullCalcOnLoad`; every window: position, size, first sheet, active tab) |
| `wb.language`, `wb.revision`, `wb.title`, ... | same |

### Rows, columns, cells

| ExcelJS | turbosheet |
| --- | --- |
| `ws.columns = [{ header, key, width }]` | same |
| `ws.getColumn(1).width = 20` | same |
| `ws.addRow([..] \| {..})` | same; returns the 1-based row number (ExcelJS returns a `Row`) |
| `ws.addRows(rows)` | same |
| `ws.getRow(3)` / `row.getCell(2)` / `row.eachCell([options], cb)` | same (`RowView`); `cb(value, colNumber, xf)` receives the value, not a cell object. `{ includeEmpty: true }` visits columns 1 to the row's last cell (empty ones get `null`), as in ExcelJS |
| `row.values`, `row.height`, `row.hidden`, `row.outlineLevel` | same getters; **no setter for `row.values`**: use `addRow` or `setCell` |
| `ws.getCell('A1').value = v` | same, or `ws.setCell('A1', v)` (faster) |
| `cell.font = {...}`, `cell.numFmt`, `cell.fill`, `cell.border`, `cell.alignment`, `cell.protection`, `cell.style` | same on the handle; or `ws.setStyle('A1', style)` |
| `cell.type` | `ws.getKind('A1')` (`CellKind`) |
| `ws.eachRow([options], cb)` | same (`cb(row, rowNumber)`); `{ includeEmpty: true }` visits every row from 1 to `ws.rowCount` |
| `ws.getColumn('A').eachCell([options], cb)` | same; `cb(value, rowNumber)` receives the value. Without options only non-empty cells are visited (ExcelJS also visits empty cells of rows that have values elsewhere); `{ includeEmpty: true }` visits rows 1 to `ws.rowCount`. `eachColumnCell(ws, 'A', [options], cb)` does the same without creating a column record |
| `ws.rowCount`, `ws.columnCount`, `ws.dimensions` | same (`dimensions` is `{ top, left, bottom, right, ref }`) |
| `ws.mergeCells('A1:C2')`, `ws.unMergeCells('A1')` | same; `ws.unmergeCells('A1')` (lower-case m) removes every merge intersecting the range; `mergeCells` throws on overlapping merges |
| `ws.views = [{ state: 'frozen', ySplit: 1 }]` | same (one view per window), or `ws.freeze(1)` |
| `ws.autoFilter = 'A1:D1'` | same |
| `ws.state = 'hidden'` | same |
| `ws.properties` (`tabColor`, `defaultRowHeight`, `outlineLevelRow`, ...) | same; `ws.tabColor` is an alias of `ws.properties.tabColor`. Outline levels are derived from the rows / columns (a property only raises them) and not read back |
| `row.collapsed` / `column.collapsed` (getters derived from outline levels) | settable flags on the summary row / column: `ws.setRowProps(n, { collapsed: true })`, `ws.getColumn('E').collapsed = true`; or `groupRows(ws, from, to, { collapsed: true })` / `groupColumns` |
| `ws.pageSetup`, `ws.headerFooter` | same shape (plain objects) |
| `ws.protect(password, options)` | `ws.protection = { password, ...options }` |
| `ws.actualRowCount`, `lastRow`, `getRows`, `getSheetValues`, `column.values` | `actualRowCount(ws)`, `lastRow(ws)`, `getRows(ws, start, n)`, `getSheetValues(ws)`, `columnValues(ws, 'A')` (functions) |
| `ws.insertRow(s)`, `spliceRows`, `spliceColumns`, `duplicateRow` | `insertRow(s)(ws, …)`, `spliceRows(ws, …)`, `spliceColumns(ws, …)`, `duplicateRow(ws, …)`, plus `deleteRows`, `insertColumns`, `deleteColumns`; formulas, merges, names, tables, … follow like in Excel ([details](./features.md#inserting-and-deleting-rows-and-columns)) |

### Values

| ExcelJS value | turbosheet |
| --- | --- |
| number, string, boolean, `Date`, `null` | same |
| `{ formula: 'SUM(A1:A2)', result: 3 }` | same |
| `{ sharedFormula: 'B1', result }` | accepted when setting a cell (`setCell`, `addRow`, `cell.value = ...`, the streaming writer): stored and written as the master's formula with relative references moved, e.g. `{ formula: 'B1+1' }` in B1 shared to C1 becomes `{ formula: 'C1+1', result }`. The master must already hold a formula (in the streaming writer: a row not yet flushed). Shared formulas in files you read are expanded to plain formulas |
| `{ richText: [{ text, font }] }` | same |
| `{ text, hyperlink, tooltip }` | same |
| `{ error: '#N/A' }` | same |

### Features

| ExcelJS | turbosheet |
| --- | --- |
| `ws.addTable({ name, ref, columns, rows })` | `addTable(ws, { name, topLeft \| ref, columns, rows })`; columns may be plain strings; header/data/totals cells are written for you |
| `cell.note = 'text' \| {...}` | `addNote(ws, 'A1', 'text' \| { text, author, visible, width, height, margins, protection, editAs })`; `protection` takes booleans (ExcelJS: `'True'` / `'False'`) |
| `cell.dataValidation = {...}` / `ws.dataValidations.add('A1:A9', {...})` | `addDataValidation(ws, { sqref: 'A1:A9', type, ... })` |
| `ws.addConditionalFormatting({ ref, rules })` | `addConditionalFormatting(ws, ref, rules)` |
| `const id = wb.addImage({ buffer \| base64, extension })`; `ws.addImage(id, 'B2:D6' \| { tl, ext })` | `const id = addImage(wb, { data \| base64, extension })`; `placeImage(ws, id, 'B2:D6' \| { tl, ext })` |
| `ws.addBackgroundImage(id)` | `ws.backgroundImage = id` |
| Gradient fills | supported, same shape (`{ type: 'gradient', gradient: 'angle' \| 'path', ... }`) |
| Charts (no API in ExcelJS; dropped on load) | `addChart(ws, { type, series, anchor, ... })`: bar, column, line, pie, doughnut, scatter, area ([details](./features.md#charts)); charts in files you read are preserved |

### I/O

```ts
// ExcelJS
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('in.xlsx');
wb.getWorksheet(1).getCell('A1').value = 'edited';
await wb.xlsx.writeFile('out.xlsx');
```

```ts
// turbosheet
import { readFileSync, writeFileSync } from 'node:fs';
import { readXlsx, writeXlsx } from 'turbosheet';

const wb = readXlsx(readFileSync('in.xlsx'));
wb.getWorksheet(1)!.getCell('A1').value = 'edited';
writeFileSync('out.xlsx', writeXlsx(wb));
```

Streaming write:

```ts
// ExcelJS
const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx', useStyles: true });
const ws = wb.addWorksheet('Data');
for (const r of rows) ws.addRow(r).commit();
await ws.commit();
await wb.commit();
```

```ts
// turbosheet
import { createWriteStream } from 'node:fs';
import { createXlsxWriter } from 'turbosheet';

const file = createWriteStream('big.xlsx');
const w = createXlsxWriter({ sink: (chunk) => new Promise((resolve) => file.write(chunk, () => resolve())) });
const ws = w.addWorksheet('Data');
for (const r of rows) ws.addRow(r);
await ws.commit();
await w.end();
file.end();
```

Streaming read:

```ts
// ExcelJS
const reader = new ExcelJS.stream.xlsx.WorkbookReader('big.xlsx', {});
for await (const sheet of reader) for await (const row of sheet) { /* row.values is 1-based */ }
```

```ts
// turbosheet
import { openAsBlob } from 'node:fs';
import { xlsxRows } from 'turbosheet';

for await (const row of xlsxRows(await openAsBlob('big.xlsx'))) {
  // row.sheet, row.row (1-based), row.values (0-based: values[0] is column A)
}
```

A Blob is read lazily by random access: the sheets stream at constant memory whatever order the archive stores its parts in. A `createReadStream(...)` also works, but files from Excel and ExcelJS store the sheets before `styles.xml`, so a stream buffers them until the metadata arrives.

CSV:

```ts
// ExcelJS: await wb.csv.readFile(path); await wb.csv.writeBuffer()
// turbosheet: CSV is independent of the workbook
import { parseCsv, writeCsv } from 'turbosheet/csv';
const rows = parseCsv(text, { parseNumbers: true });
const csv = writeCsv(rows, { escapeFormulas: true });
```

To turn a worksheet into CSV rows, use `sheetToArrays(ws, { raw: true })` from `turbosheet/json` and pass the result to `writeCsv`; to load CSV into a sheet use `arraysToSheet(ws, parseCsv(text, { parseNumbers: true }))` (empty fields leave cells empty, like ExcelJS).

ExcelJS' `csv.read` types every field; turbosheet types only what you ask for. To match it:

```ts
const rows = parseCsv(text, {
  parseNumbers: true,
  parseBooleans: true, // ExcelJS: exact 'true' / 'false'; turbosheet: any case
  parseErrors: true, // '#N/A' → { error: '#N/A' }
  dateFormats: ['MM-DD-YYYY'], // ExcelJS default list: ISO forms + MM-DD-YYYY
  localDates: true, // ExcelJS (dayjs) parses local time; turbosheet defaults to UTC
});
arraysToSheet(new Workbook({ localDates: true }).addWorksheet('Sheet1'), rows);
```

`dateFormats` uses the same `YYYY MM DD HH mm ss` tokens as dayjs (strict), ISO 8601 is always tried first, and impossible dates such as `31/02/2024` stay text. ExcelJS turns anything `Number()` accepts into a number (`007` → 7, `0x10` → 16); `parseNumbers` accepts only plain decimal numbers. ExcelJS decodes bytes as UTF-8 by default; `csvRows(source, { encoding })` takes any `TextDecoder` label, honors BOMs (UTF-16 included), and `'auto'` falls back to windows-1252.

## Behavior differences to watch for

- **Synchronous reads/writes block the event loop** for large files. For big data use the streaming writer/reader, or run `readXlsx`/`writeXlsx` in a worker.
- **Dates are UTC** (as in ExcelJS). `new Date(Date.UTC(2025, 0, 15))` is the cell value `2025-01-15`: the serial is computed from the UTC instant, and reading gives UTC-based `Date`s. Build dates with `Date.UTC`, or opt into the local wall clock with `new Workbook({ localDates: true })` and `readXlsx(bytes, { localDates: true })` / `xlsxRows(source, { localDates: true })`.
- **Style objects are shared.** `ws.getStyle('A1')` returns an interned object; `{ ...style, font: { ...style.font, bold: true } }` then `setStyle` is the way to derive a new one.
- **Errors are typed.** `readXlsx` throws `EncryptedFileError` for password-protected input, `XlsxParseError` for broken XML, `ZipLimitError` for hostile archives (default limits: 10,000 entries, 1 GiB per entry, 2 GiB total, compression ratio 1000).
- **`maxRows` / `maxCols`** work as in ExcelJS (`readXlsx(bytes, { maxRows: 1000, maxCols: 50 })`, also for `xlsxRows`) but throw `SheetLimitError` (message `Max row count (1000) exceeded in sheet "Data" (row 1001)`; fields `sheet`, `limit`, `max`). They are checked against cell positions: a cell in row 1001 fails `maxRows: 1000` even when the sheet holds fewer rows; ExcelJS counts `<row>` / `<c>` elements instead. Parsing stops at the first such cell.
- **Sheet names** are validated with Excel's rules when you add or rename a sheet.
- **Merged cells:** `mergeCells` does not clear or copy cell values; a value left in a non-top-left cell stays in the model (Excel displays only the top-left one). Put the value in the top-left cell.
- **`ws.columns = [...]` with headers writes row 1**, like ExcelJS; `addRow` then continues below it.
- **Alignment is stored with OOXML names and encoding:** ExcelJS's `vertical: 'middle'`, `textRotation: 'vertical'` and negative rotations are accepted, and read back as `'center'`, `255` and Excel's encoding (`-45` → `135`).
- **Shared formulas are expanded on read** (`{ formula: 'A4+1', result }` per cell); `{ sharedFormula }` is accepted when setting a cell but never returned.

## What turbosheet adds

- **Formula evaluation:** `evaluate(wb)` in `turbosheet/formula` computes formula results (90 functions, custom functions, cycles); see [formula.md](./formula.md).
- **Charts:** `addChart` writes native bar, column, line, pie, doughnut, scatter and area charts.
- **Encryption:** `encryptXlsx` / `decryptXlsx` / `isEncrypted` in `turbosheet/crypto` (agile read/write, standard read).
- **Round-trip preservation:** charts, pivot tables, shapes, VBA, themes and unknown XML survive load, edit, save. Inspect with `listPreservedParts`.
- **`turbosheet/lite`:** the same API with features opt-in, about 25 KB write-only (see [bundle-size.md](./bundle-size.md)).
- **Profiles and compression control:** `profile: 'low-memory' | 'balanced' | 'max-speed'`, `compression`, `strings: 'shared' | 'inline'`.
- **Zip-bomb limits** on every read path and **JSON helpers** (`sheetToObjects`, `objectsToSheet`, NDJSON).
- **Web-standard runtime support:** browsers, workers and Bun/Deno are expected to work (CI covers Node 22/24).

## Not supported yet

ExcelJS 4.4 APIs without an equivalent: image sources given as `filename`. Table edits are free functions instead of `table.addColumn` / `removeColumns` / `removeRows` / `getColumn(i).name` / `headerRow` / `totalsRow` + `commit()`: `addTableColumn`, `removeTableColumns`, `removeTableRows`, `renameTableColumn`, `setTableHeaderRow`, `setTableTotalsRow` (applied at once, no commit). Shared formulas are written as shared formulas but read back expanded (no `sharedFormula` / `shareType` values). Neither library reads or writes `.xls` / `.xlsb` or authors pivot tables; turbosheet's decryption does not cover RC4 / extensible encryption (ExcelJS has none). If you depend on one of these, stay on ExcelJS for that path or open an issue; [exceljs-parity.md](./exceljs-parity.md) has the details.
