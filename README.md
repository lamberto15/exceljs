# turbosheet

[![npm](https://img.shields.io/npm/v/turbosheet)](https://www.npmjs.com/package/turbosheet) [![CI](https://github.com/lamberto15/turbosheet/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/lamberto15/turbosheet/actions/workflows/ci.yml?query=branch%3Amain) [![minzipped size](https://img.shields.io/bundlephobia/minzip/turbosheet)](https://bundlephobia.com/package/turbosheet) [![node](https://img.shields.io/node/v/turbosheet)](https://nodejs.org) [![license](https://img.shields.io/npm/l/turbosheet)](./LICENSE)

Fast, small, streaming-first XLSX and CSV for TypeScript on every JavaScript runtime.

> Status: **0.x**. The API may change between minor versions. See [Status](#status).

## Why

- **Fast.** 100k rows x 20 columns: `writeXlsx` 1.2 s (0.6 s with `profile: 'max-speed'`) vs 6.6-7.6 s for ExcelJS `writeBuffer`; reading takes 1.3 s vs 3.4-4.8 s for `xlsx.load`.
- **Low memory.** Cells live in typed arrays, not one object per cell. The streaming writer holds only a batch of rows: heap stayed around 30-80 MB where ExcelJS `writeBuffer` peaked at 2.1 GB.
- **Small.** 36 KB (min + brotli, with dependencies) for a write-only bundle; 25 KB with `turbosheet/lite`. Tree-shakable, `sideEffects: false`.
- **Web standards only.** `Uint8Array`, Web Streams, `TextEncoder`, `CompressionStream` / `DecompressionStream`, WebCrypto. No `Buffer`, no `node:*` imports in `src/` (in Node, Bun and Deno, `node:zlib` is looked up at runtime with `process.getBuiltinModule` for faster deflate, inflate and CRC-32).
- **Safe by default.** Zip-bomb limits, no XML entity expansion, round-trip preservation of content it does not model.

Numbers are from our own benchmark against `exceljs@4.4.0` (Node v26.13.0, Apple M1 Pro, 100,000 rows x 20 columns of numbers, strings and dates; see [`test/xlsx-write.bench.test.ts`](./test/xlsx-write.bench.test.ts)). Treat them as indicative, run it on your data. Methodology and full tables: [docs/benchmarks.md](./docs/benchmarks.md); reproduce with `pnpm bench`.

| 100k rows x 20 cols     | turbosheet                               | ExcelJS 4.4.0           |
| ----------------------- | ------------------------------------- | ----------------------- |
| write, in memory        | 1.2 s `writeXlsx` (balanced)          | 6.6-7.6 s `writeBuffer` |
| write, in memory, fast  | 0.6 s `writeXlsx` (`max-speed`)       |                         |
| write, streaming        | 0.4-1.2 s `xlsxStream` / `createXlsxWriter` | 3.3-3.5 s `WorkbookWriter` |
| read into a workbook    | 1.3 s `readXlsx`                      | 3.4-4.8 s `xlsx.load`   |
| read, streaming rows    | 1.2 s `xlsxRows`                      | 1.8-3.7 s `WorkbookReader` |
| peak JS heap, in-memory write | ~150 MB (cells are partly off-heap) | 2.1 GB              |
| peak JS heap, streaming write | 30-80 MB                        | 39-94 MB                |

## Install

```sh
npm install turbosheet     # or pnpm add / yarn add / bun add
```

ESM and CommonJS builds, TypeScript types included. Releases from 0.2.0 on are published from CI with signed npm provenance. Requires Node 22 or newer (or another modern runtime, see [Runtime support](#runtime-support)).

## Quick start

### Write

```ts
import { Workbook, writeXlsx } from 'turbosheet';

const wb = new Workbook();
const ws = wb.addWorksheet('Sales');
ws.columns = [
  { header: 'Region', key: 'region', width: 14 },
  { header: 'Units', key: 'units', width: 10 },
  { header: 'Shipped', key: 'shipped', width: 12 },
];
ws.addRow({ region: 'North', units: 120, shipped: new Date(Date.UTC(2025, 0, 15)) });
ws.addRow(['South', 80, new Date(Date.UTC(2025, 1, 3))]);
ws.setCell('B4', { formula: 'SUM(B2:B3)', result: 200 });
ws.setStyle('A1', { font: { bold: true } });
ws.getCell('C2').numFmt = 'yyyy-mm-dd';
ws.freeze(1);

const bytes: Uint8Array = writeXlsx(wb); // synchronous
```

### Read

```ts
import { readXlsx } from 'turbosheet';

const back = readXlsx(bytes); // Uint8Array | ArrayBuffer, synchronous
const sheet = back.getWorksheet('Sales')!;
sheet.getValue('A2'); // 'North'
sheet.getValue(3, 2); // 80   (row, column), 1-based like Excel
sheet.eachRow((row, n) => console.log(n, row.values));
```

### Stream write (constant memory)

```ts
import { createWriteStream } from 'node:fs';
import { createXlsxWriter } from 'turbosheet';

const file = createWriteStream('big.xlsx');
const w = createXlsxWriter({
  sink: (chunk) => new Promise((resolve) => file.write(chunk, () => resolve())),
});
const sheet = w.addWorksheet('Data', {
  columns: [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 24 },
  ],
  freeze: 1,
});
for (let i = 1; i <= 5000; i++) sheet.addRow({ id: i, name: `row ${i}` });
await sheet.commit();
await w.end();
file.end();
```

`xlsxStream(wb)` returns a web `ReadableStream<Uint8Array>` for an existing workbook (handy for `new Response(stream)`). Like `writeXlsx`, it stores the workbook metadata before the sheets, so its first bytes come once the last sheet is encoded.

### Stream read

```ts
import { openAsBlob } from 'node:fs';
import { xlsxRows } from 'turbosheet';

// Accepts bytes, a Blob / File, an AsyncIterable<Uint8Array> (Node streams) or a web ReadableStream.
for await (const row of xlsxRows(await openAsBlob('big.xlsx'))) {
  // row.sheet, row.row (1-based), row.values (0-based by column: values[0] is column A)
}
// merged ranges: xlsxRows(src, { onSheetEnd: ({ sheet, merges }) => ... })
```

Bytes and Blobs are read random-access (central directory first), so sheets stream at constant memory whatever order the file stores its parts in. A stream is read in archive order: a sheet stored before `styles.xml` / `sharedStrings.xml` (files saved by Excel or ExcelJS, `createXlsxWriter` output) is buffered until they arrive.

Dates: Excel dates have no time zone. turbosheet reads and writes JS `Date`s by their UTC wall clock (`new Date(Date.UTC(2024, 0, 1))` is Jan 1). If your code builds dates with `new Date(2024, 0, 1)`, use `new Workbook({ localDates: true })` and `readXlsx(buf, { localDates: true })` / `xlsxRows(src, { localDates: true })`.

### CSV

```ts
import { csvRows, csvStream, parseCsv, writeCsv } from 'turbosheet/csv';

const text = writeCsv(
  [
    ['name', 'note'],
    ['Ann', 'said "hi", left'],
    ['Bob', '=HYPERLINK("http://evil")'],
  ],
  { escapeFormulas: true }, // prefixes ' to values starting with = + - @ TAB CR
);
parseCsv(text); // string[][]
parseCsv('a,1\r\nb,2', { parseNumbers: true }); // [['a', 1], ['b', 2]]
parseCsv('TRUE,#N/A,04/11/2019', { parseBooleans: true, parseErrors: true, dateFormats: ['DD/MM/YYYY'] });
// [[true, { error: '#N/A' }, 2019-11-04T00:00:00.000Z]]

// Streaming: csvStream(rows) -> ReadableStream<Uint8Array>; csvRows(source) -> AsyncIterable<string[]>
// csvRows reads UTF-16 files with a BOM as is; { encoding: 'auto' } also reads Excel's windows-1252 CSV
```

### JSON helpers

```ts
import { objectsToSheet, sheetToNdjson, sheetToObjects } from 'turbosheet/json';

objectsToSheet(ws, [{ name: 'Ann', age: 31 }, { name: 'Bob', age: 42 }]); // header from keys
sheetToObjects(ws); // [{ name: 'Ann', age: 31 }, { name: 'Bob', age: 42 }]
for await (const line of sheetToNdjson(ws)) process.stdout.write(line); // one JSON object per row
```

### Formulas

```ts
import { evaluate, evaluateFormula } from 'turbosheet/formula';

ws.setCell('C2', { formula: 'A2*B2' });
const report = evaluate(wb); // fills Formula.result of every formula, in dependency order
// report.circular, report.unsupported list what could not be computed
evaluateFormula(wb, 'SUMIF(A:A,">10")'); // ad-hoc
```

Results are written as cached values, so files open with correct numbers in viewers that do not recalculate. Supported functions and semantics: [docs/formula.md](./docs/formula.md).

### Password-protected files

```ts
import { decryptXlsx, encryptXlsx, isEncrypted } from 'turbosheet/crypto';

const encrypted = await encryptXlsx(writeXlsx(wb), 'hunter2'); // AES-256, like Excel 2010+
isEncrypted(encrypted); // true; readXlsx(encrypted) throws EncryptedFileError
const back = readXlsx(await decryptXlsx(encrypted, 'hunter2')); // WrongPasswordError otherwise
```

Agile (Excel 2010+) files are read and written; standard (Excel 2007) files are read only.

## Features

Formula evaluation (`turbosheet/formula`: 213 functions, structured references, `LET`, dynamic arrays that spill, dependency graph, incremental recalculation, circular-reference detection), column auto-fit (`autoFitColumns`, `estimateColumnWidths` for streaming), row/column insert, delete, splice and duplicate with Excel-style reference adjustment (`insertRows`, `deleteColumns`, `spliceRows`, ...), outline groups, tables (append, read and edit), notes, threaded comments, custom views, pictures, charts (bar, column, line, pie, doughnut, scatter, area), conditional formatting, data validation, defined names, print setup, sheet protection, rich text, hyperlinks, merged cells, frozen panes, auto-filter, shared or inline strings. Everything the model does not represent (charts read from a file, pivot tables, shapes, VBA, threaded comments, ...) is kept and written back on round trip. Details and examples: [docs/features.md](./docs/features.md).

```ts
import { addTable, addNote, addDataValidation, addConditionalFormatting } from 'turbosheet';

addTable(ws, {
  name: 'Sales',
  columns: ['Region', { name: 'Units', totalsRowFunction: 'sum' }],
  rows: [['North', 120], ['South', 80]],
  totalsRow: true,
});
addNote(ws, 'A1', { text: 'Region name', author: 'me' });
addDataValidation(ws, { sqref: 'A2:A3', type: 'list', values: ['North', 'South'] });
addConditionalFormatting(ws, 'B2:B3', [
  { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } },
]);
```

## Documentation

- [docs/api.md](./docs/api.md): API reference
- [docs/features.md](./docs/features.md): one section per feature
- [docs/migrating-from-exceljs.md](./docs/migrating-from-exceljs.md): ExcelJS to turbosheet
- [docs/exceljs-parity.md](./docs/exceljs-parity.md): every ExcelJS test, ported and classified (supported / different / missing)
- [docs/bundle-size.md](./docs/bundle-size.md): `turbosheet/lite` and size budgets
- [docs/benchmarks.md](./docs/benchmarks.md): benchmark methodology and results
- [CONTRIBUTING.md](./CONTRIBUTING.md): development setup and PR flow
- [ROADMAP.md](./ROADMAP.md): released versions and what comes next
- [docs/formula.md](./docs/formula.md): formula evaluation (`turbosheet/formula`)
- [ARCHITECTURE.md](./ARCHITECTURE.md): internals; [RELEASING.md](./RELEASING.md): release process

## Entry points

| Import             | What                                                        | Size (min + brotli, with deps) |
| ------------------ | ----------------------------------------------------------- | ------------------------------ |
| `turbosheet`          | read + write, every feature on                              | 36 KB write-only, 79 KB everything (`*`) |
| `turbosheet/lite`     | same API, features opt-in via `options.features`           | 25 KB write-only, 43 KB read + write |
| `turbosheet/csv`      | CSV parser and writer (RFC 4180, streaming)                 | 3.2 KB                         |
| `turbosheet/json`     | worksheet to/from objects, arrays, NDJSON                   | not size-budgeted yet          |
| `turbosheet/crypto`   | password-protected files (MS-OFFCRYPTO)                     | 7.2 KB                         |
| `turbosheet/formula`  | formula parser and evaluator (`evaluate(wb)` fills cached results) | 17 KB                    |

Sizes are from `pnpm size` for the 0.6.0 build (tree-shaken imports: write-only `{ Workbook, writeXlsx }`, lite read + write `{ Workbook, readXlsx, writeXlsx }`, others `*`; budgets in [docs/bundle-size.md](./docs/bundle-size.md)). The only runtime dependency is [fflate](https://github.com/101arrowz/fflate).

## Runtime support

- **Tested in CI:** Node 22 and 24 on Linux and Windows.
- **Expected to work, not yet CI-tested:** Bun, Deno, browsers, Cloudflare Workers and other edge runtimes. The code only uses web-standard APIs (`Uint8Array`, Web Streams, `TextEncoder`, WebCrypto). Where `process.getBuiltinModule('node:zlib')` exists (Node 22.3+, Bun, Deno), zip deflate, inflate and CRC-32 use it; otherwise async writers deflate with `CompressionStream`, `xlsxRows` inflates Blobs with `DecompressionStream`, and everything else uses fflate. Bug reports from these runtimes are welcome.

## Security

- **Zip bombs:** reading enforces limits on entry count, per-entry and total uncompressed size (checked while inflating, not only from headers) and compression ratio; override with `ReadOptions.limits` (`xlsxRows` applies the byte limits only to parts it holds in memory; a sheet parsed as it inflates is checked against the compression ratio). Paths with `..`, absolute paths, backslashes and NUL are rejected.
- **XML:** the tokenizer never expands DTDs or custom entities (no XXE, no billion laughs); only the five predefined entities and numeric character references are decoded.
- **CSV injection:** `writeCsv(rows, { escapeFormulas: true })` prefixes `'` to text starting with `=`, `+`, `-`, `@`, TAB or CR. It is off by default; turn it on when exporting untrusted data for spreadsheet users.
- **Encryption:** reads reject tampered files (HMAC check) and absurd password-hash iteration counts.
- Fuzz tests mutate real files against the reader, CSV parser and XML tokenizer.

## Status

This is a **0.x** release: the API may still change in minor versions, with changes listed in [CHANGELOG.md](./CHANGELOG.md).

Compatibility is verified in CI by opening generated files with LibreOffice and by reading and writing them with ExcelJS. A 17-file pack covering every feature (including an encrypted file, the streaming writer and real files re-saved after reading; see `test/smoke/excel-pack.mjs`) was opened in **Microsoft Excel for Mac 16.114** with no repair prompts, and its content was checked through Excel (tables, conditional formatting, pictures, protection, chart sheets). Excel for Windows has not been checked yet. If you ever see a "We found a problem with some content" dialog, please open an issue with the file attached.

Not supported yet: `.xls` and `.xlsb`. See the [roadmap](./ROADMAP.md).

## Credits and license

Inspired by and partly derived from [ExcelJS](https://github.com/exceljs/exceljs) by Guyon Roche (MIT). Released under the [MIT license](./LICENSE).
