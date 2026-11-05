# Roadmap

Where turbosheet stands and what is next. Update this file with every release.

## Released

| Version | Highlights |
| --- | --- |
| 0.1.0 | Clean rewrite: core model, streaming XLSX read/write, CSV, JSON, tables, conditional formatting, data validation, notes, images, protection, print setup, defined names, encryption (`/crypto`), round-trip preservation, `/lite` opt-in bundles |
| 0.2.0 | Chart writing (7 types), column auto-fit |
| 0.2.1 | Round-trip fidelity: autoFilter criteria, sort state, named styles, form controls, chartsheet order, sheetFormatPr, table query links |
| 0.3.0 | Formula evaluation (`/formula`, 87 functions), `_xlfn.` future-function prefix fix, ExcelJS parity audit + fixes ([docs/exceljs-parity.md](docs/exceljs-parity.md)) |
| 0.4.0 | Row/column structure edits (`insertRows`, `deleteRows`, `spliceRows`, `spliceColumns`, `duplicateRow`, ... with Excel-style reference adjustment); tables: `tableRows`, `addTableRows`; outline groups; iterators; default font, auto wrap, hyperlink look, ignored errors, page breaks, `setSheetStyle`, `protectSheet`; `localDates`; merges in `xlsxRows`; in-cell checkboxes, `fitImage`, fractional anchors; dynamic-array formulas without `@`; `copyWorksheet` / `copyRange`; `sortRange`, `fillTemplate`; long dropdown lists; document properties; rich-text hyperlinks; ExcelJS migration helpers; `fillFormula`, `calcProperties`, `OFFSET` / `INDIRECT` / `LOOKUP` (90 functions); gradient fills, extra font attributes, workbook views; `cellText` / `formatValue`, `richTextFromHtml`, column helpers |
| 0.5.0 | Native zlib in Node / Bun / Deno (`writeXlsx` / `readXlsx` sync deflate / inflate, off-thread deflate at the requested level in the streaming writers, `zlib.crc32`); `createXlsxWriter` defaults to `low-memory`; metadata parts stored first so `xlsxRows` streams turbosheet's own files; CSV typing (`parseBooleans`, `parseErrors`, `dateFormats`, `localDates`) and encodings; `includeEmpty` iteration, `maxRows` / `maxCols`, `{ sharedFormula }` values; several sheet views, sheet `properties`, `language` / `revision`, note margins / protection |
| 0.6.0 | Random-access `xlsxRows` for bytes and `Blob` / `File` (central directory, metadata first): Excel, ExcelJS and `createXlsxWriter` files stream at constant memory |

## Next candidates

Tracked as [GitHub issues](https://github.com/lamberto15/turbosheet/issues) (many picked from upvoted ExcelJS issues, label `from-exceljs`; none open at the moment). Larger themes:

1. **Formula engine** — statistical and financial functions (structured references, `LET`, dynamic-array spilling and incremental recalculation are done).
2. **Custom sheet views** (`<customSheetViews>`; several sheet views with split / frozen panes are supported since 0.5.0).
3. **Threaded comments** as an editable model (currently preserved raw); chart reading into the model; more chart types (combo, secondary axes, trendlines).
4. **Excel for Windows verification** of the excel pack (only Excel for Mac 16.114 verified so far).
5. Performance gate: profile the read path; Rust/WASM only if JS is clearly the bottleneck.

## Known limitations

See "Not supported yet" in the [README](README.md), [docs/exceljs-parity.md](docs/exceljs-parity.md) (🚫 rows) and the round-trip table in [docs/features.md](docs/features.md).

## How work is done

- Every change: PR into `next`, with a changeset (`pnpm changeset`) when user-visible; required checks: Node 22/24 × ubuntu/windows, quality. The maintainer merges.
- File-format changes are verified in real Microsoft Excel with the excel pack (`pnpm build && node test/smoke/excel-pack.mjs out/pack`), not just LibreOffice/ExcelJS — e.g. the `_xlfn.` bug was only visible in Excel.
- Releases: merge the bot's "chore(release): version packages" PR → CI publishes to npm over OIDC with provenance ([RELEASING.md](RELEASING.md)).
