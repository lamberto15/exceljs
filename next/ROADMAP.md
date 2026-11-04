# Roadmap

Where turbosheet stands and what is next. Update this file with every release.

## Released

| Version | Highlights |
| --- | --- |
| 0.1.0 | Clean rewrite: core model, streaming XLSX read/write, CSV, JSON, tables, conditional formatting, data validation, notes, images, protection, print setup, defined names, encryption (`/crypto`), round-trip preservation, `/lite` opt-in bundles |
| 0.2.0 | Chart writing (7 types), column auto-fit |
| 0.2.1 | Round-trip fidelity: autoFilter criteria, sort state, named styles, form controls, chartsheet order, sheetFormatPr, table query links |
| 0.3.0 | Formula evaluation (`/formula`, 87 functions), `_xlfn.` future-function prefix fix, ExcelJS parity audit + fixes ([docs/exceljs-parity.md](docs/exceljs-parity.md)) |

## Next candidates (ordered by user value)

1. **Row/column structure edits** — `insertRows`, `spliceRows`/`spliceColumns`, `duplicateRow` (≈45 ExcelJS tests; most-missed migration API). Must shift merges, formulas (reuse `translateFormula`), conditional formatting/validation ranges, tables, defined names.
2. **Document properties** — title, subject, keywords, company, manager, … (`docProps/core.xml` + `app.xml`).
3. **Formula engine** — OFFSET/INDIRECT/LOOKUP, statistical and financial functions, structured references, dynamic-array spilling, incremental recalculation.
4. **Shared-formula authoring** (`fillFormula`) and gradient fills; remaining font attributes (charset, outline, shadow).
5. **Workbook / multiple sheet views**, row/column `collapsed`.
6. **Threaded comments** as an editable model (currently preserved raw); chart reading into the model; more chart types (combo, secondary axes, trendlines).
7. **Excel for Windows verification** of the excel pack (only Excel for Mac 16.114 verified so far).
8. Performance gate: profile the read path; Rust/WASM only if JS is clearly the bottleneck.

## Known limitations

See "Not supported yet" in the [README](README.md), [docs/exceljs-parity.md](docs/exceljs-parity.md) (🚫 rows) and the round-trip table in [docs/features.md](docs/features.md).

## How work is done

- Every change: PR into `next` with a changeset (`pnpm changeset`); required checks: Node 22/24 × ubuntu/windows, Node 24 macOS, quality. The maintainer merges (admin bypass only when the macOS runner is merely queued and everything else is green).
- File-format changes are verified in real Microsoft Excel with the excel pack (`pnpm build && node test/smoke/excel-pack.mjs out/pack`), not just LibreOffice/ExcelJS — e.g. the `_xlfn.` bug was only visible in Excel.
- Releases: merge the bot's "chore(release): version packages" PR → CI publishes to npm over OIDC with provenance ([RELEASING.md](RELEASING.md)).
