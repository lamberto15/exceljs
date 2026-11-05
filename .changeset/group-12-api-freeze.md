---
"turbosheet": minor
---

API freeze before 1.0 (group 12, #95-#98). **Breaking:**

- `XlsxRow.values` is renamed to `XlsxRow.cells`. It is still 0-based (`cells[0]` is column A); the new name keeps it from being mistaken for the 1-based `RowView.values`. Replace `row.values` with `row.cells` in `xlsxRows` loops.
- Internals are no longer exported or typed. The published `.d.ts` is built with `stripInternal`.
  - Removed exports: `RowData`, `SharedStrings`, `fontMetrics`, `resolveWriteOptions` / `ResolvedWriteOptions`, `drawingEmitter`, and the `Xf`, `NumFmt` and `CellCallback` types.
  - `StyleRegistry` is exported as a type only (`wb.styles`).
  - `ws.rowData`, `ws.columnData`, `ws.setCellRaw`, `ws.addMergeRaw` and `RowView.data` still exist at runtime but are internal (not in the types, not covered by semver).
- `FeatureWriter` / `FeatureReader` are opaque in the types (only `name`). Pass the built-in feature halves to `features` as before. Custom features use internal extension points and are not covered by semver.

New: `api/report.txt`, kept current by a test, lists every export of every entry point, and `pnpm check:dts` type-checks the published declarations (run in CI). The indexing rule and the stability policy are documented in docs/api.md ("API stability").
