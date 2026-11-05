# Bundle size / `./lite`

Two entries expose the same functions and the same `Workbook` model:

| Entry                   | Features                                                                     | Use it when                                                    |
| ----------------------- | ---------------------------------------------------------------------------- | -------------------------------------------------------------- |
| `turbosheet`        | all built-ins, always on                                                     | you want everything to just work (default)                     |
| `turbosheet/lite`   | none by default; opt in per call with `features`                             | bundle size matters (browser, edge), you write plain data      |

```ts
import { Workbook, writeXlsx, readXlsx, addTable, tablesWriter, tablesReader } from 'turbosheet/lite';

const wb = new Workbook();
addTable(wb.addWorksheet('Data'), { name: 'Sales', columns: ['Region', 'Units'], rows: [['N', 1]] });
const bytes = writeXlsx(wb, { features: [tablesWriter] });        // also xlsxStream, createXlsxWriter
const back = readXlsx(bytes, { features: [tablesReader] });       // xlsxRows ignores features
```

Each feature has a write half and a read half, so a write-only bundle never contains readers:

| Feature                | Model fields                         | Write half                    | Read half                     |
| ---------------------- | ------------------------------------ | ----------------------------- | ----------------------------- |
| tables                 | `ws.tables`                          | `tablesWriter`                | `tablesReader`                |
| notes (comments)       | `ws.comments`, `ws.threadedComments`, `wb.persons` | `commentsWriter`    | `commentsReader`              |
| custom views           | `wb.customViews`, `ws.customSheetViews` | `customViewsWriter`        | `customViewsReader`           |
| images                 | `ws.images`, `ws.backgroundImage`    | `imagesWriter`                | `imagesReader`                |
| charts                 | `ws.charts`                          | `chartsWriter`                | `chartsReader`                |
| conditional formatting | `ws.conditionalFormatting`           | `conditionalFormattingWriter` | `conditionalFormattingReader` |
| data validation        | `ws.dataValidations`                 | `dataValidationWriter`        | `dataValidationReader`        |
| print setup            | `ws.pageSetup`, `ws.headerFooter`, `ws.rowBreaks`, `ws.colBreaks` | `printSetupWriter` | `printSetupReader` |
| ignored errors         | `ws.ignoredErrors`                   | `ignoredErrorsWriter`         | `ignoredErrorsReader`         |
| pivot tables           | `ws.pivotTables`                     | `pivotTablesWriter`           | — (read pivots are kept raw)  |
| sheet protection       | `ws.protection` (pulls SHA-512)      | `sheetProtectionWriter`       | `sheetProtectionReader`       |
| default theme part     | —                                    | `themeWriter`                 | — (a file's theme is kept raw) |

Helpers (`addTable`, `addNote`, `placeImage`, `hashPassword`, ...) are exported from both entries
and tree-shake individually. On the root, `features` adds custom `FeatureWriter` /
`FeatureReader`s to the built-ins (duplicates are ignored).

## Behavior in lite

- **No silent data loss on write.** Setting a feature-owned field without its writer throws, e.g.
  `Worksheet 'Data': ws.tables is set but the 'tables' feature is not enabled; pass features:
  [tablesWriter]`. Empty arrays are not data. The streaming writer throws at `commit()` for
  fields set after the rows. A workbook read by the root with charts next to pictures in a
  drawing needs `imagesWriter` (or `chartsWriter`) to write those anchors back (`ws.preserved is set ...`).
- **Reading without a feature keeps it raw.** Its worksheet elements and parts (table parts, notes
  and their VML, drawings and media, conditional formatting incl. x14 extensions, data
  validations, protection, page setup/margins/header-footer, the theme) are kept by
  `preserveUnknown` (default true) and written back unchanged by lite or root writers.
  `preserveUnknown: false` drops them.
- **No double emission.** If you read without a feature and then set its model data on the same
  sheet before writing with it, the model replaces the raw element (and the parts it linked):
  read with the feature if you want to edit existing content. Raw `sheetPr` children / `extLst`
  entries are dropped when the feature emits the same child / ext `uri`; raw legacy notes give
  way to modelled ones.
- **No default theme.** Lite writes no `xl/theme/theme1.xml` unless `themeWriter` is enabled
  (~0.5 KB brotli). Excel, LibreOffice and Numbers open such files and use their built-in
  Office theme for theme colors / fonts. `theme: false` removes it on both entries.
- **Same bytes.** Lite with all writers (`themeWriter` included) produces exactly the root's output.

## Budgets (`pnpm build && pnpm size`, minified + brotli, all dependencies)

| size-limit entry               | Import                                                         | Limit | Size     |
| ------------------------------ | -------------------------------------------------------------- | ----- | -------- |
| write-only                     | root `{ Workbook, writeXlsx }`                                 | 43.4 KB | 43.07 KB |
| read+write                     | root `*`                                                       | 91.3 KB | 91.03 KB |
| lite write-only                | lite `{ Workbook, writeXlsx }`                                 | 26.8 KB | 26.69 KB |
| lite read+write                | lite `{ Workbook, readXlsx, writeXlsx }`                       | 45.6 KB | 45.32 KB |
| lite write + tables + images   | lite + `addTable, tablesWriter, addImage, placeImage, imagesWriter` | 30.8 KB | 30.52 KB |
| lite autofit                   | lite `{ Workbook, autoFitColumns }`                            | 12.8 KB | 12.47 KB |
| lite write + charts            | lite + `addChart, chartsWriter`                                 | 31.1 KB | 30.80 KB |
| csv                            | `./csv` `*`                                                    | 3.3 KB | 3.16 KB |
| crypto                         | `./crypto` `*`                                                 | 7.7 KB | 7.18 KB |
| formula                        | `./formula` `*`                                                | 27.7 KB | 27.27 KB |

Budget policy: each limit is about the measured size + 0.5 KB at most (most are tighter), so growth is a deliberate decision: raise a limit in the same PR as the feature that needs it, with the reason. Where the bytes go in the root write-only bundle (minified, ~108 KB before brotli): worksheet model 14.5 KB, fflate (deflate) 11.3 KB, workbook / styles XML writer 11.1 KB, style registry 8 KB, sheet writer 7.8 KB, conditional formatting 6.5 KB, charts 6.3 KB, tables 4.1 KB, notes 3.6 KB; the zip glue (`write/zipout.ts`, `zip/zip.ts`, the native zlib lookup in `zip/node.ts`) ~2.3 KB. Helpers such as `sortRange`, `fillTemplate`, the row / column edits, `copyWorksheet` and `cellText` cost nothing unless imported.

What group 12 (#95-#98) costs: ~0.03 KB in the lite table / image bundle (renamed fields); the API report and `check:dts` add nothing to the bundles.

What group 20 (#115) costs: pivot table authoring +2.6 KB in the root bundles (lite: only with `pivotTablesWriter` / `addPivotTable`).

What group 19 (#113, #114) costs: the chart writer +0.7 KB (combo groups, secondary axes, trendlines, per-series labels); reading charts into the model +2.5 KB in read+write (`chartsReader`); the reader's part-claiming hook +0.2 KB in every reading bundle.

What group 18 (#111, #112) costs: the root bundles +1.5 KB write / +2.7 KB read+write (threaded comments with persons and their legacy notes in the comments feature; the custom views feature); lite bundles only pay when they import `commentsWriter` / `customViewsWriter`.

What group 17 (#109, #110) costs: `turbosheet/formula` +3.7 KB for 50 financial, date / time, text and information functions; nothing in the core bundles.

What group 16 (#107, #108) costs: `turbosheet/formula` +2.1 KB for 64 math, trigonometric and statistical functions (incl. AGGREGATE); nothing in the core bundles.

What group 15 (#104-#106) costs: `turbosheet/formula` +4.3 KB (dynamic-array evaluation and spilling, FILTER / SORT / SORTBY / UNIQUE / SEQUENCE / TRANSPOSE / RANDARRAY, structured references, LET, SUBTOTAL, incremental recalculation); every writer ~0.45 KB (LET names get `_xlpm.`, `A1#` is written as `ANCHORARRAY(A1)`, #SPILL! / #CALC! as #VALUE!, the shared spill-function list). The core bundles do not include the formula engine.

What group 14 (#103) costs: the table editing functions ~1.4 KB, only in bundles that import them (read+write imports `*`); writing the `[@Col]` shorthand as `[[#This Row],[Col]]` ~0.15 KB on every writer.

What group 13 (#99-#102) costs (brotli, ~1 KB on every xlsx entry; the limits above are the result): shared-formula output (#101: group tracking in the worksheet model plus the `t="shared"` master / follower encoder) ~0.5 KB on write; `cell.hyperlink` / `cell.tooltip`, the `addRow` / `insertRow(s)` style codes and the new sheet view fields ~0.5 KB in the model and sheet reader / writer.

What the 0.5.0 / 0.6.0 changes cost (brotli; the limits above are the result):

- Native zlib (#71, #75: the runtime `node:zlib` lookup, `zlib.crc32`, `deflateRawSync` batches in
  `writeXlsx`, `inflateRawSync` with a limit-capped `maxOutputLength` in `readXlsx`): ~0.35 KB on write
  bundles, ~0.5 KB more on read+write. The off-thread `createDeflateRaw` deflater is only in the async
  writers' modules.
- Metadata parts stored before the worksheets (#74, `ZipOut.reserve`): ~0.1 KB on write bundles.
- Sheet / workbook metadata parity (#82: several sheet views and their windows, `ws.properties` in
  `<sheetPr>` / `<sheetFormatPr>` incl. the `x14ac` namespace, `language` / `revision`, read back): ~0.4 KB
  on the core writer, ~0.45 KB on read+write. Note margins / protection / `editAs` live in the comments
  feature (root only).
- ExcelJS iteration / shared-formula compatibility (#80): ~1 KB on every bundle with a `Workbook`:
  `{ sharedFormula }` cell values need `translateFormula` in the worksheet model (~0.55 KB), plus the
  `{ includeEmpty }` paths of `eachRow` / `row.eachCell`, `column.eachCell` and the master-cell lookup.
- CSV read typing and decoding (#81: `parseBooleans`, `parseErrors`, `dateFormats`, `localDates`,
  `encoding` with BOM sniffing and `'auto'`): `./csv` from 2.38 to 3.14 KB.
- Random-access `xlsxRows` (#86: `unzipBlob`, metadata-first entry order, deferred size checks in
  `unzipLazy`): ~1.1 KB on root `*` (bundles importing `xlsxRows`), ~0.15 KB on lite read+write.
- `TEXT` fraction formats (`# ?/?`, `??/??`, `?/8`) ~0.4 KB on `./formula` (limit 17 → 17.5 KB);
  completing partial `_xlfn.` prefixes ~0.05 KB on every writer (lite read+write limit 43.4 → 43.6 KB).

Row / column structure edits (`insertRows`, `spliceRows`, ...) add ~3.7 KB to `*` and nothing to
bundles that don't import them.

`./formula` is a separate entry, so it costs nothing unless imported (the core bundles above did
not change when it was added). Its budget covers the whole evaluator, because `evaluate` reaches
every built-in through the function table: parser (~2.3 KB), engine (dependency graph, cycles,
array formulas, ~5 KB), coercion and comparison rules incl. date / number text parsing (~1.8 KB),
the `TEXT` number/date formatter (~2 KB) and 90 functions (~6 KB). The original 15 KB target was
missed; the 17.5 KB limit leaves a little room for fixes without letting it grow unnoticed.

Every writer carries the future-function table (`src/xlsx/common/xlfn.ts`, about 0.6 KB): functions
added after Excel 2007 must be stored as `_xlfn.XLOOKUP(...)` or Excel shows #NAME?. Only the ~100
undotted names are listed (every dotted built-in except ERROR.TYPE is a future function).

The root write-only bundle may grow as features are added, because lite is the small path. The lite
write-only limit is what guards against regressions in the core writer and against features leaking into lite (it has almost no headroom: keep it that way).

How the small lite bundles are achieved:

- the writer/reader cores take feature lists (`options.features`); only the root's
  `write/index.ts` / `read/index.ts` import the built-in list (`features/registry.ts`);
- the preservation write side is reached through hooks `readXlsx` attaches to `wb.preserved`,
  so write-only bundles never include it;
- the dist is unbundled (one file per module, `sideEffects: false`), so bundlers drop whole
  unused feature modules — with one shared chunk, property accesses in the feature objects'
  initializers (`SheetSlot.tableParts`) kept every feature alive.

Check what a bundle contains with esbuild's metafile, e.g.

```sh
echo "import { Workbook, writeXlsx } from '$PWD/dist/lite.js'; console.log(Workbook, writeXlsx)" > /tmp/e.mjs
npx -y esbuild@0.25.12 /tmp/e.mjs --bundle --minify --format=esm --metafile=/tmp/meta.json --outfile=/tmp/out.js
# outputs[...].inputs of /tmp/meta.json with bytesInOutput > 0: only dist/xlsx/features/types.js
# (slot constants) from features/
```

or `npx size-limit --why`.

The root `*` budget (78.8 KB) measures importing every export at once (all features, `xlsxRows`, auto-fit,
charts, the row / column edits): a worst case no app hits. The budgets that guard real-world size are the
lite ones.
