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
| notes (comments)       | `ws.comments`                        | `commentsWriter`              | `commentsReader`              |
| images                 | `ws.images`, `ws.backgroundImage`    | `imagesWriter`                | `imagesReader`                |
| charts                 | `ws.charts`                          | `chartsWriter`                | — (read charts are kept raw)  |
| conditional formatting | `ws.conditionalFormatting`           | `conditionalFormattingWriter` | `conditionalFormattingReader` |
| data validation        | `ws.dataValidations`                 | `dataValidationWriter`        | `dataValidationReader`        |
| print setup            | `ws.pageSetup`, `ws.headerFooter`, `ws.rowBreaks`, `ws.colBreaks` | `printSetupWriter` | `printSetupReader` |
| ignored errors         | `ws.ignoredErrors`                   | `ignoredErrorsWriter`         | `ignoredErrorsReader`         |
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
| write-only                     | root `{ Workbook, writeXlsx }`                                 | 35 KB | 34.75 KB |
| read+write                     | root `*`                                                       | 76 KB | 75.63 KB |
| lite write-only                | lite `{ Workbook, writeXlsx }`                                 | 23.5 KB | 23.42 KB |
| lite read+write                | lite `{ Workbook, readXlsx, writeXlsx }`                       | 42.5 KB | 42.07 KB |
| lite write + tables + images   | lite + `addTable, tablesWriter, addImage, placeImage, imagesWriter` | 27.5 KB | 27.26 KB |
| lite + autofit                 | lite `{ Workbook, autoFitColumns }`                            | 11 KB | 10.97 KB |
| lite write + charts            | lite + `addChart, chartsWriter`                                 | 27 KB | 26.82 KB |
| csv                            | `./csv` `*`                                                    | 2.9 KB | 2.38 KB |
| crypto                         | `./crypto` `*`                                                 | 7.7 KB | 7.18 KB |
| formula                        | `./formula` `*`                                                | 17 KB | 16.75 KB |

Budget policy: each limit is at most the measured size + 0.5 KB (some are tighter), so growth is a deliberate decision: raise a limit in the same PR as the feature that needs it, with the reason. Where the bytes go in the root write-only bundle (minified): worksheet model 13.5 KB, fflate (deflate) 11.6 KB, workbook / styles XML writer 10.6 KB, style registry 7.5 KB, sheet writer 7.4 KB, conditional formatting 6.6 KB, charts 6.5 KB, tables 4.2 KB. Helpers such as `sortRange`, `fillTemplate`, the row / column edits, `copyWorksheet` and `cellText` cost nothing unless imported.

Native zip helpers (#71) add ~0.15 KB to every write bundle: the runtime `node:zlib` lookup and the `zlib.crc32` hook in the zip writer (the off-thread zlib deflater is only in the async writers). The two lite combo limits were raised by 0.5 KB for it. Sync native zlib (#75: `deflateRawSync` batches in `writeXlsx`, `inflateRawSync` with limit-capped `maxOutputLength` in `readXlsx`) adds ~0.2 KB to write bundles and ~0.5 KB to read+write; the root `*` and lite read+write limits were raised by 0.4 / 0.5 KB.

Row / column structure edits (`insertRows`, `spliceRows`, ...) add ~3.7 KB to `*` and nothing to
bundles that don't import them.

`./formula` is a separate entry, so it costs nothing unless imported (the core bundles above did
not change when it was added). Its budget covers the whole evaluator, because `evaluate` reaches
every built-in through the function table: parser (~2.3 KB), engine (dependency graph, cycles,
array formulas, ~5 KB), coercion and comparison rules incl. date / number text parsing (~1.8 KB),
the `TEXT` number/date formatter (~2 KB) and 87 functions (~6 KB). The 15 KB target was missed by
about 1 KB; 16.5 KB leaves room for fixes without letting it grow unnoticed.

Every writer carries the future-function table (`src/xlsx/common/xlfn.ts`, about 0.6 KB): functions
added after Excel 2007 must be stored as `_xlfn.XLOOKUP(...)` or Excel shows #NAME?. Only the ~100
undotted names are listed (every dotted built-in except ERROR.TYPE is a future function). This is
why the two "lite write + ..." limits moved from 24 KB to 24.5 KB.

The root write-only limit is 32 KB (was 30 KB): the root may grow as features are added, because
lite is the small path. The lite write-only limit is what guards against regressions in the core
writer and against features leaking into lite (it has almost no headroom: keep it that way).

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
echo "import { Workbook, writeXlsx } from './dist/lite.js'; console.log(Workbook, writeXlsx)" > /tmp/e.mjs
npx -y esbuild@0.25.12 /tmp/e.mjs --bundle --minify --format=esm --metafile=/tmp/meta.json --outfile=/tmp/out.js
# inputs of /tmp/meta.json: only dist/xlsx/features/types.js (slot constants) from features/
```

or `npx size-limit --why`.

The root `*` budget (64 KB) measures importing every export at once (all features, auto-fit, charts): a worst case no app hits. The budgets that guard real-world size are the lite ones.

0.3.0 headroom: after formula evaluation (the `_xlfn.` table), font family/scheme and the parity fixes, lite write-only measured 20.97 KB against a 21 KB limit, so budgets were raised by ~1 KB (lite write-only 22 KB, root write-only 33 KB, lite combos 25.5 KB) to keep ordinary fixes from failing CI. The lite write-only target remains ~20 KB.
