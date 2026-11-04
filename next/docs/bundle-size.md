# Bundle size / `./lite`

Two entries expose the same functions and the same `Workbook` model:

| Entry                   | Features                                                                     | Use it when                                                    |
| ----------------------- | ---------------------------------------------------------------------------- | -------------------------------------------------------------- |
| `xlsxkit`        | all built-ins, always on                                                     | you want everything to just work (default)                     |
| `xlsxkit/lite`   | none by default; opt in per call with `features`                             | bundle size matters (browser, edge), you write plain data      |

```ts
import { Workbook, writeXlsx, readXlsx, addTable, tablesWriter, tablesReader } from 'xlsxkit/lite';

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
| conditional formatting | `ws.conditionalFormatting`           | `conditionalFormattingWriter` | `conditionalFormattingReader` |
| data validation        | `ws.dataValidations`                 | `dataValidationWriter`        | `dataValidationReader`        |
| print setup            | `ws.pageSetup`, `ws.headerFooter`    | `printSetupWriter`            | `printSetupReader`            |
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
  drawing needs `imagesWriter` to write those anchors back (`ws.preserved is set ...`).
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
| write-only                     | root `{ Workbook, writeXlsx }`                                 | 32 KB | 28.37 KB |
| read+write                     | root `*`                                                       | 60 KB | 55.15 KB |
| lite write-only                | lite `{ Workbook, writeXlsx }`                                 | 21 KB | 19.96 KB |
| lite read+write                | lite `{ Workbook, readXlsx, writeXlsx }`                       | 40 KB | 36.21 KB |
| lite write + tables + images   | lite + `addTable, tablesWriter, addImage, placeImage, imagesWriter` | 24 KB | 23.43 KB |
| csv                            | `./csv` `*`                                                    | 5 KB  | 2.38 KB  |
| crypto                         | `./crypto` `*`                                                 | 8 KB  | 7.18 KB  |

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
