# Features — how to add one

A _feature_ is an OOXML capability beyond cell data (protection, print setup, conditional
formatting, data validation, tables, comments, images, ...). Each feature is a self-contained
module in this folder; the writers and the reader drive it through the extension points in
[`types.ts`](./types.ts). Reference implementations: [`sheet-protection.ts`](./sheet-protection.ts)
(one leaf element, password hashing) and [`print-setup.ts`](./print-setup.ts) (several
elements, a `sheetPr` child, a subtree with text).

## Files to touch

| File                                                           | What you add                                                                                                                        |
| -------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `src/xlsx/features/<name>.ts` (or `<name>/index.ts` + helpers) | the feature: model types, `<name>Writer: FeatureWriter`, `<name>Reader: FeatureReader`                                              |
| `src/xlsx/features/registry.ts`                                | ONE import line + one entry in `WRITERS` and one in `READERS` (the package root's built-ins)                                        |
| `src/core/features.ts`                                         | the model field(s): one `import type` line + one documented optional field in `WorksheetFeatureFields` (or `WorkbookFeatureFields`) |
| `src/xlsx/features/index.ts`                                   | public exports (model types, helpers, `<name>Writer` / `<name>Reader`) — re-exported from the root and `./lite`                     |
| `src/xlsx/write/sheet.ts` `OWNED`                              | one `'field:<name>'` entry per worksheet field (lite: writing the field without the feature throws)                                 |
| `test/<name>.test.ts`                                          | tests, including the exceljs oracle (below)                                                                                         |

Nothing else: do not edit `read/sheet-part.ts`, `core/worksheet.ts`, ... (in `write/sheet.ts` only
the `OWNED` table). If an extension point is missing, add it to the framework in its own commit.

## Opt-in (`./lite`) and the public feature object

The package root runs every feature in `registry.ts`; the `./lite` entry runs only what the user
passes. So each feature exports its two halves, as named `const`s, from `features/index.ts`:

```ts
// features/data-validation.ts
/** Write half of the data validation feature. ... */
export const dataValidationWriter: FeatureWriter = { name: 'data-validation', ... };
/** Read half ... */
export const dataValidationReader: FeatureReader = { name: 'data-validation', ... };
```

```ts
import { writeXlsx, readXlsx, dataValidationWriter, dataValidationReader } from 'turbosheet/lite';
writeXlsx(wb, { features: [dataValidationWriter] });
readXlsx(buf, { features: [dataValidationReader] });
```

- Halves, not one `{ writer, reader }` object: an object literal keeps both halves alive, so a
  write-only user would ship the reader too (measured: +1.9 KB brotli for tables, +2.8 KB for
  tables + images). Separate `const`s tree-shake per direction.
- Naming: `<camelCaseName>Writer` / `<camelCaseName>Reader` where `name` is the kebab-case
  `FeatureWriter.name` (`'data-validation'` → `dataValidationWriter`); the missing-feature error
  derives the export name from it.
- Keep module top level pure: no calls, `new` only with `/* @__PURE__ */`. A feature module is
  only dropped from a lite bundle when nothing of it is used (module-level `sideEffects: false`
  tree-shaking on the unbundled dist).
- Feature-owned worksheet fields go into `OWNED` (`write/sheet.ts`): writing such a field while
  the feature is disabled throws `ws.<field> is set but the '<name>' feature is not enabled; pass
  features: [<name>Writer]` instead of silently dropping it. Empty arrays are not data.
- A feature that only writes (no reader needed) exports just the writer — see `theme.ts`.

## Model data

Feature data is plain optional properties on `Worksheet` / `Workbook`, declared in
`src/core/features.ts` (type-only; merged into the classes so the published `.d.ts` carries
them — TS module augmentation from feature files would not survive declaration bundling):

```ts
// src/core/features.ts
import type { DataValidation } from '../xlsx/features/data-validation.js';
export interface WorksheetFeatureFields {
  /** Data validations (features/data-validation). */
  dataValidations?: DataValidation[] | undefined;
}
```

Users write `ws.dataValidations = [...]`. Prefer plain data (no methods: methods would need
prototype patching, a side effect). Helper functions taking the worksheet are fine
(`addDataValidation(ws, ...)`), exported from `features/index.ts`. The streaming writer exposes
the same `Worksheet` as `streamSheet.worksheet` (it never holds cells), so features work there
too as long as they only need sheet-level data.

## Writing: `FeatureWriter`

```ts
export const dataValidationWriter: FeatureWriter = {
  name: 'data-validation',
  sheet: {
    emitters: [{ slot: SheetSlot.dataValidations, emit: (c) => dvXml(c.ws.dataValidations) }],
    namespaces: [], // extra xmlns declarations on <worksheet>, if any
  },
  book: {
    write(c) {
      /* workbook-level parts */
    },
    emitters: [
      /* workbook.xml children */
    ],
  },
};
```

- **Emitters** produce one worksheet element and are bound to a `SheetSlot` (the CT_Worksheet
  child order). All emitters — built-in and feature — run sorted by slot, so element order is
  correct by construction. Return `''` when there is nothing to write. One owner per slot.
  - Head slots (< `sheetData`: sheetPr, dimension, sheetViews, sheetFormatPr, cols) run before
    the rows — in `createXlsxWriter` that is when the sheet starts, so set head data before
    adding rows. Tail slots run after the rows (at `commit()` when streaming).
  - Container slots: `sheetPr` and `extLst`. Their emitters return CHILD elements
    (`<pageSetUpPr .../>`, `<ext uri="{...}">…</ext>`) and the framework wraps them once;
    `order` sorts children (`SheetPrOrder`).
  - Emitters run once per sheet per write; keep them pure (derive output from the model).
- **`SheetWriteCtx`** (argument `c`): `ws`, `workbook`, `styles`, `index`, `path`, `streaming`,
  `tabSelected`, plus package services:
  - `c.addRel(type, target, external?)` → `"rIdN"`. `type` is a short officeDocument
    relationship name (`'comments'`, `'vmlDrawing'`, `'table'`, `'drawing'`) or a full URI.
    `target` is relative to the sheet (`'../comments1.xml'`). Hyperlinks share the id space.
  - `c.addPart(path, data, contentType?)` — string or bytes; written right after the sheet part
    in every writer (writeXlsx, xlsxStream, createXlsxWriter). `contentType` adds an Override.
  - `c.addDefault(ext, contentType)` — Default content type (e.g. `'png'`, `'image/png'`).
  - `c.nextIndex(family)` — 1, 2, ... unique per package (`comments${n}.xml`, `table${n}.xml`).
  - A part that itself has relationships (a drawing → images) writes its own rels part with
    `addPart('xl/drawings/_rels/drawing1.xml.rels', ...)`.
- **Drawing anchors** (pictures, charts, shapes): a sheet has ONE drawing part, so no feature
  owns the `drawing` slot. List the shared `drawingEmitter` (`./drawing.ts`) in `emitters` and
  your anchors in `sheet.anchors: [{ order, has(ws), emit(d) }]`: the writer runs the emitter
  once per sheet (shared emitter objects are deduplicated) and it merges every enabled
  feature's anchors (by `order`: images 0, charts 1) plus the anchors kept from a read file
  into `xl/drawings/drawingN.xml`. `emit` gets a `DrawingWriteCtx`: `d.drawingRel(type,
  target)` (relationships of the drawing part), `d.shapeId()` (unique `cNvPr` ids, above the
  kept anchors'), `d.drawingPath`; `anchorXml` / `toAnchor` build and normalize anchors.
- **Styles**: cell styles via `c.styles.intern(style)`; differential formats (conditional
  formatting, tables) via `c.styles.internDxf(style)` → `dxfId` (written to `<dxfs>`). Intern
  dxfs before `styles.xml` is written: emitters run before it in every writer.
- **Book level**: `book.write(ctx)` runs once after all sheets (add workbook parts/rels with
  `ctx.addPart` / `ctx.addRel`); `book.emitters` add workbook.xml children by `BookSlot`.

## Reading: `FeatureReader`

```ts
export const dataValidationReader: FeatureReader = {
  name: 'data-validation',
  sheet: {
    elements: ['dataValidations'], // paths below <worksheet>; 'sheetPr/x' = sheetPr child
    open(ctx, el, attrs) {
      // the element's own start tag
      const list: DataValidation[] = [];
      ctx.ws.dataValidations = list;
      return {
        // receives the subtree (local names), then end()
        open(name, a) {
          /* copy what you need from `a` — attrs objects are shared */
        },
        text(t) {},
        close(name) {},
        end() {},
      };
    },
    finish(ctx) {
      /* after the whole sheet: load related parts via ctx.rels + ctx.readText */
    },
  },
  book(ctx) {
    /* after all sheets */
  },
};
```

- The sheet parser delegates whole subtrees by element path, never anything inside
  `<sheetData>` (the cell hot loop is untouched). Several features may claim the same path
  (`'extLst/ext'`): events are fanned out — check the `uri` attribute.
- `SheetReadCtx`: `ws`, `workbook`, `path`, `rels` (the sheet's relationships, `rel.path` is the
  resolved zip path), `readText(path)` / `readBytes(path)` (lazy inflate),
  `relsOf(path)` (e.g. a drawing's rels), `xf(s)` (file style index → registry xf),
  `dxf(id)` (file dxfId → registry dxf), `scratch` (per-sheet Map for state between `open` and
  `finish`; key it by your reader object).
- Readers run in `readXlsx` (DOM mode) only; `xlsxRows` ignores features.

## Rules

- **Imports**: only `src/core/*`, `src/xml/*`, `./types.ts`, `./util.ts` (bool/num/attrs/
  base64/escape helpers, rich text runs: `runsXml` / `tElem` / `StringItemBuilder`) and your own
  files. Never import runtime code from `src/xlsx/read/*`
  or `src/xlsx/write/*`: every built-in writer ships in the root write-only bundle and every
  reader in the read bundle (and in any lite bundle that enables the feature). Never import
  `registry.ts` from a module `./lite` reaches.
- Web APIs only (no `node:*`, no `Buffer`); no top-level side effects (`/* @__PURE__ */` for
  module-level `new`).
- Size budget (`pnpm build && pnpm size`, see `docs/bundle-size.md`): root write-only ≤ 32 KB,
  read+write ≤ 60 KB brotli, shared by all features; lite write-only ≤ 21 KB (target ~20) (no feature may
  leak into it) — keep emitters string-template based and tables small.
- Escape everything user-provided: `escapeAttr` for attributes, `escapeText` for text (or
  `attrs({...})`, which escapes and skips undefined/false).

## Round-trip preservation

`readXlsx` keeps whatever no reader consumes (`preserveUnknown`, default true; code in
`read/preserve.ts`, model in `common/preserved.ts`, writer in `write/preserve.ts`, reached via
the `w` hooks on the kept data) and `writeXlsx` / `xlsxStream` write it back (`createXlsxWriter`
never has any). A feature disabled at read time (`./lite`) claims nothing, so everything it
would have modelled is kept raw and round-trips unchanged — your reader must not be needed for
that. What that means for a feature:

- **Claiming = owning.** A worksheet element listed in `elements` is consumed: it is not kept,
  and the relationships its `r:id` attributes reference are not kept either. Elements nobody
  claims (and `extLst/ext` entries every claimant declined by returning `undefined`) are kept
  raw and re-emitted in their `SheetSlot`. Claim only what you regenerate.
- **One owner per slot**: a kept element is written only when no emitter produced its slot
  (container slots `sheetPr` / `extLst` merge, minus kept children an emitter produced: same
  `SheetPrOrder`, same ext `uri`). Drop-in support for a new element = claim it. If the kept
  element links parts (`r:id`), add the slot to `OWNER` in `write/preserve.ts` so they are not
  written orphaned when the model takes over (raw and model only meet when a workbook read
  without the feature is written with it).
- **Never emit defaults over kept content**: an emitter that writes something without model data
  (print setup's default `<pageMargins>`) must return `''` when the sheet kept that element.
- **dxf references** in kept XML are remapped (`*dxfId` attributes of sheet fragments and of
  table / pivot parts) because the reader interns, and may merge, the file's dxfs.
- **Relationships no element references** (pivotTable, threadedComments, queryTable, ...) are
  kept; a feature that models such a part must make the preserve layer skip it only when it
  produced model data (see the `comments` exclusion in `Preserver.finishSheet`: raw when
  `ws.comments` is unset, so a lite read keeps the notes).
- **Mixed parts** are decided in `read/preserve.ts`: drawings keep their non-picture anchors
  (`ws.preserved.drawing`, merged by the shared `drawingEmitter` of `./drawing.ts`, which any
  feature contributing drawing anchors enables); a legacy VML with only note shapes is dropped
  (comments regenerate it); its other shapes (form controls) on a sheet with notes are kept in
  `ws.preserved.vml` and merged by the comments writer (`presVml`); an unchanged table whose part
  has relationships (query table) is written from its kept part by the tables writer
  (`presTable`). Kept elements next to a model field: `<autoFilter>` criteria and
  `<sheetFormatPr>` attributes are merged with the built-in emitters' output in `write/preserve.ts`.
- Part names: kept parts keep their path unless the writer generates it; `nextIndex(family)`
  skips numbers taken by kept parts (by basename: `drawing1` reserves drawing 1).
- Invalidation: kept content is not updated when cells change (Excel recalculates charts /
  pivots). Renaming a sheet drops its kept content that references other parts; removing a
  sheet drops everything only it referenced. `calcChain.xml` is always dropped.

## Testing (exceljs oracle)

`exceljs` 4.4 is a devDependency and the reference for semantics and compatibility. For each
feature test, in `test/<name>.test.ts`:

1. **Write → ours**: build a `Workbook`, set the feature data, `writeXlsx`, `readXlsx`, compare.
   Repeat with `xlsxStream` and `createXlsxWriter` (set data on `streamSheet.worksheet`).
2. **Write → exceljs**: load our bytes with
   `await new ExcelJS.Workbook().xlsx.load(buf as any)` and assert exceljs sees the same values.
3. **exceljs → ours**: create the feature with exceljs, `await x.xlsx.writeBuffer()`, read it
   with `readXlsx`, assert.
4. **Raw XML**: `unzipSync(buf)` and check element order/attributes when it matters.

See `test/features.test.ts` for all four patterns (it also registers a dummy feature to cover
every hook). Run `npx tsc --noEmit && npx vitest run && npx biome check .` before committing.
