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
| `src/xlsx/features/registry.ts`                                | ONE import line + one entry in `WRITERS` and one in `READERS`                                                                       |
| `src/core/features.ts`                                         | the model field(s): one `import type` line + one documented optional field in `WorksheetFeatureFields` (or `WorkbookFeatureFields`) |
| `src/xlsx/features/index.ts`                                   | public exports (model types, helpers) — re-exported from the package root                                                           |
| `test/<name>.test.ts`                                          | tests, including the exceljs oracle (below)                                                                                         |

Nothing else: do not edit `write/sheet.ts`, `read/sheet-part.ts`, `core/worksheet.ts`, ... If an
extension point is missing, add it to the framework in its own commit.

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
  base64/escape helpers) and your own files. Never import runtime code from `src/xlsx/read/*`
  or `src/xlsx/write/*`: every registered writer ships in the write-only bundle and every reader
  in the read bundle.
- Web APIs only (no `node:*`, no `Buffer`); no top-level side effects (`/* @__PURE__ */` for
  module-level `new`).
- Size budget (`pnpm build && pnpm size`): write-only ≤ 30 KB, read+write ≤ 60 KB brotli, shared
  by all features — keep emitters string-template based and tables small.
- Escape everything user-provided: `escapeAttr` for attributes, `escapeText` for text (or
  `attrs({...})`, which escapes and skips undefined/false).

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
