# ExcelJS parity

Which ExcelJS 4.4 behaviours turbosheet supports, measured against the ExcelJS test suite. Every ExcelJS integration, issue, PR and end-to-end test, and the behaviour-level unit tests, was ported to turbosheet's API in [`test/exceljs-parity/`](../test/exceljs-parity/) (one file per area; each test cites its original with `// exceljs: <spec> — "<title>"`). exceljs 4.4.0 is used there as an oracle: many tests also check that exceljs loads what turbosheet writes, and [`interop.test.ts`](../test/exceljs-parity/interop.test.ts) reads exceljs-written files.

Status legend:

- ✅ **supported**: the ported test passes.
- 🐛 **turbosheet bug**: the port failed because turbosheet was wrong. All are fixed (see [bugs found](#bugs-found-and-fixed)); the ported test is the regression test.
- 🚫 **not supported**: the feature is missing; the port is `it.skip` with the reason.
- ➖ **intentionally different**: API or semantics differ by design (sync I/O, 0-based arrays, immutable interned styles, ...); the port asserts turbosheet's behaviour and the comment explains the difference.
- ❓ **not applicable**: tests ExcelJS internals (Anchor, Range, colCache caches, StreamBuf, ...) or Node file APIs.

Counts are per `it(...)` in the ExcelJS sources (a loop such as the three "Hidden" fixtures or the nine `slideFormula` cases counts once). XForm unit tests (`spec/unit/xlsx/**`, serializer internals) are out of scope and not counted.

## Summary

| Group | ✅ | 🐛 | 🚫 | ➖ | ❓ | total |
| --- | --- | --- | --- | --- | --- | --- |
| end-to-end | 0 | 1 | 0 | 0 | 0 | 1 |
| integration | 72 | 7 | 23 | 45 | 6 | 153 |
| issues | 23 | 0 | 1 | 3 | 0 | 27 |
| prs | 9 | 0 | 1 | 3 | 0 | 13 |
| unit | 35 | 0 | 57 | 21 | 64 | 177 |
| **all** | 139 | 8 | 82 | 72 | 70 | 371 |

Most ➖ tests pass with a small API translation (for example `ws.setRowProps(2, { style })` for `row.font = ...`). Ported tests that also cover the turbosheet side of a difference live next to them.

### By area

| Area (test file) | ✅ | 🐛 | 🚫 | ➖ | ❓ | total |
| --- | --- | --- | --- | --- | --- | --- |
| [conditional-formatting](../test/exceljs-parity/conditional-formatting.test.ts) | 1 | 0 | 0 | 1 | 0 | 2 |
| [csv](../test/exceljs-parity/csv.test.ts) | 0 | 0 | 0 | 2 | 0 | 2 |
| [data-validation](../test/exceljs-parity/data-validation.test.ts) | 2 | 0 | 0 | 1 | 0 | 3 |
| [defined-names](../test/exceljs-parity/defined-names.test.ts) | 2 | 0 | 1 | 5 | 1 | 9 |
| [formulas](../test/exceljs-parity/formulas.test.ts) | 7 | 1 | 8 | 0 | 0 | 16 |
| [hyperlinks](../test/exceljs-parity/hyperlinks.test.ts) | 0 | 1 | 2 | 0 | 0 | 3 |
| [images](../test/exceljs-parity/images.test.ts) | 11 | 0 | 0 | 1 | 18 | 30 |
| [issues](../test/exceljs-parity/issues.test.ts) | 23 | 0 | 1 | 3 | 0 | 27 |
| [merges](../test/exceljs-parity/merges.test.ts) | 7 | 0 | 1 | 7 | 0 | 15 |
| [misc](../test/exceljs-parity/misc.test.ts) | 27 | 0 | 5 | 7 | 4 | 43 |
| [notes](../test/exceljs-parity/notes.test.ts) | 1 | 0 | 3 | 0 | 0 | 4 |
| [print](../test/exceljs-parity/print.test.ts) | 3 | 0 | 0 | 0 | 0 | 3 |
| [prs](../test/exceljs-parity/prs.test.ts) | 9 | 0 | 1 | 3 | 0 | 13 |
| [rows-columns](../test/exceljs-parity/rows-columns.test.ts) | 21 | 0 | 48 | 21 | 0 | 90 |
| [streaming-read](../test/exceljs-parity/streaming-read.test.ts) | 3 | 0 | 5 | 1 | 0 | 9 |
| [streaming-write](../test/exceljs-parity/streaming-write.test.ts) | 5 | 0 | 0 | 4 | 0 | 9 |
| [styles](../test/exceljs-parity/styles.test.ts) | 5 | 0 | 0 | 10 | 0 | 15 |
| [tables](../test/exceljs-parity/tables.test.ts) | 0 | 0 | 7 | 2 | 0 | 9 |
| [values](../test/exceljs-parity/values.test.ts) | 12 | 6 | 0 | 4 | 0 | 22 |
| (internals, no port) | 0 | 0 | 0 | 0 | 47 | 47 |

## Bugs found and fixed

| Bug | Found by | Fix |
| --- | --- | --- |
| Negative `alignment.textRotation` (the ExcelJS convention, -90..-1) was written as `textRotation="-45"`, which is outside OOXML's 0..180/255 range, so Excel reports a corrupt file. Now stored in Excel's encoding (90 + \|deg\|, e.g. -45 → 135); values outside the range are dropped. | `workbook.spec.js` "xlsx file" (test book), `styles.test.ts` regression | `fix(styles)` |
| Negative `alignment.indent` was written as `indent="-1"` (unsigned in OOXML). Now dropped. | test book bad alignments | `fix(styles)` |
| A 6-digit `argb` (`'FF0000'`) was written as `rgb="FF0000"` (ST_UnsignedIntHex needs 8 digits) and read back as `FFFF0000`, so the model changed on a round trip. Now padded with an opaque alpha in the model and in every `rgb` attribute (styles, rich text, tab color). | test book border `thickRainbow`, streaming "rich text" | `fix(styles)` |
| `Worksheet.addRow` was not atomic: a rejected value (e.g. an ExcelJS `{ sharedFormula }`) threw after the row record was created, so every later `addRow` landed one row lower (in-memory and streaming writer). | `workbook-xlsx-writer.spec.js` "shared formula" | `fix(core)` |
| `Workbook.removeWorksheet` did not update sheet-scoped defined names (`localSheetId` is a tab index): print area / print titles / local names of later sheets moved to a neighbour or vanished, and the removed sheet's names were written against another sheet. | defined-names port (extra check next to `workbook.spec.js` "defined names") | `fix(core)` |

| The hyperlink of a **formula cell** was skipped by the reader (a cell holds one value kind) and lost on a read → write round trip. It is now kept raw in `ws.preserved` (relationship ids remapped) and written back while the cell still holds the formula it was read with; changing or clearing the cell drops it. | `workbook-xlsx-reader.spec.js` "should contain the linked url" (`formulas.xlsx` A1) | `fix(read)` |

No failure was caused by the `_xlfn.` future-function prefix work happening on another branch: `formulas.xlsx` keeps `_xlfn.CONCAT(...)` exactly as stored, which is what ExcelJS expects.

## Notable unsupported features

- Row/column structure edits: `insertRow(s)`, `spliceRows`, `spliceColumns`, `duplicateRow`, row splice (≈45 ExcelJS tests).
- Shared-formula authoring (`fillFormula`, `{ sharedFormula }` values); shared formulas in files are expanded on read.
- Gradient fills; font `charset`, `outline`, `shadow`, `condense`, `extend` (dropped on read, so a read → write round trip loses them).
- Workbook views (`wb.views`), several sheet views per sheet, `activePane`, row/column `collapsed`, sheet `properties` (defaultRowHeight, outlineLevelCol/Row).
- Document properties beyond creator / lastModifiedBy / created / modified (company, manager, title, subject, keywords, category, description, language, revision, contentStatus).
- Note margins / protection / editAs; page breaks (kept raw on round trip but not modelled); pivot table authoring; table editing API (`table.commit()`); `maxRows` / `maxCols` read limits; `getRows`, `getSheetValues`, `column.eachCell`, `eachRow({ includeEmpty })`, `actualRowCount`.
- CSV: custom read date formats (ISO 8601 only), boolean / error typing, non-UTF-8 encodings.

## Additions made for parity

Decided by the maintainer after the first audit and implemented on the same branch:

- **Font `family` and `scheme`** are modelled (`Font.family?: number`, `Font.scheme?: 'major' | 'minor' | 'none'`): read from styles, rich text and notes, written after the font name, part of the style identity. The default font is Calibri 11, family 2, scheme minor (as Excel writes it); a partial font without its own name inherits them. Editing and saving an Excel file keeps `<scheme val="minor"/>`.
- **ExcelJS alignment aliases** are accepted: `vertical: 'middle'` (stored as `'center'`) and `textRotation: 'vertical'` (stored as 255), next to negative rotations.
- **Empty CSV fields leave cells empty**: `arraysToSheet` / `objectsToSheet` treat `''` as no cell (option `emptyAs: 'string'` keeps empty-string cells), so `arraysToSheet(ws, parseCsv(text))` matches ExcelJS.
- **Hyperlinks on formula cells** are preserved raw (see the bug table).

## Remaining notes for the maintainer

- `removeWorksheet` drops names scoped to the removed sheet; preserved workbook XML that refers to sheets by index is not adjusted (none found in the fixtures).
- A kept formula-cell hyperlink is matched by object identity: assigning an equal formula object also drops it (documented in features.md).

## Every ExcelJS test

Area = file in `test/exceljs-parity/` (`(skip)` = `it.skip` there). Spec paths are relative to `spec/`.

### `end-to-end/express.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| downloads a workbook | 🐛 | values | web `Response(xlsxStream(wb))` instead of express; failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |

### `integration/gold.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Styles | ✅ | values | empty in exceljs |
| Values | ➖ | values | no `lastColumn`/`lastRow` objects; values identical |

### `integration/issues/issue-1027-sheet-defined-data-validate.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 1027 - Broken due to Cannot set property 'marked' of undefined error | ✅ | issues | whole-column list validation |

### `integration/issues/issue-1328-xlsx-worksheet-reader-date.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 1328 - should emit row with Date Object | ✅ | issues | xlsxRows returns Date |

### `integration/issues/issue-1339-speciel-cell-file.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 1339 - Special cell value results invalid file | ✅ | issues | streaming writer |

### `integration/issues/issue-1364-incorrect-worksheet-name-streaming-xlsx.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 1364 - Incorrect Worksheet Name on Streaming XLSX Reader | ✅ | issues | xlsxRows `row.sheet` |

### `integration/issues/issue-163-xlsx-read-file.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 163 - Error while using xslx readFile method | ✅ | issues |  |

### `integration/issues/issue-1669-optional-custom-autofilter-on-table.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 1669 - optional autofilter and custom autofilter on tables | ✅ | issues | also round-trips |

### `integration/issues/issue-176-unexpected-xml-node.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 176 - Unexpected xml node in parseOpen | ✅ | issues |  |

### `integration/issues/issue-1842-dataValidations-memory-overload.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| when loading an in memory buffer | ✅ | issues | validations modelled per sqref, not expanded per cell |
| when using readFile | ✅ | issues | no `ignoreNodes`: lite without `dataValidationReader` keeps the XML raw |

### `integration/issues/issue-2125-spliceRows-last-row.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 2125 - spliceRows remove last row | 🚫 | issues (skip) | `spliceRows` not supported |

### `integration/issues/issue-219-1904-dates.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Reading 1904.xlsx | ✅ | issues |  |
| Writing and Reading | ✅ | issues |  |

### `integration/issues/issue-234-vertical-tab-char.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 234 - Broken XLSX because of "vertical tab" ascii character in a cell | ➖ | issues | control characters are kept as Excel `_xHHHH_` escapes (exact round trip) instead of stripped; file stays valid |

### `integration/issues/issue-257-sheet-order.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 257 - worksheet order is not respected | ✅ | issues |  |

### `integration/issues/issue-275-hyperlink-query-param.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 275 - hyperlink with query arguments corrupts workbook | ✅ | issues | streaming writer; exceljs oracle agrees |

### `integration/issues/issue-539-contentType-element.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Reading 1904.xlsx | ✅ | issues |  |

### `integration/issues/issue-623-borders-for-merged-cell.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 623 - Issue with borders for merged cell when rewriting an excel workbook | ✅ | issues |  |

### `integration/issues/issue-703-speciel-cell-file.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 703 - Special cell value results invalid file | ✅ | issues |  |

### `integration/issues/issue-771-data-validation-without-type.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 771 - Issue with dataValidation without type and with formula1 or formula2 | ✅ | issues | missing type reads as `any` |

### `integration/issues/issue-877-hyperlink-no-text.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 877 - hyperlink without text crashes on write | ✅ | issues |  |

### `integration/issues/issue-880-malformed-comment.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 880 - malformed comment crashes on write | ✅ | issues |  |

### `integration/issues/issue-988-missing-table-autofilter.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 988 - table without autofilter model | ✅ | issues |  |

### `integration/issues/issue-991-csv-read-dates.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 991 - differentiates between strings with leading numbers and dates when reading csv files | ➖ | issues | ISO 8601 only, parsed as UTC; US `11-04-2019` stays text |

### `integration/issues/issue-995-encoding.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| issue 995 - encoding option works fine | ✅ | issues | CSV is always UTF-8 (no `encoding` option) |

### `integration/issues/shared-formula.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| copied cells should have the right formulas | ➖ | issues | shared formulas are expanded on read (`{ formula: "A4+1" }`, no `sharedFormula`/`shareType`) |
| copied cells should have the right types | ✅ | issues |  |
| copied cells should have the same fields | ✅ | issues |  |

### `integration/pr/pr-896/test-pr-896.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| should store cell text and comment with leading new line | ✅ | prs |  |

### `integration/pr/test-pr-1157.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 1204 - Read and write data validation should be successful | ➖ | prs | formulas read back as text (`formula1: "2"`), one entry per sqref |

### `integration/pr/test-pr-1220.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 1220 - The worksheet should not be undefined | ✅ | prs |  |

### `integration/pr/test-pr-1262.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 1262 - protect should work with streaming workbook writer | ✅ | prs | `s.worksheet.protection = { password, spinCount }` |

### `integration/pr/test-pr-1334.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 1334 - Fix the error that comment does not delete at spliceColumn | 🚫 | prs (skip) | `spliceColumns` not supported |

### `integration/pr/test-pr-1431.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 1431 - streaming reader should handle rich text within shared strings | ✅ | prs |  |

### `integration/pr/test-pr-1487.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Reading 1904.xlsx | ➖ | prs | no `lastColumn`; `columnCount` |

### `integration/pr/test-pr-1576.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Reading test-issue-1575.xlsx | ✅ | prs |  |

### `integration/pr/test-pr-2244.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 2244- Fix xlsx.writeFile() not catching error when error occurs | ➖ | prs | no `filename` image source; missing data / unknown image id throws |

### `integration/pr/test-pr-567.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Should be able to read this file | ✅ | prs |  |

### `integration/pr/test-pr-728.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| pull request 728 - Read worksheet hidden state | ✅ | prs |  |

### `integration/pr/test-pr-896.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Should preserve leading and trailing whitespace | ✅ | prs |  |
| Should preserve newlines | ✅ | prs |  |

### `integration/workbook-xlsx-reader.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Should anchor height equals to row height | ❓ | images (skip) | exceljs Anchor internals |
| Should anchor height equals to row height | ❓ | images (skip) | exceljs Anchor internals |
| Should anchor width equals to column width when custom | ❓ | images (skip) | exceljs Anchor internals |
| Should anchor width equals to column width when custom | ❓ | images (skip) | exceljs Anchor internals |
| Should integer part of col equals nativeCol | ✅ | images | anchors use whole cell indexes + offsets |
| Should integer part of col equals nativeCol | ✅ | images | anchors use whole cell indexes + offsets |
| Should integer part of row equals nativeRow | ✅ | images | anchors use whole cell indexes + offsets |
| Should integer part of row equals nativeRow | ✅ | images | anchors use whole cell indexes + offsets |
| edit styles of single row instead of all | ✅ | styles | immutable interned styles make this safe by design |
| should bail out if the file contains more cells than the limit | 🚫 | streaming-read (skip) | no `maxRows`/`maxCols` (break out of `xlsxRows`) |
| should bail out if the file contains more rows than the limit | 🚫 | streaming-read (skip) | no `maxRows`/`maxCols` (break out of `xlsxRows`) |
| should bail out if the file contains more rows than the limit | 🚫 | streaming-read (skip) | no `maxRows`/`maxCols` (break out of `xlsxRows`) |
| should be classified as a formula cell | ✅ | formulas / hyperlinks | hyperlinked formula: kind is Formula (➖) |
| should be classified as a formula cell | ✅ | formulas / hyperlinks | hyperlinked formula: kind is Formula (➖) |
| should contain the linked url | 🐛 | hyperlinks | hyperlink on a formula cell was dropped on read and lost on round trip; now kept raw and written back while the cell holds that formula (`fix(read)`) |
| should decode the underscore | ✅ | misc |  |
| should fail fast on a huge file | 🚫 | streaming-read (skip) | no `maxRows`/`maxCols` (break out of `xlsxRows`) |
| should fail fast on a huge file | 🚫 | streaming-read (skip) | no `maxRows`/`maxCols` (break out of `xlsxRows`) |
| should have text corresponding to the evaluated formula result | ✅ | formulas / hyperlinks |  |
| should have text corresponding to the evaluated formula result | ✅ | formulas / hyperlinks |  |
| should have the formula source | ✅ | formulas / hyperlinks | `_xlfn.` prefix kept as stored |
| should have the formula source | ✅ | formulas / hyperlinks | `_xlfn.` prefix kept as stored |
| should not break | ✅ | misc |  |
| should not crash | ✅ | defined-names |  |
| should parse fine if the limit is not exceeded | ✅ | streaming-read |  |
| should parse fine if the limit is not exceeded | ✅ | streaming-read |  |
| should parse fine if the limit is not exceeded | ✅ | streaming-read |  |
| should reject the promise with the sax error | ✅ | misc | `XlsxParseError` |
| xlsx file | ➖ | streaming-read | strings/dates resolved, rows without values skipped, no row heights |

### `integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| A lot of sheets | ✅ | streaming-write |  |
| Cell annotation supports setting margins and protection properties | 🚫 | notes (skip) | note margins/protection/editAs not modelled |
| Without styles | ✅ | values |  |
| addRow | ✅ | streaming-write |  |
| auto filter | ✅ | streaming-write |  |
| creates sheets with correct names | ➖ | streaming-write | default `SheetN` |
| defined names | ➖ | defined-names | `addDefinedName(w.workbook, …)` |
| does not escape special xml characters | ✅ | values |  |
| rich text | ➖ | values | 6-digit argb reads back with alpha `FF` |
| serializes and deserializes dataValidations | ✅ | data-validation |  |
| serializes row styles and columns properly | ➖ | styles | row style given to `addRow`; styled column without width reads Excel default width |
| shared formula | 🐛 | formulas | `{ sharedFormula }` is not a value (➖); the rejected row used to stay half-written and shift later rows — fixed (`fix(core)`) |
| with background image | ✅ | images |  |
| with background image where worksheet is commited in advance | ✅ | images |  |
| with conditional formatting | ✅ | conditional-formatting |  |
| with conditional formatting that contains numFmt (#1814) | ➖ | conditional-formatting | rule `style.numFmt` is the format code (exceljs `{ id, formatCode }`) |
| with zip compression option | ✅ | data-validation |  |
| writes notes | ✅ | notes |  |
| xlsx file | ✅ | values | row-level styles only (no per-cell styles in a streamed row) |

### `integration/workbook/images.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| image extensions should not be case sensitive | ✅ | images |  |
| stores background image | ✅ | images |  |
| stores embedded image and hyperlink | ✅ | images |  |
| stores embedded image with hyperlinks | ✅ | images |  |
| stores embedded image with one-cell-anchor | ✅ | images |  |
| stores embedded image with oneCell | ➖ | images | cell + pixel offsets instead of fractional cells |

### `integration/workbook/pivot-tables.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| if pivot table NOT added, then certain xml and rels files are not added | ✅ | misc |  |
| if pivot table added, then certain xml and rels files are added | 🚫 | misc (skip) | pivot authoring not supported (read pivots are preserved) |

### `integration/workbook/styles.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| in-cell formats properly in xlsx file | ✅ | styles | rich text identical, incl. font `family`/`scheme` |
| null cells retain style | ✅ | styles |  |
| row styles and columns properly | ➖ | styles | row style via `setRowProps`; values identical |
| sets col styles | ➖ | styles | row/column style applies to new/empty cells only; existing cells keep theirs |
| sets row styles | ➖ | styles | row/column style applies to new/empty cells only; existing cells keep theirs |

### `integration/workbook/workbook.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| CSV file and its configuration | ➖ | csv | delimiter + `dateFormat` supported; custom read date formats are not (ISO only) |
| Duplicate rows replacing properly | 🚫 | rows-columns (skip) | `duplicateRow` not supported |
| Duplicate rows shifting properly | 🚫 | rows-columns (skip) | `duplicateRow` not supported |
| Duplicate rows with height properly | 🚫 | rows-columns (skip) | `duplicateRow` not supported |
| Duplicate rows with styles properly | 🚫 | rows-columns (skip) | `duplicateRow` not supported |
| a lot of sheets to xlsx file | ✅ | misc |  |
| auto filter | ✅ | misc |  |
| company, manager, etc | 🚫 | misc (skip) | only creator/lastModifiedBy/created/modified are modelled |
| creator, lastModifiedBy, etc | ✅ | misc |  |
| csv file | ➖ | csv | via `sheetToArrays`/`writeCsv`/`parseCsv`/`arraysToSheet`; hyperlinks as text, booleans/errors untyped |
| dataValidations | ➖ | data-validation | sqref-based model, defined name instead of `cell.name`, formulas read as text |
| defined names | ➖ | defined-names | no `cell.name(s)`/`addName`/`getRanges`; `addDefinedName` with ranges; sheet names always quoted |
| empty string | ✅ | values |  |
| empty strings | ✅ | values |  |
| frozen panes | ➖ | misc | one sheet view per sheet; only modelled view fields |
| language, revision and contentStatus | 🚫 | misc (skip) | only creator/lastModifiedBy/created/modified are modelled |
| multiple book views | 🚫 | misc (skip) | `wb.views` not supported |
| printTitlesColumn | ✅ | print | `setPrintTitles` / `getPrintTitles` |
| printTitlesRow | ✅ | print | `setPrintTitles` / `getPrintTitles` |
| printTitlesRowAndColumn | ✅ | print | `setPrintTitles` / `getPrintTitles` |
| serialises and deserialises properly | ➖ | merges | slave cells read `null` (no merge proxy) |
| serialises split panes | ➖ | misc | one view per sheet; `activePane` not modelled |
| shared formula | 🚫 | formulas | no `fillFormula`; plain formulas round-trip |
| sheets with correct names | ➖ | misc | default name is `SheetN` (exceljs `sheetN`) |
| spliced meat and ham | 🚫 | rows-columns (skip) | splice not supported |
| styles | ➖ | merges | merging does not copy the master style; set it per cell |
| throw an error for wrong data type | ✅ | misc | different message |
| throws an error when csv file not found | ❓ | misc/csv (skip) | no file system API |
| throws an error when xlsx file not found | ❓ | misc/csv (skip) | no file system API |
| title, subject, etc | 🚫 | misc (skip) | only creator/lastModifiedBy/created/modified are modelled |
| xlsx file | 🐛 | values | failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |
| xlsx file with best compression | 🐛 | values | `compression` 9/default/1/0; failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |
| xlsx file with default compression | 🐛 | values | `compression` 9/default/1/0; failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |
| xlsx file with fast compression | 🐛 | values | `compression` 9/default/1/0; failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |
| xlsx file with no compression | 🐛 | values | `compression` 9/default/1/0; failed on negative `textRotation` (written as invalid `textRotation="-45"`), negative `indent` and 6-digit RGB colors; fixed (`fix(styles)`). test book round trip (`test-book.ts`); subset: no gradient fills / font outline / workbook views / sharedFormula values / merged-cell value proxy |

### `integration/worksheet-xlsx-writer.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| adds column headers | ✅ | rows-columns |  |
| adds column headers by letter | ➖ | rows-columns | no `defn`/`headers`; setting `header` does not write row 1 |
| adds column headers by number | ➖ | rows-columns | no `defn`/`headers`; setting `header` does not write row 1 |
| adds columns | ✅ | rows-columns |  |
| adds multiple row breaks | 🚫 | rows-columns (skip) | page breaks not modelled (kept raw) |
| adds rows by contiguous array | ✅ | rows-columns |  |
| adds rows by object | ✅ | rows-columns |  |
| adds rows by sparse array | ➖ | rows-columns | arrays are 0-based (index 0 = column A) |
| assigns cell types properly | ➖ | streaming-write | streamed rows cannot be read back; checked in the written file |
| does not allow overlapping merges | ✅ | merges |  |
| references the same top-left value | ➖ | merges |  |
| sets col styles | ➖ | styles | see styles.spec |
| sets row styles | ➖ | styles | see styles.spec |
| stores shared string values properly | ➖ | streaming-write | streamed rows cannot be read back; checked in the written file |
| stores values properly | ➖ | streaming-write | streamed rows cannot be read back; checked in the written file |

### `integration/worksheet.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Should not break when importing a .numbers file | ✅ | misc | workbook without worksheets |
| Should not break when importing an Excel file that contains a chartsheet | ✅ | misc |  |
| Should set hidden attribute correctly (${file}) | ✅ | rows-columns | 3 fixtures |
| adds column headers | ✅ | rows-columns |  |
| adds column headers by letter | ➖ | rows-columns | no `defn`/`headers` |
| adds column headers by number | ➖ | rows-columns | no `defn`/`headers` |
| adds columns | ✅ | rows-columns |  |
| adds rows by contiguous array | ➖ | rows-columns | no `getRows`; `eachCell` gets values |
| adds rows by object | ➖ | rows-columns | no `getRows`; `eachCell` gets values |
| adds rows by sparse array | ➖ | rows-columns | 0-based arrays; no `getRows` |
| adds rows with style option | 🚫 | rows-columns (skip) | `insertRow(s)` / style codes not supported |
| assigns cell types properly | ✅ | values |  |
| does not allow overlapping merges | ✅ | merges |  |
| does not allow overlapping merges | ✅ | merges |  |
| inserts rows by contiguous array | ➖ | rows-columns | no `getRows`; `eachCell` gets values |
| inserts rows by object | ➖ | rows-columns | no `getRows`; `eachCell` gets values |
| inserts rows by sparse array | ➖ | rows-columns | 0-based arrays; no `getRows` |
| inserts rows with style option | 🚫 | rows-columns (skip) | `insertRow(s)` / style codes not supported |
| iterates over collumn cells | 🚫 | rows-columns (skip) | `column.eachCell` not supported |
| iterates over rows | ➖ | rows-columns | no `includeEmpty` |
| keep first 31 characters | ➖ | rows-columns | throws instead of truncating |
| merges and unmerges | ✅ | merges | `unmergeCells` |
| merges styles | ➖ | merges | no style copy on merge |
| puts the lotion in the basket | ✅ | rows-columns |  |
| references the same top-left value | ➖ | merges | no merge proxy |
| returns sheet values | ➖ | rows-columns | `sheetToArrays` (dense, 0-based) |
| returns undefined when row range is less than 1 | 🚫 | rows-columns (skip) | `getRows` not supported |
| save the original name | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| sets col styles | ➖ | styles | see styles.spec |
| sets row styles | ➖ | styles | see styles.spec |
| should style of the inserted row with inherited style be mutable | 🚫 | rows-columns (skip) | `insertRow(s)` / style codes not supported |
| stores shared string values properly | ✅ | values |  |
| stores values properly | ➖ | values | a formula `result` cannot be a Date |
| thrown an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| throws an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| throws an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| throws an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| throws an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| when empty should thrown an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |
| when isn't string should thrown an error | ✅ | rows-columns | messages differ; >31 chars throws instead of truncating (see next) |

### `unit/doc/anchor.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| should colWidth equals 640000 when column has not set custom width | ❓ | images (skip) | exceljs Anchor class |
| should colWidth equals 640000 when worksheet is undefined | ❓ | images (skip) | exceljs Anchor class |
| should colWidth equals column width | ❓ | images (skip) | exceljs Anchor class |
| should integer part of col and colOff should be always equals | ❓ | images (skip) | exceljs Anchor class |
| should integer part of row and rowOff should be always equals | ❓ | images (skip) | exceljs Anchor class |
| should recalculate col | ❓ | images (skip) | exceljs Anchor class |
| should recalculate row | ❓ | images (skip) | exceljs Anchor class |
| should rowHeight equals 180000 when row has not set height | ❓ | images (skip) | exceljs Anchor class |
| should rowHeight equals 180000 when worksheet is undefined | ❓ | images (skip) | exceljs Anchor class |
| should rowHeight equals row height | ❓ | images (skip) | exceljs Anchor class |
| should update colWidth | ❓ | images (skip) | exceljs Anchor class |
| should update nativeColOff after col has been changed | ❓ | images (skip) | exceljs Anchor class |
| should update nativeRowOff after row has been changed | ❓ | images (skip) | exceljs Anchor class |
| should update rowHeight | ❓ | images (skip) | exceljs Anchor class |

### `unit/doc/cell.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Cell comments supports setting margins, protection, and position properties | 🚫 | notes (skip) | note margins/protection/editAs |
| can set comment | 🚫 | notes (skip) | note margins/protection/editAs |
| doesn't upgrade from non-string to hyperlink | 🚫 | hyperlinks (skip) | no `cell.hyperlink` setter |
| escapes dangerous html | ❓ | — | Cell class internals (no `cell.html`, no Merge type) |
| has effective types | ➖ | values | no `effectiveType`; inspect `result` |
| inherits column styles | ✅ | styles |  |
| inherits row styles | ✅ | styles |  |
| merges | ❓ | — | Cell class internals (no `cell.html`, no Merge type) |
| shares formulas | 🚫 | formulas (skip) |  |
| stores values | ✅ | values |  |
| upgrades from string to hyperlink | 🚫 | hyperlinks (skip) | no `cell.hyperlink` setter |
| validates options on construction | ❓ | — | Cell class internals (no `cell.html`, no Merge type) |

### `unit/doc/column.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| creates by defn | ❓ | — | Column model internals (`values`, `defn`, model) |
| creates model | ❓ | — | Column model internals (`values`, `defn`, model) |
| gets column values | ❓ | — | Column model internals (`values`, `defn`, model) |
| maintains properties | ❓ | — | Column model internals (`values`, `defn`, model) |
| sets column values | ❓ | — | Column model internals (`values`, `defn`, model) |
| sets default column width | ❓ | — | Column model internals (`values`, `defn`, model) |
| sets sparse column values | ❓ | — | Column model internals (`values`, `defn`, model) |
| sets sparse column values | ❓ | — | Column model internals (`values`, `defn`, model) |

### `unit/doc/defined-names.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| adds names for cells | ➖ | defined-names | range-based API |
| creates matrix from model | ❓ | defined-names (skip) |  |
| gets the right ranges for a name | ➖ | defined-names | range-based API |
| removes names for cells | ➖ | defined-names | range-based API |
| skips values with invalid range | ✅ | defined-names | bogus names kept as-is |
| splices | 🚫 | defined-names (skip) |  |

### `unit/doc/range.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| constructs as expected | ❓ | — | Range class (address math covered by col-cache ports) |
| detects containment | ❓ | — | Range class (address math covered by col-cache ports) |
| detects intersections | ❓ | — | Range class (address math covered by col-cache ports) |
| doesn't always include the default row/col | ❓ | — | Range class (address math covered by col-cache ports) |
| expands properly | ❓ | — | Range class (address math covered by col-cache ports) |
| has a valid default value | ❓ | — | Range class (address math covered by col-cache ports) |

### `unit/doc/row.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| builds a model | ❓ | — | Row model internals |
| builds from model | ❓ | — | Row model internals |
| counts cells | ❓ | — | Row model internals |
| iterates over cells | ❓ | — | Row model internals |
| remove almost to end | 🚫 | rows-columns (skip) | row splice |
| remove and insert fewer | 🚫 | rows-columns (skip) | row splice |
| remove and insert more | 🚫 | rows-columns (skip) | row splice |
| remove and insert replacements | 🚫 | rows-columns (skip) | row splice |
| remove only | 🚫 | rows-columns (skip) | row splice |
| remove past end | 🚫 | rows-columns (skip) | row splice |
| remove to end | 🚫 | rows-columns (skip) | row splice |
| stores cells | ❓ | — | Row model internals |
| stores values by whole row | ❓ | — | Row model internals |

### `unit/doc/workbook-writer.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| returns undefined for non-existant sheet | ✅ | streaming-write |  |

### `unit/doc/workbook.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| assigns cell types properly | ✅ | misc / values |  |
| assigns rich text | ✅ | misc / values |  |
| inserts duplicates | 🚫 | rows-columns (skip) |  |
| overwrites with duplicates | 🚫 | rows-columns (skip) |  |
| returns undefined for non-existant sheet | ✅ | misc / values |  |
| returns undefined for sheet 0 | ✅ | misc / values |  |
| returns undefined for sheet 0 after accessing wb.worksheets or wb.eachSheet  | ✅ | misc / values |  |
| serialises to model | ❓ | — | skipped in exceljs |
| stores shared string values properly | ✅ | misc / values |  |

### `unit/doc/worksheet-table.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| adds a column | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| adds a row | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| creates a table | ➖ | tables | `topLeft` (exceljs `ref`), no `totalsRowResult`/`filterButton` |
| moves the table | ➖ | tables | placed at C2 directly (no `table.commit()`) |
| removes a column | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| removes a row | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| removes header | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| removes totals | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |
| renames a column | 🚫 | tables (skip) | table editing API (`commit`, add/remove rows/columns) |

### `unit/doc/worksheet-writer.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| generates valid xml even when there is no data | ✅ | streaming-write |  |

### `unit/doc/worksheet.merge.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| does not allow overlapping merges | ✅ | merges |  |
| does not allow overlapping merges | ✅ | merges |  |
| merges and unmerges | ✅ | merges |  |
| merges styles | ➖ | merges |  |
| preserves merges after row inserts | 🚫 | merges (skip) |  |
| references the same top-left value | ➖ | merges |  |

### `unit/doc/worksheet.pagebreak.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| adds multiple row breaks | 🚫 | rows-columns (skip) |  |

### `unit/doc/worksheet.shared-formula.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Fills formula across using 1D array values | 🚫 | formulas (skip) | `fillFormula` |
| Fills formula down and across using 1D array values | 🚫 | formulas (skip) | `fillFormula` |
| Fills formula down using 1D array values | 🚫 | formulas (skip) | `fillFormula` |
| Fills formula using 2D array values | 🚫 | formulas (skip) | `fillFormula` |
| Fills formula using function | 🚫 | formulas (skip) | `fillFormula` |
| Translates formulae to slave cells | 🚫 | formulas (skip) | `fillFormula` |

### `unit/doc/worksheet.styles.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| sets col styles | ➖ | styles | row/column style applies to new/empty cells only; existing cells keep theirs |
| sets row styles | ➖ | styles | row/column style applies to new/empty cells only; existing cells keep theirs |

### `unit/doc/worksheet.values.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Insert defined names | 🚫 | rows-columns (skip) | splice |
| Insert defined names | 🚫 | rows-columns (skip) | splice |
| Insert style | 🚫 | rows-columns (skip) | splice |
| Insert style | 🚫 | rows-columns (skip) | splice |
| Remove and insert fewer | 🚫 | rows-columns (skip) | splice |
| Remove and insert fewer | 🚫 | rows-columns (skip) | splice |
| Remove and insert more | 🚫 | rows-columns (skip) | splice |
| Remove and insert more | 🚫 | rows-columns (skip) | splice |
| Remove and insert same | 🚫 | rows-columns (skip) | splice |
| Remove and insert same | 🚫 | rows-columns (skip) | splice |
| Remove defined names | 🚫 | rows-columns (skip) | splice |
| Remove defined names | 🚫 | rows-columns (skip) | splice |
| Remove only | 🚫 | rows-columns (skip) | splice |
| Remove style | 🚫 | rows-columns (skip) | splice |
| Remove style | 🚫 | rows-columns (skip) | splice |
| Replace defined names | 🚫 | rows-columns (skip) | splice |
| Replace defined names | 🚫 | rows-columns (skip) | splice |
| Replace style | 🚫 | rows-columns (skip) | splice |
| Replace style | 🚫 | rows-columns (skip) | splice |
| Splices almost to end | 🚫 | rows-columns (skip) | splice |
| Splices past end | 🚫 | rows-columns (skip) | splice |
| Splices to end | 🚫 | rows-columns (skip) | splice |
| adds column headers | ✅ | rows-columns |  |
| adds column headers by letter | ➖ | rows-columns |  |
| adds column headers by number | ➖ | rows-columns |  |
| adds columns | ✅ | rows-columns |  |
| adds rows by contiguous array | ✅ | rows-columns |  |
| adds rows by object | ✅ | rows-columns |  |
| adds rows by sparse array | ➖ | rows-columns |  |
| assigns cell types properly | ✅ | values |  |
| calculates columnCount and actualColumnCount | ➖ | rows-columns | no `actualRowCount`/`actualColumnCount` |
| calculates rowCount and actualRowCount | ➖ | rows-columns | no `actualRowCount`/`actualColumnCount` |
| handles column keys | 🚫 | rows-columns (skip) | splice |
| iterates over collumn cells | 🚫 | rows-columns (skip) |  |
| iterates over rows | ➖ | rows-columns |  |
| returns sheet values | ➖ | rows-columns |  |
| splices columns | 🚫 | rows-columns (skip) | splice |
| stores shared string values properly | ✅ | values |  |
| stores values properly | ✅ | values |  |

### `unit/doc/worksheet.views.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| adjusts collapsed property of columns | 🚫 | rows-columns (skip) | `collapsed` not modelled |
| adjusts collapsed property of row | 🚫 | rows-columns (skip) | `collapsed` not modelled |
| sets outline levels via column headers | ✅ | rows-columns |  |

### `unit/utils/cell-matrix.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| findCell only returns known cells | ❓ | — | exceljs internals |
| getCell always returns a cell | ❓ | — | exceljs internals |

### `unit/utils/col-cache.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| caches values | ❓ | misc (skip) | cache internals |
| convert [sheetName!][$]col[$]row[[$]col[$]row] into address or range structures | ➖ | misc | sheet prefix accepted but not returned |
| converts letters to numbers | ✅ | misc | 0-based helpers |
| converts numbers to letters | ✅ | misc | 0-based helpers |
| decodes addresses | ✅ | misc | 0-based helpers |
| decodes addresses and ranges | ✅ | misc | 0-based helpers |
| gets address structures (and caches them) | ❓ | misc (skip) | cache internals |
| throws when out of bounds | ✅ | misc | 0-based helpers |
| tolerates a missing column number | ➖ | misc | lone column/row is a range (`decodeRange`), `decodeAddr` throws |
| tolerates a missing row number | ➖ | misc | lone column/row is a range (`decodeRange`), `decodeAddr` throws |
| validates addresses properly | ✅ | misc | 0-based helpers |

### `unit/utils/copy-style.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| should copy a style deeply | ❓ | — | exceljs internals |
| should copy fill.stops deeply | ❓ | — | exceljs internals |
| should return the argument if a falsy value passed | ❓ | — | exceljs internals |

### `unit/utils/shared-formula.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| ${args[0]} from ${args[1]} to ${args[2]} | ✅ | formulas | `translateFormula`, 9 cases |

### `unit/utils/shared-strings.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Does not escape values | ❓ | — | exceljs internals |
| Stores and shares string values | ❓ | — | exceljs internals |

### `unit/utils/stream-buf.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| handle unsupported type of chunk | ❓ | — | exceljs internals |
| handles buffers | ❓ | — | exceljs internals |
| signals end | ❓ | — | exceljs internals |
| writes StringBuf chunks | ❓ | — | exceljs internals |
| writes strings as UTF8 | ❓ | — | exceljs internals |

### `unit/utils/string-buf.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| grows properly | ❓ | — | exceljs internals |
| resets | ❓ | — | exceljs internals |
| writes strings as UTF8 | ❓ | — | exceljs internals |

### `unit/utils/under-dash.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| works on complex arrays | ❓ | — | exceljs internals |
| works on complex objects | ❓ | — | exceljs internals |
| works on simple values | ❓ | — | exceljs internals |

### `unit/utils/utils.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| '${fmt}' a date | ✅ | misc |  |
| '${fmt}' is not a date | ✅ | misc |  |
| encodes xml text | ➖ | misc | `escapeText`/`escapeAttr`; `'` and DEL not escaped |
| should convert date to excel properly | ✅ | misc |  |
| should not lost millisecond precision when parsing excel date | ✅ | misc |  |
| should round to the nearest millisecond when parsing excel date | ✅ | misc |  |

### `unit/utils/xml-stream.spec.js`

| Test | Status | Area | Notes |
| --- | --- | --- | --- |
| Writes simple XML doc | ❓ | — | exceljs internals |
| Writes text in XML doc | ❓ | — | exceljs internals |
| attributes are escaped | ❓ | — | exceljs internals |
| rolls back | ❓ | — | exceljs internals |
| text is escaped | ❓ | — | exceljs internals |
