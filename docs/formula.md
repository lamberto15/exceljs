# Formula evaluation (`turbosheet/formula`)

`turbosheet/formula` parses Excel formulas and computes them against a `Workbook`. It is a
separate entry point: the core bundles (`turbosheet`, `turbosheet/lite`) never include it.

```ts
import { readXlsx, writeXlsx } from 'turbosheet';
import { evaluate, evaluateCell, evaluateFormula, track } from 'turbosheet/formula';

const wb = readXlsx(bytes);
const report = evaluate(wb); // every formula cell gets a fresh Formula.result
// report: { evaluated, circular: ['Sheet1!A1', ...], unsupported: ['NPV', ...], errors: [...] }
const out = writeXlsx(wb); // the results are written as cached values (<v>)

evaluateCell(wb, 'Sheet1', 'D10'); // one cell and what it depends on
evaluateFormula(wb, 'SUM(B2:B9)*Rate', { sheet: 'Sheet1' }); // ad-hoc text, nothing written

const calc = track(wb); // evaluate once, keep the dependency graph
calc.set('Sheet1', 'B2', 5); // recalculates only what depends on B2 → ['Sheet1!D10', ...]
```

Contents: [API](#api) - [Incremental recalculation](#incremental-recalculation) - [Semantics](#semantics) - [Structured references](#structured-references) - [Dynamic arrays](#dynamic-arrays) - [Functions](#functions) - [TEXT formats](#text-formats) - [Limits](#limits) - [Performance](#performance)

## API

```ts
evaluate(wb, options?): EvaluationReport
evaluateCell(wb, sheet: string | Worksheet, address: string, options?): FormulaResult | null
evaluateFormula(wb, text, { sheet?, address?, array?, now? }?): FormulaResult | FormulaResult[][]
parseFormula(text): FormulaNode                 // throws FormulaParseError (message + position)
createEvaluator({ functions?, now? }): Evaluator // .evaluate / .evaluateCell / .evaluateFormula / .registerFunction
supportedFunctions(): string[]

type FormulaResult = number | string | boolean | CellError
```

`EvaluateOptions` (for `evaluate` and `evaluateCell`):

| Option | Default | Meaning |
| --- | --- | --- |
| `sheets` | all | Evaluate the formulas of these sheets (names or `Worksheet`s). Formulas they depend on in other sheets are evaluated too (and written). |
| `writeResults` | `true` (`evaluate`), `false` (`evaluateCell`) | Store results into `Formula.result`. With `false` the report has `results: Map<'Sheet1!A1', FormulaResult>` instead and the workbook is not modified. |
| `maxIterations` | `0` | Iterative calculation for circular references (Excel's "Enable iterative calculation"): at most this many passes over each cycle, starting from the cached results. `0`: circular cells get `#REF!`. |
| `maxChange` | `0.001` | Iteration stops when no value of the cycle changes by more than this. |
| `onUnsupported(name, address)` | | Called once per unsupported function name, with the first cell using it. |
| `now` | `() => new Date()` | Clock for `TODAY()` / `NOW()` (inject it in tests). |

`EvaluationReport`: `evaluated` (formula cells computed), `circular` (addresses on a cycle),
`unsupported` (sorted function names; their calls return `#NAME?`), `errors` (formulas that do not
parse, and exceptions thrown by custom functions: `{ address, message }`), `results` (when
`writeResults` is `false`).

Evaluation never throws because of a formula: problems become Excel errors in the result
(`#NAME?`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#N/A`, `#NUM!`). `evaluateFormula` throws
`FormulaParseError` for text that does not parse, and `RangeError` for an unknown sheet.

### Incremental recalculation

```ts
track(wb, options?): Recalc        // full evaluation (results written), dependency graph kept
calc.set(sheet, address, value)    // set a cell, recalculate its dependents; returns their addresses
calc.update([{ sheet, address }])  // after editing cells yourself (ws.setCell ...)
calc.recalc()                      // full evaluation again (after structural edits)
calc.report                        // EvaluationReport of the last full evaluation
```

After a change, the formulas whose references contain a changed cell, or a recalculated formula
cell, are re-evaluated in dependency order; nothing else is touched. Formulas using volatile or
indirect functions (`INDIRECT`, `OFFSET`, `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`,
`CELL`, `INFO`) are recalculated on every change. A full evaluation is done instead when the
changed cell is or becomes a formula, lies in an array / spill area, or when a formula on a
circular reference or a dynamic-array formula is affected (its spill area may change). Cell
moves (insertRows, sortRange, ...) and formula edits made outside `set` need `recalc()`.
Circular references are detected in both modes (`report.circular`; their cells get `#REF!`, or
iterate with `maxIterations`).

### Custom functions

Functions are registered per evaluator instance; nothing global is mutated.

```ts
import { createEvaluator, toNumber, toText } from 'turbosheet/formula';

const ev = createEvaluator({
  functions: { DOUBLE: (args) => toNumber(args[0]!) * 2 },
  now: () => new Date(2024, 0, 15),
});
ev.registerFunction('GREET', (args) => `Hello ${toText(args[0]!)}`);
ev.evaluate(wb);
```

A function gets its arguments already evaluated: a reference arrives as a `FormulaRange`
(`sheet`, `top`, `left`, `bottom`, `right`, `rows`, `cols`, `at(i, j)`), an array constant as a
`FormulaArray` (`rows`, `cols`, `data`, `at(i, j)`), an omitted argument as `null`. The helpers
`toNumber`, `toText`, `toBoolean` (Excel coercion; a one-cell range gives its value),
`scalarValues` (every value of a range or array) and `isError` are exported. Return a value or a
`CellError` (`{ error: '#N/A' }`); a thrown `CellError` becomes the result too (the helpers throw
`#VALUE!` this way). Any other exception gives `#VALUE!` and an entry in `report.errors`. The
second argument is the `FunctionContext`: `workbook`, `sheet`, `row` / `col` (0-based host cell),
`array` (array-evaluation mode), `date1904`, `now()`.

## Semantics

- **Order and cycles.** Formula cells form a dependency graph (static references, ranges and
  defined names). It is evaluated dependencies first (Tarjan's strongly connected components,
  iterative, so 100k-long chains do not overflow the stack). Cells on a cycle (including
  `A1 = A1 + 1`) get `#REF!` (as Google Sheets does; Excel shows 0 and a warning) and are listed
  in `report.circular`; cells depending on them get the error through normal propagation.
- **Formulas as stored by the reader.** Shared formulas are already expanded per cell. Array
  formulas (`Formula.array`, e.g. `'C1:C3'`) are evaluated in array mode; the master cell gets the
  top-left value and, with `writeResults`, the other cells of the area are overwritten with the
  rest of the array (a smaller result repeats a one-wide dimension, else `#N/A`, as Excel fills).
- **Ranges in single-value positions (implicit intersection).** Non-array formulas follow
  legacy Excel semantics, which is how files store them: `=A1:A3*2` in row 2 uses `A2`; with no
  intersecting row / column the result is `#VALUE!`. The same applies to function arguments that
  take one value (`ABS(A1:A3)`) and to operators inside functions (`SUM(A1:A3*2)` intersects;
  Excel 365 shows it as `SUM(@A1:A3*2)`). Inside array formulas and in `SUMPRODUCT` arguments,
  operators and one-value functions work element-wise with Excel broadcasting. `@` is the explicit
  intersection operator.
- **Conversions.** Arithmetic converts text that looks like a number (`"1,000"`, `" 12 "`,
  `"50%"`, `"$5"`, `"(3)"`, `"1e3"`), a date (`"2024-01-31"`, `"1/31/2024"` (US order),
  `"31-Jan-2024"`, `"Jan 31, 2024"`) or a time (`"13:30"`, `"1:30 PM"`); other text is `#VALUE!`
  (`"TRUE"` is not converted). `TRUE` / `FALSE` are 1 / 0. A blank cell is 0 in arithmetic, `""`
  in text and FALSE in logic. SUM-style functions ignore text, logicals and blanks inside ranges
  but convert direct arguments (`SUM("5", TRUE)` = 6; `SUM(A1:A2)` with `"5"` and `TRUE` = 0).
  Numbers convert to text like Excel's General format with up to 15 significant digits
  (`1/3&""` = `"0.333333333333333"`, `1E+21`, `1E-07`).
- **Comparisons.** Text compares case-insensitively; different types order number < text <
  logical (`1 < "0"` and `"z" < FALSE` are TRUE); blank compares as 0, `""` or FALSE; numbers are
  equal when they agree to 15 significant digits (`0.1+0.2=0.3` is TRUE). Lookups and criteria
  match case-insensitively; `MATCH`, `VLOOKUP` / `HLOOKUP` exact mode, `XLOOKUP` / `XMATCH`
  match mode 2, `COUNTIF`-family criteria and `SEARCH` understand the wildcards `*`, `?` and `~`
  (escape).
- **Precision.** Calculations use doubles; results are rounded to 15 significant digits when
  stored (as Excel displays and stores them), so `=0.1+0.2` stores `0.3`. Display formatting is
  not applied. Non-finite results are `#NUM!`; text results longer than 32767 characters are
  `#VALUE!`.
- **Results.** A formula that evaluates to a blank (`=Z99`, `=IF(TRUE,)`) gives 0, like Excel.
  A reference result gives the referenced value (with implicit intersection).
- **Dates** are serial numbers in the workbook's date system (`wb.date1904` is respected by
  `DATE`, `YEAR`, `TEXT`, ...). In the 1900 system the fictitious 1900-02-29 (serial 60) is
  reported as 1900-03-01 by `YEAR` / `MONTH` / `DAY` / `TEXT`; every other serial matches Excel.
  `TODAY()` and `NOW()` use the local wall-clock time of the injectable clock and are recomputed
  on every `evaluate`.
- **Names.** Sheet names are case-insensitive; quoted names (`'My Sheet'!A1`) are supported.
  Defined names resolve sheet-local before workbook scope, may hold ranges, constants or formulas,
  and may reference other names (16 levels). An unknown name is `#NAME?`; a name whose value
  contains `#REF!` evaluates to `#REF!`.

## Structured references

Table references resolve against `ws.tables` (any sheet; table names are case-insensitive):
`Sales[Amount]` (data rows), `Sales[]`, `Sales[#All]`, `Sales[#Headers]`, `Sales[#Totals]`,
`Sales[#Data]`, combinations (`Sales[[#Headers],[#Data],[Amount]]`), column ranges
(`Sales[[Qty]:[Amount]]`), the current row (`Sales[@Amount]`, `Sales[[#This Row],[Amount]]`,
`[@Amount]` inside the table) and escaped names (`Sales[Unit'#]`). Dependencies follow them
(a calculated column depends on its own row only). An unknown table or column, or a missing
header / totals row, is `#REF!`; `@` in a row that has no data row of the table is `#VALUE!`.
`LET(name, value, ..., calculation)` binds names left to right (values may be ranges); the
`_xlpm.` prefix files use is accepted, and the writers add it.

## Dynamic arrays

A dynamic-array formula (`Formula.dynamic`, or one using a spilling function such as `FILTER`,
`SORT`, `UNIQUE`, `SEQUENCE`, `TRANSPOSE` — see [features](./features.md)) is evaluated in array
mode and its result spills: `evaluate` writes the values into the cells below / right of it and
sets `Formula.array` to the spilled range, so the file opens with them. When the area is not
empty (other than the formula's own previous values), off the sheet, or overlaps the area of an
earlier dynamic formula, the result is `#SPILL!` (written to files as `#VALUE!`, as Excel's plain
value for it). `A1#` refers to the spilled range (`SUM(A1#)`, `ROWS(A1#)`; `#REF!` when A1
holds no dynamic formula or did not spill); files store it as `ANCHORARRAY(A1)`, which is read
the same way. Formulas reading spilled cells see the spilled values: when a spill area changes,
the workbook is evaluated again with the new areas (usually one extra pass). Values left from
a previous, larger spill are cleared.

## Functions

99 built-in functions. `formatValue(value, numFmt)` / `cellText(ws, addr)` apply number formats like Excel's display (the `TEXT` renderer; General shows at most 11 characters). `supportedFunctions()` lists them; anything else evaluates to `#NAME?`
and is reported in `unsupported`. Files store newer functions with an `_xlfn.` prefix; it is
stripped (`_xlfn.CONCAT` is `CONCAT`).

| Category | Functions | Notes |
| --- | --- | --- |
| Math | `SUBTOTAL` `SUM` `PRODUCT` `AVERAGE` `MIN` `MAX` `COUNT` `COUNTA` `COUNTBLANK` `SUMIF` `SUMIFS` `AVERAGEIF` `AVERAGEIFS` `COUNTIF` `COUNTIFS` `SUMPRODUCT` `ROUND` `ROUNDUP` `ROUNDDOWN` `INT` `MOD` `ABS` `POWER` `SQRT` `PI` `SIGN` | Criteria: `5`, `"5"`, `">=5"`, `"<>x"`, `"a*"`, `"?b"`, `"~*"`, `""` (blank or empty text), `"="` (blank), `"<>"` (non-blank), `TRUE`, `"#N/A"`, dates (`">=1/1/2024"`); numbers stored as text match numeric criteria for `=` / `<>`. `SUMIF`'s sum range takes the criteria range's shape; the `*IFS` ranges must have the same shape (`#VALUE!`). `SUMPRODUCT` treats non-numbers as 0. `ROUND` rounds half away from zero after a 15-digit correction (`ROUND(1.005,2)` = 1.01). `MIN`/`MAX` of nothing is 0, `AVERAGE` of nothing `#DIV/0!`. `SUBTOTAL` 1-11 / 101-111 over references skips cells holding `SUBTOTAL` / `AGGREGATE` formulas; 101-111 also skip hidden rows. |
| Logical | `IF` `IFS` `IFERROR` `IFNA` `LET` `AND` `OR` `NOT` `XOR` `SWITCH` `TRUE` `FALSE` | `IF`, `IFS`, `IFERROR`, `IFNA`, `SWITCH` and `CHOOSE` only evaluate the branch they return. `IF` with an array condition (array formulas) picks element-wise. `AND`/`OR`/`XOR` ignore text and blanks in ranges and give `#VALUE!` when there is no logical value. |
| Lookup | `VLOOKUP` `HLOOKUP` `XLOOKUP` `LOOKUP` `INDEX` `MATCH` `XMATCH` `CHOOSE` `OFFSET` `INDIRECT` `ROW` `COLUMN` `ROWS` `COLUMNS` | Approximate modes use Excel's binary search on sorted data (`MATCH` -1: descending). `XLOOKUP`/`XMATCH`: match modes 0, -1, 1, 2 and search modes 1, -1 (2 / -2 are searched linearly). `XLOOKUP` and `INDEX` return references (`SUM(INDEX(A1:B5,0,2))`, `XLOOKUP(...):XLOOKUP(...)`). `INDEX` with a row or column of 0 returns the whole column / row. `OFFSET` and `INDIRECT` return references (negative `OFFSET` heights / widths extend up / left; `INDIRECT` takes A1 text, sheet-qualified refs and defined names, not R1C1). `LOOKUP`: vector and array forms, approximate match on sorted data. |
| Text | `CONCAT` `CONCATENATE` `TEXTJOIN` `LEFT` `RIGHT` `MID` `LEN` `LOWER` `UPPER` `PROPER` `TRIM` `SUBSTITUTE` `REPLACE` `FIND` `SEARCH` `TEXT` `VALUE` `EXACT` `REPT` | `FIND` is case-sensitive, `SEARCH` case-insensitive with wildcards. `TRIM` removes ASCII spaces only. `TEXT`: see [TEXT formats](#text-formats). `VALUE` uses the conversions above. Lengths count UTF-16 code units (as Excel). |
| Date | `DATE` `TODAY` `NOW` `YEAR` `MONTH` `DAY` `WEEKDAY` `EDATE` `EOMONTH` `DATEDIF` `DAYS` | `DATE` normalizes month/day overflow and adds 1900 to years below 1900. `WEEKDAY` types 1, 2, 3, 11-17. `DATEDIF` units `Y` `M` `D` `MD` `YM` `YD` (`MD` follows Excel's formula and can be negative at month ends, as in Excel). |
| Dynamic arrays | `FILTER` `SORT` `SORTBY` `UNIQUE` `SEQUENCE` `TRANSPOSE` `RANDARRAY` | Results spill (see [Dynamic arrays](#dynamic-arrays)). `FILTER` with no match and no fallback, `UNIQUE` / `SEQUENCE` with an empty result give `#CALC!`. `SORT` / `SORTBY` order numbers < text < logicals < errors, blanks last, stable; several sort keys (`SORT(A1:C9,{1,2},{1,-1})`), `by_col`. `UNIQUE` compares text case-insensitively; `exactly_once`. |
| Info | `ISBLANK` `ISNUMBER` `ISTEXT` `ISNONTEXT` `ISLOGICAL` `ISERROR` `ISERR` `ISNA` `NA` | A formula cell is never blank (`ISBLANK` of a cell holding `=""` is FALSE). |
| Other | `SINGLE` | What files store for the `@` operator. |

## TEXT formats

`TEXT(value, format)` supports:

- digit placeholders `0` `#` `?`, the decimal point, the thousands separator (`#,##0`), scaling
  commas after the last digit (`0.0,,"M"`), percent (`0%`, `0.0%`), scientific (`0.00E+00`),
  literal text (`"text"`, `\x`, `$ - + ( ) : / space`), `_x` (a space) and `*x` (ignored);
  positional literals between placeholders (`000-00-0000`);
- sections `positive;negative;zero;text` (`#,##0.00;(#,##0.00)`, `0;-0;"zero"`, `@`);
- `General`;
- dates and times: `yyyy` `yy` `mmmmm` `mmmm` `mmm` `mm` `m` `dddd` `ddd` `dd` `d` `hh` `h`
  `mm` / `m` as minutes (after an hour or before seconds) `ss` `s` `AM/PM` `A/P`, elapsed `[h]`
  `[m]` `[s]` and fractional seconds (`ss.00`). English month and day names.

Tested formats include `"0"`, `"0.00"`, `"#,##0"`, `"#,##0.00"`, `"$#,##0.00"`, `"0%"`,
`"yyyy-mm-dd"`, `"dd/mm/yyyy"`, `"mmm d, yyyy"`, `"dddd, mmmm d"`, `"h:mm AM/PM"`, `"[h]:mm"`,
`"0.00E+00"`, and fractions: `# ?/?` (whole part + fraction), `?/?` (improper), up to three-digit
denominators (`# ??/??`, `???/???`) or a fixed one (`# ?/8`). Like Excel, the fraction is the
last continued-fraction convergent whose denominator fits (0.14159 as `??/??` is `1/7`, not
`14/99`); the results were checked against Excel for Mac. Not supported: colors and locale tags
are ignored (`[Red]`, `[$-409]`) and conditions (`[>100]`) are not implemented.

## Limits

- Not supported (parse error → `#NAME?` and an entry in `report.errors`): external workbooks (`[1]Sheet1!A1`), 3-D references (`Sheet1:Sheet3!A1`), the
  union (`(A1,B1)`) and intersection (space) operators. R1C1 notation is not supported either
  (`R1C1` reads as an unknown name: `#NAME?`).
- Not implemented (→ `#NAME?`, listed in `unsupported`): every function not in the table, notably
  statistical, financial and engineering functions, `LAMBDA`, and the newer array functions
  (`VSTACK`, `TAKE`, `TEXTSPLIT`, ...). Register your own with `createEvaluator({ functions })`.
- A formula returning an array spills only when it is a dynamic-array formula (see
  [Dynamic arrays](#dynamic-arrays)); otherwise legacy implicit intersection applies, as in a
  file without dynamic-array metadata.
- Formulas are expected in the file format's syntax (en-US: `,` argument separator, `.` decimal
  point), which is what `readXlsx` returns and Excel writes. Function names come without the
  `_xlfn.` prefix files use for newer functions (the readers strip it, the writers add it back);
  the evaluator also accepts prefixed names.
- **en-US conventions.** Text-to-number conversion (`VALUE`, arithmetic, criteria) and `TEXT`
  use en-US rules regardless of the machine's locale: `.` decimal point, `,` thousands
  separator, `$` currency, US date order (`1/31/2024`), English month and day names. Excel
  applies the user's locale instead, so `VALUE("$1,000")` or `TEXT(x,"#,##0.00")` can give a
  different result (or `#VALUE!`) in Excel on a machine with other regional settings.
- Volatile functions are recomputed on every `evaluate` and on every `track(...).set` change.
- Implicit intersection uses the host cell's row / column regardless of the referenced sheet.

## Performance

The evaluator caches parsed formulas per text (shared by all workbooks evaluated with the same
evaluator), iterates ranges directly over the row storage without materializing them, keeps a
dense copy of a column once it is read by an aggregate or a lookup, and indexes columns used by
repeated exact-match lookups. Numbers from `node` 26 on an Apple M1 Pro:

| Workload | Time |
| --- | --- |
| 100k `=A{n}*2+B{n}` (parse, evaluate, write results) | ~0.2 s |
| same, parsed formulas already cached | ~0.09 s |
| 100k-long dependency chain `=A{n-1}+1` | ~0.14 s |
| 10k `VLOOKUP` (exact) + 10k `SUMIF` over 10k rows | ~1.5 s |

Incremental recalculation (`track(wb).set(...)`) on 100k formulas (50k rows of `=A{n}*2+1` and a
running total `=C{n-1}+B{n}`), `node` 26 on Apple silicon; a full `evaluate` of the same sheet
takes ~56 ms:

| Change | Recalculated | Time |
| --- | --- | --- |
| a cell nothing depends on | 0 | ~4 ms |
| the last input | 2 | ~5 ms |
| the middle input | 25,002 | ~18 ms |
| the first input | 50,001 | ~32 ms |

The fixed part is one pass over the formulas in dependency order, testing each formula's
references against the changed cells.
| 1000 x `SUM(A:A)/COUNT(A:A)` over 100k numbers | ~1.5 s |

Bundle: about 17.2 KB minified and brotli-compressed for the whole entry (`pnpm size`, limit
17.5 KB); the core bundles are unchanged.
