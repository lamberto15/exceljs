# Benchmarks

turbosheet versus `exceljs@4.4.0` on a 100,000 row x 20 column sheet. These numbers are indicative: they
come from one machine and one dataset, so run `pnpm bench` on your own hardware and data before drawing
conclusions.

## Methodology

- **Data.** Write benchmark: 100,000 rows x 20 columns, 8 numbers, 6 short strings, 3 dates, 3 styled numbers,
  and every 10th row bold. Read benchmark: a 7.9 MB file of the same size (numbers, strings, booleans) written
  by ExcelJS.
- **Timing.** Wall-clock `performance.now()` around one run of each case, after a forced GC. Each case is a
  single sample, so expect noise (see Variance).
- **Memory.** A 5 ms sampler records the peak above the pre-case baseline. The write benchmark reports JS
  heap (`heapUsed`) and RSS; the read benchmark reports heap plus `ArrayBuffer`s, because decompressed bytes
  live outside the V8 heap. Peak RSS rarely shrinks, so later cases under-report it; run a case alone with
  `-t` to isolate it.
- **Fairness.** Both libraries write the same cell values and styles. ExcelJS streaming runs with
  `useStyles: true`, once with shared strings and once with inline strings. turbosheet is shown under each of
  its three profiles (`balanced`, `low-memory`, `max-speed`); file size differs between profiles
  (`max-speed` compresses less).
- **Sources.** [`test/xlsx-write.bench.test.ts`](../test/xlsx-write.bench.test.ts) and the gated `perf` test in
  [`test/xlsx-read.test.ts`](../test/xlsx-read.test.ts). [`bench/run.mjs`](../bench/run.mjs) runs both and
  prints the tables below.

## Machine

Node v26.13.0, Apple M1 Pro, macOS (Darwin 25.6.0, arm64), 2026-11-04.

## Write

| case                                                 | time ms | size MB | peak heap MB | peak RSS MB |
| ---------------------------------------------------- | ------: | ------: | -----------: | ----------: |
| turbosheet `writeXlsx` (balanced)                    |    2134 |    8.77 |          147 |         292 |
| turbosheet `writeXlsx` (low-memory)                  |    2092 |    8.77 |          145 |          21 |
| turbosheet `writeXlsx` (max-speed)                   |    1357 |   11.54 |          147 |          24 |
| turbosheet `createXlsxWriter` (balanced)             |    1390 |    8.37 |           77 |          64 |
| turbosheet `createXlsxWriter` (low-memory)           |    1211 |    8.46 |           47 |          14 |
| turbosheet `createXlsxWriter` (max-speed)            |    1226 |    8.46 |           46 |          10 |
| turbosheet `createXlsxWriter` + `flush` (balanced)   |    1354 |    8.37 |           46 |           4 |
| turbosheet `createXlsxWriter` + `flush` (low-memory) |    1230 |    8.46 |           31 |           1 |
| turbosheet `createXlsxWriter` + `flush` (max-speed)  |    1196 |    8.46 |           31 |           1 |
| turbosheet `xlsxStream` (balanced)                   |    1057 |    8.37 |          159 |         131 |
| turbosheet `xlsxStream` (low-memory)                 |     899 |    8.46 |          147 |           5 |
| turbosheet `xlsxStream` (max-speed)                  |     918 |    8.46 |          146 |           5 |
| ExcelJS `xlsx.writeBuffer`                           |    6685 |    8.51 |         2113 |        2091 |
| ExcelJS `WorkbookWriter` (shared strings)            |    3485 |   11.39 |           73 |           0 |
| ExcelJS `WorkbookWriter` (inline strings)            |    3393 |   10.89 |           39 |           0 |

`xlsxStream` builds the whole workbook in memory and then streams the compressed output, so its heap is the
in-memory workbook's. `createXlsxWriter` with `flush` is the low-memory streaming path: rows are written as
they are added.

## Read

| case                                   | time ms | peak heap + buffers MB |
| -------------------------------------- | ------: | ---------------------: |
| turbosheet `readXlsx` (DOM)            |    1213 |                    139 |
| turbosheet `xlsxRows`                  |    1170 |                     83 |
| turbosheet `xlsxRows` (low-memory)     |    2189 |                     27 |
| turbosheet `xlsxRows` (metadata first) |    1186 |                      7 |
| ExcelJS `xlsx.load` (DOM)              |    3441 |                   1188 |
| ExcelJS `WorkbookReader` (streaming)   |    1793 |                     70 |

"Metadata first" is the same data with workbook metadata (`styles.xml`, `sharedStrings.xml`) stored before
the sheet, as `writeXlsx` writes it, so `xlsxRows` needs no buffering. Excel 16.x and ExcelJS store the sheet
first. The plain `xlsxRows` row reads the ExcelJS-written order, which buffers some
data.

## Variance

Over the runs made while preparing this page, ExcelJS varied more than turbosheet: `xlsx.load` took 3.4-4.7 s
and `WorkbookReader` 1.8-3.7 s, while turbosheet's read and write times moved by about 10% or less. The README
quotes ranges for those cases for this reason.

## Reproduce

```sh
pnpm bench
```

Or run the pieces individually (add `BENCH_ROWS=20000` to the write benchmark for a quicker pass):

```sh
RUN_BENCH=1 NODE_OPTIONS=--expose-gc npx vitest run test/xlsx-write.bench.test.ts --silent=false
RUN_BENCH=1 npx vitest run test/xlsx-read.test.ts -t perf --silent=false --pool=forks --execArgv=--expose-gc
```
