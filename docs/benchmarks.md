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

Node v26.13.0, Apple M1 Pro, macOS (Darwin 25.6.0, arm64), 2026-11-05 (after native zlib, #71 / #75).

## Write

| case                                                 | time ms | size MB | peak heap MB | peak RSS MB |
| ---------------------------------------------------- | ------: | ------: | -----------: | ----------: |
| turbosheet `writeXlsx` (balanced)                    |    1222 |    8.38 |          150 |         280 |
| turbosheet `writeXlsx` (low-memory)                  |    1061 |    8.49 |          155 |          22 |
| turbosheet `writeXlsx` (max-speed)                   |     557 |   11.16 |          150 |          19 |
| turbosheet `createXlsxWriter` (balanced)             |    1177 |    8.37 |           71 |          54 |
| turbosheet `createXlsxWriter` (low-memory)           |    1060 |    8.46 |           65 |          18 |
| turbosheet `createXlsxWriter` (max-speed)            |     535 |   11.15 |           64 |           1 |
| turbosheet `createXlsxWriter` + `flush` (balanced)   |    1184 |    8.37 |           47 |           9 |
| turbosheet `createXlsxWriter` + `flush` (low-memory) |     975 |    8.46 |           35 |           0 |
| turbosheet `createXlsxWriter` + `flush` (max-speed)  |     528 |   11.15 |           51 |           0 |
| turbosheet `xlsxStream` (balanced)                   |    1023 |    8.37 |          161 |          13 |
| turbosheet `xlsxStream` (low-memory)                 |     858 |    8.46 |          147 |         132 |
| turbosheet `xlsxStream` (max-speed)                  |     438 |   11.15 |          164 |          19 |
| ExcelJS `xlsx.writeBuffer`                           |    6598 |    8.51 |         2113 |        2093 |
| ExcelJS `WorkbookWriter` (shared strings)            |    3446 |   11.39 |           64 |          28 |
| ExcelJS `WorkbookWriter` (inline strings)            |    3284 |   10.89 |           42 |           3 |

`xlsxStream` builds the whole workbook in memory and then streams the compressed output, so its heap is the
in-memory workbook's. `createXlsxWriter` with `flush` is the low-memory streaming path: rows are written as
they are added. In Node, deflate runs on native zlib (sync batches in `writeXlsx`, off-thread in the
streaming writers), so `max-speed` (level 1) is now about twice as fast as the default level.

## Read

| case                                   | time ms | peak heap + buffers MB |
| -------------------------------------- | ------: | ---------------------: |
| turbosheet `readXlsx` (DOM)            |    1213 |                    142 |
| turbosheet `xlsxRows`                  |    1195 |                     85 |
| turbosheet `xlsxRows` (low-memory)     |    2220 |                     28 |
| turbosheet `xlsxRows` (metadata first) |    1189 |                      4 |
| ExcelJS `xlsx.load` (DOM)              |    3628 |                    721 |
| ExcelJS `WorkbookReader` (streaming)   |    1785 |                     72 |

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
