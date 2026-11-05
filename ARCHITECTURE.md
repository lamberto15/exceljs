# Architecture & module contracts

Single published package, subpath exports for tree-shaking (`.`, `./lite`, `./csv`, `./json`, `./crypto`,
`./formula`).
`.` is batteries-included (every XLSX feature on); `./lite` is the same API with features opt-in
(`options.features`, see "src/xlsx" below and `docs/bundle-size.md`). The build is unbundled
(tsdown `unbundle`: one dist file per source module) so consumer bundlers can drop whole unused
modules (`sideEffects: false`); a single shared chunk would keep every feature object alive.
Shared value/style types live in `src/core/types.ts` — import from there, never redefine.

Rules for every module:
- Web-standard APIs only in `src/` (Uint8Array, TextEncoder/TextDecoder, Web Streams, async iterators). No `node:*` imports, no `Buffer`.
  The one exception is `src/zip/node.ts`: it looks up `node:zlib` at runtime with `process.getBuiltinModule`
  (Node >= 22.3, Bun, Deno 2), and every caller falls back to Web APIs / fflate when that fails.
- Only runtime dependency: `fflate`.
- No per-item allocation in hot loops where avoidable; prefer string chunk arrays, typed arrays, char-code scanning.
- `sideEffects: false` — no top-level side effects.
- TS strict + `isolatedDeclarations`: exported functions need explicit return types.
- Tests in `test/<module>.test.ts` (vitest). ExcelJS's specs are ported in `test/exceljs-parity/` (each test cites
  the original) and its sample files live in `test/fixtures/exceljs/`.

## src/xml — XML writer + streaming tokenizer
```ts
export function escapeText(s: string): string;      // & < >  (+ strip invalid XML 1.0 chars)
export function escapeAttr(s: string): string;      // & < > " \n \r \t
export class XmlWriter {                              // appends to string[] chunks
  constructor(onFlush?: (chunk: string) => void, flushAt?: number);
  raw(s: string): this;
  open(tag: string, attrs?: Record<string, string | number | boolean | undefined>): this; // '<tag a="..">'
  leaf(tag: string, attrs?: ..., text?: string): this; // self-closing or with text
  close(tag: string): this;
  text(s: string): this;
  toString(): string;     // remaining buffered content
  flush(): void;          // emits buffered content via onFlush
}
// Pull/push tokenizer over string chunks (chunks may split anywhere, incl. inside tags/entities/UTF-16 pairs).
export interface XmlHandler {
  open?(name: string, attrs: XmlAttrs, selfClosing: boolean): void;
  close?(name: string): void;
  text?(text: string): void;  // entities decoded; CDATA delivered as text
}
export type XmlAttrs = Record<string, string>; // null-prototype object, entities decoded
export class XmlTokenizer { constructor(h: XmlHandler); write(chunk: string): void; end(): void; }
export function parseXml(xml: string, h: XmlHandler): void;
```
Security: never expand DTD/custom entities (ignore `<!DOCTYPE ...>` entirely); only the 5 predefined + numeric char refs. Names keep their prefix (`x:c`) — consumers strip if needed via exported `localName(name)`.

## src/zip — fflate wrapper, native zlib, safety limits
```ts
export interface ZipEntryIn { path: string; data: Uint8Array | string; level?: number }
export function zipSync(entries: ZipEntryIn[], level: number): Uint8Array;
export class ZipStreamWriter {                       // streaming, emits Uint8Array chunks
  constructor(onData: (chunk: Uint8Array, final: boolean) => void, level: number,
              opts?: { compressor?: AsyncCompressor });
  file(path: string, level?: number): { push(data: Uint8Array | string, final?: boolean): void };
  ready(maxBytes: number): Promise<void> | undefined; idle(): Promise<void>; end(): void;
}
export function unzipSync(data: Uint8Array, limits?: Partial<Limits>): Map<string, Uint8Array>;
export function unzipStream(source: AsyncIterable<Uint8Array>, limits?: Partial<Limits>):
  AsyncGenerator<{ path: string; chunks: AsyncIterable<Uint8Array>; unbounded(): void }>;
export function unzipLazy(data: Uint8Array, limits?: Partial<Limits>, opts?: { deferSizeCheck? }):
  Map<string, LazyEntry>;
//   LazyEntry { path, size, compressedSize, compression, bytes(): Uint8Array,
//               chunks(chunkSize?, unbounded?): Iterable<Uint8Array> }  — central directory only; inflate on demand
export function unzipBlob(blob: Blob, limits?, opts?): Promise<Map<string, BlobEntry>>;
//   BlobEntry: same fields, chunks(unbounded?): AsyncIterable — reads blob.slice() ranges on demand
export function deflateStream(level, onData): ByteStream;   // raw DEFLATE push streams, so no
export function inflateStream(onData): ByteStream;          // module outside src/zip imports fflate
export function nativeCompressor(): AsyncCompressor | undefined; // + nativeDeflateAvailable()
export class ZipLimitError extends Error {}  // ZipPathError extends it (unsafe / duplicate paths)
```
Limit accounting lives in `src/zip/limits.ts` (`Budget`) and is shared by every unzip path.
`unbounded()` (unzipStream) / `unbounded = true` (lazy and Blob `chunks`) is for callers that keep
nothing: only `maxCompressionRatio` applies then (`xlsxRows` uses it for sheets it parses as they inflate).

Native compression (`src/zip/node.ts`, `src/zip/native.ts`; fflate everywhere else):
- Sync zip writing (`writeXlsx`): entries are deflated with `zlib.deflateRawSync` in ~1 MiB sync-flushed
  batches, else fflate. CRC-32 uses `zlib.crc32` when present, in every writer.
- Async writers (`xlsxStream`, `createXlsxWriter`) pass `{ compressor: nativeCompressor() }` to
  `ZipStreamWriter` (kept out of the sync path so write-only bundles don't ship it): entries are deflated
  off the main thread by `zlib.createDeflateRaw({ level })`, else by `CompressionStream('deflate-raw')`
  (no level knob), inside fflate's Zip container (data descriptors). Level 0 still stores. Async mode
  exposes `ready(maxBytes)` (backpressure) and `idle()` (all output emitted).
- `unzipLazy` (`readXlsx`, `xlsxRows` on bytes) inflates entries up to 64 MiB with `zlib.inflateRawSync`,
  limits enforced through `maxOutputLength`; larger entries and other runtimes stream through fflate.
  `unzipBlob` inflates with `DecompressionStream('deflate-raw')` when present, else fflate.
Enforce `Limits` (from types.ts): entry count, per-entry & total uncompressed bytes (checked while inflating, not just from headers), compression ratio; reject paths with `..`, absolute paths, backslashes, NUL.

## src/csv — RFC 4180 streaming read/write
```ts
export interface CsvReadOptions { delimiter?: string; quote?: string; bom?: boolean /*strip, default true*/;
  parseNumbers?, parseDates?, parseBooleans?, parseErrors?: boolean; dateFormats?: readonly string[];
  localDates?: boolean; maxFieldBytes?: number; skipEmptyLines?: boolean /* default true */ }
// any typing option turns rows into CsvValue[] (string | number | boolean | Date | CellError)
export class CsvParser<R> { constructor(onRow: (row: R) => void, o?: CsvReadOptions); write(chunk: string): void; end(): void; }
export function parseCsv(text: string, o?: CsvReadOptions): string[][] | CsvValue[][];
export function csvRows(source: AsyncIterable<string | Uint8Array> | Iterable<...> | ReadableStream<Uint8Array>,
  o?: CsvReadOptions & { encoding?: string /* TextDecoder label or 'auto'; BOM wins */ }): AsyncIterable<string[] | CsvValue[]>;
export interface CsvWriteOptions { delimiter?: string; quote?: string; eol?: string; bom?: boolean;
  escapeFormulas?: boolean /* prefix ' to = + - @ \t \r, default false */; dateFormat?: (d: Date) => string }
export function formatCsvRow(row: readonly CellValue[], o?: CsvWriteOptions): string;
export function writeCsv(rows: Iterable<readonly CellValue[]>, o?): string;
export function csvStream(rows: AsyncIterable<readonly CellValue[]> | Iterable<...>, o?): ReadableStream<Uint8Array>;
```
Handles quoted fields with delimiters/newlines/escaped quotes, CRLF/LF/CR, chunks split anywhere
(byte chunks are decoded chunk-safely). Errors are `CsvError`.

## src/core — workbook model
- `Workbook` → `Worksheet[]`, shared `StyleRegistry`. `SharedStrings` (`core/sharedstrings.ts`) is built by the
  writers (strings `'shared'`) and the reader, not kept on the model.
- Cell storage is row-major (XLSX is row-major) and object-free per cell:
  each row keeps parallel growable typed arrays (`col: Uint16Array`, `kind: Uint8Array`, `num: Float64Array`, `style: Uint32Array`)
  plus a sparse `obj: unknown[]` slot only for strings/formulas/rich values. Rows live in a sparse array indexed by row number.
- `StyleRegistry.intern(style: Style): number` dedupes by canonical key → xf index; also exposes fonts/fills/borders/numFmts tables for writing.
- Addresses: 0-based `(row, col)` internally; helpers `encodeCol(0)="A"`, `decodeCol`, `encodeAddr`, `decodeAddr("B3")`, `decodeRange("A1:C3")`.
- Dates: `dateToSerial(d, date1904, local?)`, `serialToDate(n, date1904, local?)` (UTC based unless `local`, Excel 1900 leap-year bug honored).
- Built-in number formats table (ids 0–49) and `isDateFormat(fmt)` detection.
- `StyleRegistry.internDxf(style)` → dxf index (`<dxfs>`, for conditional formatting / tables).
- `Workbook.definedNames: DefinedName[]` (sheet scope = index into `worksheets`; `_xlnm._FilterDatabase`
  is managed via `worksheet.autoFilter`). Reader/feature helpers: `ws._ensureRow(r0, capacity)`,
  `RowProps.styleIndex`, `ws.addMergeRaw(range)` (unvalidated); `mergeCells` overlap checks are indexed.
- Feature-owned optional fields (`ws.pageSetup`, `ws.protection`, ...) are declared type-only in
  `src/core/features.ts` and merged into `Worksheet` / `Workbook`.

## src/xlsx — read / write / features
- `write/`: `writeXlsx` (sync; `zlib.deflateRawSync` batches in Node / Bun / Deno, else fflate),
  `xlsxStream` and `createXlsxWriter` (async, off-thread native deflate when available). Default profile
  `balanced` (shared strings), except `createXlsxWriter`: `low-memory` (inline strings, constant memory).
  Sheet XML = head emitters + `encodeRow` (shared hot loop) + tail emitters. Package order: each sheet
  part, then its feature parts and rels; workbook-level parts come from `bookParts()`, shared by all
  writers. `writeXlsx` / `xlsxStream` reserve the metadata entries (`[Content_Types].xml`, `_rels/.rels`,
  workbook.xml + rels, styles, sst) at the start of the archive (`ZipOut.reserve`; the zip holds later
  output back until they are filled) so streaming readers get them before the sheets;
  `createXlsxWriter` writes every workbook-level part at `end()`, after the sheets.
  The implementations (`book.ts`, `stream-writer.ts`) are feature-agnostic: they run the
  `FeatureWriter`s in `options.features` (resolved to `ResolvedWriteOptions.writers`; one emitter
  plan per write, `sheetPlan`). `write/index.ts` (the root API) wraps them with every built-in.
- `read/`: `readXlsx` (DOM; lazy unzip, one part inflated at a time) and `xlsxRows` (streaming rows).
  `xlsxRows` reads bytes and `Blob`s by the central directory (`unzipLazy` / `unzipBlob`), metadata
  parts first, then each sheet inflated as it is parsed; streams go through `unzipStream` in archive
  order, buffering a sheet that arrives before the metadata. Both enforce `maxRows` / `maxCols`
  (`SheetLimitError`).
  `readXlsx` throws `EncryptedFileError` (a `XlsxParseError`) for OLE/CFB input, pointing at `./crypto`.
  Same split: `read/reader.ts` runs `options.features` readers, `read/index.ts` adds all built-ins.
- `features/`: extension points (`types.ts`), one module per feature exporting its `<name>Writer` /
  `<name>Reader` halves (public, for `./lite` opt-in; the default theme is `themeWriter`), and the
  root-only list of built-ins (`registry.ts`: `WRITERS`, `READERS`, never imported by `./lite`).
  Every worksheet element is produced by a `SheetEmitter` bound to a CT_Worksheet slot
  (`SheetSlot`), built-ins included, so element order is global. Features also add package parts,
  relationships and content types, and consume worksheet subtrees when reading (never inside
  `sheetData`). Writing a feature-owned model field while that feature is disabled throws
  (`checkFeatures` in `write/sheet.ts`: a property check per sheet). How to add a feature:
  `src/xlsx/features/README.md`.
- Round-trip preservation (`ReadOptions.preserveUnknown`, default true): `read/preserve.ts`
  captures parts / XML nobody models into `wb.preserved` / `ws.preserved` (model:
  `common/preserved.ts`: lazy raw parts + their rels, fragments with rId placeholders);
  `write/preserve.ts` names and re-emits them. The writer core reaches it only through the
  `w` hooks readXlsx attaches to the kept data, so write-only bundles never include it. Features
  disabled at read time (`./lite`) are not claimed, so their content falls through to
  preservation and round-trips raw. See "Round-trip preservation" in the features README.

## src/formula — formula parser and evaluator (`./formula`)
```ts
export function parseFormula(text: string): FormulaNode;      // throws FormulaParseError
export function evaluate(wb, opts?): EvaluationReport;        // also evaluateCell / evaluateFormula
export function createEvaluator({ functions?, now? }): Evaluator;
```
- `parser.ts`: hand-written Pratt parser (Excel precedence: `:` > unary > `%` > `^` > `* /` > `+ -` >
  `&` > comparisons; left-associative). Nodes are plain objects with 0-based coordinates.
- `evaluator.ts`: `Evaluator` (function table + parse cache keyed by formula text) and a per-call
  `Run`: indexes the formula cells of every sheet (node ids are contiguous per sheet and sorted by
  row, then column, so a cell's id is two binary searches away), collects static references per
  parsed formula, and runs an iterative Tarjan SCC over the nodes reachable from the roots —
  each SCC is evaluated as soon as it is popped, which is dependencies-first. Singleton SCCs are
  computed; cycles get `#REF!` (or iterate with `maxIterations`). A reference reached only
  dynamically is evaluated on demand (re-entry = circular). Results are stored at the end
  (`{ ...formula, result }` in the slot; array formulas also fill their area via `setCellRaw`).
- Values (`values.ts`): scalars (`null` = blank, `CellError` objects), `FormulaRange` (lazy
  reference, reads through the run) and `FormulaArray`. `EvalContext.operand` dereferences
  (implicit intersection unless in array mode); function arguments stay references. Built-ins
  throw `CellError`s; the call boundary turns them into results.
- Ranges: `Run.each` walks `rowData` directly (formula cells advance a node-id cursor in step with
  the slots); single-column ranges of 16+ rows go through a per-run dense column cache, which
  lookups also use (plus an exact-match index built on the second lookup into the same lane).
  The cache is bypassed while iterating cycles.
- Imports only `core/dates.ts` from the core (and core types), never the Worksheet code
  (`CellKind` values are inlined), so it adds nothing to the core bundles.

## src/crypto — password-protected files (`./crypto`, MS-OFFCRYPTO)
```ts
export function isEncrypted(data: Uint8Array): boolean;  // CFB signature + EncryptionInfo stream
export function decryptXlsx(data: Uint8Array, password: string,
  opts?: { verifyIntegrity?: boolean /* default true */; limits?: Partial<Limits>; maxSpinCount?: number }
): Promise<Uint8Array>;                                   // inner zip bytes
export function encryptXlsx(zipBytes: Uint8Array, password: string,
  opts?: { cipher?: 'AES-128' | 'AES-192' | 'AES-256'; hash?: 'SHA1' | 'SHA256' | 'SHA384' | 'SHA512'; spinCount?: number }
): Promise<Uint8Array>;                                   // CFB container
export class WrongPasswordError, IntegrityError, UnsupportedEncryptionError, CfbError;

const wb = readXlsx(await decryptXlsx(bytes, pw));
const out = await encryptXlsx(writeXlsx(wb), pw);
```
- `cfb.ts`: minimal MS-CFB. Reads v3/v4; every sector id range-checked, chains walked at most
  once per sector (cycle detection), directory tree walk tracks visited entries, stream sizes
  capped by `Limits.maxEntryBytes`; errors are `CfbError`. Writes v3 (mini stream below 4096
  bytes, DIFAT when the FAT exceeds 109 sectors, balanced all-black directory trees).
- `agile.ts`: agile encryption 4.4 (read + write). Write defaults match Excel 2010+: AES-256-CBC,
  SHA-512, spinCount 100000, HMAC `dataIntegrity`, Excel's exact XML (namespaces, `\r\n` after
  the declaration). The `\x06DataSpaces` storage (Version, DataSpaceMap, StrongEncryptionDataSpace,
  StrongEncryptionTransform/\x06Primary) is written too: it is constant (~500 bytes) and Office
  always writes it, so we don't depend on Excel tolerating its absence.
- `standard.ts`: standard encryption 2.2/3.2/4.2 (Excel 2007; AES-ECB + SHA-1), read only.
  RC4 CryptoAPI and extensible encryption throw `UnsupportedEncryptionError`.
- Hashing: the password spin loop (100k tiny digests) uses synchronous hashes (`core/sha512.ts`,
  shared with sheet protection, plus SHA-1/256 in `crypto/hash.ts`): ~120 ms per derivation vs
  ~900 ms with one awaited `crypto.subtle.digest` per iteration. AES and the package HMAC use
  WebCrypto. WebCrypto AES-CBC always applies PKCS#7 while Office segments are unpadded: encrypt
  drops the trailing all-padding block; decrypt appends a crafted block `AES(C_last ^ 0x10*16)`
  that decrypts to a full padding block. The package is decrypted with one call, then the first
  block of every 4096-byte segment is corrected for its own IV. AES-ECB (standard) = CBC with a
  zero IV, then XOR each block with the previous ciphertext block.
