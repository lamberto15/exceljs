# Architecture & module contracts

Single published package, subpath exports for tree-shaking (`.`, `./csv`, `./json`, `./crypto`).
Shared value/style types live in `src/core/types.ts` — import from there, never redefine.

Rules for every module:
- Web-standard APIs only in `src/` (Uint8Array, TextEncoder/TextDecoder, Web Streams, async iterators). No `node:*` imports, no `Buffer`.
- Only runtime dependency: `fflate`.
- No per-item allocation in hot loops where avoidable; prefer string chunk arrays, typed arrays, char-code scanning.
- `sideEffects: false` — no top-level side effects.
- TS strict + `isolatedDeclarations`: exported functions need explicit return types.
- Tests in `test/<module>.test.ts` (vitest). Old exceljs specs in `../spec/` are reference material for edge cases.

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

## src/zip — fflate wrapper with safety limits
```ts
export interface ZipEntryIn { path: string; data: Uint8Array | string; level?: number }
export function zipSync(entries: ZipEntryIn[], level: number): Uint8Array;
export class ZipStreamWriter {                       // streaming, emits Uint8Array chunks
  constructor(onData: (chunk: Uint8Array, final: boolean) => void, level: number);
  file(path: string, level?: number): { push(data: Uint8Array | string, final?: boolean): void };
  end(): void;
}
export function unzipSync(data: Uint8Array, limits?: Partial<Limits>): Map<string, Uint8Array>;
export function unzipStream(source: AsyncIterable<Uint8Array>, limits?: Partial<Limits>):
  AsyncIterable<{ path: string; chunks: AsyncIterable<Uint8Array> }>;  // optional/may be phase 2
export function unzipLazy(data: Uint8Array, limits?: Partial<Limits>): Map<string, LazyEntry>;
//   LazyEntry { path, size, compressedSize, compression, bytes(): Uint8Array,
//               chunks(chunkSize?): Iterable<Uint8Array> }  — central directory only; inflate on demand
export function deflateStream(level, onData): ByteStream;   // raw DEFLATE push streams, so no
export function inflateStream(onData): ByteStream;          // module outside src/zip imports fflate
export class ZipLimitError extends Error {}
```
Limit accounting lives in `src/zip/limits.ts` (`Budget`) and is shared by every unzip path.
Async writers (`xlsxStream`, `createXlsxWriter`) pass `{ compressor: nativeCompressor() }` to
`ZipStreamWriter` (kept out of the sync path so write-only bundles don't ship it): entries
are deflated by `CompressionStream('deflate-raw')` when the runtime has it (off the main thread in
Node), inside fflate's Zip container (data descriptors). Level 0 still stores; sync `writeXlsx` always
uses fflate. Async mode exposes `ready(maxBytes)` (backpressure) and `idle()` (all output emitted).
Enforce `Limits` (from types.ts): entry count, per-entry & total uncompressed bytes (checked while inflating, not just from headers), compression ratio; reject paths with `..`, absolute paths, backslashes, NUL.

## src/csv — RFC 4180 streaming read/write
```ts
export interface CsvReadOptions { delimiter?: string; quote?: string; bom?: boolean /*strip, default true*/; 
  parseNumbers?: boolean; parseDates?: boolean; maxFieldBytes?: number }
export class CsvParser { constructor(onRow: (row: string[]) => void, o?: CsvReadOptions); write(chunk: string): void; end(): void; }
export function parseCsv(text: string, o?: CsvReadOptions): string[][];
export function csvRows(source: AsyncIterable<string | Uint8Array> | ReadableStream<Uint8Array>, o?): AsyncIterable<string[]>;
export interface CsvWriteOptions { delimiter?: string; quote?: string; eol?: string; bom?: boolean;
  escapeFormulas?: boolean /* prefix ' to = + - @ \t \r, default false */; dateFormat?: (d: Date) => string }
export function formatCsvRow(row: readonly CellValue[], o?: CsvWriteOptions): string;
export function writeCsv(rows: Iterable<readonly CellValue[]>, o?): string;
export function csvStream(rows: AsyncIterable<readonly CellValue[]> | Iterable<...>, o?): ReadableStream<Uint8Array>;
```
Handles quoted fields with delimiters/newlines/escaped quotes, CRLF/LF/CR, chunks split anywhere.

## src/core — workbook model
- `Workbook` → `Worksheet[]`, shared `StyleRegistry`, optional `SharedStrings`.
- Cell storage is row-major (XLSX is row-major) and object-free per cell:
  each row keeps parallel growable typed arrays (`col: Uint16Array`, `kind: Uint8Array`, `num: Float64Array`, `style: Uint32Array`)
  plus a sparse `obj: unknown[]` slot only for strings/formulas/rich values. Rows live in a sparse array indexed by row number.
- `StyleRegistry.intern(style: Style): number` dedupes by canonical key → xf index; also exposes fonts/fills/borders/numFmts tables for writing.
- Addresses: 0-based `(row, col)` internally; helpers `encodeCol(0)="A"`, `decodeCol`, `encodeAddr`, `decodeAddr("B3")`, `decodeRange("A1:C3")`.
- Dates: `dateToSerial(d, date1904)`, `serialToDate(n, date1904)` (UTC based, Excel 1900 leap-year bug honored).
- Built-in number formats table (ids 0–49) and `isDateFormat(fmt)` detection.
- `StyleRegistry.internDxf(style)` → dxf index (`<dxfs>`, for conditional formatting / tables).
- `Workbook.definedNames: DefinedName[]` (sheet scope = index into `worksheets`; `_xlnm._FilterDatabase`
  is managed via `worksheet.autoFilter`). Reader/feature helpers: `ws._ensureRow(r0, capacity)`,
  `RowProps.styleIndex`, `ws.addMergeRaw(range)` (unvalidated); `mergeCells` overlap checks are indexed.
- Feature-owned optional fields (`ws.pageSetup`, `ws.protection`, ...) are declared type-only in
  `src/core/features.ts` and merged into `Worksheet` / `Workbook`.

## src/xlsx — read / write / features
- `write/`: `writeXlsx` (sync, fflate), `xlsxStream` and `createXlsxWriter` (async, native deflate when
  available). Sheet XML = head emitters + `encodeRow` (shared hot loop) + tail emitters. Package order:
  each sheet part, then its feature parts and rels; workbook-level parts (workbook.xml, styles, sst,
  rels, `[Content_Types].xml`, docProps) last, from `bookParts()` shared by all writers.
- `read/`: `readXlsx` (DOM; lazy unzip, one part inflated at a time) and `xlsxRows` (streaming rows).
  `readXlsx` throws `EncryptedFileError` (a `XlsxParseError`) for OLE/CFB input, pointing at `./crypto`.
- `features/`: extension points (`types.ts`) and the static registry (`registry.ts`: `WRITERS`,
  `READERS`). Every worksheet element is produced by a `SheetEmitter` bound to a CT_Worksheet slot
  (`SheetSlot`), built-ins included, so element order is global. Features also add package parts,
  relationships and content types, and consume worksheet subtrees when reading (never inside
  `sheetData`). How to add a feature: `src/xlsx/features/README.md`.
- Round-trip preservation (`ReadOptions.preserveUnknown`, default true): `read/preserve.ts`
  captures parts / XML nobody models into `wb.preserved` / `ws.preserved` (model:
  `common/preserved.ts`: lazy raw parts + their rels, fragments with rId placeholders);
  `write/preserve.ts` (small: it ships in write-only bundles) names and re-emits them. See
  "Round-trip preservation" in the features README.

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
