# Architecture & module contracts

Single published package, subpath exports for tree-shaking (`.`, `./csv`, `./json`).
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
Async writers (`xlsxStream`, `createXlsxWriter`) pass `{ native: true }` to `ZipStreamWriter`: entries
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
