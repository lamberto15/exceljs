export type { LazyEntry } from './lazy.js';
export { unzipLazy } from './lazy.js';
export { assertSafePath, ZipLimitError, ZipPathError } from './limits.js';
export { nativeDeflateAvailable } from './native.js';
export { type ByteSink, type ByteStream, deflateStream, inflateStream } from './stream.js';
export type { ZipEntryIn, ZipEntryOut, ZipStreamOptions } from './zip.js';
export { unzipStream, unzipSync, ZipStreamWriter, zipSync } from './zip.js';
