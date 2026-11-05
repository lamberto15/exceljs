// Safety limits shared by every unzip path (unzipSync, unzipStream, unzipLazy): entry count,
// unsafe entry names, per-entry / total decompressed bytes and compression ratio — always
// measured on ACTUAL decompressed output, never trusted from headers alone.

import type { Limits } from '../core/types.js';
import { DEFAULT_LIMITS } from '../core/types.js';

/** Thrown when a configured limit (entries, bytes, ratio) is exceeded or the archive is malformed. */
export class ZipLimitError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ZipLimitError';
  }
}

/** Thrown for unsafe entry names (path traversal, absolute paths, ...). */
export class ZipPathError extends ZipLimitError {
  constructor(message: string) {
    super(message);
    this.name = 'ZipPathError';
  }
}

/** Ratio checks only start once an entry has produced this many bytes. */
const RATIO_MIN_BYTES = 1024 * 1024;

export function assertSafePath(path: string): void {
  if (
    path.length === 0 ||
    path.includes('\0') ||
    path.includes('\\') ||
    path.startsWith('/') ||
    /^[A-Za-z]:/.test(path) ||
    path.split('/').includes('..')
  ) {
    throw new ZipPathError(`Unsafe zip entry path: ${JSON.stringify(path)}`);
  }
}

/** Limit accounting for one archive. */
export class Budget {
  readonly limits: Limits;
  private entries = 0;
  /** Decompressed bytes produced so far (all entries). */
  private total = 0;

  constructor(limits?: Partial<Limits>) {
    this.limits = { ...DEFAULT_LIMITS, ...limits };
  }

  /**
   * Validate an entry header: count, name, declared size. Returns false for directory entries
   * (nothing to extract).
   */
  enter(name: string, declared: number | undefined): boolean {
    const L = this.limits;
    if (++this.entries > L.maxEntries) {
      throw new ZipLimitError(`Too many zip entries (max ${L.maxEntries})`);
    }
    assertSafePath(name);
    if (name.endsWith('/')) return false;
    if (declared !== undefined && declared > L.maxEntryBytes) {
      throw new ZipLimitError(`Zip entry too large: ${name}`);
    }
    return true;
  }

  /**
   * Account output of entry `name`: it has produced `out` bytes in total from `compressed` input
   * bytes, `fresh` of which were not accounted before. `ratio` = false for stored entries.
   */
  output(
    name: string,
    out: number,
    fresh: number,
    compressed: number,
    ratio: boolean,
    capped = true,
  ): void {
    const L = this.limits;
    if (capped) {
      this.total += fresh;
      if (out > L.maxEntryBytes) {
        throw new ZipLimitError(`Zip entry too large: ${name} (max ${L.maxEntryBytes} bytes)`);
      }
      if (this.total > L.maxTotalBytes) {
        throw new ZipLimitError(`Zip total uncompressed size exceeds ${L.maxTotalBytes} bytes`);
      }
    }
    if (ratio && out > RATIO_MIN_BYTES && out / Math.max(compressed, 1) > L.maxCompressionRatio) {
      throw new ZipLimitError(
        `Zip entry ${name} exceeds max compression ratio ${L.maxCompressionRatio}`,
      );
    }
  }
}
