// Write option resolution: profile → defaults, explicit options win.

import type { Profile, WriteOptions } from '../../core/types.js';
import type { FeatureWriter } from '../features/types.js';

export interface ResolvedWriteOptions {
  readonly profile: Profile;
  /** Deflate level 0-9 (0 = store). */
  readonly compression: number;
  readonly strings: 'shared' | 'inline';
  /** Only set when given explicitly (writeXlsx uses the workbook's own date system). */
  readonly date1904: boolean | undefined;
  readonly theme: boolean;
  /** Feature writers to run (`features`, minus the theme when `theme: false`). */
  readonly writers: readonly FeatureWriter[];
}

const PROFILES: Record<Profile, { compression: number; strings: 'shared' | 'inline' }> = {
  'low-memory': { compression: 6, strings: 'inline' },
  balanced: { compression: 6, strings: 'shared' },
  'max-speed': { compression: 1, strings: 'inline' },
};

/**
 * Resolve write options. Profiles: low-memory {compression 6, inline strings},
 * balanced (default) {compression 6, shared strings}, max-speed {compression 1, inline strings}.
 */
export function resolveWriteOptions(opts: WriteOptions = {}): ResolvedWriteOptions {
  const profile = opts.profile ?? 'balanced';
  const base = PROFILES[profile];
  if (!base) throw new TypeError(`Unknown write profile: ${String(profile)}`);
  let compression = opts.compression ?? base.compression;
  if (!(compression >= 0 && compression <= 9)) {
    throw new RangeError(`compression must be 0-9, got ${compression}`);
  }
  compression |= 0;
  const strings = opts.strings ?? base.strings;
  if (strings !== 'shared' && strings !== 'inline') {
    throw new TypeError(`strings must be 'shared' or 'inline', got ${String(strings)}`);
  }
  const theme = opts.theme !== false;
  const f = opts.features ?? [];
  const writers = theme ? f : f.filter((w) => w.name !== 'theme');
  return { profile, compression, strings, date1904: opts.date1904, theme, writers };
}
