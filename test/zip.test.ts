// @ts-ignore -- @types/node is not installed; tests only
import { readdirSync, readFileSync } from 'node:fs';
import {
  unzipSync as rawUnzipSync,
  zipSync as rawZipSync,
  strFromU8,
  strToU8,
  Zip,
  ZipDeflate,
} from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  deflateStream,
  inflateStream,
  unzipLazy,
  unzipStream,
  unzipSync,
  ZipLimitError,
  ZipPathError,
  ZipStreamWriter,
  zipSync,
} from '../src/zip/index.js';
import { Inflight, NativeDeflateFile, nativeCompressor } from '../src/zip/native.js';
import { nodeZlib } from '../src/zip/node.js';

async function collect(
  src: AsyncIterable<Uint8Array>,
  limits?: Parameters<typeof unzipStream>[1],
): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const e of unzipStream(src, limits)) {
    const parts: Uint8Array[] = [];
    for await (const c of e.chunks) parts.push(c);
    const total = parts.reduce((n, p) => n + p.length, 0);
    const buf = new Uint8Array(total);
    let o = 0;
    for (const p of parts) buf.set(p, o), (o += p.length);
    out.set(e.path, buf);
  }
  return out;
}

async function* chunked(data: Uint8Array, seed = 1): AsyncGenerator<Uint8Array> {
  let s = seed;
  const rnd = (): number => (s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  for (let i = 0; i < data.length; ) {
    const n = 1 + Math.floor(rnd() * 37);
    yield data.slice(i, i + n);
    i += n;
  }
}

async function* whole(data: Uint8Array): AsyncGenerator<Uint8Array> {
  yield data;
}

const sample = [
  { path: 'a.txt', data: 'hello' },
  { path: 'dir/b.xml', data: '<x>é ü 日本</x>'.repeat(100) },
  { path: 'empty.txt', data: '' },
  { path: 'bin.dat', data: Uint8Array.from({ length: 5000 }, (_, i) => (i * 7) & 255) },
];

describe('zipSync / unzipSync', () => {
  it('round-trips', () => {
    const m = unzipSync(zipSync(sample, 6));
    expect([...m.keys()]).toEqual(sample.map((s) => s.path));
    expect(strFromU8(m.get('a.txt')!)).toBe('hello');
    expect(strFromU8(m.get('dir/b.xml')!)).toBe('<x>é ü 日本</x>'.repeat(100));
    expect(m.get('empty.txt')!.length).toBe(0);
    expect(m.get('bin.dat')).toEqual(sample[3]!.data);
  });

  it('level 0 stores, level 6 compresses', () => {
    const e = [{ path: 'z.txt', data: 'abc'.repeat(10000) }];
    const stored = zipSync(e, 0);
    const deflated = zipSync(e, 6);
    expect(stored.length).toBeGreaterThan(30000);
    expect(deflated.length).toBeLessThan(1000);
    expect(strFromU8(unzipSync(stored).get('z.txt')!)).toBe('abc'.repeat(10000));
    expect(strFromU8(unzipSync(deflated).get('z.txt')!)).toBe('abc'.repeat(10000));
  });
});

describe('ZipStreamWriter', () => {
  it('writes many files incrementally and round-trips', async () => {
    const chunks: Uint8Array[] = [];
    let finals = 0;
    const w = new ZipStreamWriter((c, f) => (chunks.push(c), f && finals++), 6);
    for (let i = 0; i < 200; i++) {
      const f = w.file(`f/${i}.txt`, i % 2 ? 0 : undefined);
      f.push(`part1-${i}|`);
      f.push(strToU8(`part2-${i}`), true);
    }
    expect(chunks.length).toBeGreaterThan(10); // emitted incrementally
    w.end();
    expect(finals).toBe(1);
    const total = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of chunks) total.set(c, o), (o += c.length);
    const m = unzipSync(total);
    expect(m.size).toBe(200);
    expect(strFromU8(m.get('f/7.txt')!)).toBe('part1-7|part2-7');
    expect(strFromU8(m.get('f/8.txt')!)).toBe('part1-8|part2-8');
  });
});

describe('unzipStream', () => {
  it('round-trips with random tiny chunks', async () => {
    const zip = zipSync(sample, 6);
    for (const seed of [1, 2, 3]) {
      const m = await collect(chunked(zip, seed));
      expect([...m.keys()]).toEqual(sample.map((s) => s.path));
      expect(strFromU8(m.get('a.txt')!)).toBe('hello');
      expect(m.get('bin.dat')).toEqual(sample[3]!.data);
      expect(m.get('empty.txt')!.length).toBe(0);
    }
  });

  it('reads archives with data descriptors (fflate streaming Zip) in small chunks', async () => {
    const parts: Uint8Array[] = [];
    const z = new Zip((err, c) => {
      if (err) throw err;
      parts.push(c);
    });
    for (let i = 0; i < 5; i++) {
      const f = new ZipDeflate(`s${i}.txt`, { level: 6 });
      z.add(f);
      f.push(strToU8(`data-${i}-`.repeat(500)), true);
    }
    z.end();
    const zip = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of parts) zip.set(p, o), (o += p.length);
    const m = await collect(chunked(zip, 9));
    expect(m.size).toBe(5);
    expect(strFromU8(m.get('s3.txt')!)).toBe('data-3-'.repeat(500));
    expect(unzipSync(zip).size).toBe(5);
  });

  it('can skip entries without consuming them', async () => {
    const zip = zipSync(sample, 6);
    const seen: string[] = [];
    for await (const e of unzipStream(chunked(zip))) seen.push(e.path);
    expect(seen).toEqual(sample.map((s) => s.path));
  });

  it('enforces limits and paths', async () => {
    const zip = zipSync(sample, 6);
    await expect(collect(whole(zip), { maxEntries: 2 })).rejects.toBeInstanceOf(ZipLimitError);
    const bomb = zipSync([{ path: 'z', data: new Uint8Array(50_000_000) }], 9);
    await expect(collect(chunked(bomb.subarray(0, 100)))).rejects.toThrow();
    await expect(collect(whole(bomb))).rejects.toThrow(/ratio/);
    const bad = rawZipSync({ '../evil.txt': strToU8('x') });
    await expect(collect(whole(bad))).rejects.toBeInstanceOf(ZipPathError);
  });
});

describe('limits', () => {
  it('too many entries', () => {
    const zip = zipSync(sample, 6);
    expect(() => unzipSync(zip, { maxEntries: 3 })).toThrow(ZipLimitError);
    expect(unzipSync(zip, { maxEntries: 4 }).size).toBe(4);
  });

  it('entry too large (actual output)', () => {
    const zip = zipSync([{ path: 'a', data: 'x'.repeat(10_000) }], 6);
    expect(() => unzipSync(zip, { maxEntryBytes: 5000 })).toThrow(/too large/);
  });

  it('total too large', () => {
    const zip = zipSync(
      [
        { path: 'a', data: 'x'.repeat(6000) },
        { path: 'b', data: 'y'.repeat(6000) },
      ],
      6,
    );
    expect(() => unzipSync(zip, { maxTotalBytes: 10_000 })).toThrow(/total/i);
  });

  it('ratio bomb: 50MB of zeros', () => {
    const zip = zipSync([{ path: 'zeros.bin', data: new Uint8Array(50_000_000) }], 9);
    expect(zip.length).toBeLessThan(100_000);
    expect(() => unzipSync(zip)).toThrow(/ratio/);
    // passes when the ratio limit is relaxed
    expect(unzipSync(zip, { maxCompressionRatio: 1e9 }).get('zeros.bin')!.length).toBe(50_000_000);
  });

  it('does not flag small highly compressible entries', () => {
    const zip = zipSync([{ path: 'a.xml', data: ' '.repeat(500_000) }], 9);
    expect(unzipSync(zip).get('a.xml')!.length).toBe(500_000);
  });

  it('lying headers do not bypass limits', () => {
    const zip = zipSync([{ path: 'a', data: 'x'.repeat(100_000) }], 6).slice();
    // zero out declared uncompressed size in local header (offset 22) and central dir
    const dv = new DataView(zip.buffer);
    dv.setUint32(22, 1, true);
    const cd = zip.length - 22 - 46 - 1; // central dir header starts 46+name bytes before EOCD
    dv.setUint32(cd + 24, 1, true);
    expect(() => unzipSync(zip, { maxEntryBytes: 1000 })).toThrow(ZipLimitError);
  });

  it.each(['../a', 'a/../../b', '/abs', 'C:/x', 'C:x', 'a\\b', 'a\0b'])(
    'rejects unsafe path %j',
    (p) => {
      const zip = rawZipSync({ [p]: strToU8('x') });
      expect(() => unzipSync(zip)).toThrow(ZipPathError);
    },
  );
});

describe('real xlsx', () => {
  it('reads a file from the old repo test data', async () => {
    const dir = new URL('./fixtures/exceljs/data/', import.meta.url);
    const f = readdirSync(dir).find((n: string) => n.endsWith('.xlsx'))!;
    const bytes = new Uint8Array(readFileSync(new URL(f, dir)));
    const m = unzipSync(bytes);
    expect(m.has('[Content_Types].xml')).toBe(true);
    expect(m.has('xl/workbook.xml')).toBe(true);
    const s = await collect(chunked(bytes, 5));
    expect([...s.keys()].sort()).toEqual([...m.keys()].sort());
    expect(s.get('xl/workbook.xml')).toEqual(m.get('xl/workbook.xml'));
  });
});

describe('regressions', () => {
  it('stored (level 0) entries larger than 1MB are not flagged as ratio bombs', () => {
    const big = new Uint8Array(3 * 1024 * 1024); // all zeros: ratio check would fire if applied
    const z = zipSync([{ path: 'big.bin', data: big }], 0);
    const m = unzipSync(z);
    expect(m.get('big.bin')?.length).toBe(big.length);
  });
});

describe('unzipLazy', () => {
  const join = (parts: Iterable<Uint8Array>): Uint8Array => {
    const arr = [...parts];
    const out = new Uint8Array(arr.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of arr) out.set(p, o), (o += p.length);
    return out;
  };

  it('indexes without inflating and matches unzipSync', () => {
    for (const level of [0, 6]) {
      const zip = zipSync(sample, level);
      const lazy = unzipLazy(zip);
      const eager = unzipSync(zip);
      expect([...lazy.keys()]).toEqual([...eager.keys()]);
      for (const [k, v] of eager) {
        const e = lazy.get(k)!;
        expect(e.size).toBe(v.length);
        expect(e.compression).toBe(level === 0 ? 0 : 8);
        expect(e.bytes()).toEqual(v);
        expect(join(e.chunks(1))).toEqual(v);
      }
    }
  });

  it('chunks() inflates incrementally', () => {
    const data = strToU8(
      Array.from({ length: 60_000 }, (_, i) => `<c r="A${i}"><v>${i * 3.7}</v></c>`).join(''),
    );
    const e = unzipLazy(zipSync([{ path: 'big.bin', data }], 6)).get('big.bin')!;
    const parts = [...e.chunks(16 * 1024)];
    expect(parts.length).toBeGreaterThan(3);
    expect(Math.max(...parts.map((p) => p.length))).toBeLessThan(data.length);
    expect(join(parts)).toEqual(data);
  });

  it('enforces limits on actual output', () => {
    const zip = zipSync(sample, 6);
    expect(() => unzipLazy(zip, { maxEntries: 3 })).toThrow(ZipLimitError);
    const big = zipSync([{ path: 'a', data: 'x'.repeat(100_000) }], 6).slice();
    const dv = new DataView(big.buffer);
    dv.setUint32(22, 1, true);
    dv.setUint32(big.length - 22 - 46 - 1 + 24, 1, true); // lie in the central directory too
    const e = unzipLazy(big, { maxEntryBytes: 1000 }).get('a')!;
    expect(() => e.bytes()).toThrow(/too large/);
    const bomb = zipSync([{ path: 'zeros.bin', data: new Uint8Array(50_000_000) }], 9);
    expect(() => unzipLazy(bomb).get('zeros.bin')!.bytes()).toThrow(/ratio/);
    expect(() => [...unzipLazy(bomb).get('zeros.bin')!.chunks()]).toThrow(/ratio/);
    const two = zipSync(
      [
        { path: 'a', data: 'x'.repeat(6000) },
        { path: 'b', data: 'y'.repeat(6000) },
      ],
      6,
    );
    const m = unzipLazy(two, { maxTotalBytes: 10_000 });
    m.get('a')!.bytes();
    m.get('a')!.bytes(); // re-reading the same entry is not double counted
    expect(() => m.get('b')!.bytes()).toThrow(/total/i);
  });

  it.each(['../a', '/abs', 'a\\b'])('rejects unsafe path %j', (p) => {
    expect(() => unzipLazy(rawZipSync({ [p]: strToU8('x') }))).toThrow(ZipPathError);
  });

  it('rejects non-zip data and truncated archives', () => {
    expect(() => unzipLazy(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toThrow(ZipLimitError);
    expect(unzipLazy(new Uint8Array(0)).size).toBe(0);
    const zip = zipSync(sample, 6);
    expect(() => unzipLazy(zip.subarray(0, zip.length - 30))).toThrow(ZipLimitError);
  });

  it('deflateStream / inflateStream round-trip', () => {
    const src = strToU8('<x>hello</x>'.repeat(10_000));
    const comp: Uint8Array[] = [];
    const d = deflateStream(1, (c) => void comp.push(c));
    d.push(src.subarray(0, 5000));
    d.push(src.subarray(5000), true);
    const outp: Uint8Array[] = [];
    const inf = inflateStream((c) => void outp.push(c));
    for (let i = 0; i < comp.length; i++) inf.push(comp[i]!, i === comp.length - 1);
    expect(join(outp)).toEqual(src);
  });
});

describe('native deflate (#71)', () => {
  // independent table CRC-32, to check the CRCs the writers store
  const TABLE = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (d: Uint8Array): number => {
    let c = -1;
    for (let i = 0; i < d.length; i++) c = TABLE[(c ^ d[i]!) & 255]! ^ (c >>> 8);
    return ~c >>> 0;
  };
  /** CRC-32 stored in each central directory entry. */
  const storedCrcs = (zip: Uint8Array): Map<string, number> => {
    const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    let e = zip.length - 22;
    while (dv.getUint32(e, true) !== 0x06054b50) e--;
    const out = new Map<string, number>();
    let p = dv.getUint32(e + 16, true);
    for (let i = dv.getUint16(e + 10, true); i > 0; i--) {
      const n = dv.getUint16(p + 28, true);
      const name = strFromU8(zip.subarray(p + 46, p + 46 + n));
      out.set(name, dv.getUint32(p + 16, true));
      p += 46 + n + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
    }
    return out;
  };
  const text = (n: number): Uint8Array => {
    let s = '';
    for (let i = 0; i < n; i++) s += `<row r="${i}"><c><v>${(i * 7919) % 100003}</v></c></row>`;
    return strToU8(s);
  };
  const write = async (level: number, compressor?: ReturnType<typeof nativeCompressor>) => {
    const parts: Uint8Array[] = [];
    const w = new ZipStreamWriter((c) => void parts.push(c), level, { compressor });
    const data = text(20_000);
    const a = w.file('a.xml');
    for (let i = 0; i < data.length; i += 10_000) a.push(data.subarray(i, i + 10_000));
    a.push(new Uint8Array(0), true);
    w.file('b.xml', 0).push('stored', true);
    w.file('c.xml').push('', true);
    w.end();
    await w.idle();
    const out = new Uint8Array(parts.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of parts) out.set(c, o), (o += c.length);
    return { out, data };
  };

  it('honors the compression level and stores correct CRCs', async () => {
    const sizes: number[] = [];
    for (const level of [1, 6, 9]) {
      for (const comp of [nativeCompressor(), undefined]) {
        const { out, data } = await write(level, comp);
        const files = unzipSync(out);
        expect(files.get('a.xml')).toEqual(data);
        expect(strFromU8(files.get('b.xml')!)).toBe('stored');
        const crcs = storedCrcs(out);
        expect(crcs.get('a.xml')).toBe(crc32(data));
        expect(crcs.get('b.xml')).toBe(crc32(strToU8('stored')));
        expect(crcs.get('c.xml')).toBe(0);
        if (comp) sizes.push(out.length);
      }
    }
    // node:zlib path: level 1 < 6 <= 9 in compression ratio
    expect(sizes[0]!).toBeGreaterThan(sizes[1]!);
    expect(sizes[1]!).toBeGreaterThanOrEqual(sizes[2]!);
  });

  it('CompressionStream entries (browser path) are valid', async () => {
    const q = new Inflight();
    const parts: Uint8Array[] = [];
    const z = new Zip((err, c) => {
      if (err) throw err;
      parts.push(c);
    });
    const f = new NativeDeflateFile('x.xml', q);
    z.add(f);
    const data = text(5000);
    f.push(data.subarray(0, 1000));
    f.push(data.subarray(1000), true);
    z.end();
    await q.idle();
    const out = new Uint8Array(parts.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of parts) out.set(c, o), (o += c.length);
    expect(unzipSync(out).get('x.xml')).toEqual(data);
    expect(storedCrcs(out).get('x.xml')).toBe(crc32(data));
  });
});

describe('native sync zlib (#75)', () => {
  const big = (): Uint8Array => {
    let s = '';
    for (let i = 0; i < 120_000; i++) s += `<c r="A${i}" t="s"><v>${(i * 7919) % 100003}</v></c>`;
    return strToU8(s); // > 1 MiB: several sync-flushed batches
  };

  it('is available in this runtime', () => {
    expect(nodeZlib()?.deflateRawSync).toBeTypeOf('function');
  });

  it('sync writer output (multi-batch) decodes with fflate and honors the level', () => {
    const data = big();
    expect(data.length).toBeGreaterThan(3 << 20);
    const sizes: number[] = [];
    for (const level of [1, 9]) {
      const parts: Uint8Array[] = [];
      const w = new ZipStreamWriter((c) => void parts.push(c), level);
      const f = w.file('s.xml');
      for (let i = 0; i < data.length; i += 100_000) f.push(data.subarray(i, i + 100_000));
      f.push(new Uint8Array(0), true);
      w.file('e.xml').push('', true);
      w.end();
      const out = new Uint8Array(parts.reduce((n, c) => n + c.length, 0));
      let o = 0;
      for (const c of parts) out.set(c, o), (o += c.length);
      const raw = rawUnzipSync(out); // fflate's own inflater, independent of zlib
      expect(raw['s.xml']).toEqual(data);
      expect(raw['e.xml']!.length).toBe(0);
      expect(unzipLazy(out).get('s.xml')!.bytes()).toEqual(data);
      sizes.push(out.length);
    }
    expect(sizes[0]!).toBeGreaterThan(sizes[1]!);
  });

  it('lazy native inflate: corrupt data still throws, lying sizes fall back to streaming', () => {
    const zip = zipSync([{ path: 'a', data: big() }], 6).slice();
    const start = 30 + 1; // local header + name 'a'
    for (let i = 100; i < 140; i++) zip[start + i] = 0xff; // garble the deflate stream
    expect(() => unzipLazy(zip).get('a')!.bytes()).toThrow();
    expect(() => [...unzipLazy(zip).get('a')!.chunks()]).toThrow();

    // central directory says 10 bytes; actual output is 70 MB (> the 64 MiB native cap)
    const zeros = new Uint8Array(70 << 20);
    const z2 = zipSync([{ path: 'z', data: zeros }], 1).slice();
    const dv2 = new DataView(z2.buffer);
    let e = z2.length - 22;
    while (dv2.getUint32(e, true) !== 0x06054b50) e--;
    const cd = dv2.getUint32(e + 16, true);
    dv2.setUint32(cd + 24, 10, true);
    let n = 0;
    const limits = { maxCompressionRatio: 1e9 };
    for (const c of unzipLazy(z2, limits)
      .get('z')!
      .chunks(1 << 20))
      n += c.length;
    expect(n).toBe(zeros.length);
    expect(() => [...unzipLazy(z2).get('z')!.chunks()]).toThrow(/ratio/);
  });
});
