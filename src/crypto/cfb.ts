// Minimal OLE Compound File Binary (MS-CFB) reader/writer — just enough for the encrypted-OOXML
// container (MS-OFFCRYPTO 2.3.4): read named streams of v3 (512-byte sectors) and v4 (4096-byte
// sectors) files, and write a v3 file from a list of `[path, bytes]` streams.
//
// The reader treats every number in the file as hostile: sector ids are range-checked, every chain
// walk is bounded by the sector count (catches FAT cycles), the directory tree walk tracks visited
// entries, and stream sizes are capped by `Limits.maxEntryBytes`.

import { DEFAULT_LIMITS, type Limits } from '../core/types.js';

export class CfbError extends Error {
  override name = 'CfbError';
}

const END = 0xfffffffe; // ENDOFCHAIN
const FREE = 0xffffffff; // FREESECT / NOSTREAM
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const CUTOFF = 4096; // mini stream cutoff
const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** True when `data` starts with the CFB signature `D0 CF 11 E0 A1 B1 1A E1`. */
export function isCfb(data: Uint8Array): boolean {
  if (data.length < 512) return false;
  for (let i = 0; i < 8; i++) if (data[i] !== SIG[i]) return false;
  return true;
}

export interface CfbEntry {
  /** Full path, storages joined by `/` (root entry excluded), e.g. `\x06DataSpaces/Version`. */
  path: string;
  /** 1 = storage, 2 = stream. */
  type: number;
  start: number;
  size: number;
}

export interface Cfb {
  entries: CfbEntry[];
  /** Stream bytes by path, or undefined when no such stream exists. */
  read(path: string): Uint8Array | undefined;
}

const fail = (msg: string): never => {
  throw new CfbError(`Invalid compound file: ${msg}`);
};

export function readCfb(data: Uint8Array, limits?: Partial<Limits>): Cfb {
  const maxBytes = limits?.maxEntryBytes ?? DEFAULT_LIMITS.maxEntryBytes;
  if (!isCfb(data)) fail('missing signature');
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const u32 = (o: number): number => dv.getUint32(o, true);
  const major = dv.getUint16(26, true);
  const shift = dv.getUint16(30, true);
  if (dv.getUint16(28, true) !== 0xfffe) fail('bad byte order');
  if (!((major === 3 && shift === 9) || (major === 4 && shift === 12))) fail('bad version');
  if (dv.getUint16(32, true) !== 6 || u32(56) !== CUTOFF) fail('bad mini stream parameters');
  const ss = 1 << shift;
  const per = ss >> 2; // u32 entries per sector
  const nSec = Math.ceil((data.length - ss) / ss);
  const sector = (id: number): Uint8Array => {
    if (id >= nSec) fail(`sector ${id} out of range`);
    const o = (id + 1) * ss;
    return data.subarray(o, Math.min(o + ss, data.length));
  };

  // FAT sector ids: 109 in the header, the rest in the DIFAT chain.
  const nFat = u32(44);
  if (nFat > nSec) fail('FAT larger than file');
  const fatIds: number[] = [];
  for (let i = 0; i < 109 && fatIds.length < nFat; i++) fatIds.push(u32(76 + i * 4));
  for (let d = u32(68), n = 0; fatIds.length < nFat; n++) {
    if (d === END || d === FREE || n >= nSec) fail('broken DIFAT chain');
    const off = (d + 1) * ss;
    sector(d);
    for (let i = 0; i < per - 1 && fatIds.length < nFat; i++) fatIds.push(u32(off + i * 4));
    d = off + ss <= data.length ? u32(off + ss - 4) : END;
  }
  const fat = new Uint32Array(nFat * per);
  fatIds.forEach((id, i) => {
    const s = sector(id);
    if (s.length < ss) fail('truncated FAT sector');
    for (let j = 0; j < per; j++) fat[i * per + j] = dv.getUint32((id + 1) * ss + j * 4, true);
  });

  // Follow a chain (each sector at most once — catches cycles); `need` stops early once enough.
  const chain = (start: number, table: Uint32Array, need = Infinity): number[] => {
    const out: number[] = [];
    const seen = new Uint8Array(table.length);
    for (let s = start; s !== END && out.length < need; s = table[s]!) {
      if (s >= table.length || seen[s]) fail('broken sector chain');
      seen[s] = 1;
      out.push(s);
    }
    return out;
  };
  const gather = (ids: number[], size: number, get: (id: number) => Uint8Array, unit: number) => {
    if (ids.length * unit < size) fail('stream shorter than its size');
    const out = new Uint8Array(size);
    for (let i = 0, o = 0; o < size; i++) {
      const b = get(ids[i]!);
      const n = Math.min(unit, size - o);
      if (b.length < n) fail('truncated sector');
      out.set(b.subarray(0, n), o);
      o += n;
    }
    return out;
  };
  const big = (start: number, size: number): Uint8Array =>
    gather(chain(start, fat, Math.ceil(size / ss)), size, sector, ss);

  // Directory.
  const dirIds = chain(u32(48), fat);
  const dir = gather(dirIds, dirIds.length * ss, sector, ss);
  const ddv = new DataView(dir.buffer);
  const count = dir.length >> 7;
  if (count === 0 || dir[66] !== 5) fail('missing root entry');
  const sizeOf = (o: number): number => {
    const hi = major === 4 ? ddv.getUint32(o + 124, true) : 0;
    const n = hi * 0x100000000 + ddv.getUint32(o + 120, true);
    if (n > maxBytes) fail('stream exceeds size limit');
    return n;
  };

  // Mini stream (in the root entry's chain) and mini FAT.
  const miniFatIds = chain(u32(60), fat, u32(64));
  const miniFat = new Uint32Array(miniFatIds.length * per);
  miniFatIds.forEach((id, i) => {
    const s = sector(id);
    for (let j = 0; j < per && j * 4 < s.length; j++)
      miniFat[i * per + j] = dv.getUint32((id + 1) * ss + j * 4, true);
  });
  let mini: Uint8Array | undefined;
  const miniStream = (): Uint8Array => (mini ??= big(ddv.getUint32(116, true), sizeOf(0)));

  // Walk the red-black trees (any shape accepted), building paths.
  const entries: CfbEntry[] = [];
  const seen = new Uint8Array(count);
  const walk = (sid: number, prefix: string): void => {
    const stack = [sid];
    while (stack.length) {
      const id = stack.pop()!;
      if (id === FREE) continue;
      if (id >= count || seen[id]) fail('bad directory tree');
      seen[id] = 1;
      const o = id << 7;
      const nameLen = ddv.getUint16(o + 64, true);
      if (nameLen > 64 || nameLen & 1) fail('bad entry name');
      let name = '';
      for (let i = 0; i + 2 < nameLen; i += 2)
        name += String.fromCharCode(ddv.getUint16(o + i, true));
      const type = dir[o + 66]!;
      stack.push(ddv.getUint32(o + 68, true), ddv.getUint32(o + 72, true));
      if (type !== 1 && type !== 2) continue;
      const path = prefix + name;
      entries.push({ path, type, start: ddv.getUint32(o + 116, true), size: sizeOf(o) });
      if (type === 1) walk(ddv.getUint32(o + 76, true), `${path}/`);
    }
  };
  seen[0] = 1;
  walk(ddv.getUint32(76, true), '');

  return {
    entries,
    read(path) {
      const e = entries.find((x) => x.type === 2 && x.path === path);
      if (!e) return undefined;
      if (e.size >= CUTOFF) return big(e.start, e.size);
      const ms = miniStream();
      const ids = chain(e.start, miniFat, Math.ceil(e.size / 64));
      return gather(ids, e.size, (id) => ms.subarray(id * 64, id * 64 + 64), 64);
    },
  };
}

// ---------------------------------------------------------------------------------------------
// Writer (v3, 512-byte sectors). Layout: FAT | DIFAT | directory | mini FAT | mini stream | streams.

interface Node {
  name: string;
  data?: Uint8Array;
  kids: Node[];
  sid: number;
  start: number;
}

const cmp = (a: Node, b: Node): number =>
  a.name.length - b.name.length ||
  (a.name.toUpperCase() < b.name.toUpperCase()
    ? -1
    : a.name.toUpperCase() > b.name.toUpperCase()
      ? 1
      : 0);

/** Build a v3 compound file holding the given streams (paths use `/` between storages). */
export function writeCfb(files: ReadonlyArray<readonly [string, Uint8Array]>): Uint8Array {
  const root: Node = { name: 'Root Entry', kids: [], sid: 0, start: END };
  const nodes = [root];
  for (const [path, data] of files) {
    let at = root;
    const parts = path.split('/');
    parts.forEach((name, i) => {
      let n = at.kids.find((k) => k.name === name);
      if (!n) {
        if (name.length > 31) throw new CfbError(`CFB name too long: ${name}`);
        n = { name, kids: [], sid: nodes.length, start: END };
        if (i === parts.length - 1) n.data = data;
        nodes.push(n);
        at.kids.push(n);
      }
      at = n;
    });
  }
  const streams = nodes.filter((n) => n.data);
  const small = streams.filter((n) => n.data!.length < CUTOFF);
  const large = streams.filter((n) => n.data!.length >= CUTOFF);
  const sec = (n: number, u: number): number => Math.ceil(n / u);
  const miniSecs = small.reduce((a, n) => a + sec(n.data!.length, 64), 0);
  const nDir = sec(nodes.length, 4);
  const nMiniFat = sec(miniSecs, 128);
  const nMini = sec(miniSecs, 8);
  const base = nDir + nMiniFat + nMini + large.reduce((a, n) => a + sec(n.data!.length, 512), 0);
  let nFat = 1;
  let nDifat = 0;
  while (base + nFat + nDifat > nFat * 128) {
    nFat++;
    nDifat = nFat > 109 ? sec(nFat - 109, 127) : 0;
  }
  const total = nFat + nDifat + base;
  const out = new Uint8Array((total + 1) * 512);
  const dv = new DataView(out.buffer);
  const put = (o: number, v: number): void => dv.setUint32(o, v, true);
  const fat = new Uint32Array(nFat * 128).fill(FREE);
  const run = (table: Uint32Array, start: number, n: number): number => {
    for (let i = 0; i < n; i++) table[start + i] = i === n - 1 ? END : start + i + 1;
    return n ? start : END;
  };
  for (let i = 0; i < nFat; i++) fat[i] = FATSECT;
  for (let i = 0; i < nDifat; i++) fat[nFat + i] = DIFSECT;
  let next = nFat + nDifat;
  const dirStart = run(fat, next, nDir);
  next += nDir;
  const miniFatStart = run(fat, next, nMiniFat);
  next += nMiniFat;
  root.start = run(fat, next, nMini);
  const miniBase = (next + 1) * 512;
  next += nMini;
  const miniFat = new Uint32Array(nMiniFat * 128).fill(FREE);
  let m = 0;
  for (const n of small) {
    const k = sec(n.data!.length, 64);
    n.start = run(miniFat, m, k);
    out.set(n.data!, miniBase + m * 64);
    m += k;
  }
  for (const n of large) {
    const k = sec(n.data!.length, 512);
    n.start = run(fat, next, k);
    out.set(n.data!, (next + 1) * 512);
    next += k;
  }

  // Header.
  out.set(SIG);
  dv.setUint16(24, 0x3e, true);
  dv.setUint16(26, 3, true);
  dv.setUint16(28, 0xfffe, true);
  dv.setUint16(30, 9, true);
  dv.setUint16(32, 6, true);
  put(44, nFat);
  put(48, dirStart);
  put(56, CUTOFF);
  put(60, miniFatStart);
  put(64, nMiniFat);
  put(68, nDifat ? nFat : END);
  put(72, nDifat);
  for (let i = 0; i < 109; i++) put(76 + i * 4, i < nFat ? i : FREE);
  // DIFAT sectors: 127 FAT ids + next pointer each.
  for (let d = 0; d < nDifat; d++) {
    const o = (nFat + d + 1) * 512;
    for (let i = 0; i < 127; i++) {
      const f = 109 + d * 127 + i;
      put(o + i * 4, f < nFat ? f : FREE);
    }
    put(o + 508, d === nDifat - 1 ? END : nFat + d + 1);
  }
  for (let i = 0; i < fat.length; i++) put(512 + i * 4, fat[i]!);
  const mfo = (miniFatStart + 1) * 512;
  for (let i = 0; i < miniFat.length; i++) put(mfo + i * 4, miniFat[i]!);

  // Directory: each storage's children form a balanced binary tree (all black — readers don't
  // validate colouring) ordered by length, then case-insensitive name.
  const dirOff = (dirStart + 1) * 512;
  for (let i = 0; i < nDir * 4; i++) {
    const o = dirOff + i * 128;
    put(o + 68, FREE);
    put(o + 72, FREE);
    put(o + 76, FREE);
  }
  const tree = (kids: Node[], lo: number, hi: number): number => {
    if (lo > hi) return FREE;
    const mid = (lo + hi) >> 1;
    const o = dirOff + kids[mid]!.sid * 128;
    put(o + 68, tree(kids, lo, mid - 1));
    put(o + 72, tree(kids, mid + 1, hi));
    return kids[mid]!.sid;
  };
  for (const n of nodes) {
    const o = dirOff + n.sid * 128;
    for (let i = 0; i < n.name.length; i++) dv.setUint16(o + i * 2, n.name.charCodeAt(i), true);
    dv.setUint16(o + 64, (n.name.length + 1) * 2, true);
    out[o + 66] = n === root ? 5 : n.data ? 2 : 1;
    out[o + 67] = 1;
    n.kids.sort(cmp);
    put(o + 76, tree(n.kids, 0, n.kids.length - 1));
    put(o + 116, n.data || n === root ? n.start : 0);
    put(o + 120, n === root ? miniSecs * 64 : (n.data?.length ?? 0));
  }
  return out;
}
