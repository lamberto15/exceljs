// Repo conventions enforced as tests, so mistakes fail fast on every OS.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const TEST_DIR = fileURLToPath(new URL('.', import.meta.url));

function testFiles(dir: string): string[] {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out.push(...testFiles(p));
    else if (/\.(ts|mjs|js)$/.test(name)) out.push(p);
  }
  return out;
}

describe('test conventions', () => {
  it('never derives file paths from URL.pathname (breaks on Windows: "/D:/...")', () => {
    const offenders = testFiles(TEST_DIR).filter((f) =>
      /import\.meta\.url\)\s*\.pathname/.test(readFileSync(f, 'utf8')),
    );
    expect(offenders, 'use fileURLToPath(new URL(..., import.meta.url)) instead').toEqual([]);
  });
});
