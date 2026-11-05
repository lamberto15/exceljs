// The public API, as a committed report (api/report.txt): every name each entry point exports,
// marked `value` (present at runtime) or `type`. A change to the API shows up as a diff of that
// file in review; update it with `npx vitest run test/api-surface.test.ts -u` when intended.
// (`pnpm check:dts` also type-checks the published .d.ts after `pnpm build`.)

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const ENTRIES: Record<string, string> = {
  turbosheet: 'src/index.ts',
  'turbosheet/lite': 'src/lite.ts',
  'turbosheet/csv': 'src/csv/index.ts',
  'turbosheet/json': 'src/json/index.ts',
  'turbosheet/crypto': 'src/crypto/index.ts',
  'turbosheet/formula': 'src/formula/index.ts',
  'turbosheet/xlsb': 'src/xlsb/index.ts',
  'turbosheet/xls': 'src/xls/index.ts',
};

const toTs = (from: string, spec: string): string =>
  resolve(dirname(from), spec).replace(/\.js$/, '.ts');

/** Names a module exports (following `export *` and re-export lists; `@internal` ones skipped). */
function names(file: string, seen = new Set<string>()): Set<string> {
  const out = new Set<string>();
  if (seen.has(file)) return out;
  seen.add(file);
  const src = readFileSync(file, 'utf8');
  for (const m of src.matchAll(/^export \* from '([^']+)';/gm)) {
    for (const n of names(toTs(file, m[1]!), seen)) out.add(n);
  }
  for (const m of src.matchAll(/^export (?:type )?\{([^}]*)\}(?: from '[^']+')?;/gm)) {
    for (const part of m[1]!.split(',')) {
      const p = part.trim().replace(/^type\s+/, '');
      if (!p) continue;
      const as = /\bas\s+(\w+)$/.exec(p);
      out.add(as ? as[1]! : p);
    }
  }
  // declarations, unless their JSDoc says @internal
  for (const m of src.matchAll(
    /(\/\*\*(?:[^*]|\*(?!\/))*\*\/\s*)?^export (?:declare )?(?:abstract )?(?:async )?(?:function\*?|const|let|class|interface|type|enum) (\w+)/gm,
  )) {
    if (m[1]?.includes('@internal')) continue;
    out.add(m[2]!);
  }
  return out;
}

describe('public API', () => {
  it('matches the committed report (api/report.txt)', async () => {
    let report = '';
    for (const [name, file] of Object.entries(ENTRIES)) {
      const runtime = new Set(Object.keys(await import(resolve(ROOT, file))));
      const all = names(resolve(ROOT, file));
      for (const k of runtime) all.add(k);
      report += `# ${name}\n`;
      for (const n of [...all].sort()) report += `${runtime.has(n) ? 'value' : 'type '} ${n}\n`;
      report += '\n';
    }
    await expect(report).toMatchFileSnapshot('../api/report.txt');
  });
});
