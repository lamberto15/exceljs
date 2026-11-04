// Runs the write and read benchmarks (100k rows x 20 cols, turbosheet vs exceljs 4.4.0)
// and prints machine info plus markdown tables. Usage: pnpm bench
// Needs the devDependencies installed; takes about a minute.
import { spawnSync } from 'node:child_process';
import os from 'node:os';

function run(args, extraEnv = {}) {
  const r = spawnSync('npx', ['vitest', 'run', ...args, '--silent=false'], {
    env: { ...process.env, RUN_BENCH: '1', ...extraEnv },
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
    shell: process.platform === 'win32',
  });
  if (r.status !== 0) {
    process.stderr.write(r.stdout + r.stderr);
    process.exit(r.status ?? 1);
  }
  return r.stdout;
}

const cpu = os.cpus()[0]?.model ?? 'unknown CPU';
console.log(`Node ${process.version}, ${cpu}, ${os.type()} ${os.release()} (${os.arch()})\n`);

console.log('### Write (100k rows x 20 cols)\n');
const write = run(['test/xlsx-write.bench.test.ts'], { NODE_OPTIONS: '--expose-gc' });
console.log(
  write
    .split('\n')
    .filter((l) => l.startsWith('|'))
    .join('\n'),
);

console.log('\n### Read (100k rows x 20 cols)\n');
const read = run([
  'test/xlsx-read.test.ts',
  '-t',
  'perf',
  '--pool=forks',
  '--execArgv=--expose-gc',
]);
const rows = read
  .split('\n')
  .map((l) => l.match(/^(.+?)\s+(\d+) ms\s+peak heap\+buffers \+(\d+) MB/))
  .filter(Boolean);
console.log('| case | time ms | peak heap+buffers MB |\n|---|---:|---:|');
for (const m of rows) console.log(`| ${m[1].trim()} | ${m[2]} | ${m[3]} |`);
