// Generates sample files from the built package for external validation (LibreOffice in CI).
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook, writeXlsx } from '../../dist/index.js';

const out = process.argv[2] ?? 'out';
mkdirSync(out, { recursive: true });

const wb = new Workbook();
const ws = wb.addWorksheet('Smoke');
ws.columns = [
  { header: 'n', key: 'n' },
  { header: 's', key: 's', width: 24 },
  { header: 'd', key: 'd' },
];
for (let i = 0; i < 1000; i++)
  ws.addRow({ n: i, s: `text ${i} <&>`, d: new Date(Date.UTC(2024, 0, 1 + i)) });
ws.setCell('E1', { formula: 'SUM(A2:A1001)', result: 499500 }, { font: { bold: true } });
ws.mergeCells('F1:G2');
ws.freeze(1);

for (const [name, opts] of Object.entries({
  balanced: {},
  'low-memory': { profile: 'low-memory' },
  'max-speed': { profile: 'max-speed' },
  stored: { compression: 0 },
})) {
  writeFileSync(join(out, `${name}.xlsx`), writeXlsx(wb, opts));
}
console.log(`wrote smoke files to ${out}`);
