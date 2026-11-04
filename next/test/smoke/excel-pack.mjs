// Generates a pack of .xlsx files covering every feature, for a manual check in real Microsoft
// Excel (Excel is stricter than LibreOffice/exceljs). Run after `pnpm build`:
//   node test/smoke/excel-pack.mjs [outDir]      (default: out/excel-pack)
// Then open each file in Excel and note any "We found a problem with some content..." prompt.
// The checklist written next to the files says what to look at in each one.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { encryptXlsx } from '../../dist/crypto.js';
import {
  addChart,
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addImage,
  addNote,
  addTable,
  createXlsxWriter,
  placeImage,
  readXlsx,
  setPrintArea,
  setPrintTitles,
  Workbook,
  writeXlsx,
} from '../../dist/index.js';

const out = process.argv[2] ?? 'out/excel-pack';
mkdirSync(out, { recursive: true });
const SPEC = fileURLToPath(new URL('../../../spec/integration/data/', import.meta.url));
const checklist = [];
const save = (name, bytes, check) => {
  writeFileSync(join(out, name), bytes);
  checklist.push(`- [ ] **${name}** — ${check}`);
};

// 8x8 red PNG (valid, tiny).
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFWEXHbQSACj/P8Fu7N9hAAAAAElFTkSuQmCC';

const bold = { font: { bold: true } };
const header = {
  font: { bold: true, color: { argb: 'FFFFFFFF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: { bottom: { style: 'medium', color: { argb: 'FF000000' } } },
};

// 01 — values, styles, layout -------------------------------------------------------------
function valuesWorkbook() {
  const wb = new Workbook();
  wb.creator = 'excel-pack';
  const ws = wb.addWorksheet('Values', { tabColor: { argb: 'FF00B050' } });
  ws.columns = [
    { header: 'Kind', key: 'k', width: 18 },
    { header: 'Value', key: 'v', width: 34 },
    { header: 'Note', key: 'n', width: 40 },
  ];
  for (let c = 1; c <= 3; c++) ws.setStyle(1, c, header);
  const rows = [
    ['integer', 42],
    ['negative float', -1234.5678],
    ['big number', 1e21],
    ['tiny number', 5e-324],
    ['percent (0.00%)', 0.1234],
    ['currency', 1234567.891],
    ['date', new Date(Date.UTC(2024, 1, 29))],
    ['date+time', new Date(Date.UTC(2024, 11, 31, 23, 59, 30))],
    ['boolean TRUE', true],
    ['boolean FALSE', false],
    ['error', { error: '#N/A' }],
    ['formula', { formula: 'B2*2', result: 84 }],
    ['string formula', { formula: 'UPPER("abc")', result: 'ABC' }],
    ['array formula', { formula: 'SUM(B2:B3*2)', result: -2385.1356, array: 'B15' }],
    ['escaping', 'a<b & c>"d" \'e\''],
    ['whitespace', '  leading and trailing  '],
    ['unicode', 'émoji 🚀 中文 עברית'],
    ['control chars', 'tab\there\u0001bell'],
    ['literal _x0041_', 'keep _x0041_ literally'],
    ['multiline', 'line 1\nline 2'],
    [
      'rich text',
      {
        richText: [
          { text: 'bold ', font: { bold: true } },
          { text: 'red', font: { color: { argb: 'FFFF0000' } } },
        ],
      },
    ],
    ['hyperlink', { text: 'example.com', hyperlink: 'https://example.com', tooltip: 'open' }],
    ['internal link', { text: 'go to Other!A1', hyperlink: '#Other!A1' }],
  ];
  for (const [k, v] of rows) ws.addRow({ k, v });
  ws.setStyle('B6', { numFmt: '0.00%' });
  ws.setStyle('B7', { numFmt: '"$"#,##0.00' });
  ws.setStyle('B22', { alignment: { wrapText: true } });
  ws.setCell('C2', 'italic, strike, underline, 14pt', {
    font: { italic: true, strike: true, underline: 'double', size: 14 },
  });
  ws.setCell('C3', 'all borders + rotated', {
    border: {
      top: { style: 'thin' },
      left: { style: 'dashed' },
      bottom: { style: 'double' },
      right: { style: 'thick', color: { argb: 'FFFF0000' } },
    },
    alignment: { textRotation: 45 },
  });
  ws.setCell('C4', 'merged C4:C6');
  ws.mergeCells('C4:C6');
  ws.setRowProps(9, { height: 30 });
  ws.setRowProps(20, { hidden: true });
  ws.freeze(1, 1);
  ws.autoFilter = 'A1:B24';
  const other = wb.addWorksheet('Other');
  other.setCell('A1', 'target of the internal link');
  const hidden = wb.addWorksheet('Hidden', { state: 'hidden' });
  hidden.setCell('A1', 'this sheet is hidden');
  return wb;
}
save(
  '01-values-styles.xlsx',
  writeXlsx(valuesWorkbook()),
  'every value type, fonts/fills/borders/alignment, merges, frozen pane at B2, autofilter, 3 sheets (one hidden), green tab, row 20 hidden, both links work.',
);

// 02 — 1904 date system -------------------------------------------------------------------
{
  const wb = new Workbook({ date1904: true });
  const ws = wb.addWorksheet('1904');
  ws.addRow(['date in 1904 system', new Date(Date.UTC(2024, 0, 15))]);
  save(
    '02-date1904.xlsx',
    writeXlsx(wb),
    'B1 shows 2024-01-15 (File > Options > Advanced shows "Use 1904 date system" checked).',
  );
}

// 03 — tables ------------------------------------------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Tables');
  addTable(ws, {
    name: 'Sales',
    topLeft: 'B2',
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Units', totalsRowFunction: 'sum' },
      { name: 'Price', totalsRowFunction: 'average' },
      {
        name: 'Revenue',
        totalsRowFunction: 'sum',
        calculatedColumnFormula: 'Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]',
      },
    ],
    rows: [
      ['North', 10, 2.5],
      ['South', 7, 3],
      ['East', 12, 1.75],
      ['West', 3, 9.99],
    ],
    totalsRow: true,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
  });
  addTable(ws, {
    name: 'NoStyle',
    topLeft: 'H2',
    columns: ['A', 'B'],
    rows: [
      [1, 2],
      [3, 4],
    ],
    style: { theme: 'TableStyleLight1' },
    headerStyle: { font: { italic: true } },
  });
  save(
    '03-tables.xlsx',
    writeXlsx(wb),
    'two tables; Sales has a totals row (sum/average) and a calculated Revenue column; filter buttons on headers.',
  );
}

// 04 — conditional formatting ------------------------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('CF');
  for (let i = 1; i <= 20; i++) ws.addRow([i, 21 - i, i % 5 === 0 ? 'apple pie' : 'banana', i * 3]);
  addConditionalFormatting(ws, 'A1:A20', [
    {
      type: 'cellIs',
      operator: 'greaterThan',
      formulae: [15],
      style: {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
        font: { color: { argb: 'FF9C0006' } },
      },
    },
  ]);
  addConditionalFormatting(ws, 'B1:B20', [
    {
      type: 'colorScale',
      colors: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
    },
  ]);
  addConditionalFormatting(ws, 'C1:C20', [
    {
      type: 'containsText',
      text: 'apple',
      style: { font: { bold: true, color: { argb: 'FF0070C0' } } },
    },
  ]);
  addConditionalFormatting(ws, 'D1:D20', [
    { type: 'dataBar', color: { argb: 'FF638EC6' }, gradient: false },
  ]);
  addConditionalFormatting(ws, 'E1:E20', [{ type: 'iconSet', iconSet: '3TrafficLights1' }]);
  for (let i = 1; i <= 20; i++) ws.setCell(i, 5, i);
  addConditionalFormatting(ws, 'F1:F20', [{ type: 'top10', rank: 3, style: bold }]);
  for (let i = 1; i <= 20; i++) ws.setCell(i, 6, (i * 7) % 13);
  save(
    '04-conditional-formatting.xlsx',
    writeXlsx(wb),
    'A>15 red; B colour scale; C "apple" bold blue; D solid data bars; E traffic lights; F top 3 bold. Home > Conditional Formatting > Manage Rules lists them.',
  );
}

// 05 — data validation --------------------------------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('DV');
  ws.setCell('A1', 'pick a fruit →');
  addDataValidation(ws, {
    sqref: 'B1',
    type: 'list',
    values: ['Apple', 'Banana', 'Cherry'],
    showErrorMessage: true,
    error: 'Choose from the list',
  });
  ws.setCell('A2', 'whole number 1–10 →');
  addDataValidation(ws, {
    sqref: 'B2',
    type: 'whole',
    operator: 'between',
    formula1: 1,
    formula2: 10,
    showInputMessage: true,
    promptTitle: 'Number',
    prompt: 'Enter 1 to 10',
  });
  ws.setCell('A3', 'date in 2024 →');
  addDataValidation(ws, {
    sqref: 'B3',
    type: 'date',
    operator: 'between',
    formula1: new Date(Date.UTC(2024, 0, 1)),
    formula2: new Date(Date.UTC(2024, 11, 31)),
  });
  ws.setCell('A4', 'list from range D1:D3 →');
  ws.setCell('D1', 'red');
  ws.setCell('D2', 'green');
  ws.setCell('D3', 'blue');
  addDataValidation(ws, { sqref: 'B4', type: 'list', formula1: '$D$1:$D$3' });
  save(
    '05-data-validation.xlsx',
    writeXlsx(wb),
    'B1 drop-down of fruits (error on other input); B2 input prompt + 1–10 rule; B3 dates in 2024; B4 drop-down from D1:D3.',
  );
}

// 06 — notes ---------------------------------------------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Notes');
  ws.columns = [{ width: 30 }, { width: 10 }];
  ws.setCell('A1', 'hover me');
  addNote(ws, 'A1', 'A plain note');
  ws.setCell('B3', 'rich');
  addNote(ws, 'B3', {
    text: { richText: [{ text: 'Bold ', font: { bold: true } }, { text: 'and normal' }] },
    author: 'Pack',
    visible: true,
  });
  addNote(ws, 'D6', 'note on an empty cell');
  save(
    '06-notes.xlsx',
    writeXlsx(wb),
    'red note markers on A1, B3, D6; B3 note is always visible and has bold text; boxes sit next to their cells.',
  );
}

// 07 — images --------------------------------------------------------------------------------
{
  const wb = new Workbook();
  const id = addImage(wb, { base64: PNG, extension: 'png' });
  const s1 = wb.addWorksheet('Images');
  s1.setCell('A1', 'two-cell anchor B2:D6, one-cell anchor at F2 (80x40 px)');
  placeImage(s1, id, 'B2:D6');
  placeImage(
    s1,
    id,
    { tl: 'F2', ext: { width: 80, height: 40 } },
    { hyperlink: 'https://example.com', description: 'red square' },
  );
  const s2 = wb.addWorksheet('Shared');
  s2.setCell('A1', 'same image reused on a second sheet');
  placeImage(s2, id, 'B2:C3');
  save(
    '07-images.xlsx',
    writeXlsx(wb),
    'red squares: stretched over B2:D6, small one at F2 (clickable link), and on sheet "Shared".',
  );
}

// 08 — protection, print setup, defined names --------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Protected');
  ws.setCell('A1', 'locked (sheet protected, password: test123)');
  ws.setCell('A2', 'unlocked — you can edit this', { protection: { locked: false } });
  ws.protection = { password: 'test123', formatCells: false };
  ws.pageSetup = {
    orientation: 'landscape',
    paperSize: 9,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };
  ws.headerFooter = { oddHeader: '&CExcel pack', oddFooter: '&LPage &P of &N&R&D' };
  for (let i = 3; i <= 60; i++) ws.addRow([`row ${i}`, i]);
  setPrintArea(ws, 'A1:B60');
  setPrintTitles(ws, { rows: '1:2' });
  addDefinedName(wb, 'TaxRate', '0.2');
  addDefinedName(wb, 'Numbers', "'Protected'!$B$3:$B$60");
  ws.setCell('D1', { formula: 'SUM(Numbers)*TaxRate', result: 0 });
  save(
    '08-protection-print.xlsx',
    writeXlsx(wb),
    'sheet protected (Review > Unprotect asks for test123), A2 editable; Print Preview: landscape A4, header/footer, rows 1–2 repeat, print area A1:B60; D1 uses names TaxRate/Numbers (Formulas > Name Manager).',
  );
}

// 09/10 — streaming writer ---------------------------------------------------------------------
async function streamed(name, opts, rows) {
  const w = createXlsxWriter(opts);
  const chunks = [];
  const reader = w.readable.getReader();
  const pump = (async () => {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) return;
      chunks.push(value);
    }
  })();
  const s = w.addWorksheet('Stream', {
    columns: [
      { header: 'id', key: 'id', width: 10 },
      { header: 'text', key: 't', width: 20 },
      { header: 'when', key: 'd', width: 14 },
      { header: 'value', key: 'v', width: 12 },
    ],
    freeze: 1,
  });
  for (let i = 0; i < rows; i++) {
    s.addRow({
      id: i + 1,
      t: `row ${i + 1}`,
      d: new Date(Date.UTC(2024, 0, 1 + (i % 365))),
      v: Math.round(Math.sin(i) * 1e4) / 100,
    });
    if (i % 5000 === 0) await s.flush();
  }
  await s.commit();
  await w.end();
  await pump;
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const bytes = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) {
    bytes.set(c, o);
    o += c.length;
  }
  return bytes;
}
save(
  '09-stream-100k.xlsx',
  await streamed('09', { profile: 'balanced' }, 100_000),
  'streaming writer, 100,000 rows × 4 cols, shared strings, header frozen; last row id = 100000.',
);
save(
  '10-stream-inline.xlsx',
  await streamed('10', { profile: 'max-speed' }, 20_000),
  'streaming writer, max-speed profile (inline strings): 20,000 rows.',
);

// 11 — encrypted ----------------------------------------------------------------------------------
save(
  '11-encrypted-test123.xlsx',
  await encryptXlsx(writeXlsx(valuesWorkbook()), 'test123'),
  'Excel asks for a password: test123 → opens like 01-values-styles.',
);

// 12 — read → re-save (round-trip preservation) --------------------------------------------------
for (const [file, what] of [
  ['chart-sheet.xlsx', 'a chart sheet and its chart still there'],
  ['1519293514-KRISHNAPATNAM_LINE_UP.xlsx', 'picture and text-box shapes still there'],
  ['test-issue-877.xlsx', 'header/footer picture (Page Layout view) still there'],
  ['images.xlsx', 'images still placed as in the original'],
  ['formulas.xlsx', 'formulas still calculate'],
]) {
  try {
    const src = readFileSync(join(SPEC, file));
    save(
      `12-resaved-${file}`,
      writeXlsx(readXlsx(new Uint8Array(src))),
      `${what} (compare with the original in spec/integration/data/${file}).`,
    );
  } catch (e) {
    checklist.push(`- [ ] 12-resaved-${file} — SKIPPED: ${e.message}`);
  }
}

// 13 — everything on one sheet -----------------------------------------------------------------
{
  const wb = valuesWorkbook();
  const ws = wb.getWorksheet('Values');
  addNote(ws, 'A2', 'note');
  addConditionalFormatting(ws, 'B2:B3', [
    { type: 'cellIs', operator: 'lessThan', formulae: [0], style: bold },
  ]);
  addDataValidation(ws, { sqref: 'E2', type: 'list', values: ['x', 'y'] });
  placeImage(ws, addImage(wb, { base64: PNG, extension: 'png' }), 'E5:F8');
  addTable(wb.getWorksheet('Other'), {
    name: 'TblOther',
    topLeft: 'C3',
    columns: ['a', 'b'],
    rows: [[1, 2]],
  });
  save(
    '13-everything.xlsx',
    writeXlsx(wb),
    'all features together (note, CF, drop-down at E2, image E5:F8, table on "Other") — mainly: opens without a repair prompt.',
  );
}

// 14 — charts: one sheet of data + one chart per type ------------------------------------------
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Chart Data');
  ws.addRow(['Month', 'North', 'South', 'X', 'Y']);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  for (const [i, m] of months.entries())
    ws.addRow([m, 10 + i * 5, 30 - i * 3, i + 1, (i + 1) ** 2]);
  const d = "'Chart Data'!";
  const two = [
    { name: { ref: `${d}$B$1` }, categories: `${d}$A$2:$A$7`, values: `${d}$B$2:$B$7` },
    { name: { ref: `${d}$C$1` }, categories: `${d}$A$2:$A$7`, values: `${d}$C$2:$C$7` },
  ];
  const at = (i) =>
    `${'GP'[i % 2]}${2 + Math.floor(i / 2) * 18}:${'NW'[i % 2]}${17 + Math.floor(i / 2) * 18}`;
  addChart(ws, { type: 'column', title: 'Column (clustered)', series: two, anchor: at(0) });
  addChart(ws, {
    type: 'bar',
    title: 'Bar (stacked)',
    stacked: 'stacked',
    series: two,
    anchor: at(1),
    legend: { position: 'b' },
  });
  addChart(ws, {
    type: 'line',
    title: 'Line',
    series: [
      { ...two[0], color: { argb: 'FFC00000' } },
      { ...two[1], smooth: true, marker: false },
    ],
    anchor: at(2),
    axes: { x: { title: 'Month' }, y: { title: 'Units', min: 0 } },
  });
  addChart(ws, {
    type: 'pie',
    title: 'Pie',
    series: [two[0]],
    anchor: at(3),
    dataLabels: { showPercent: true },
  });
  addChart(ws, { type: 'doughnut', title: 'Doughnut', series: [two[1]], anchor: at(4) });
  addChart(ws, {
    type: 'scatter',
    title: 'Scatter',
    series: [{ name: 'y = x²', xValues: `${d}$D$2:$D$7`, values: `${d}$E$2:$E$7` }],
    anchor: at(5),
    legend: false,
  });
  addChart(ws, {
    type: 'area',
    title: 'Area (stacked)',
    stacked: 'stacked',
    series: two,
    anchor: at(6),
  });
  save(
    '14-charts.xlsx',
    writeXlsx(wb),
    'seven charts on "Chart Data" (clustered column, stacked bar, line with a red series + a smooth markerless one and axis titles, pie with % labels, doughnut, scatter y=x² without legend, stacked area), each titled with its type and showing the Jan–Jun data.',
  );
}

writeFileSync(
  join(out, 'CHECKLIST.md'),
  `# Excel compatibility check

Open each file in **Microsoft Excel** (desktop; note the version + OS below).
For each file: (1) does Excel show "We found a problem with some content … Do you want us to try to recover?" — if yes, click Yes and **copy the repair log text** (it names the broken part); (2) does it look as described.

Excel version / OS: ______________________

${checklist.join('\n')}

Report back: the file names that prompted a repair (+ the repair log text), and anything that looks wrong.
`,
);
console.log(`wrote ${checklist.length} checks to ${out}`);
