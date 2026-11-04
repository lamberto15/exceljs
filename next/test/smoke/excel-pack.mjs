// Generates a pack of .xlsx files covering every feature, for a manual check in real Microsoft
// Excel (Excel is stricter than LibreOffice/exceljs). Run after `pnpm build`:
//   node test/smoke/excel-pack.mjs [outDir]      (default: out/excel-pack)
// Then open each file in Excel and note any "We found a problem with some content..." prompt.
// The checklist written next to the files says what to look at in each one.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { encryptXlsx } from '../../dist/crypto.js';
import { evaluate } from '../../dist/formula.js';
import {
  addChart,
  addConditionalFormatting,
  addDataValidation,
  addDefinedName,
  addImage,
  addNote,
  addTable,
  addTableRows,
  createXlsxWriter,
  deleteRows,
  duplicateRow,
  groupColumns,
  groupRows,
  insertColumns,
  insertRows,
  placeImage,
  readXlsx,
  setPrintArea,
  setPrintTitles,
  setSheetStyle,
  spliceRows,
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

// 16 — formulas computed by turbosheet/formula (cached values) ------------------------------------
{
  const wb = new Workbook();
  const data = wb.addWorksheet('Data');
  data.addRow(['Product', 'Region', 'Qty', 'Price', 'Date']);
  const rows = [
    ['Apples', 'North', 5, 1.2, Date.UTC(2024, 0, 15)],
    ['Bananas', 'South', 12, 0.5, Date.UTC(2024, 1, 29)],
    ['Cherries', 'North', 3, 4.75, Date.UTC(2024, 2, 1)],
    ['Dates', 'East', 8, 3.1, Date.UTC(2023, 11, 31)],
    ['Elderberries', 'South', 0, 6, Date.UTC(2024, 5, 30)],
  ];
  for (const [p, r, q, pr, d] of rows) data.addRow([p, r, q, pr, new Date(d)]);
  addDefinedName(wb, 'Qty', 'Data!$C$2:$C$6');
  addDefinedName(wb, 'Rate', '0.2');
  const ws = wb.addWorksheet('Formulas');
  ws.columns = [
    { header: 'What', width: 26 },
    { header: 'Formula', width: 58 },
    { header: 'Excel computes', width: 22 },
    { header: 'turbosheet computed', width: 22 },
    { header: 'Match', width: 8 },
  ];
  for (let c = 1; c <= 5; c++) ws.setStyle(1, c, header);
  const formulas = [
    ['arithmetic, precedence', '-2^2+10/4*3-1'],
    ['15-digit result', '0.1+0.2'],
    ['concatenation', '"Total: "&SUM(Data!C2:C6)&" units"'],
    ['SUM / defined name', 'SUM(Qty)*(1+Rate)'],
    ['SUMIF wildcard', 'SUMIF(Data!A2:A6,"*e*",Data!C2:C6)'],
    ['SUMIFS', 'SUMIFS(Data!C2:C6,Data!B2:B6,"North",Data!C2:C6,">3")'],
    ['SUMPRODUCT', 'SUMPRODUCT(Data!C2:C6,Data!D2:D6)'],
    ['AVERAGEIF', 'AVERAGEIF(Data!B2:B6,"South",Data!D2:D6)'],
    [
      'COUNTIF / COUNTIFS',
      'COUNTIF(Data!B2:B6,"<>North")+COUNTIFS(Data!C2:C6,">0",Data!D2:D6,"<5")',
    ],
    [
      'COUNT / COUNTA / COUNTBLANK',
      'COUNT(Data!A1:E6)&"/"&COUNTA(Data!A1:E6)&"/"&COUNTBLANK(Data!F1:F6)',
    ],
    ['MIN / MAX / AVERAGE', 'MAX(Data!D2:D6)-MIN(Data!D2:D6)+AVERAGE(Data!C2:C6)'],
    ['ROUND family', 'ROUND(2.345,2)+ROUNDUP(-3.14159,1)+ROUNDDOWN(31415.92654,-2)'],
    ['INT / MOD / ABS / POWER / SQRT', 'INT(-8.9)+MOD(-3,2)+ABS(-4)+POWER(2,10)+SQRT(16)'],
    ['PRODUCT', 'PRODUCT(Data!C2:C4)'],
    [
      'IF / AND / OR / NOT / XOR',
      'IF(AND(Data!C2>1,OR(Data!B2="x",NOT(FALSE))),XOR(TRUE,FALSE),"no")',
    ],
    ['IFS / SWITCH', 'IFS(Data!C3>10,"big",TRUE,"small")&SWITCH(Data!B4,"North","-N","-?")'],
    ['IFERROR / IFNA', 'IFERROR(1/0,"div0")&IFNA(MATCH("zz",Data!A2:A6,0),"-none")'],
    ['VLOOKUP exact', 'VLOOKUP("Cherries",Data!A2:D6,4,FALSE)'],
    ['VLOOKUP approximate', 'VLOOKUP(7,{0,"low";5,"mid";10,"high"},2)'],
    ['HLOOKUP', 'HLOOKUP("Qty",Data!A1:E6,3,FALSE)'],
    ['XLOOKUP not found', 'XLOOKUP("Figs",Data!A2:A6,Data!C2:C6,"missing")'],
    ['INDEX / MATCH', 'INDEX(Data!A2:E6,MATCH("Dates",Data!A2:A6,0),2)'],
    ['XMATCH next larger', 'XMATCH(4,Data!D2:D6,1)'],
    ['CHOOSE / ROW / COLUMNS', 'CHOOSE(2,"a","b","c")&ROW(Data!C4)&COLUMNS(Data!A1:E1)'],
    [
      'text functions',
      'PROPER(LOWER("hELLO wORLD"))&"|"&TRIM("  a   b ")&"|"&LEFT("abc",2)&RIGHT("abc",1)&MID("abcdef",3,2)',
    ],
    ['SUBSTITUTE / REPLACE', 'SUBSTITUTE("a-b-c","-","+",2)&REPLACE("2009",3,2,"10")'],
    [
      'FIND / SEARCH / LEN',
      'FIND("M","Miriam McGovern",3)+SEARCH("margin","Profit Margin")+LEN("hello")',
    ],
    ['TEXTJOIN / CONCAT', 'TEXTJOIN(", ",TRUE,Data!A2:A4)&CONCAT("-",UPPER("x"))'],
    ['TEXT number formats', 'TEXT(1234.567,"0")&" "&TEXT(0.256,"0%")&" "&TEXT(5,"000")'],
    [
      'TEXT date formats',
      'TEXT(Data!E2,"yyyy-mm-dd")&" "&TEXT(Data!E3,"dd/mm/yyyy")&" "&TEXT(Data!E4,"mmm d, yyyy")',
    ],
    ['VALUE / EXACT / REPT', 'VALUE("1000")+EXACT("a","A")+LEN(REPT("ab",3))'],
    ['DATE / YEAR / MONTH / DAY', 'DATE(YEAR(Data!E3),MONTH(Data!E3)+1,DAY(Data!E3))'],
    ['WEEKDAY / EDATE / EOMONTH', 'WEEKDAY(Data!E2)+EDATE(Data!E2,1)-EOMONTH(Data!E2,0)'],
    ['DATEDIF / DAYS', 'DATEDIF(Data!E5,Data!E6,"M")+DAYS(Data!E6,Data!E5)'],
    [
      'IS functions',
      'ISBLANK(Data!F1)&ISNUMBER(Data!C2)&ISTEXT(Data!A2)&ISERROR(1/0)&ISNA(NA())&ISERR(NA())',
    ],
    ['error result', 'Data!A2+1'],
    ['implicit intersection', 'Data!C2:C6*2'],
    ['comparison rules', '("a"<"B")&(1<"0")&("z"<FALSE)'],
  ];
  formulas.forEach(([what, f], i) => {
    const r = i + 2;
    ws.addRow([
      what,
      `=${f}`,
      { formula: f },
      null,
      { formula: `IF(ISERROR(C${r}),ISERROR(D${r}),IFERROR(C${r}=D${r},FALSE))` },
    ]);
  });
  evaluate(wb);
  // copy turbosheet's results into the static column D, then re-evaluate the Match column
  for (let i = 0; i < formulas.length; i++) ws.setCell(i + 2, 4, ws.getValue(i + 2, 3).result);
  evaluate(wb);
  save(
    '16-formulas.xlsx',
    writeXlsx(wb),
    `"Formulas" sheet: ${formulas.length} formulas whose cached values were computed by turbosheet/formula. Values shown before Excel recalculates match after pressing Ctrl+Alt+F9: column D is turbosheet's result as a plain value, so after recalculation column C (Excel) must equal D and the Match column must be TRUE on every row (all 38). Future functions (XOR, IFS, SWITCH, IFNA, XLOOKUP, XMATCH, TEXTJOIN, CONCAT, DAYS) must not show #NAME?. The examples avoid currency and thousands/decimal separators; TEXT date codes (yyyy, mmm) assume an English Excel UI.`,
  );
}

// 17: row / column structure edits; every reference must still point at the same data
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  const sum = wb.addWorksheet('Summary');
  addTable(ws, {
    name: 'Sales',
    ref: 'A1:D6',
    columns: [
      'Item',
      'Qty',
      'Price',
      {
        name: 'Total',
        calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
      },
    ],
    rows: [
      ['a', 1, 10],
      ['b', 2, 20],
      ['c', 3, 30],
      ['d', 4, 40],
      ['e', 5, 50],
    ],
  });
  ws.setCell('F1', { formula: 'SUM(B2:B6)' });
  ws.setCell('F2', { formula: 'B5*C5' });
  ws.setCell('F7', { formula: 'SUM($C$2:$C$6)' });
  ws.setCell('H1', 'merged');
  ws.mergeCells('H1:I2');
  ws.setCell('A9', 'dup');
  ws.setCell('B9', 7);
  ws.setCell('C9', { formula: 'B9*2' });
  ws.addRow(['tail', 100]);
  addConditionalFormatting(ws, 'B2:B6', [{ type: 'expression', formulae: ['B2>=3'], style: bold }]);
  addDataValidation(ws, { sqref: 'C2:C6', type: 'whole', operator: 'greaterThan', formula1: 'B2' });
  addNote(ws, 'A4', 'note on item c');
  addDefinedName(wb, 'QtyCol', 'Data!$B$2:$B$6');
  sum.setCell('A1', { formula: 'SUM(Data!B2:B6)' });
  sum.setCell('A2', { formula: 'Data!B4' });
  sum.setCell('A3', { formula: 'SUM(QtyCol)' });
  sum.setCell('A4', { formula: "'Data'!C6" });
  addChart(sum, {
    type: 'column',
    title: 'Qty',
    series: [{ categories: 'Data!$A$2:$A$6', values: 'Data!$B$2:$B$6' }],
    anchor: 'C1:J15',
  });

  insertRows(ws, 4, [['ins', 6, 60]]); // inside the table: Total filled
  deleteRows(ws, 3); // item b
  insertColumns(ws, 'C', 1); // inside the table: new column "Column5"
  duplicateRow(ws, 9, 2, true); // copies of "dup" in rows 10-11, formula translated
  spliceRows(ws, 12, 1, ['spliced', 1]); // replaces the moved "tail" row
  save(
    '17-structure-edits.xlsx',
    writeXlsx(wb),
    `Rows/columns inserted, deleted, duplicated and spliced after building the sheet. "Data": table Sales is A1:E6 with columns Item, Qty, Column5, Price, Total (items a, ins, c, d, e; Total = Qty*Price on every row, also "ins"). G1 = 19 (SUM(B2:B6)), G2 = 160 (B5*D5 → item d), G7 = 190 (prices); merged block I1:J2; bold CF on Qty >= 3; note on A4 (item c); rows 9-11 "dup" 7 with D = 14 (D10 = B10*2, D11 = B11*2); row 12 "spliced". "Summary": A1 = 19, A2 = 3 (item c's Qty), A3 = 19 (defined name QtyCol), A4 = 50; the chart shows a, ins, c, d, e. No repair prompt.`,
  );
}

// 18: table rows appended, outline groups (collapsed)
{
  const wb = new Workbook();
  const ws = wb.addWorksheet('Tables');
  addTable(ws, {
    name: 'Orders',
    columns: [
      { name: 'Item', width: 16, totalsRowLabel: 'Total' },
      { name: 'Qty', totalsRowFunction: 'sum' },
      { name: 'Price' },
      {
        name: 'Amount',
        calculatedColumnFormula: 'Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]',
      },
    ],
    rows: [
      ['a', 1, 10],
      ['b', 2, 20],
    ],
    totalsRow: true,
  });
  ws.setCell('A8', 'below the table');
  addTableRows(ws, 'Orders', [['c', 3, 30], { Item: 'd', Qty: 4, Price: 40 }]);
  const g = wb.addWorksheet('Outline');
  g.addRow(['Region', 'Q1', 'Q2', 'Q3', 'Total']);
  for (const r of ['N', 'S', 'E', 'W'])
    g.addRow([r, 1, 2, 3, { formula: `SUM(B${g.rowCount + 1}:D${g.rowCount + 1})` }]);
  g.addRow([
    'All',
    { formula: 'SUM(B2:B5)' },
    { formula: 'SUM(C2:C5)' },
    { formula: 'SUM(D2:D5)' },
    { formula: 'SUM(E2:E5)' },
  ]);
  groupRows(g, 2, 5);
  groupRows(g, 3, 4, { collapsed: true });
  groupColumns(g, 'B', 'D', { collapsed: true });
  save(
    '18-table-rows-outline.xlsx',
    writeXlsx(wb),
    `"Tables": table Orders A1:D6 with items a, b, c, d (two appended), Amount = Qty*Price on every row (90 and 160 for c and d), totals row 6 (Qty 10), column A wider; "below the table" moved to A10. "Outline": rows 2-5 grouped (level 1), rows 3-4 a collapsed level-2 group (hidden, + button on row 5); columns B-D grouped and collapsed (hidden, + button above E). Expanding shows the data. No repair prompt.`,
  );
}

// 19: Excel-like polish: default font, wrap for line breaks, hyperlink look, ignored errors,
// page breaks, whole-sheet style (unlocked) + protection
{
  const wb = new Workbook({ defaultFont: { name: 'Arial', size: 10 } });
  const ws = wb.addWorksheet('Polish');
  ws.addRow(['ID', 'Note', 'Link']);
  for (let i = 1; i <= 60; i++)
    ws.addRow([String(10000 + i), i === 1 ? 'first line\nsecond line' : `row ${i}`]);
  ws.setCell('C2', { text: 'example.com', hyperlink: 'https://example.com' });
  ws.ignoredErrors = [{ sqref: 'A2:A61', numberStoredAsText: true }];
  ws.rowBreaks = [30];
  setSheetStyle(ws, { protection: { locked: false } });
  ws.setStyle('A1', { ...ws.getStyle('A1'), protection: { locked: true } });
  ws.protection = { sheet: true };
  save(
    '19-polish.xlsx',
    writeXlsx(wb),
    `"Polish": whole sheet in Arial 10 (Home > Font shows Arial for empty cells too); B2 shows two lines (wrapped); C2 is blue underlined and opens example.com; column A IDs have no green "number stored as text" triangles; Page Break Preview shows a manual break after row 30; the sheet is protected but every cell except A1 can be edited (A1 refuses edits). No repair prompt.`,
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
