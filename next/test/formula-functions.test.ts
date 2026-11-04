// Built-in functions of turbosheet/formula. Most cases are the examples from Microsoft's
// Excel function documentation (support.microsoft.com), plus edge cases.

import { describe, expect, it } from 'vitest';
import { type CellValue, Workbook } from '../src/core/index.js';
import { evaluateFormula, type FormulaResult, supportedFunctions } from '../src/formula/index.js';

/** A workbook whose first sheet holds `rows` from A1 (null = empty cell). */
function sheet(rows: CellValue[][] = []): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Sheet1');
  rows.forEach((row, r) => {
    row.forEach((v, c) => {
      if (v !== null && v !== undefined) ws.setCell(r + 1, c + 1, v);
    });
  });
  return wb;
}

function check(wb: Workbook, cases: [string, FormulaResult | FormulaResult[][]][]): void {
  for (const [f, want] of cases) {
    const got = evaluateFormula(wb, f);
    if (typeof want === 'number' && typeof got === 'number' && !Number.isInteger(want)) {
      expect(got, f).toBeCloseTo(want, 6);
    } else expect(got, f).toEqual(want);
  }
}

const E = (error: string): FormulaResult => ({ error }) as FormulaResult;

describe('math', () => {
  it('SUM / PRODUCT / AVERAGE / MIN / MAX', () => {
    const wb = sheet([[-5], [15], [30], ['5'], [true], [null], [{ formula: '1/0' }]]);
    check(wb, [
      ['SUM(A1:A3)', 40],
      ['SUM(A2:A3,15)', 60],
      ['SUM(A1:A6)', 40], // text and TRUE in ranges are ignored
      ['SUM("5",2,TRUE)', 8], // direct arguments are converted
      ['SUM("x")', E('#VALUE!')],
      ['SUM(A1:A7)', E('#DIV/0!')],
      ['SUM()', 0],
      ['SUM({1,2;3,4})', 10],
      ['PRODUCT(5,15,30)', 2250],
      ['PRODUCT(A1:A3)', -2250],
      ['PRODUCT(A1:A3,2)', -4500],
      ['PRODUCT(Z1:Z3)', 0],
      ['AVERAGE(10,7,9,27,2)', 11],
      ['AVERAGE(A1:A6)', 40 / 3],
      ['AVERAGE(Z1:Z9)', E('#DIV/0!')],
      ['MIN(10,7,9,27,2)', 2],
      ['MAX(10,7,9,27,2)', 27],
      ['MAX(A1:A3,30)', 30],
      ['MIN(A1:A6)', -5],
      ['MAX(Z1:Z9)', 0],
    ]);
  });

  it('COUNT / COUNTA / COUNTBLANK', () => {
    // Excel docs data: Data, 12/8/2008, (blank), 19, 22.24, TRUE, #DIV/0!
    const wb = sheet([
      ['Data'],
      [new Date(Date.UTC(2008, 11, 8))],
      [null],
      [19],
      [22.24],
      [true],
      [{ error: '#DIV/0!' }],
      [{ formula: '""' }],
    ]);
    check(wb, [
      ['COUNT(A2:A7)', 3],
      ['COUNT(A5:A7)', 1],
      ['COUNT(A2:A7,2)', 4],
      ['COUNT(1,"2","x",TRUE)', 3],
      ['COUNTA(A1:A7)', 6],
      ['COUNTA(A1:A8)', 7], // "" from a formula counts
      ['COUNTA(1,,"")', 2],
      ['COUNTBLANK(A1:A8)', 2], // the blank and the ""
      ['COUNTBLANK(B1:B10)', 10],
    ]);
  });

  it('SUMIF (Excel docs examples)', () => {
    const wb = sheet([
      ['Property Value', 'Commission', 'Data'],
      [100000, 7000, 250000],
      [200000, 14000],
      [300000, 21000],
      [400000, 28000],
      [],
      ['Category', 'Food', 'Sales'],
      ['Vegetables', 'Tomatoes', 2300],
      ['Vegetables', 'Celery', 5500],
      ['Fruits', 'Oranges', 800],
      [null, 'Butter', 400],
      ['Vegetables', 'Carrots', 4200],
      ['Fruits', 'Apples', 1200],
    ]);
    check(wb, [
      ['SUMIF(A2:A5,">160000",B2:B5)', 63000],
      ['SUMIF(A2:A5,">160000")', 900000],
      ['SUMIF(A2:A5,300000,B2:B5)', 21000],
      ['SUMIF(A2:A5,">"&C2,B2:B5)', 49000],
      ['SUMIF(A8:A13,"Fruits",C8:C13)', 2000],
      ['SUMIF(A8:A13,"Vegetables",C8:C13)', 12000],
      ['SUMIF(B8:B13,"*es",C8:C13)', 4300],
      ['SUMIF(A8:A13,"",C8:C13)', 400],
      ['SUMIF(A8:A13,"fruits",C8)', 2000], // sum_range takes the criteria range's shape
      ['SUMIF(A:A,"Fruits",C:C)', 2000],
      ['SUMIF(A:A,"<>Fruits",C:C)', 262400], // includes A2 (C2 = 250000) and the blank A11
      ['SUMIF(A2:A5,"<>"&A3,B2:B5)', 56000],
    ]);
  });

  it('SUMIFS / COUNTIF / COUNTIFS / AVERAGEIF / AVERAGEIFS (Excel docs examples)', () => {
    const wb = sheet([
      ['Qty', 'Product', 'Salesperson'],
      [5, 'Apples', 'Tom'],
      [4, 'Apples', 'Sarah'],
      [15, 'Artichokes', 'Tom'],
      [3, 'Artichokes', 'Sarah'],
      [22, 'Bananas', 'Tom'],
      [12, 'Bananas', 'Sarah'],
      [10, 'Carrots', 'Tom'],
      [33, 'Carrots', 'Sarah'],
    ]);
    check(wb, [
      ['SUMIFS(A2:A9,B2:B9,"=A*",C2:C9,"Tom")', 20],
      ['SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")', 30],
      ['SUMIFS(A2:A9,B2:B9,"Apples")', 9],
      ['SUMIFS(A2:A9,B2:B8,"Apples")', E('#VALUE!')], // shapes must agree
      ['AVERAGEIFS(A2:A9,C2:C9,"Sarah",A2:A9,">3")', 49 / 3],
      ['AVERAGEIFS(A2:A9,C2:C9,"Nobody")', E('#DIV/0!')],
      ['COUNTIFS(C2:C9,"Tom",A2:A9,">=10")', 3],
    ]);
    const fruit = sheet([
      ['Data', 'Data'],
      ['apples', 32],
      ['oranges', 54],
      ['peaches', 75],
      ['apples', 86],
    ]);
    check(fruit, [
      ['COUNTIF(A2:A5,"apples")', 2],
      ['COUNTIF(A2:A5,A4)', 1],
      ['COUNTIF(A2:A5,A2)+COUNTIF(A2:A5,A3)', 3],
      ['COUNTIF(B2:B5,">55")', 2],
      ['COUNTIF(B2:B5,"<>"&B4)', 3],
      ['COUNTIF(B2:B5,">=32")-COUNTIF(B2:B5,">85")', 3],
      ['COUNTIF(A2:A5,"*")', 4],
      ['COUNTIF(A2:A5,"?????es")', 2],
      ['COUNTIF(A2:A5,"APPLES")', 2],
      ['COUNTIF(B2:B5,54)', 1],
      ['COUNTIF(B2:B5,"54")', 1],
      ['COUNTIF(A1:B5,"<>")', 10],
      ['COUNTIF(A1:C5,"")', 5], // blank cells of column C
      ['COUNTIF(C:C,"")', 1048576],
      ['COUNTIF(A:A,"<>apples")', 1048574],
    ]);
    const yes = sheet([
      ['Salesperson', 'Q1', 'Q2', 'Q3'],
      ['Davidoski', 'Yes', 'No', 'No'],
      ['Burke', 'Yes', 'Yes', 'No'],
      ['Sundaram', 'Yes', 'Yes', 'Yes'],
      ['Levitan', 'No', 'Yes', 'Yes'],
    ]);
    check(yes, [
      ['COUNTIFS(B2:D2,"=Yes")', 1],
      ['COUNTIFS(B2:B5,"=Yes",C2:C5,"=Yes")', 2],
      ['COUNTIFS(B5:D5,"=Yes",B3:D3,"=Yes")', 1],
      ['COUNTIFS(B2:B5,"Yes",C2:C5)', E('#VALUE!')],
    ]);
    const prop = sheet([
      ['Property Value', 'Commission'],
      [100000, 7000],
      [200000, 14000],
      [300000, 21000],
      [400000, 28000],
    ]);
    check(prop, [
      ['AVERAGEIF(B2:B5,"<23000")', 14000],
      ['AVERAGEIF(A2:A5,"<250000")', 150000],
      ['AVERAGEIF(A2:A5,"<95000")', E('#DIV/0!')],
      ['AVERAGEIF(A2:A5,">250000",B2:B5)', 24500],
    ]);
  });

  it('criteria: booleans, errors, dates, numbers stored as text', () => {
    const wb = sheet([
      [true, 1],
      [false, 2],
      ['TRUE', 4],
      [{ error: '#N/A' }, 8],
      [new Date(Date.UTC(2024, 0, 15)), 16],
      ['1', 32],
      [1, 64],
    ]);
    check(wb, [
      ['SUMIF(A1:A7,TRUE,B1:B7)', 1],
      ['SUMIF(A1:A7,"TRUE",B1:B7)', 1],
      ['SUMIF(A1:A7,"#N/A",B1:B7)', 8],
      ['SUMIF(A1:A7,">=1/1/2024",B1:B7)', 16],
      ['SUMIF(A1:A7,"=2024-01-15",B1:B7)', 16],
      ['SUMIF(A1:A7,1,B1:B7)', 96],
      ['SUMIF(A1:A7,"<5",B1:B7)', 64],
      ['COUNTIF(A1:A7,"~*")', 0],
    ]);
  });

  it('SUMPRODUCT', () => {
    const wb = sheet([
      [3, 4, 2, 7],
      [8, 6, 6, 7],
      [1, 9, 5, 3],
      ['x', null, 1, 1],
    ]);
    check(wb, [
      ['SUMPRODUCT(A1:B3,C1:D3)', 156],
      ['SUMPRODUCT(A1:B4,C1:D4)', 156], // text and blanks count as 0
      ['SUMPRODUCT({1,2,3})', 6],
      ['SUMPRODUCT(A1:A3,C1:D3)', E('#VALUE!')],
      ['SUMPRODUCT((A1:A3>2)*(B1:B3))', 10],
      ['SUMPRODUCT(--(A1:A3>2))', 2],
      ['SUMPRODUCT(A1:A2,{1;#DIV/0!})', E('#DIV/0!')],
    ]);
  });

  it('ROUND family, INT, MOD, ABS, POWER, SQRT', () => {
    check(sheet(), [
      ['ROUND(2.15,1)', 2.2],
      ['ROUND(2.149,1)', 2.1],
      ['ROUND(-1.475,2)', -1.48],
      ['ROUND(21.5,-1)', 20],
      ['ROUND(626.3,-3)', 1000],
      ['ROUND(1.98,-1)', 0],
      ['ROUND(-50.55,-2)', -100],
      ['ROUND(1.005,2)', 1.01],
      ['ROUND(2.5,0)', 3],
      ['ROUND(-2.5,0)', -3],
      ['ROUND("3.7",0)', 4],
      ['ROUNDUP(3.2,0)', 4],
      ['ROUNDUP(76.9,0)', 77],
      ['ROUNDUP(3.14159,3)', 3142 / 1000],
      ['ROUNDUP(-3.14159,1)', -3.2],
      ['ROUNDUP(31415.92654,-2)', 31500],
      ['ROUNDUP(0.1+0.2,1)', 0.3],
      ['ROUNDDOWN(3.2,0)', 3],
      ['ROUNDDOWN(76.9,0)', 76],
      ['ROUNDDOWN(3.14159,3)', 3.141],
      ['ROUNDDOWN(-3.14159,1)', -3.1],
      ['ROUNDDOWN(31415.92654,-2)', 31400],
      ['INT(8.9)', 8],
      ['INT(-8.9)', -9],
      ['MOD(3,2)', 1],
      ['MOD(-3,2)', 1],
      ['MOD(3,-2)', -1],
      ['MOD(-3,-2)', -1],
      ['MOD(5,0)', E('#DIV/0!')],
      ['ABS(2)', 2],
      ['ABS(-2)', 2],
      ['ABS("x")', E('#VALUE!')],
      ['POWER(5,2)', 25],
      ['POWER(98.6,3.2)', 2401077.22206958],
      ['POWER(4,5/4)', 5.65685424949238],
      ['POWER(-1,0.5)', E('#NUM!')],
      ['SQRT(16)', 4],
      ['SQRT(-16)', E('#NUM!')],
      ['SQRT(ABS(-16))', 4],
    ]);
  });
});

describe('logical', () => {
  it('IF / IFS / IFERROR / IFNA', () => {
    const wb = sheet([
      [1500, 900, 210, 35, 93],
      [500, 900, 55, 0, 85],
      [500, 925, '', 23, 69],
    ]);
    check(wb, [
      ['IF(A1>B1,"Over Budget","Within Budget")', 'Over Budget'],
      ['IF(A2>B2,"Over Budget","Within Budget")', 'Within Budget'],
      ['IF(A1="Yes",1,2)', 2],
      ['IF(A1>B1,C1-D1)', 175],
      ['IF(A2>B2,C1)', false],
      ['IF("TRUE",1,2)', 1],
      ['IF("x",1,2)', E('#VALUE!')],
      ['IF(1/0,1,2)', E('#DIV/0!')],
      ['IF(TRUE,1,1/0)', 1], // lazy
      ['IFS(E1>89,"A",E1>79,"B",E1>69,"C",E1>59,"D",TRUE,"F")', 'A'],
      ['IFS(E2>89,"A",E2>79,"B")', 'B'],
      ['IFS(E3>89,"A",E3>79,"B")', E('#N/A')],
      ['IFERROR(C1/D1,"Error in calculation")', 6],
      ['IFERROR(C2/D2,"Error in calculation")', 'Error in calculation'],
      ['IFERROR(C3/D3,"Error in calculation")', 'Error in calculation'],
      ['IFERROR(1,1/0)', 1],
      ['IFNA(NA(),"none")', 'none'],
      ['IFNA(1/0,"none")', E('#DIV/0!')],
      ['IFNA(VLOOKUP("x",A1:B3,2,FALSE),"Not found")', 'Not found'],
    ]);
  });

  it('AND / OR / XOR / NOT / SWITCH / TRUE / FALSE', () => {
    const wb = sheet([
      [50, 'x', true, null],
      [101, 'y', false, null],
    ]);
    check(wb, [
      ['AND(TRUE,TRUE)', true],
      ['AND(TRUE,FALSE)', false],
      ['AND(2+2=4,2+3=5)', true],
      ['AND(A1>1,A1<100)', true],
      ['AND(C1:C2)', false],
      ['AND(B1:B2)', E('#VALUE!')], // no logical values
      ['AND(C1:D2,TRUE)', false],
      ['AND("TRUE",1)', true],
      ['AND("x")', E('#VALUE!')],
      ['AND(1/0,TRUE)', E('#DIV/0!')],
      ['OR(A1>1,A1<100)', true],
      ['OR(A2<1,A2>200)', false],
      ['OR(C1:C2)', true],
      ['OR(FALSE,0)', false],
      ['XOR(3>0,2<9)', false],
      ['XOR(3>12,4>6)', false],
      ['XOR(3>0,4>6)', true],
      ['XOR(TRUE,TRUE,TRUE)', true],
      ['NOT(FALSE)', true],
      ['NOT(A1>100)', true],
      ['NOT(0)', true],
      ['SWITCH(WEEKDAY(DATE(2024,1,15)),1,"Sunday",2,"Monday",3,"Tuesday","No match")', 'Monday'],
      ['SWITCH(99,1,"a",2,"b","No match")', 'No match'],
      ['SWITCH(99,1,"a",2,"b")', E('#N/A')],
      ['SWITCH("B","a",1,"b",2)', 2],
      ['SWITCH(1,"1","text",1,"number")', 'number'],
      ['TRUE()', true],
      ['FALSE()', false],
    ]);
  });
});

describe('lookup', () => {
  const parts = (): Workbook =>
    sheet([
      ['ID', 'Name', 'Dept', 'Salary'],
      [101, 'Barnhill', 'Sales', 70000],
      [102, 'Bauer', 'IT', 82000],
      [103, 'Hughes', 'Sales', 64000],
      [105, 'Ito', 'Ops', 59000],
      [109, 'Ratcliffe', 'IT', 91000],
    ]);

  it('VLOOKUP / HLOOKUP', () => {
    check(parts(), [
      ['VLOOKUP(103,A2:D6,2,FALSE)', 'Hughes'],
      ['VLOOKUP(104,A2:D6,2,FALSE)', E('#N/A')],
      ['VLOOKUP(104,A2:D6,2,TRUE)', 'Hughes'],
      ['VLOOKUP(104,A2:D6,2)', 'Hughes'],
      ['VLOOKUP(200,A2:D6,4)', 91000],
      ['VLOOKUP(1,A2:D6,4)', E('#N/A')],
      ['VLOOKUP("bauer",B2:D6,3,FALSE)', 82000],
      ['VLOOKUP("Ba*",B2:D6,2,FALSE)', 'Sales'],
      ['VLOOKUP("?to",B2:D6,2,FALSE)', 'Ops'],
      ['VLOOKUP(103,A2:D6,5,FALSE)', E('#REF!')],
      ['VLOOKUP(103,A2:D6,0,FALSE)', E('#VALUE!')],
      ['VLOOKUP("103",A2:D6,2,FALSE)', E('#N/A')], // text does not match numbers
      ['VLOOKUP(105,A:D,2,FALSE)', 'Ito'],
      ['VLOOKUP(1/0,A2:D6,2,FALSE)', E('#DIV/0!')],
    ]);
    const h = sheet([
      ['Axles', 'Bearings', 'Bolts'],
      [4, 4, 9],
      [5, 7, 10],
      [6, 8, 11],
    ]);
    check(h, [
      ['HLOOKUP("Axles",A1:C4,2,TRUE)', 4],
      ['HLOOKUP("Bearings",A1:C4,3,FALSE)', 7],
      ['HLOOKUP("B",A1:C4,3,TRUE)', 5],
      ['HLOOKUP("Bolts",A1:C4,4)', 11],
      ['HLOOKUP(3,{1,2,3;"a","b","c";"d","e","f"},2,TRUE)', 'c'],
    ]);
  });

  it('MATCH / XMATCH', () => {
    const wb = sheet([
      ['Product', 'Count'],
      ['Bananas', 25],
      ['Oranges', 38],
      ['Apples', 40],
      ['Pears', 41],
    ]);
    check(wb, [
      ['MATCH(39,B2:B5,1)', 2],
      ['MATCH(41,B2:B5,0)', 4],
      ['MATCH(40,B2:B5,-1)', E('#N/A')],
      ['MATCH(50,{60,50,40},-1)', 2],
      ['MATCH(45,{60,50,40},-1)', 2],
      ['MATCH(39,B2:B5)', 2],
      ['MATCH("?ears",A2:A5,0)', 4],
      ['MATCH("oranges",A2:A5,0)', 2],
      ['MATCH(1,A2:B5,0)', E('#N/A')], // not a vector
      ['MATCH("b",{"a","b","c"},0)', 2],
      ['MATCH(10,B2:B5,1)', E('#N/A')],
      ['XMATCH("Apples",A2:A5)', 3],
      ['XMATCH(39,B2:B5,1)', 3],
      ['XMATCH(39,B2:B5,-1)', 2],
      ['XMATCH("*s",A2:A5,2)', 1],
      ['XMATCH("*s",A2:A5,2,-1)', 4],
      ['XMATCH(100,B2:B5)', E('#N/A')],
    ]);
  });

  it('XLOOKUP', () => {
    check(parts(), [
      ['XLOOKUP(105,A2:A6,B2:B6)', 'Ito'],
      ['XLOOKUP(105,A2:A6,B2:D6)', [['Ito', 'Ops', 59000]]],
      ['XLOOKUP("hughes",B2:B6,D2:D6)', 64000],
      ['XLOOKUP(104,A2:A6,B2:B6)', E('#N/A')],
      ['XLOOKUP(104,A2:A6,B2:B6,"Not found")', 'Not found'],
      ['XLOOKUP(104,A2:A6,B2:B6,,-1)', 'Hughes'],
      ['XLOOKUP(104,A2:A6,B2:B6,,1)', 'Ito'],
      ['XLOOKUP("Sales",C2:C6,B2:B6,,0,-1)', 'Hughes'],
      ['XLOOKUP("R*",B2:B6,A2:A6,,2)', 109],
      ['XLOOKUP("Salary",A1:D1,A6:D6)', 91000],
      ['XLOOKUP(105,A2:A6,B2:B5)', E('#VALUE!')],
      ['SUM(XLOOKUP(102,A2:A6,D2:D6):XLOOKUP(105,A2:A6,D2:D6))', 205000],
    ]);
  });

  it('INDEX / CHOOSE / ROW / COLUMN / ROWS / COLUMNS', () => {
    const wb = sheet([
      ['Apples', 'Lemons'],
      ['Bananas', 'Pears'],
      [1, 2],
      [3, 4],
    ]);
    check(wb, [
      ['INDEX(A1:B2,2,2)', 'Pears'],
      ['INDEX(A1:B2,2,1)', 'Bananas'],
      ['INDEX({1,2;3,4},1,2)', 2],
      ['INDEX(A1:A4,3)', 1],
      ['INDEX(A1:B1,2)', 'Lemons'], // a single row: the index is the column
      ['SUM(INDEX(A3:B4,0,2))', 6],
      ['SUM(INDEX(A3:B4,2,0))', 7],
      ['INDEX(A1:B2,3,1)', E('#REF!')],
      ['INDEX(A1:B2,-1,1)', E('#VALUE!')],
      ['ISBLANK(INDEX(A1:C1,3))', true],
      ['INDEX(A1:B4,MATCH("Bananas",A1:A4,0),2)', 'Pears'],
      ['CHOOSE(2,"1st","2nd","3rd","Finished")', '2nd'],
      ['CHOOSE(4,A1,A2,A3,"Finished")', 'Finished'],
      ['SUM(CHOOSE(2,A3:A4,B3:B4))', 6],
      ['CHOOSE(5,1,2)', E('#VALUE!')],
      ['CHOOSE(1,1,1/0)', 1],
      ['ROW(C10)', 10],
      ['ROW(C10:D20)', 10],
      ['COLUMN(D10)', 4],
      ['COLUMN(B:D)', 2],
      ['ROWS(C1:E4)', 4],
      ['COLUMNS(C1:E4)', 3],
      ['ROWS({1,2,3;4,5,6})', 2],
      ['COLUMNS({1,2,3;4,5,6})', 3],
      ['ROWS(1)', 1],
    ]);
    expect(evaluateFormula(wb, 'ROW(A1:A3)', { array: true })).toEqual([[1], [2], [3]]);
    expect(evaluateFormula(wb, 'COLUMN()', { address: 'E7' })).toBe(5);
  });
});

describe('text', () => {
  it('CONCAT / CONCATENATE / TEXTJOIN', () => {
    const wb = sheet([
      ['a', 'b'],
      [null, 'd'],
      [1.5, true],
    ]);
    check(wb, [
      [
        'CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")',
        'The sun will come up tomorrow.',
      ],
      ['CONCAT(A1:B3)', 'abd1.5TRUE'],
      ['CONCAT(A1:B2,"-",1/4)', 'abd-0.25'],
      [
        'CONCATENATE("Stream population for ","species"," is ",32,"/mile.")',
        'Stream population for species is 32/mile.',
      ],
      ['CONCATENATE(A1," ",B1)', 'a b'],
      ['A1&B1', 'ab'],
      ['TEXTJOIN(", ",TRUE,A1:B3)', 'a, b, d, 1.5, TRUE'],
      ['TEXTJOIN(", ",FALSE,A1:B3)', 'a, b, , d, 1.5, TRUE'],
      ['TEXTJOIN("-",TRUE,"x","","y")', 'x-y'],
      ['TEXTJOIN("",TRUE,{"a","b";"c","d"})', 'abcd'],
      ['CONCAT(NA())', E('#N/A')],
      ['LEN(REPT("a",32767)&"b")', E('#VALUE!')],
    ]);
  });

  it('LEFT / RIGHT / MID / LEN / LOWER / UPPER / PROPER / TRIM', () => {
    check(sheet([['Sale Price', 'Sweden', 'Fluid Flow']]), [
      ['LEFT(A1,4)', 'Sale'],
      ['LEFT(B1)', 'S'],
      ['LEFT(B1,0)', ''],
      ['LEFT(B1,99)', 'Sweden'],
      ['LEFT(B1,-1)', E('#VALUE!')],
      ['LEFT(12345,2)', '12'],
      ['RIGHT(A1,5)', 'Price'],
      ['RIGHT("Stock Number")', 'r'],
      ['RIGHT(A1,0)', ''],
      ['MID(C1,1,5)', 'Fluid'],
      ['MID(C1,7,20)', 'Flow'],
      ['MID(C1,20,5)', ''],
      ['MID(C1,0,5)', E('#VALUE!')],
      ['LEN("Phoenix, AZ")', 11],
      ['LEN("")', 0],
      ['LEN("     One   ")', 11],
      ['LEN(Z1)', 0],
      ['LOWER("E. E. Cummings")', 'e. e. cummings'],
      ['UPPER("total")', 'TOTAL'],
      ['PROPER("this is a TITLE")', 'This Is A Title'],
      ['PROPER("2-way street")', '2-Way Street'],
      ['PROPER("76BudGet")', '76Budget'],
      ['TRIM(" First Quarter    Earnings ")', 'First Quarter Earnings'],
    ]);
  });

  it('SUBSTITUTE / REPLACE / FIND / SEARCH / EXACT / REPT / VALUE', () => {
    check(sheet([['Statements', 'Profit Margin', 'margin']]), [
      ['SUBSTITUTE("Sales Data","Sales","Cost")', 'Cost Data'],
      ['SUBSTITUTE("Quarter 1, 2008","1","2",1)', 'Quarter 2, 2008'],
      ['SUBSTITUTE("Quarter 1, 2011","1","2",3)', 'Quarter 1, 2012'],
      ['SUBSTITUTE("aaa","a","b",9)', 'aaa'],
      ['SUBSTITUTE("abc","","x")', 'abc'],
      ['REPLACE("abcdefghijk",6,5,"*")', 'abcde*k'],
      ['REPLACE("2009",3,2,"10")', '2010'],
      ['REPLACE("123456",1,3,"@")', '@456'],
      ['FIND("M","Miriam McGovern")', 1],
      ['FIND("m","Miriam McGovern")', 6],
      ['FIND("M","Miriam McGovern",3)', 8],
      ['FIND("z","abc")', E('#VALUE!')],
      ['FIND("","abc")', 1],
      ['SEARCH("e",A1,6)', 7],
      ['SEARCH(C1,B1)', 8],
      ['REPLACE(B1,SEARCH(C1,B1),6,"Amount")', 'Profit Amount'],
      ['SEARCH("p?o",B1)', 1],
      ['SEARCH("~?","a?b")', 2],
      ['SEARCH("x",B1)', E('#VALUE!')],
      ['EXACT("word","word")', true],
      ['EXACT("Word","word")', false],
      ['EXACT("w ord","word")', false],
      ['REPT("*-",3)', '*-*-*-'],
      ['REPT("-",10)', '----------'],
      ['REPT("x",0)', ''],
      ['VALUE("$1,000")', 1000],
      ['VALUE("16:48:00")-VALUE("12:00:00")', 0.2],
      ['VALUE("abc")', E('#VALUE!')],
      ['VALUE(12)', 12],
      ['VALUE("(5)")', -5],
    ]);
  });

  it('TEXT', () => {
    check(sheet(), [
      ['TEXT(1234.567,"0")', '1235'],
      ['TEXT(1234.567,"0.00")', '1234.57'],
      ['TEXT(1234.567,"#,##0")', '1,235'],
      ['TEXT(1234.567,"#,##0.00")', '1,234.57'],
      ['TEXT(1234.567,"$#,##0.00")', '$1,234.57'],
      ['TEXT(0.285,"0%")', '29%'],
      ['TEXT(0.285,"0.0%")', '28.5%'],
      ['TEXT(DATE(2024,3,5),"yyyy-mm-dd")', '2024-03-05'],
      ['TEXT(DATE(2024,3,5),"dd/mm/yyyy")', '05/03/2024'],
      ['TEXT(DATE(2024,3,5),"mmm d, yyyy")', 'Mar 5, 2024'],
      ['TEXT(DATE(2024,3,5),"dddd, mmmm d")', 'Tuesday, March 5'],
      ['TEXT(DATE(2024,3,5),"m/d/yy")', '3/5/24'],
      ['TEXT(0.75,"hh:mm:ss")', '18:00:00'],
      ['TEXT(0.75,"h:mm AM/PM")', '6:00 PM'],
      ['TEXT(1.5,"[h]:mm")', '36:00'],
      ['TEXT(45306.5104166667,"yyyy-mm-dd hh:mm")', '2024-01-15 12:15'],
      ['TEXT(1234.5,"0.00E+00")', '1.23E+03'],
      ['TEXT(0.00012,"0.0E+0")', '1.2E-4'],
      ['TEXT(-1234.5,"#,##0.00;(#,##0.00)")', '(1,234.50)'],
      ['TEXT(-5,"0")', '-5'],
      ['TEXT(0,"0.00;-0.00;""zero""")', 'zero'],
      ['TEXT(1234567,"#,##0.0,,""M""")', '1.2M'],
      ['TEXT(123456789,"000-00-0000")', '123-45-6789'],
      ['TEXT(5,"000")', '005'],
      ['TEXT(0.5,"#.##")', '.5'],
      ['TEXT(3,"0.0#")', '3.0'],
      ['TEXT(1/3,"General")', '0.333333333333333'],
      ['TEXT("abc","0")', 'abc'],
      ['TEXT("12","0.0")', '12.0'],
      ['TEXT("abc","@!")', 'abc!'],
      ['TEXT(TRUE,"0")', 'TRUE'],
      ['TEXT(1234.5,"[Red]#,##0")', '1,235'],
    ]);
  });
});

describe('date', () => {
  it('DATE / YEAR / MONTH / DAY', () => {
    check(sheet(), [
      ['DATE(2008,1,2)', 39449],
      ['DATE(108,1,2)', 39449], // years below 1900 are offset
      ['DATE(2008,14,2)', 39846], // month overflow → Feb 2 2009
      ['DATE(2008,1,-15)', 39432], // Dec 16 2007
      ['DATE(1900,1,1)', 1],
      ['DATE(1900,3,1)', 61],
      ['DATE(-1,1,1)', E('#NUM!')],
      ['DATE(10000,1,1)', E('#NUM!')],
      ['YEAR(DATE(2008,7,5))', 2008],
      ['YEAR("7/5/2008")', 2008],
      ['MONTH("15-Apr-2011")', 4],
      ['DAY("15-Apr-2011")', 15],
      ['DAY(DATE(2024,2,29))', 29],
      ['YEAR(-1)', E('#NUM!')],
      ['MONTH("x")', E('#VALUE!')],
    ]);
  });

  it('WEEKDAY / EDATE / EOMONTH / DATEDIF / DAYS', () => {
    check(sheet(), [
      ['WEEKDAY("2/14/2008")', 5],
      ['WEEKDAY("2/14/2008",2)', 4],
      ['WEEKDAY("2/14/2008",3)', 3],
      ['WEEKDAY("2/14/2008",11)', 4],
      ['WEEKDAY("2/14/2008",16)', 6],
      ['WEEKDAY("2/14/2008",17)', 5],
      ['WEEKDAY(1)', 1],
      ['WEEKDAY(1,9)', E('#NUM!')],
      ['EDATE("15-Jan-2011",1)', 40589],
      ['EDATE("15-Jan-2011",-1)', 40527],
      ['EDATE("15-Jan-2011",2)', 40617],
      ['EDATE(DATE(2024,1,31),1)', 45351], // Feb 29 2024
      ['EOMONTH("1-Jan-2011",1)', 40602],
      ['EOMONTH("1-Jan-2011",-3)', 40482],
      ['EOMONTH(DATE(2024,1,15),0)', 45322],
      ['DATEDIF("1/1/2001","1/1/2003","Y")', 2],
      ['DATEDIF("6/1/2001","8/15/2002","D")', 440],
      ['DATEDIF("6/1/2001","8/15/2002","YD")', 75],
      ['DATEDIF("6/1/2001","8/15/2002","MD")', 14],
      ['DATEDIF("6/1/2001","8/15/2002","M")', 14],
      ['DATEDIF("6/1/2001","8/15/2002","YM")', 2],
      ['DATEDIF("1/25/2024","3/5/2024","MD")', 9],
      ['DATEDIF("1/1/2003","1/1/2001","Y")', E('#NUM!')],
      ['DATEDIF("1/1/2001","1/1/2003","X")', E('#NUM!')],
      ['DAYS("15-MAR-2021","1-FEB-2021")', 42],
      ['DAYS(DATE(2021,12,31),DATE(2021,1,1))', 364],
      ['DAYS("x",1)', E('#VALUE!')],
    ]);
  });
});

describe('info', () => {
  it('IS functions and NA', () => {
    const wb = sheet([
      [null, 'Gold', 4, { error: '#REF!' }, { error: '#N/A' }, true, { formula: '""' }],
    ]);
    check(wb, [
      ['ISBLANK(A1)', true],
      ['ISBLANK(B1)', false],
      ['ISBLANK(G1)', false],
      ['ISNUMBER(C1)', true],
      ['ISNUMBER("4")', false],
      ['ISNUMBER(B1)', false],
      ['ISTEXT(B1)', true],
      ['ISTEXT(G1)', true],
      ['ISTEXT(C1)', false],
      ['ISNONTEXT(C1)', true],
      ['ISLOGICAL(F1)', true],
      ['ISERROR(D1)', true],
      ['ISERROR(E1)', true],
      ['ISERROR(B1)', false],
      ['ISERR(D1)', true],
      ['ISERR(E1)', false],
      ['ISNA(E1)', true],
      ['ISNA(D1)', false],
      ['ISNA(MATCH(99,C1:C1,0))', true],
      ['NA()', E('#N/A')],
    ]);
    expect(evaluateFormula(wb, 'ISNUMBER(A1:C1)', { array: true })).toEqual([[false, false, true]]);
  });

  it('every function in the coverage table is implemented', () => {
    const names = supportedFunctions();
    for (const f of [
      'SUM SUMIF SUMIFS SUMPRODUCT AVERAGE AVERAGEIF AVERAGEIFS MIN MAX COUNT COUNTA COUNTBLANK',
      'COUNTIF COUNTIFS ROUND ROUNDUP ROUNDDOWN INT MOD ABS POWER SQRT PRODUCT',
      'IF IFS IFERROR IFNA AND OR NOT XOR SWITCH TRUE FALSE',
      'VLOOKUP HLOOKUP XLOOKUP INDEX MATCH XMATCH CHOOSE ROW COLUMN ROWS COLUMNS',
      'CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN LOWER UPPER PROPER TRIM SUBSTITUTE REPLACE',
      'FIND SEARCH TEXT VALUE EXACT REPT',
      'DATE TODAY NOW YEAR MONTH DAY WEEKDAY EDATE EOMONTH DATEDIF DAYS',
      'ISBLANK ISNUMBER ISTEXT ISERROR ISNA ISERR NA',
    ]
      .join(' ')
      .split(' ')) {
      expect(names, f).toContain(f);
    }
  });
});
