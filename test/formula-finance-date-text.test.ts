// Financial, date / time, text and information functions (#109, #110). The expected values were
// checked against Excel for Mac (the same workbook evaluated by Excel, compared cell by cell),
// except the marked ones, where Excel for Mac follows its locale / Mac Roman and turbosheet
// follows en-US / Windows-1252 (see docs/formula.md).

import { describe, expect, it } from 'vitest';
import { evaluate } from '../src/formula/index.js';
import { Workbook } from '../src/index.js';

const CASES: [string, unknown][] = [
  ['PMT(0.08/12,10,10000)', -1037.03208935916],
  ['PMT(0.08/12,10,10000,0,1)', -1030.16432717798],
  ['PMT(0,12,1200)', -100],
  ['PMT(0.05,0,100)', { error: '#NUM!' }],
  ['FV(0.06/12,10,-200,-500,1)', 2581.40337406014],
  ['FV(0,10,-100)', 1000],
  ['PV(0.08/12,240,500)', -59777.1458511878],
  ['PV(0.05,10,0,1000)', -613.913253540759],
  ['IPMT(0.1/12,1,36,8000)', -66.6666666666667],
  ['IPMT(0.1,3,3,8000)', -292.447129909366],
  ['IPMT(0.1/12,2,36,8000,0,1)', -64.5332989183138],
  ['IPMT(0.1/12,1,36,8000,0,1)', 0],
  ['IPMT(0.1,4,3,8000)', { error: '#NUM!' }],
  ['PPMT(0.1/12,1,24,2000)', -75.6231860083666],
  ['PPMT(0.08,10,10,200000)', -27598.0534624214],
  ['NPER(0.12/12,-100,-1000,10000,1)', 59.6738656742946],
  ['NPER(0.01,-100,1000)', 10.5886444594232],
  ['NPER(0,-100,1000)', 10],
  ['NPV(0.1,-10000,3000,4200,6800)', 1188.44341233522],
  ['NPV(0.08,A2:A4)+A1', 1776.66006198242],
  ['IRR(A1:A4)', 0.163405600688989],
  ['IRR({-70000,12000,15000,18000,21000})', -0.0212448482734109],
  ['IRR({-70000,12000,15000},-0.1)', -0.443506941334741],
  ['IRR({1,2,3})', { error: '#NUM!' }],
  ['XNPV(0.09,B1:B5,C1:C5)', 2086.64760203153],
  ['XIRR(B1:B5,C1:C5)', 0.373362533518831],
  ['XIRR(B1:B5,C1:C5,0.5)', 0.373362533518832],
  ['RATE(48,-200,8000)', 0.00770147248823321],
  ['RATE(48,-200,8000)*12', 0.0924176698587985],
  ['RATE(10,-1000,10000,0,1)', 2.86387415473983e-12],
  ['RATE(4,100,1000)', { error: '#NUM!' }],
  ['SLN(30000,7500,10)', 2250],
  ['DB(1000000,100000,6,1,7)', 186083.333333333],
  ['DB(1000000,100000,6,7,7)', 15845.0984738481],
  ['DB(1000000,100000,6,3)', 147939.759],
  ['DDB(2400,300,10,1)', 480],
  ['DDB(2400,300,10,10)', 22.1225472000001],
  ['DDB(2400,300,120,1,2)', 40],
  ['DDB(2400,300,10,2,1.5)', 306],
  ['EFFECT(0.0525,4)', 0.0535426673707582],
  ['EFFECT(0.05,0)', { error: '#NUM!' }],
  ['NOMINAL(0.053543,4)', 0.052500319868356],
  ['TIME(12,30,0)', 0.520833333333333],
  ['TIME(27,0,0)', 0.125],
  ['TIME(0,-1,0)', { error: '#NUM!' }],
  ['TIME(0,750,0)', 0.520833333333333],
  ['HOUR(0.75)', 18],
  ['HOUR("6:45 PM")', 18],
  ['MINUTE(0.7852)', 50],
  ['SECOND(TIME(1,2,59.6))', 59],
  ['SECOND(0.5000057)', 0],
  ['DATEVALUE("8/22/2011")', 40777], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['DATEVALUE("22-MAY-2011")', 40685],
  ['DATEVALUE(1)', { error: '#VALUE!' }],
  ['TIMEVALUE("2:24 AM")', 0.1],
  ['TIMEVALUE("22-Aug-2011 6:35 AM")', 0.274305555554747],
  ['WEEKNUM(DATE(2012,3,9))', 10],
  ['WEEKNUM(DATE(2012,3,9),2)', 11],
  ['WEEKNUM(DATE(2024,1,1),1)', 1],
  ['WEEKNUM(DATE(2024,12,31),21)', 1],
  ['WEEKNUM(DATE(2021,1,3),15)', 1],
  ['WEEKNUM(DATE(2024,1,1),3)', { error: '#NUM!' }],
  ['ISOWEEKNUM(DATE(2012,3,9))', 10],
  ['ISOWEEKNUM(DATE(2021,1,1))', 53],
  ['ISOWEEKNUM(DATE(2020,12,31))', 53],
  ['NETWORKDAYS(DATE(2024,12,20),DATE(2025,1,10))', 16],
  ['NETWORKDAYS(DATE(2024,12,20),DATE(2025,1,10),D1:D2)', 14],
  ['NETWORKDAYS(DATE(2025,1,10),DATE(2024,12,20))', -16],
  ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),7)', 23],
  ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),11)', 27],
  ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0000110")', 23],
  ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),2,D1:D2)', 22],
  ['WORKDAY(DATE(2024,12,20),10,D1:D2)', 45664],
  ['WORKDAY(DATE(2025,1,10),-10,D1:D2)', 45652],
  ['WORKDAY.INTL(DATE(2024,1,1),30,"0000011")', 45334],
  ['WORKDAY.INTL(DATE(2024,1,1),10,7)', 45306],
  ['WORKDAY.INTL(DATE(2024,1,1),10,"1111111")', { error: '#VALUE!' }],
  ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))', 0.580555555555556],
  ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)', 0.576502732240437],
  ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)', 0.578082191780822],
  ['YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),0)', 1],
  ['YEARFRAC(DATE(2011,12,15),DATE(2013,3,1),1)', 1.20985401459854],
  ['YEARFRAC(DATE(2011,6,15),DATE(2012,3,1),1)', 0.710382513661202],
  ['YEARFRAC(DATE(2012,1,31),DATE(2012,3,31),4)', 0.166666666666667],
  ['YEARFRAC(DATE(2012,1,31),DATE(2012,3,31),0)', 0.166666666666667],
  ['YEARFRAC(DATE(2012,7,30),DATE(2012,1,1),2)', 0.586111111111111],
  ['CHAR(65)', 'A'],
  ['CHAR(128)', '€'], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['CHAR(169)', '©'],
  ['CODE("A")', 65],
  ['CODE("€")', 128], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['CODE("é")', 233], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['UNICHAR(9731)', '☃'],
  ['UNICODE("☃")', 9731],
  ['LEN(CLEAN(CHAR(9)&"x"&CHAR(10)))', 1],
  ['FIXED(1234.567,1)', '1,234.6'],
  ['FIXED(1234.567,-1)', '1,230'],
  ['FIXED(-1234.567,-1,TRUE)', '-1230'],
  ['FIXED(44.332)', '44.33'],
  ['FIXED(2.5,0)', '3'],
  ['DOLLAR(1234.567,2)', '$1,234.57'], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['DOLLAR(-1234.567,-2)', '($1,200)'], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['DOLLAR(0.125,2)', '$0.13'], // en-US / Windows (Excel for Mac: locale currency, Mac Roman, stricter date text)
  ['NUMBERVALUE("2.500,27",",",".")', 2500.27],
  ['NUMBERVALUE("3.5%")', 0.035],
  ['NUMBERVALUE(" 1 234 ")', 1234],
  ['NUMBERVALUE("1,2,3",",",".")', { error: '#VALUE!' }],
  ['T("x")', 'x'],
  ['T(5)', ''],
  ['T(E3)', ''],
  ['TEXTBEFORE("Red riding hood","d")', 'Re'],
  ['TEXTBEFORE("Red riding hood","d",2)', 'Red ri'],
  ['TEXTBEFORE("Red riding hood","d",-1)', 'Red riding hoo'],
  ['TEXTBEFORE("Red riding hood","D",1,1)', 'Re'],
  ['TEXTBEFORE("abc","z")', { error: '#N/A' }],
  ['TEXTBEFORE("abc","z",1,0,0,"none")', 'none'],
  ['TEXTBEFORE("a-b","-",2,0,1)', 'a-b'],
  ['TEXTAFTER("Red riding hood","d")', ' riding hood'],
  ['TEXTAFTER("Red riding hood","d",-2)', 'ing hood'],
  ['TEXTAFTER("a,b;c",{",",";"},2)', 'c'],
  ['ROWS(TEXTSPLIT("a,b,c",","))', 1],
  ['COLUMNS(TEXTSPLIT("a,b,c",","))', 3],
  ['INDEX(TEXTSPLIT("a,b;c,d,e",",",";"),2,3)', 'e'],
  ['INDEX(TEXTSPLIT("a,b;c",",",";"),2,2)', { error: '#N/A' }],
  ['COLUMNS(TEXTSPLIT("a,,b",",",,TRUE))', 2],
  ['N(5)', 5],
  ['N("x")', 0],
  ['N(TRUE)', 1],
  ['N(DATE(2024,1,1))', 45292],
  ['N(E2)', 0],
  ['TYPE(1)', 1],
  ['TYPE("a")', 2],
  ['TYPE(TRUE)', 4],
  ['TYPE(#N/A)', 16],
  ['TYPE({1,2})', 64],
  ['TYPE(A1:A2)', 16],
  ['ISEVEN(-2.5)', true],
  ['ISODD(3)', true],
  ['ISODD(-3.9)', true],
  ['ISEVEN("x")', { error: '#VALUE!' }],
  ['ISREF(A1)', true],
  ['ISREF(5)', false],
  ['ISREF(A1:B2)', true],
  ['ISFORMULA(E1)', true],
  ['ISFORMULA(E2)', false],
  ['ERROR.TYPE(#N/A)', 7],
  ['ERROR.TYPE(1/0)', 2],
  ['ERROR.TYPE(5)', { error: '#N/A' }],
  ['ADDRESS(2,3)', '$C$2'],
  ['ADDRESS(2,3,2)', 'C$2'],
  ['ADDRESS(2,3,2,FALSE)', 'R2C[3]'],
  ['ADDRESS(2,3,4,TRUE,"My Sheet")', "'My Sheet'!C2"],
  ['ADDRESS(2,3,1,FALSE,"S")', 'S!R2C3'],
  ['ADDRESS(1,27,3)', '$AA1'],
  ['HYPERLINK("http://x.com","site")', 'site'],
  ['HYPERLINK("http://x.com")', 'http://x.com'],
];

describe('financial, date, text and information functions (Excel-checked)', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const set = (r: number, c: number, v: number | Date) => {
    ws.setCell(r, c, v);
  };
  [-10000, 3000, 4200, 6800].forEach((v, i) => {
    set(i + 1, 1, v);
  });
  [-10000, 2750, 4250, 3250, 2750].forEach((v, i) => {
    set(i + 1, 2, v);
  });
  ['2008-01-01', '2008-03-01', '2008-10-30', '2009-02-15', '2009-04-01'].forEach((d, i) => {
    set(i + 1, 3, new Date(`${d}T00:00:00Z`));
  });
  ['2024-12-25', '2025-01-01'].forEach((d, i) => {
    set(i + 1, 4, new Date(`${d}T00:00:00Z`));
  });
  ws.setCell('E1', { formula: 'A1' });
  ws.setCell('E2', 'text');
  ws.setCell('E3', true);
  CASES.forEach(([f], i) => {
    ws.setCell(i + 1, 8, { formula: f });
  });
  const report = evaluate(wb);

  it('evaluates without unsupported functions or errors', () => {
    expect(report.unsupported).toEqual([]);
    expect(report.errors).toEqual([]);
  });

  it.each(CASES.map((c, i) => [c[0], c[1], i] as const))('%s = %j', (_f, want, i) => {
    const got = (ws.getValue(i + 1, 8) as { result: unknown }).result;
    if (typeof want === 'number') expect(got).toBeCloseTo(want, 6);
    else expect(got).toEqual(want);
  });
});
