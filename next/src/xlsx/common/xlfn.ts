// "Future functions": functions added after Excel 2007 are stored in OOXML formulas with a
// `_xlfn.` prefix (FILTER and SORT with `_xlfn._xlws.`); Excel hides the prefix in its UI and
// shows #NAME? when it is missing. The model is prefix-free: writers add the prefix, readers
// strip it (only when it is the canonical one, so round trips are exact).
// Source: Microsoft, "Issue: An _xlfn. prefix is displayed in front of a formula" and MS-XLSX.
// Every dotted built-in name is a future function except ERROR.TYPE (BETA.DIST, STDEV.S,
// NETWORKDAYS.INTL, ...), so only the undotted ones are listed. LET / LAMBDA parameter names
// (`_xlpm.`) are left as they are.

const UNDOTTED =
  ' ACOT ACOTH AGGREGATE ANCHORARRAY ARABIC ARRAYTOTEXT BASE BITAND BITLSHIFT BITOR BITRSHIFT' +
  ' BITXOR BYCOL BYROW CHOOSECOLS CHOOSEROWS COMBINA CONCAT COT COTH CSC CSCH DAYS DECIMAL DROP' +
  ' ENCODEURL EXPAND FIELDVALUE FILTERXML FORMULATEXT GAMMA GAUSS GROUPBY HSTACK IFNA IFS IMAGE' +
  ' IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET' +
  ' MAKEARRAY MAP MAXIFS MINIFS MUNIT NUMBERVALUE PDURATION PERCENTOF PERMUTATIONA PHI PIVOTBY' +
  ' QUERYSTRING RANDARRAY REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE' +
  ' SHEET SHEETS SINGLE SORTBY STOCKHISTORY SWITCH TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT' +
  ' TOCOL TOROW TRIMRANGE UNICHAR UNICODE UNIQUE VALUETOTEXT VSTACK WEBSERVICE WRAPCOLS WRAPROWS' +
  ' XLOOKUP XMATCH XOR ';

/** The prefix a function name needs in a file ('' for none). */
function prefixOf(name: string): string {
  const n = name.toUpperCase();
  if (n === 'FILTER' || n === 'SORT') return '_xlfn._xlws.';
  return (n.includes('.') ? n !== 'ERROR.TYPE' : UNDOTTED.includes(` ${n} `)) ? '_xlfn.' : '';
}

// quoted strings / sheet names and [structured refs] are skipped; an identifier followed by '('
const ADD = /("(?:[^"]|"")*"|'(?:[^']|'')*'|\[[^\]]*\])|(?<![\w.\\])[A-Za-z][\w.]*(?=\()/g;
const STRIP =
  /("(?:[^"]|"")*"|'(?:[^']|'')*'|\[[^\]]*\])|(?<![\w.\\])(_xlfn\.(?:_xlws\.)?|_xlws\.)([A-Za-z][\w.]*)(?=\()/gi;

/** Add the `_xlfn.` / `_xlfn._xlws.` prefix to future functions (idempotent, case-insensitive). */
export function addXlfn(formula: string): string {
  return formula.indexOf('(') < 0
    ? formula
    : formula.replace(ADD, (m, q?: string) => (q ? m : prefixOf(m) + m));
}

/** Remove canonical future-function prefixes (`_xlfn.XLOOKUP(` → `XLOOKUP(`). */
export function stripXlfn(formula: string): string {
  return !/_xl/i.test(formula)
    ? formula
    : formula.replace(STRIP, (m, q?: string, p?: string, name?: string) =>
        q || prefixOf(name!) !== p!.toLowerCase() ? m : name!,
      );
}
