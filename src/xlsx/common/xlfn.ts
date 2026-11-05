// "Future functions": functions added after Excel 2007 are stored in OOXML formulas with a
// `_xlfn.` prefix (FILTER and SORT with `_xlfn._xlws.`); Excel hides the prefix in its UI and
// shows #NAME? when it is missing. The model is prefix-free: writers add the prefix, readers
// strip it (only when it is the canonical one, so round trips are exact).
// Source: Microsoft, "Issue: An _xlfn. prefix is displayed in front of a formula" and MS-XLSX.
// Every dotted built-in name is a future function except ERROR.TYPE (BETA.DIST, STDEV.S,
// NETWORKDAYS.INTL, ...), so only the undotted ones are listed. LET names are written with the
// `_xlpm.` prefix files need (Excel shows a blank / broken cell without it); readers keep it.

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

// quoted strings / sheet names and [structured refs] are skipped; an identifier followed by '(',
// with any (partial) prefix the caller already wrote
const ADD =
  /("(?:[^"]|"")*"|'(?:[^']|'')*'|\[[^\]]*\])|(?<![\w.\\])(_xlfn\.(?:_xlws\.)?|_xlws\.)?([A-Za-z][\w.]*)(?=\()/gi;
const STRIP =
  /("(?:[^"]|"")*"|'(?:[^']|'')*'|\[[^\]]*\])|(?<![\w.\\])(_xlfn\.(?:_xlws\.)?|_xlws\.)([A-Za-z][\w.]*)(?=\()/gi;

/**
 * Add the `_xlfn.` / `_xlfn._xlws.` prefix to future functions (idempotent, case-insensitive).
 * A partial prefix written by the caller is completed (`_xlfn.SORT(` → `_xlfn._xlws.SORT(`);
 * a prefix on a name that needs none is left as written.
 */
export function addXlfn(formula: string): string {
  let f = formula.includes('[@') ? expandThisRow(formula) : formula;
  if (/LET\s*\(/i.test(f)) f = addXlpm(f);
  if (f.includes('#')) f = anchorArray(f);
  return f.indexOf('(') < 0
    ? f
    : f.replace(ADD, (m, q?: string, p?: string, name?: string) => {
        if (q) return m;
        const want = prefixOf(name!);
        return want && want !== p?.toLowerCase() ? want + name! : m;
      });
}

/**
 * Structured references' `@` shorthand (`T[@Col]`, `T[@[Col]]`, `[@Col]`, `T[@]`) is UI-only:
 * files must say `[#This Row]` (Excel drops a table whose calculated column uses `@`, and shows
 * #REF! in cells). Strings and quoted sheet names are left alone.
 */
export function expandThisRow(f: string): string {
  let out = '';
  for (let i = 0; i < f.length; i++) {
    const ch = f[i]!;
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length && !(f[j] === ch && f[j + 1] !== ch)) j += f[j] === ch ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j;
    } else if (ch === '[' && f[i + 1] === '@') {
      let depth = 0;
      let j = i;
      for (; j < f.length; j++) {
        if (f[j] === "'") j++;
        else if (f[j] === '[') depth++;
        else if (f[j] === ']' && --depth === 0) break;
      }
      const rest = f.slice(i + 2, j);
      out +=
        rest === ''
          ? '[#This Row]'
          : rest.startsWith('[')
            ? `[[#This Row],${rest}]`
            : `[[#This Row],[${rest}]]`;
      i = j;
    } else out += ch;
  }
  return out;
}

/** Remove canonical future-function prefixes (`_xlfn.XLOOKUP(` → `XLOOKUP(`). */
export function stripXlfn(formula: string): string {
  return !/_xl/i.test(formula)
    ? formula
    : formula.replace(STRIP, (m, q?: string, p?: string, name?: string) =>
        q || prefixOf(name!) !== p!.toLowerCase() ? m : name!,
      );
}

const TOKEN =
  /("(?:[^"]|"")*")|('(?:[^']|'')*')|(\[(?:[^[\]']|'.|\[(?:[^\]']|'.)*\])*\])|([A-Za-z_\\][\w.\\]*)|(\s+|[\s\S])/g;

/** Prefix LET names (and their uses inside the LET) with `_xlpm.` (idempotent). */
export function addXlpm(f: string): string {
  const toks: string[] = [];
  const ident: boolean[] = [];
  for (const m of f.matchAll(TOKEN)) {
    toks.push(m[0]);
    ident.push(m[4] !== undefined);
  }
  const next = (i: number): number => {
    let j = i + 1;
    while (j < toks.length && /^\s+$/.test(toks[j]!)) j++;
    return j;
  };
  for (let i = 0; i < toks.length; i++) {
    if (!ident[i] || !/^(?:_xlfn\.)?LET$/i.test(toks[i]!)) continue;
    const open = next(i);
    if (toks[open] !== '(') continue;
    // split the arguments at depth 1; find the closing parenthesis
    const args: number[][] = [[]];
    let depth = 0;
    let close = toks.length;
    for (let j = open + 1; j < toks.length; j++) {
      const t = toks[j]!;
      if (t === '(' || t === '{') depth++;
      else if ((t === ')' || t === '}') && depth > 0) depth--;
      else if (t === ')') {
        close = j;
        break;
      } else if (t === ',' && depth === 0) {
        args.push([]);
        continue;
      }
      if (!/^\s+$/.test(t)) args[args.length - 1]!.push(j);
    }
    const names = new Set<string>();
    for (let a = 0; a + 1 < args.length; a += 2) {
      const arg = args[a]!;
      if (arg.length === 1 && ident[arg[0]!])
        names.add(toks[arg[0]!]!.toLowerCase().replace(/^_xlpm\./, ''));
    }
    for (let j = open + 1; j < close; j++) {
      if (!ident[j] || toks[j - 1] === '!' || toks[next(j)] === '(' || toks[next(j)] === '!')
        continue;
      const t = toks[j]!;
      if (!/^_xlpm\./i.test(t) && names.has(t.toLowerCase())) toks[j] = `_xlpm.${t}`;
    }
  }
  return toks.join('');
}

/**
 * Spill references (`A1#`, `Sheet1!$B$2#`) are stored as `ANCHORARRAY(A1)` in files (Excel
 * removes a formula with `#`); the caller's prefix pass then adds `_xlfn.`.
 */
export function anchorArray(f: string): string {
  return f.replace(ANCHOR, (m, q?: string, sheet?: string, cell?: string) =>
    q ? m : `ANCHORARRAY(${sheet ?? ''}${cell})`,
  );
}

const ANCHOR =
  /("(?:[^"]|"")*")|(?<![\w.$])((?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+)#/g;
