// Generators for the small, non-sheet XLSX parts. Each returns a complete XML document string.

import type { SharedStrings } from '../../core/sharedstrings.js';
import type { StyleRegistry, Xf } from '../../core/styles.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  Color,
  DefinedName,
  Fill,
  Font,
  Style,
} from '../../core/types.js';
import type { SheetState } from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = `${NS_R}/`;
const CT = 'application/vnd.openxmlformats-officedocument.';

/** Per-sheet facts needed by workbook-level parts. */
export interface SheetEntry {
  readonly name: string;
  readonly sheetId: number;
  readonly state: SheetState;
  readonly autoFilter: string | undefined;
}

export interface BookInfo {
  readonly sheets: readonly SheetEntry[];
  readonly date1904: boolean;
  readonly sharedStrings: boolean;
  readonly formulas: boolean;
  readonly activeTab: number;
  readonly creator: string | undefined;
  readonly lastModifiedBy: string | undefined;
  readonly created: Date | undefined;
  readonly modified: Date | undefined;
  /** User defined names (wb.definedNames); autoFilter names are added by workbookXml. */
  readonly definedNames?: readonly DefinedName[] | undefined;
}

export const sheetPath = (i: number): string => `xl/worksheets/sheet${i + 1}.xml`;
export const sheetRelsPath = (i: number): string => `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;

/** [Content_Types].xml; `defaults` / `overrides` come from features (PackageParts). */
export function contentTypesXml(
  sheetCount: number,
  sharedStrings: boolean,
  defaults: ReadonlyMap<string, string> = new Map(),
  overrides: readonly [string, string][] = [],
): string {
  let s = `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`;
  for (const [ext, ct] of defaults) {
    if (ext !== 'rels' && ext !== 'xml') {
      s += `<Default Extension="${escapeAttr(ext)}" ContentType="${escapeAttr(ct)}"/>`;
    }
  }
  s += `<Override PartName="/xl/workbook.xml" ContentType="${CT}spreadsheetml.sheet.main+xml"/>`;
  for (const [path, ct] of overrides) {
    s += `<Override PartName="/${escapeAttr(path)}" ContentType="${escapeAttr(ct)}"/>`;
  }
  for (let i = 0; i < sheetCount; i++) {
    s += `<Override PartName="/${sheetPath(i)}" ContentType="${CT}spreadsheetml.worksheet+xml"/>`;
  }
  s += `<Override PartName="/xl/styles.xml" ContentType="${CT}spreadsheetml.styles+xml"/>`;
  if (sharedStrings) {
    s += `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}spreadsheetml.sharedStrings+xml"/>`;
  }
  return `${s}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="${CT}extended-properties+xml"/></Types>`;
}

export function rootRelsXml(): string {
  return `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL}officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${NS_PKG_REL}/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

function isoDate(d: Date | undefined): string | undefined {
  if (!d) return undefined;
  const t = d.getTime();
  if (t - t !== 0) return undefined;
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function corePropsXml(b: BookInfo): string {
  let s = `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`;
  if (b.creator !== undefined) s += `<dc:creator>${escapeText(b.creator)}</dc:creator>`;
  if (b.lastModifiedBy !== undefined) {
    s += `<cp:lastModifiedBy>${escapeText(b.lastModifiedBy)}</cp:lastModifiedBy>`;
  }
  const c = isoDate(b.created);
  if (c) s += `<dcterms:created xsi:type="dcterms:W3CDTF">${c}</dcterms:created>`;
  const m = isoDate(b.modified);
  if (m) s += `<dcterms:modified xsi:type="dcterms:W3CDTF">${m}</dcterms:modified>`;
  return `${s}</cp:coreProperties>`;
}

export function appPropsXml(b: BookInfo): string {
  const n = b.sheets.length;
  let titles = '';
  for (const sh of b.sheets) titles += `<vt:lpstr>${escapeText(sh.name)}</vt:lpstr>`;
  return `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${n}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`;
}

/** 'My Sheet' → "'My Sheet'" (always quoted, inner apostrophes doubled). */
function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

/** "A1:D10" → "$A$1:$D$10". */
function absRef(ref: string): string {
  return ref.replace(/\$/g, '').replace(/([A-Za-z]+)(\d+)/g, '$$$1$$$2');
}

function definedNameXml(
  name: string,
  sheet: number | undefined,
  hidden: boolean,
  value: string,
  comment?: string,
): string {
  return `<definedName name="${escapeAttr(name)}"${sheet !== undefined ? ` localSheetId="${sheet}"` : ''}${
    comment ? ` comment="${escapeAttr(comment)}"` : ''
  }${hidden ? ' hidden="1"' : ''}>${escapeText(value)}</definedName>`;
}

/**
 * workbook.xml. `extra`: feature elements as [BookSlot, xml] pairs, merged into CT_Workbook
 * order with the built-ins (slot 18 = extLst children, wrapped in one <extLst>).
 */
export function workbookXml(b: BookInfo, extra: readonly [number, string][] = []): string {
  const parts: [number, string][] = [
    [2, b.date1904 ? '<workbookPr date1904="1"/>' : '<workbookPr/>'],
    [
      4,
      `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"${
        b.activeTab > 0 ? ` firstSheet="0" activeTab="${b.activeTab}"` : ''
      }/></bookViews>`,
    ],
  ];
  let sheets = '<sheets>';
  let names = '';
  const user = b.definedNames ?? [];
  for (const n of user) {
    if (
      n.localSheetId !== undefined &&
      !(n.localSheetId >= 0 && n.localSheetId < b.sheets.length)
    ) {
      continue; // scope points at no sheet: Excel would reject the file
    }
    names += definedNameXml(n.name, n.localSheetId, n.hidden === true, n.value, n.comment);
  }
  for (let i = 0; i < b.sheets.length; i++) {
    const sh = b.sheets[i]!;
    sheets += `<sheet name="${escapeAttr(sh.name)}" sheetId="${sh.sheetId}"${
      sh.state !== 'visible' ? ` state="${sh.state}"` : ''
    } r:id="rId${i + 1}"/>`;
    if (
      sh.autoFilter &&
      !user.some((n) => n.name === '_xlnm._FilterDatabase' && n.localSheetId === i)
    ) {
      names += definedNameXml(
        '_xlnm._FilterDatabase',
        i,
        true,
        `${quoteSheet(sh.name)}!${absRef(sh.autoFilter)}`,
      );
    }
  }
  parts.push([5, `${sheets}</sheets>`]);
  if (names) parts.push([8, `<definedNames>${names}</definedNames>`]);
  parts.push([
    9,
    b.formulas ? '<calcPr calcId="191029" fullCalcOnLoad="1"/>' : '<calcPr calcId="191029"/>',
  ]);
  let ext = '';
  for (const e of extra) {
    if (e[0] === 18) ext += e[1];
    else if (e[1]) parts.push(e);
  }
  parts.sort((x, y) => x[0] - y[0]);
  let s = `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`;
  for (const p of parts) s += p[1];
  if (ext) s += `<extLst>${ext}</extLst>`;
  return `${s}</workbook>`;
}

/** Workbook relationships; `extra` = feature rel elements (ids from sheetCount + 3, see BOOK_REL_BASE). */
export function workbookRelsXml(sheetCount: number, sharedStrings: boolean, extra = ''): string {
  let s = `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">`;
  for (let i = 0; i < sheetCount; i++) {
    s += `<Relationship Id="rId${i + 1}" Type="${REL}worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
  }
  s += `<Relationship Id="rId${sheetCount + 1}" Type="${REL}styles" Target="styles.xml"/>`;
  if (sharedStrings) {
    s += `<Relationship Id="rId${sheetCount + 2}" Type="${REL}sharedStrings" Target="sharedStrings.xml"/>`;
  }
  return `${s}${extra}</Relationships>`;
}

export interface LinkRec {
  readonly ref: string;
  readonly target: string;
  readonly tooltip: string | undefined;
}

export function sharedStringsXml(sst: SharedStrings): string {
  const list = sst.strings;
  const parts: string[] = [
    `${XML_DECL}<sst xmlns="${NS_MAIN}" count="${sst.count}" uniqueCount="${list.length}">`,
  ];
  let buf = '';
  for (let i = 0; i < list.length; i++) {
    buf += `<si>${tElem(list[i]!)}</si>`;
    if (buf.length > 65536) {
      parts.push(buf);
      buf = '';
    }
  }
  parts.push(buf, '</sst>');
  return parts.join('');
}

/** `<t>` element, with xml:space="preserve" only when the text has leading/trailing whitespace. */
export function tElem(v: string): string {
  const n = v.length;
  if (n === 0) return '<t></t>';
  const a = v.charCodeAt(0);
  const z = v.charCodeAt(n - 1);
  const e = escapeText(v);
  return a === 32 || a === 9 || a === 10 || a === 13 || z === 32 || z === 9 || z === 10 || z === 13
    ? `<t xml:space="preserve">${e}</t>`
    : `<t>${e}</t>`;
}

// ---- styles.xml ------------------------------------------------------------------------

export function colorAttrs(c: Color): string {
  let s = '';
  if (c.argb !== undefined) s += ` rgb="${escapeAttr(c.argb)}"`;
  else if (c.theme !== undefined) s += ` theme="${c.theme}"`;
  else if (c.indexed !== undefined) s += ` indexed="${c.indexed}"`;
  if (c.tint !== undefined && c.tint !== 0) s += ` tint="${c.tint}"`;
  return s;
}

/** Font children in Excel's order. `nameTag` is 'name' in styles.xml and 'rFont' in rich text. */
export function fontBody(f: Font, nameTag: 'name' | 'rFont'): string {
  let s = '';
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.underline) {
    const u = f.underline === true ? 'single' : f.underline;
    s += u === 'single' ? '<u/>' : `<u val="${u}"/>`;
  }
  if (f.vertAlign) s += `<vertAlign val="${f.vertAlign}"/>`;
  if (f.size !== undefined) s += `<sz val="${f.size}"/>`;
  if (f.color) s += `<color${colorAttrs(f.color)}/>`;
  if (f.name !== undefined) s += `<${nameTag} val="${escapeAttr(f.name)}"/>`;
  return s;
}

function fillXml(f: Fill): string {
  const p = f.pattern || 'none';
  if (!f.fgColor && !f.bgColor) return `<fill><patternFill patternType="${escapeAttr(p)}"/></fill>`;
  let s = `<fill><patternFill patternType="${escapeAttr(p)}">`;
  if (f.fgColor) s += `<fgColor${colorAttrs(f.fgColor)}/>`;
  s += f.bgColor
    ? `<bgColor${colorAttrs(f.bgColor)}/>`
    : p === 'solid'
      ? '<bgColor indexed="64"/>'
      : '';
  return `${s}</patternFill></fill>`;
}

function edgeXml(tag: string, e: BorderEdge | undefined): string {
  if (!e?.style) return `<${tag}/>`;
  return e.color
    ? `<${tag} style="${e.style}"><color${colorAttrs(e.color)}/></${tag}>`
    : `<${tag} style="${e.style}"/>`;
}

function borderXml(b: Border): string {
  const d = b.diagonal;
  let attrs = '';
  if (d?.up) attrs += ' diagonalUp="1"';
  if (d?.down) attrs += ' diagonalDown="1"';
  return `<border${attrs}>${edgeXml('left', b.left)}${edgeXml('right', b.right)}${edgeXml(
    'top',
    b.top,
  )}${edgeXml('bottom', b.bottom)}${edgeXml('diagonal', d)}</border>`;
}

function alignmentXml(a: Alignment): string {
  let s = '<alignment';
  if (a.horizontal) s += ` horizontal="${a.horizontal}"`;
  if (a.vertical) s += ` vertical="${a.vertical}"`;
  if (a.textRotation) s += ` textRotation="${a.textRotation}"`;
  if (a.wrapText) s += ' wrapText="1"';
  if (a.indent) s += ` indent="${a.indent}"`;
  if (a.shrinkToFit) s += ' shrinkToFit="1"';
  return `${s}/>`;
}

/** One `<dxf>` (CT_Dxf child order: font, numFmt, fill, alignment, border, protection). */
function dxfXml(d: Style, reg: StyleRegistry): string {
  let s = '<dxf>';
  if (d.font) s += `<font>${fontBody(d.font, 'name')}</font>`;
  if (d.numFmt) {
    s += `<numFmt numFmtId="${reg.numFmtId(d.numFmt)}" formatCode="${escapeAttr(d.numFmt)}"/>`;
  }
  if (d.fill) s += fillXml(d.fill);
  if (d.alignment) s += alignmentXml(d.alignment);
  if (d.border) s += borderXml(d.border);
  const p = d.protection;
  if (p) {
    s += `<protection${p.locked !== undefined ? ` locked="${p.locked ? 1 : 0}"` : ''}${
      p.hidden !== undefined ? ` hidden="${p.hidden ? 1 : 0}"` : ''
    }/>`;
  }
  return `${s}</dxf>`;
}

function xfXml(x: Xf): string {
  let s = `<xf numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"`;
  if (x.numFmtId) s += ' applyNumberFormat="1"';
  if (x.fontId) s += ' applyFont="1"';
  if (x.fillId) s += ' applyFill="1"';
  if (x.borderId) s += ' applyBorder="1"';
  const a = x.alignment;
  const p = x.protection;
  if (a) s += ' applyAlignment="1"';
  if (p) s += ' applyProtection="1"';
  if (!a && !p) return `${s}/>`;
  s += '>';
  if (a) s += alignmentXml(a);
  if (p) {
    s += `<protection${p.locked === false ? ' locked="0"' : ''}${p.hidden ? ' hidden="1"' : ''}/>`;
  }
  return `${s}</xf>`;
}

export function stylesXml(reg: StyleRegistry): string {
  let s = `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">`;
  const nf = reg.numFmts;
  if (nf.length > 0) {
    s += `<numFmts count="${nf.length}">`;
    for (const f of nf) {
      s += `<numFmt numFmtId="${f.id}" formatCode="${escapeAttr(f.formatCode)}"/>`;
    }
    s += '</numFmts>';
  }
  s += `<fonts count="${reg.fonts.length}">`;
  for (const f of reg.fonts) s += `<font>${fontBody(f, 'name')}</font>`;
  s += `</fonts><fills count="${reg.fills.length}">`;
  for (const f of reg.fills) s += fillXml(f);
  s += `</fills><borders count="${reg.borders.length}">`;
  for (const b of reg.borders) s += borderXml(b);
  s += `</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${reg.xfs.length}">`;
  for (const x of reg.xfs) s += xfXml(x);
  s +=
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
  const dx = reg.dxfs;
  s += `<dxfs count="${dx.length}">`;
  for (const d of dx) s += dxfXml(d, reg);
  return `${s}</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
}
