// Generators for the small, non-sheet XLSX parts. Each returns a complete XML document string.

import type { SharedStrings } from '../../core/sharedstrings.js';
import type { StyleRegistry, Xf } from '../../core/styles.js';
import type {
  Alignment,
  Border,
  BorderEdge,
  DefinedName,
  Fill,
  Style,
  WorkbookView,
} from '../../core/types.js';
import type { Workbook } from '../../core/workbook.js';
import type { SheetState } from '../../core/worksheet.js';
import { escapeAttr, escapeText } from '../../xml/index.js';
import { colorAttrs, fontBody, tElem } from '../common/rich-text.js';
import { addXlfn } from '../common/xlfn.js';

export { colorAttrs, fontBody, tElem };

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
  /** Write `fullCalcOnLoad` (formulas present, or wb.calcProperties). */
  readonly formulas: boolean;
  readonly activeTab: number;
  /** wb.views[0]: window settings (its activeTab wins over the derived one). */
  readonly view?: WorkbookView | undefined;
  readonly creator: string | undefined;
  readonly lastModifiedBy: string | undefined;
  readonly created: Date | undefined;
  readonly modified: Date | undefined;
  /** Document properties (title, subject, ...); see WorkbookFeatureFields. */
  readonly props?: Readonly<Record<string, string | undefined>> | undefined;
  /** User defined names (wb.definedNames); autoFilter names are added by workbookXml. */
  readonly definedNames?: readonly DefinedName[] | undefined;
  /** Preserved (round-trip) pieces: raw workbookPr / calcPr attributes, extra `<sheet>`s. */
  readonly prAttrs?: string | undefined;
  readonly calcAttrs?: string | undefined;
  /** Kept non-worksheet `<sheet>`s placed before worksheet i (index = sheets.length: last). */
  readonly extraSheets?: readonly string[] | undefined;
  /** Worksheet index → position in `<sheets>` (activeTab, localSheetId); identity if absent. */
  readonly at?: readonly number[] | undefined;
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
  if (!overrides.some((o) => o[0] === 'xl/workbook.xml')) {
    s += `<Override PartName="/xl/workbook.xml" ContentType="${CT}spreadsheetml.sheet.main+xml"/>`;
  }
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

/** Package relationships; `extra` = more relationship elements (ids from rId4). */
export function rootRelsXml(extra = ''): string {
  return `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL}officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${NS_PKG_REL}/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}extended-properties" Target="docProps/app.xml"/>${extra}</Relationships>`;
}

function isoDate(d: Date | undefined): string | undefined {
  if (!d) return undefined;
  const t = d.getTime();
  if (t - t !== 0) return undefined;
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** The workbook's document properties for BookInfo.props (undefined when none is set). */
export function docProps(wb: Workbook): Record<string, string | undefined> | undefined {
  const o = {
    title: wb.title,
    subject: wb.subject,
    keywords: wb.keywords,
    description: wb.description,
    category: wb.category,
    contentStatus: wb.contentStatus,
    company: wb.company,
    manager: wb.manager,
  };
  return Object.values(o).some((v) => v) ? o : undefined;
}

/** Workbook property → docProps/core.xml element. */
const CORE_PROPS = [
  ['title', 'dc:title'],
  ['subject', 'dc:subject'],
  ['keywords', 'cp:keywords'],
  ['description', 'dc:description'],
  ['category', 'cp:category'],
  ['contentStatus', 'cp:contentStatus'],
] as const;

export function corePropsXml(b: BookInfo): string {
  let s = `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`;
  const p = b.props;
  if (p) {
    for (const [k, tag] of CORE_PROPS) {
      const v = p[k];
      if (v) s += `<${tag}>${escapeText(v)}</${tag}>`;
    }
  }
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
  return `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application>${b.props?.manager ? `<Manager>${escapeText(b.props.manager)}</Manager>` : ''}${b.props?.company ? `<Company>${escapeText(b.props.company)}</Company>` : ''}<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${n}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`;
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
  }${hidden ? ' hidden="1"' : ''}>${escapeText(addXlfn(value))}</definedName>`;
}

/**
 * workbook.xml. `extra`: feature elements as [BookSlot, xml] pairs, merged into CT_Workbook
 * order with the built-ins (slot 18 = extLst children, wrapped in one <extLst>).
 */
export function workbookXml(b: BookInfo, extra: readonly [number, string][] = []): string {
  const at = b.at;
  const v = b.view;
  const pos = (i: number): number => (at ? (at[i] ?? i) : i);
  const tab = pos(v?.activeTab ?? b.activeTab);
  const first = v?.firstSheet !== undefined ? pos(v.firstSheet) : tab > 0 ? 0 : undefined;
  const x = b.extraSheets;
  const parts: [number, string][] = [
    [2, `<workbookPr${b.date1904 ? ' date1904="1"' : ''}${b.prAttrs ?? ''}/>`],
    [
      4,
      `<bookViews><workbookView${v?.visibility && v.visibility !== 'visible' ? ` visibility="${v.visibility}"` : ''} xWindow="${v?.x ?? 0}" yWindow="${v?.y ?? 0}" windowWidth="${v?.width ?? 28800}" windowHeight="${v?.height ?? 12300}"${
        first !== undefined ? ` firstSheet="${first}"` : ''
      }${tab > 0 ? ` activeTab="${tab}"` : ''}/></bookViews>`,
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
    sheets += `${x?.[i] ?? ''}<sheet name="${escapeAttr(sh.name)}" sheetId="${sh.sheetId}"${
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
  parts.push([5, `${sheets}${x?.[b.sheets.length] ?? ''}</sheets>`]);
  // scopes are worksheet indices: positions in <sheets> when kept chartsheets are interleaved
  if (at)
    names = names.replace(
      /(<definedName name="[^"]*" localSheetId=")(\d+)/g,
      (_, p, i) => p + at[+i],
    );
  if (names) parts.push([8, `<definedNames>${names}</definedNames>`]);
  parts.push([
    9,
    `<calcPr${b.calcAttrs ?? ' calcId="191029"'}${b.formulas ? ' fullCalcOnLoad="1"' : ''}/>`,
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

// ---- styles.xml ------------------------------------------------------------------------

function fillXml(f: Fill): string {
  if (f.type === 'gradient') {
    const a =
      f.gradient === 'path'
        ? ` type="path" left="${f.center?.left ?? 0}" right="${f.center?.left ?? 0}" top="${f.center?.top ?? 0}" bottom="${f.center?.top ?? 0}"`
        : ` degree="${f.degree ?? 0}"`;
    let s = `<fill><gradientFill${a}>`;
    for (const st of f.stops)
      s += `<stop position="${st.position}"><color${colorAttrs(st.color)}/></stop>`;
    return `${s}</gradientFill></fill>`;
  }
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

/** xl/metadata.xml: one cell-metadata record (cm="1") marking dynamic-array formulas. */
export function dynamicArrayMetadataXml(): string {
  return `${XML_DECL}<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
}

const XFPB = 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';

/** xl/featurePropertyBag/featurePropertyBag.xml: the checkbox bag that `xfComplement i="0"` maps to. */
export function checkboxBagXml(): string {
  return `${XML_DECL}<FeaturePropertyBags xmlns="${XFPB}"><bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag><bag type="XFComplement"><bagId k="XFControls">1</bagId></bag><bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag></FeaturePropertyBags>`;
}

function xfXml(x: Xf, named: unknown): string {
  let s = `<xf numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="${(named && x.xfId) || 0}"`;
  if (x.numFmtId) s += ' applyNumberFormat="1"';
  if (x.fontId) s += ' applyFont="1"';
  if (x.fillId) s += ' applyFill="1"';
  if (x.borderId) s += ' applyBorder="1"';
  const a = x.alignment;
  const p = x.protection;
  if (a) s += ' applyAlignment="1"';
  if (p) s += ' applyProtection="1"';
  if (!a && !p && !x.checkbox) return `${s}/>`;
  s += '>';
  if (a) s += alignmentXml(a);
  if (p) {
    s += `<protection${p.locked === false ? ' locked="0"' : ''}${p.hidden ? ' hidden="1"' : ''}/>`;
  }
  if (x.checkbox) {
    s += `<extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="${XFPB}"><xfpb:xfComplement i="0"/></ext></extLst>`;
  }
  return `${s}</xf>`;
}

/**
 * styles.xml; `tableStyles` / `tail` (colors, extLst) / `named` (cellStyleXfs, cellStyles):
 * preserved raw pieces. Xf `xfId`s are written only with the named styles they refer to.
 */
export function stylesXml(
  reg: StyleRegistry,
  tableStyles?: string,
  tail = '',
  named?: readonly [string, string],
): string {
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
  s += `</borders>${named?.[0] ?? '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'}<cellXfs count="${reg.xfs.length}">`;
  for (const x of reg.xfs) s += xfXml(x, named);
  s += `</cellXfs>${named?.[1] ?? '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'}`;
  const dx = reg.dxfs;
  s += `<dxfs count="${dx.length}">`;
  for (const d of dx) s += dxfXml(d, reg);
  return `${s}</dxfs>${
    tableStyles ??
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
  }${tail}</styleSheet>`;
}
