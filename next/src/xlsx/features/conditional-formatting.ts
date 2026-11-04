// Feature: conditional formatting — <conditionalFormatting> (ECMA-376 Part 1 §18.3.1.18) plus
// the Excel 2010 extension (<extLst><ext><x14:conditionalFormattings>) for data bar settings
// and the x14-only icon sets.
//
//   addConditionalFormatting(ws, 'A1:A100', [
//     { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } },
//     { type: 'containsText', text: 'error', style: { fill: { type: 'pattern', pattern: 'solid',
//       bgColor: { argb: 'FFFFC7CE' } } } },
//     { type: 'colorScale', colors: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }] },
//     { type: 'dataBar', color: { argb: 'FF638EC6' }, gradient: false },
//     { type: 'iconSet', iconSet: '3Arrows', showValue: false },
//   ]);
//
// - `style` becomes a differential format (styles.xml <dxfs>, via StyleRegistry.internDxf);
//   read back it is the registry dxf.
// - `priority` (1 = evaluated first) is auto-assigned when absent: after the highest explicit
//   priority of the sheet, in document order. The model is never mutated at write time.
// - Text / blank / error / time-period rules get the formula Excel expects (relative to the
//   top-left cell of the first range of `ref`) unless `formulae` is given. On read, a formula
//   equal to the generated one is dropped again, so round trips are lossless.
// - colorScale / dataBar / iconSet `cfvo` default to min..max (iconSet: equal percent steps).
// - Data bars: the classic <dataBar> is always written; when an Excel 2010 setting is used
//   (gradient: false, border, negative / axis colors, axisPosition, direction, autoMin/autoMax
//   cfvo) a linked x14 dataBar is added in the sheet extLst, which Excel 2010+ honours. Icon
//   sets '3Stars' / '3Triangles' / '5Boxes' / 'NoIcons' exist only in x14 and are written there.
//   Other x14-only rules (e.g. expression rules referencing other sheets) are not read.

import { decodeRange, encodeAddr } from '../../core/address.js';
import type { Color, Style } from '../../core/types.js';
import type { Worksheet } from '../../core/worksheet.js';
import type { XmlAttrs } from '../../xml/index.js';
import {
  type FeatureReader,
  type FeatureWriter,
  type SheetReadCtx,
  SheetSlot,
  type SheetWriteCtx,
  type SubtreeHandler,
} from './types.js';
import { attrs, bool, escapeAttr, escapeText, num } from './util.js';

export type CfvoType =
  | 'num'
  | 'percent'
  | 'max'
  | 'min'
  | 'formula'
  | 'percentile'
  | 'autoMin'
  | 'autoMax';

/** Conditional format value object: a threshold of a color scale, data bar or icon set. */
export interface Cfvo {
  type: CfvoType;
  /** Number or formula (not for min/max/autoMin/autoMax). */
  value?: number | string | undefined;
  /** Icon sets: threshold is ">=" (default true); false = ">". */
  gte?: boolean | undefined;
}

export type CfOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

export type CfTimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

export type CfIconSetName =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters'
  | '3Stars'
  | '3Triangles'
  | '5Boxes'
  | 'NoIcons';

interface CfRuleBase {
  /** Evaluation order, 1 = first; unique per sheet. Auto-assigned when absent. */
  priority?: number | undefined;
  /** Stop evaluating lower-priority rules when this one matches. */
  stopIfTrue?: boolean | undefined;
}

interface CfStyled extends CfRuleBase {
  /** Format applied when the rule matches (a differential format). */
  style?: Style | undefined;
}

/** Cell value compared with one formula (two for between / notBetween). */
export interface CfCellIsRule extends CfStyled {
  type: 'cellIs';
  operator: CfOperator;
  /** Operands: formulas ("$B$1", "\"text\"") or numbers. */
  formulae: (string | number)[];
}

/** A formula evaluated for the top-left cell, relative to it for the others. */
export interface CfExpressionRule extends CfStyled {
  type: 'expression';
  formulae: string[];
}

/** Top / bottom N items (or N percent). */
export interface CfTop10Rule extends CfStyled {
  type: 'top10';
  /** N (default 10). */
  rank?: number | undefined;
  percent?: boolean | undefined;
  bottom?: boolean | undefined;
}

export interface CfAboveAverageRule extends CfStyled {
  type: 'aboveAverage';
  /** false = below average (default true). */
  aboveAverage?: boolean | undefined;
  /** Include the average itself. */
  equalAverage?: boolean | undefined;
  /** Number of standard deviations above/below the average. */
  stdDev?: number | undefined;
}

export interface CfValuesRule extends CfStyled {
  type: 'duplicateValues' | 'uniqueValues';
}

export interface CfTextRule extends CfStyled {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
  text: string;
  /** Override of the generated formula. */
  formulae?: string[] | undefined;
}

export interface CfBlankErrorRule extends CfStyled {
  type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors';
  formulae?: string[] | undefined;
}

export interface CfTimePeriodRule extends CfStyled {
  type: 'timePeriod';
  timePeriod: CfTimePeriod;
  formulae?: string[] | undefined;
}

export interface CfColorScaleRule extends CfRuleBase {
  type: 'colorScale';
  /** Thresholds (default min, max — or min, 50th percentile, max for three colors). */
  cfvo?: Cfvo[] | undefined;
  /** One color per threshold (2 or 3). */
  colors: Color[];
}

export interface CfDataBarRule extends CfRuleBase {
  type: 'dataBar';
  /** [shortest, longest] (default min, max). */
  cfvo?: Cfvo[] | undefined;
  /** Bar (fill) color. */
  color: Color;
  /** Show the cell value next to the bar (default true). */
  showValue?: boolean | undefined;
  /** Bar length bounds in percent of the cell width (Excel default 10 / 90). */
  minLength?: number | undefined;
  maxLength?: number | undefined;
  // Excel 2010 (x14) settings — any of them adds the x14 extension.
  /** Gradient fill (default true); false = solid bars. */
  gradient?: boolean | undefined;
  border?: boolean | undefined;
  borderColor?: Color | undefined;
  negativeFillColor?: Color | undefined;
  negativeBorderColor?: Color | undefined;
  axisColor?: Color | undefined;
  axisPosition?: 'auto' | 'middle' | 'none' | undefined;
  direction?: 'context' | 'leftToRight' | 'rightToLeft' | undefined;
  negativeBarColorSameAsPositive?: boolean | undefined;
  negativeBarBorderColorSameAsPositive?: boolean | undefined;
}

export interface CfIconSetRule extends CfRuleBase {
  type: 'iconSet';
  /** Default '3TrafficLights1'. */
  iconSet?: CfIconSetName | undefined;
  /** One threshold per icon (default: equal percent steps from 0). */
  cfvo?: Cfvo[] | undefined;
  reverse?: boolean | undefined;
  /** Show the cell value next to the icon (default true). */
  showValue?: boolean | undefined;
}

export type CfRule =
  | CfCellIsRule
  | CfExpressionRule
  | CfTop10Rule
  | CfAboveAverageRule
  | CfValuesRule
  | CfTextRule
  | CfBlankErrorRule
  | CfTimePeriodRule
  | CfColorScaleRule
  | CfDataBarRule
  | CfIconSetRule;

export interface ConditionalFormatting {
  /** Target cells: space-separated refs ("A1:A10 C1:C10"). */
  ref: string;
  rules: CfRule[];
}

/** Add conditional formatting rules for `ref`; returns the new block. */
export function addConditionalFormatting(
  ws: Worksheet,
  ref: string,
  rules: CfRule[],
): ConditionalFormatting {
  const cf: ConditionalFormatting = { ref, rules };
  ws.conditionalFormatting ??= [];
  ws.conditionalFormatting.push(cf);
  return cf;
}

// ---------------------------------------------------------------- shared

const X14_NS = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
const XM_NS = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
/** Sheet extLst entry holding x14:conditionalFormattings. */
const X14_CF = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
/** cfRule extLst entry linking a classic rule to its x14 twin. */
const X14_REF = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';
const X14_ICONS = ['3Stars', '3Triangles', '5Boxes', 'NoIcons'];
const TEXT_OPS: Record<string, string> = {
  containsText: 'containsText',
  notContainsText: 'notContains',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
};

/** The formula Excel writes for text / blank / error / time-period rules. */
function autoFormula(r: CfRule, ref: string): string | undefined {
  const first = ref.trim().split(/\s+/)[0]!;
  let c: string;
  try {
    const rg = decodeRange(first.slice(first.lastIndexOf('!') + 1));
    c = encodeAddr(rg.top, rg.left);
  } catch {
    return undefined;
  }
  const t = 'text' in r ? `"${(r.text ?? '').replace(/"/g, '""')}"` : '';
  const d = `ROUNDDOWN(${c},0)`;
  const f = `FLOOR(${c},1)`;
  switch (r.type) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${t},${c})))`;
    case 'notContainsText':
      return `ISERROR(SEARCH(${t},${c}))`;
    case 'beginsWith':
      return `LEFT(${c},LEN(${t}))=${t}`;
    case 'endsWith':
      return `RIGHT(${c},LEN(${t}))=${t}`;
    case 'containsBlanks':
      return `LEN(TRIM(${c}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${c}))>0`;
    case 'containsErrors':
      return `ISERROR(${c})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${c}))`;
    case 'timePeriod':
      switch (r.timePeriod) {
        case 'today':
          return `${f}=TODAY()`;
        case 'yesterday':
          return `${f}=TODAY()-1`;
        case 'tomorrow':
          return `${f}=TODAY()+1`;
        case 'last7Days':
          return `AND(TODAY()-${f}<=6,${f}<=TODAY())`;
        case 'thisMonth':
          return `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`;
        case 'lastMonth':
          return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`;
        case 'nextMonth':
          return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`;
        case 'thisWeek':
          return `AND(TODAY()-${d}<=WEEKDAY(TODAY())-1,${d}-TODAY()<=7-WEEKDAY(TODAY()))`;
        case 'lastWeek':
          return `AND(TODAY()-${d}>=(WEEKDAY(TODAY())),TODAY()-${d}<(WEEKDAY(TODAY())+7))`;
        case 'nextWeek':
          return `AND(${d}-TODAY()>(7-WEEKDAY(TODAY())),${d}-TODAY()<(15-WEEKDAY(TODAY())))`;
      }
  }
  return undefined;
}

// ---------------------------------------------------------------- writing

function colorXml(tag: string, c: Color | undefined): string {
  if (!c) return '';
  return `<${tag}${attrs({
    rgb: c.argb,
    theme: c.argb === undefined ? c.theme : undefined,
    indexed: c.argb === undefined && c.theme === undefined ? c.indexed : undefined,
    tint: c.tint || undefined,
  })}/>`;
}

function defaultCfvo(r: CfColorScaleRule | CfDataBarRule | CfIconSetRule): Cfvo[] {
  if (r.cfvo?.length) return r.cfvo;
  if (r.type === 'iconSet') {
    const n = r.iconSet === 'NoIcons' ? 3 : +(r.iconSet ?? '3')[0]!;
    return Array.from({ length: n }, (_, i) => ({
      type: 'percent' as const,
      value: Math.round((i * 100) / n),
    }));
  }
  const mid: Cfvo[] =
    r.type === 'colorScale' && r.colors.length > 2 ? [{ type: 'percentile', value: 50 }] : [];
  return [{ type: 'min' }, ...mid, { type: 'max' }];
}

/** Classic cfvo (autoMin/autoMax exist only in x14 → min/max). */
function cfvoXml(v: Cfvo): string {
  const t = v.type === 'autoMin' ? 'min' : v.type === 'autoMax' ? 'max' : v.type;
  return `<cfvo${attrs({
    type: t,
    val: t === 'min' || t === 'max' ? undefined : String(v.value ?? 0),
    gte: v.gte === false ? '0' : undefined,
  })}/>`;
}

function x14Cfvo(v: Cfvo): string {
  const a = attrs({ type: v.type, gte: v.gte === false ? '0' : undefined });
  return v.value === undefined
    ? `<x14:cfvo${a}/>`
    : `<x14:cfvo${a}><xm:f>${escapeText(String(v.value))}</xm:f></x14:cfvo>`;
}

function needsX14(r: CfRule): boolean {
  if (r.type === 'iconSet') return X14_ICONS.includes(r.iconSet ?? '');
  if (r.type !== 'dataBar') return false;
  return (
    r.gradient === false ||
    !!r.border ||
    !!(r.borderColor || r.negativeFillColor || r.negativeBorderColor || r.axisColor) ||
    !!r.axisPosition ||
    !!r.direction ||
    r.negativeBarColorSameAsPositive !== undefined ||
    r.negativeBarBorderColorSameAsPositive !== undefined ||
    !!r.cfvo?.some((v) => v.type === 'autoMin' || v.type === 'autoMax')
  );
}

/** Priorities in document order: explicit ones kept, the rest numbered after the highest. */
function priorities(list: readonly ConditionalFormatting[]): number[] {
  let next = 0;
  for (const cf of list) for (const r of cf.rules) next = Math.max(next, r.priority ?? 0);
  const out: number[] = [];
  for (const cf of list) for (const r of cf.rules) out.push(r.priority ?? ++next);
  return out;
}

/** Deterministic x14 rule id (emitters must be pure): sheet index + rule ordinal. */
function x14Id(sheet: number, n: number): string {
  const h = (v: number, w: number) => v.toString(16).toUpperCase().padStart(w, '0');
  return `{F7189C44-${h(sheet & 0xffff, 4)}-4000-8000-${h(n, 12)}}`;
}

function formulas(fs: readonly (string | number)[] | undefined): string {
  let s = '';
  for (const f of fs ?? []) {
    const v = typeof f === 'number' ? String(f) : f.charCodeAt(0) === 61 ? f.slice(1) : f;
    s += `<formula>${escapeText(v)}</formula>`;
  }
  return s;
}

function ruleXml(c: SheetWriteCtx, r: CfRule, ref: string, priority: number, id: string): string {
  const a: Record<string, string | number | boolean | undefined> = { type: r.type };
  let body = '';
  if ('style' in r && r.style) a.dxfId = c.styles.internDxf(r.style);
  a.priority = priority;
  a.stopIfTrue = r.stopIfTrue;
  switch (r.type) {
    case 'cellIs':
      a.operator = r.operator;
      body = formulas(r.formulae);
      break;
    case 'expression':
      body = formulas(r.formulae);
      break;
    case 'top10':
      a.percent = r.percent;
      a.bottom = r.bottom;
      a.rank = r.rank ?? 10;
      break;
    case 'aboveAverage':
      a.aboveAverage = r.aboveAverage === false ? '0' : undefined;
      a.stdDev = r.stdDev;
      a.equalAverage = r.equalAverage;
      break;
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      a.operator = TEXT_OPS[r.type];
      a.text = r.text;
      body = formulas(r.formulae ?? [autoFormula(r, ref) ?? '']);
      break;
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      body = formulas(r.formulae ?? [autoFormula(r, ref) ?? '']);
      break;
    case 'timePeriod':
      a.timePeriod = r.timePeriod;
      body = formulas(r.formulae ?? [autoFormula(r, ref) ?? '']);
      break;
    case 'colorScale': {
      let cs = '';
      for (const v of defaultCfvo(r)) cs += cfvoXml(v);
      for (const col of r.colors) cs += colorXml('color', col);
      body = `<colorScale>${cs}</colorScale>`;
      break;
    }
    case 'dataBar': {
      let db = '';
      for (const v of defaultCfvo(r)) db += cfvoXml(v);
      body = `<dataBar${attrs({
        minLength: r.minLength,
        maxLength: r.maxLength,
        showValue: r.showValue === false ? '0' : undefined,
      })}>${db}${colorXml('color', r.color)}</dataBar>`;
      if (id) {
        body += `<extLst><ext uri="${X14_REF}" ${X14_NS}><x14:id>${id}</x14:id></ext></extLst>`;
      }
      break;
    }
    case 'iconSet': {
      let is = '';
      for (const v of defaultCfvo(r)) is += cfvoXml(v);
      body = `<iconSet${attrs({
        iconSet: r.iconSet,
        showValue: r.showValue === false ? '0' : undefined,
        reverse: r.reverse,
      })}>${is}</iconSet>`;
      break;
    }
  }
  return body ? `<cfRule${attrs(a)}>${body}</cfRule>` : `<cfRule${attrs(a)}/>`;
}

function emitMain(c: SheetWriteCtx): string {
  const list = c.ws.conditionalFormatting;
  if (!list?.length) return '';
  const pr = priorities(list);
  let s = '';
  let n = 0;
  for (const cf of list) {
    let rules = '';
    for (const r of cf.rules) {
      const i = n++;
      const x = needsX14(r);
      if (x && r.type === 'iconSet') continue; // x14-only icon set: extLst only
      rules += ruleXml(c, r, cf.ref, pr[i]!, x ? x14Id(c.index, i) : '');
    }
    if (rules)
      s += `<conditionalFormatting sqref="${escapeAttr(cf.ref)}">${rules}</conditionalFormatting>`;
  }
  return s;
}

function emitExt(c: SheetWriteCtx): string {
  const list = c.ws.conditionalFormatting;
  if (!list?.length) return '';
  const pr = priorities(list);
  let s = '';
  let n = 0;
  for (const cf of list) {
    let rules = '';
    for (const r of cf.rules) {
      const i = n++;
      if (!needsX14(r)) continue;
      const id = x14Id(c.index, i);
      let cv = '';
      for (const v of defaultCfvo(r as CfDataBarRule | CfIconSetRule)) cv += x14Cfvo(v);
      if (r.type === 'dataBar') {
        rules += `<x14:cfRule type="dataBar" id="${id}"><x14:dataBar${attrs({
          minLength: r.minLength,
          maxLength: r.maxLength,
          border: r.border,
          gradient: r.gradient === false ? '0' : undefined,
          direction: r.direction,
          negativeBarColorSameAsPositive: r.negativeBarColorSameAsPositive,
          negativeBarBorderColorSameAsPositive:
            r.negativeBarBorderColorSameAsPositive === false ? '0' : undefined,
          axisPosition: r.axisPosition,
        })}>${cv}${colorXml('x14:borderColor', r.borderColor)}${colorXml(
          'x14:negativeFillColor',
          r.negativeFillColor,
        )}${colorXml('x14:negativeBorderColor', r.negativeBorderColor)}${colorXml(
          'x14:axisColor',
          r.axisColor,
        )}</x14:dataBar></x14:cfRule>`;
      } else if (r.type === 'iconSet') {
        rules += `<x14:cfRule type="iconSet" priority="${pr[i]}" id="${id}"><x14:iconSet${attrs({
          iconSet: r.iconSet,
          showValue: r.showValue === false ? '0' : undefined,
          reverse: r.reverse,
        })}>${cv}</x14:iconSet></x14:cfRule>`;
      }
    }
    if (rules) {
      s += `<x14:conditionalFormatting ${XM_NS}>${rules}<xm:sqref>${escapeText(cf.ref)}</xm:sqref></x14:conditionalFormatting>`;
    }
  }
  return s
    ? `<ext uri="${X14_CF}" ${X14_NS}><x14:conditionalFormattings>${s}</x14:conditionalFormattings></ext>`
    : '';
}

export const conditionalFormattingWriter: FeatureWriter = {
  name: 'conditional-formatting',
  sheet: {
    emitters: [
      { slot: SheetSlot.conditionalFormatting, emit: emitMain },
      { slot: SheetSlot.extLst, emit: emitExt },
    ],
  },
};

// ---------------------------------------------------------------- reading

type Raw = Record<string, unknown>;

function colorOf(a: XmlAttrs): Color {
  const c: { argb?: string; theme?: number; tint?: number; indexed?: number } = {};
  if (a.rgb) c.argb = a.rgb;
  const th = num(a.theme);
  if (th !== undefined) c.theme = th;
  const ix = num(a.indexed);
  if (ix !== undefined) c.indexed = ix;
  const ti = num(a.tint);
  if (ti) c.tint = ti;
  return c;
}

function cfvoOf(a: XmlAttrs): Cfvo {
  const v: Cfvo = { type: a.type as CfvoType };
  if (a.val !== undefined) v.value = num(a.val) ?? a.val;
  if (!bool(a.gte, true)) v.gte = false;
  return v;
}

/** Shared attributes of classic and x14 <dataBar> / <iconSet>. */
function barIconAttrs(r: Raw, a: XmlAttrs): void {
  const n = num(a.minLength);
  if (n !== undefined) r.minLength = n;
  const m = num(a.maxLength);
  if (m !== undefined) r.maxLength = m;
  if (!bool(a.showValue, true)) r.showValue = false;
  if (a.iconSet) r.iconSet = a.iconSet;
  if (bool(a.reverse)) r.reverse = true;
}

/** x14 dataBar ids (from the classic rules' extLst) → rule, per sheet. */
function idMap(ctx: SheetReadCtx): Map<string, Raw> {
  let m = ctx.scratch.get(conditionalFormattingReader) as Map<string, Raw> | undefined;
  if (!m) {
    m = new Map();
    ctx.scratch.set(conditionalFormattingReader, m);
  }
  return m;
}

function mainHandler(ctx: SheetReadCtx, sqref: string): SubtreeHandler {
  const cf: ConditionalFormatting = { ref: sqref, rules: [] };
  const ids = idMap(ctx);
  let r: Raw | undefined;
  let fs: string[] = [];
  let cfvo: Cfvo[] = [];
  let colors: Color[] = [];
  let cap: string | undefined;
  let text = '';
  let id = '';
  return {
    open(name, a) {
      switch (name) {
        case 'cfRule': {
          r = { type: a.type };
          fs = [];
          cfvo = [];
          colors = [];
          id = '';
          if (a.dxfId !== undefined) {
            const st = ctx.workbook.styles.dxfs[ctx.dxf(a.dxfId)];
            if (st) r.style = st;
          }
          const p = num(a.priority);
          if (p !== undefined) r.priority = p;
          if (bool(a.stopIfTrue)) r.stopIfTrue = true;
          switch (a.type) {
            case 'cellIs':
              r.operator = a.operator;
              break;
            case 'top10':
              if (bool(a.percent)) r.percent = true;
              if (bool(a.bottom)) r.bottom = true;
              r.rank = num(a.rank) ?? 10;
              break;
            case 'aboveAverage':
              if (!bool(a.aboveAverage, true)) r.aboveAverage = false;
              if (bool(a.equalAverage)) r.equalAverage = true;
              if (a.stdDev !== undefined) r.stdDev = num(a.stdDev);
              break;
            case 'timePeriod':
              r.timePeriod = a.timePeriod;
              break;
            case 'containsText':
            case 'notContainsText':
            case 'beginsWith':
            case 'endsWith':
              r.text = a.text ?? '';
          }
          return;
        }
        case 'formula':
        case 'id':
          cap = name;
          text = '';
          return;
        case 'dataBar':
        case 'iconSet':
          if (r) barIconAttrs(r, a);
          return;
        case 'cfvo':
          cfvo.push(cfvoOf(a));
          return;
        case 'color':
          colors.push(colorOf(a));
      }
    },
    text(t) {
      if (cap) text += t;
    },
    close(name) {
      if (name === cap) {
        if (cap === 'formula') fs.push(text);
        else id = text.trim();
        cap = undefined;
        return;
      }
      if (name !== 'cfRule' || !r) return;
      switch (r.type) {
        case 'colorScale':
          r.cfvo = cfvo;
          r.colors = colors;
          break;
        case 'dataBar':
          r.cfvo = cfvo;
          if (colors[0]) r.color = colors[0];
          if (id) ids.set(id, r);
          break;
        case 'iconSet':
          r.cfvo = cfvo;
          break;
        default:
          if (fs.length && !(fs.length === 1 && fs[0] === autoFormula(r as never, sqref))) {
            r.formulae = fs;
          }
      }
      cf.rules.push(r as unknown as CfRule);
      r = undefined;
    },
    end() {
      ctx.ws.conditionalFormatting ??= [];
      ctx.ws.conditionalFormatting.push(cf);
    },
  };
}

const X14_COLORS = ['borderColor', 'negativeFillColor', 'negativeBorderColor', 'axisColor'];

/** x14:conditionalFormattings: merge data bar settings, add x14-only icon sets. */
function extHandler(ctx: SheetReadCtx): SubtreeHandler {
  const ids = idMap(ctx);
  let recs: Raw[] = [];
  let r: Raw | undefined;
  let cv: Cfvo | undefined;
  let cap: string | undefined;
  let text = '';
  let ref = '';
  return {
    open(name, a) {
      if (name === 'conditionalFormatting') {
        recs = [];
        ref = '';
      } else if (name === 'cfRule') {
        r = { type: a.type, id: a.id, cfvo: [] };
        const p = num(a.priority);
        if (p !== undefined) r.priority = p;
      } else if (!r) {
        if (name === 'sqref') {
          cap = name;
          text = '';
        }
      } else if (name === 'dataBar' || name === 'iconSet') {
        barIconAttrs(r, a);
        if (bool(a.border)) r.border = true;
        if (!bool(a.gradient, true)) r.gradient = false;
        if (bool(a.negativeBarColorSameAsPositive)) r.negativeBarColorSameAsPositive = true;
        if (!bool(a.negativeBarBorderColorSameAsPositive, true)) {
          r.negativeBarBorderColorSameAsPositive = false;
        }
        if (a.axisPosition && a.axisPosition !== 'auto') r.axisPosition = a.axisPosition;
        if (a.direction && a.direction !== 'context') r.direction = a.direction;
      } else if (name === 'cfvo') {
        cv = cfvoOf(a);
        (r.cfvo as Cfvo[]).push(cv);
      } else if (name === 'f' && cv) {
        cap = name;
        text = '';
      } else if (X14_COLORS.includes(name)) r[name] = colorOf(a);
    },
    text(t) {
      if (cap) text += t;
    },
    close(name) {
      if (name === cap) {
        if (cap === 'sqref') ref = text;
        else if (cv) cv.value = num(text) ?? text;
        cap = undefined;
      } else if (name === 'cfvo') cv = undefined;
      else if (name === 'cfRule') {
        if (r) recs.push(r);
        r = undefined;
      } else if (name === 'conditionalFormatting') {
        for (const x of recs) {
          const { id, type, cfvo, ...rest } = x;
          const main = ids.get(id as string);
          if (main && type === 'dataBar') {
            Object.assign(main, rest);
            if ((cfvo as Cfvo[]).length) main.cfvo = cfvo;
          } else if (type === 'iconSet') {
            ctx.ws.conditionalFormatting ??= [];
            const list = ctx.ws.conditionalFormatting;
            let cf = list.find((c) => c.ref === ref);
            if (!cf) {
              cf = { ref, rules: [] };
              list.push(cf);
            }
            cf.rules.push({ type, cfvo, ...rest } as CfIconSetRule);
          }
        }
      }
    },
  };
}

export const conditionalFormattingReader: FeatureReader = {
  name: 'conditional-formatting',
  sheet: {
    elements: ['conditionalFormatting', 'extLst/ext'],
    open(ctx, el, a) {
      if (el === 'conditionalFormatting') return mainHandler(ctx, a.sqref ?? '');
      return a.uri === X14_CF ? extHandler(ctx) : undefined;
    },
  },
};
