/**
 * XML Conditional Format Loader
 *
 * Loads conditional formatting rules from XLSX worksheet XML.
 */

import { ConditionalFormat, ConditionalFormatCollection } from '../features/ConditionalFormat';
import type { LoaderFontData, LoaderFillData, LoaderBorderData } from './XmlLoader';

function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}

function attr(elem: any, name: string, defaultVal = ''): string {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : defaultVal;
}

function attrOpt(elem: any, name: string): string | null {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : null;
}

export class ConditionalFormatXmlLoader {
  private _dxfStyles: any[];

  constructor(dxfStyles: any[]) {
    this._dxfStyles = dxfStyles;
  }

  loadConditionalFormatting(wsRoot: any, collection: ConditionalFormatCollection): void {
    const cfElems = ensureArray(wsRoot?.conditionalFormatting);

    for (const cfElem of cfElems) {
      const sqref = attr(cfElem, 'sqref', '');
      if (!sqref) continue;

      const rules = ensureArray(cfElem.cfRule);
      for (const rule of rules) {
        const cf = collection.add();
        cf.range = sqref;
        this._loadCfRule(cf, rule);
      }
    }

    // Apply DXF styles
    this._applyDxfStyles(collection);
  }

  private _loadCfRule(cf: ConditionalFormat, rule: any): void {
    const ruleType = attr(rule, 'type', '');
    cf.type = ruleType || null;
    cf.priority = parseInt(attr(rule, 'priority', '0'), 10);

    const stopIfTrue = attr(rule, 'stopIfTrue', '');
    cf.stopIfTrue = stopIfTrue === '1' || stopIfTrue === 'true';

    const dxfId = attrOpt(rule, 'dxfId');
    if (dxfId !== null) {
      cf._dxfId = parseInt(dxfId, 10);
    }

    // Operator
    const op = attrOpt(rule, 'operator');
    if (op) cf.operator = op;

    // Text
    const text = attrOpt(rule, 'text');
    if (text) cf.textFormula = text;

    // TimePeriod (date rules)
    const timePeriod = attrOpt(rule, 'timePeriod');
    if (timePeriod) cf.dateOperator = timePeriod;

    // aboveAverage
    const aboveAvg = attrOpt(rule, 'aboveAverage');
    if (ruleType === 'aboveAverage') {
      cf.above = aboveAvg !== '0';
    }

    // stdDev
    const stdDev = attrOpt(rule, 'stdDev');
    if (stdDev) cf.stdDev = parseInt(stdDev, 10);

    // top/bottom
    if (ruleType === 'top10') {
      const bottom = attrOpt(rule, 'bottom');
      cf.top = bottom !== '1';
      const percent = attrOpt(rule, 'percent');
      cf.percent = percent === '1';
      const rank = attrOpt(rule, 'rank');
      if (rank) cf.rank = parseInt(rank, 10);
    }

    // Duplicate/unique
    if (ruleType === 'duplicateValues') cf.duplicate = true;
    if (ruleType === 'uniqueValues') cf.duplicate = false;

    // Formulas
    const formulas = ensureArray(rule.formula);
    if (formulas.length > 0) {
      cf.formula1 = String(formulas[0]);
    }
    if (formulas.length > 1) {
      cf.formula2 = String(formulas[1]);
    }

    // Color scale
    if (rule.colorScale) {
      this._loadColorScale(cf, rule.colorScale);
    }

    // Data bar
    if (rule.dataBar) {
      this._loadDataBar(cf, rule.dataBar);
    }

    // Icon set
    if (rule.iconSet) {
      this._loadIconSet(cf, rule.iconSet);
    }

    // Map ruleType to text/date operators
    if (ruleType === 'containsText') cf.textOperator = 'containsText';
    else if (ruleType === 'notContainsText') cf.textOperator = 'notContains';
    else if (ruleType === 'beginsWith') cf.textOperator = 'beginsWith';
    else if (ruleType === 'endsWith') cf.textOperator = 'endsWith';
    else if (ruleType === 'containsBlanks') cf.textOperator = 'containsBlanks';
    else if (ruleType === 'notContainsBlanks') cf.textOperator = 'notContainsBlanks';
    else if (ruleType === 'containsErrors') cf.textOperator = 'containsErrors';
    else if (ruleType === 'notContainsErrors') cf.textOperator = 'notContainsErrors';
  }

  private _loadColorScale(cf: ConditionalFormat, elem: any): void {
    const colors = ensureArray(elem.color);
    if (colors.length === 2) {
      cf.colorScaleType = '2color';
      cf.minColor = attrOpt(colors[0], 'rgb') ?? attrOpt(colors[0], 'theme');
      cf.maxColor = attrOpt(colors[1], 'rgb') ?? attrOpt(colors[1], 'theme');
    } else if (colors.length >= 3) {
      cf.colorScaleType = '3color';
      cf.minColor = attrOpt(colors[0], 'rgb') ?? attrOpt(colors[0], 'theme');
      cf.midColor = attrOpt(colors[1], 'rgb') ?? attrOpt(colors[1], 'theme');
      cf.maxColor = attrOpt(colors[2], 'rgb') ?? attrOpt(colors[2], 'theme');
    }
  }

  private _loadDataBar(cf: ConditionalFormat, elem: any): void {
    const colors = ensureArray(elem.color);
    if (colors.length > 0) {
      cf.barColor = attrOpt(colors[0], 'rgb');
    }
    if (colors.length > 1) {
      cf.negativeColor = attrOpt(colors[1], 'rgb');
    }
    const showValue = attrOpt(elem, 'showValue');
    if (showValue === '0') cf.showIconOnly = true;
  }

  private _loadIconSet(cf: ConditionalFormat, elem: any): void {
    const iconSet = attrOpt(elem, 'iconSet');
    if (iconSet) cf.iconSetType = iconSet;

    const reverse = attrOpt(elem, 'reverse');
    if (reverse === '1') cf.reverseIcons = true;

    const showValue = attrOpt(elem, 'showValue');
    if (showValue === '0') cf.showIconOnly = true;
  }

  private _applyDxfStyles(collection: ConditionalFormatCollection): void {
    for (const cf of collection) {
      if (cf._dxfId !== null && cf._dxfId >= 0 && cf._dxfId < this._dxfStyles.length) {
        const dxfData = this._dxfStyles[cf._dxfId];
        this._applyDxfData(cf, dxfData);
      }
    }
  }

  private _applyDxfData(cf: ConditionalFormat, dxfData: any): void {
    if (!dxfData) return;

    // Apply font
    if (dxfData.font) {
      const fd = dxfData.font as LoaderFontData;
      if (fd.name) cf.font.name = fd.name;
      if (fd.size) cf.font.size = fd.size;
      if (fd.color) cf.font.color = fd.color;
      if (fd.bold !== undefined) cf.font.bold = fd.bold;
      if (fd.italic !== undefined) cf.font.italic = fd.italic;
      if (fd.underline !== undefined) cf.font.underline = fd.underline;
      if (fd.strikethrough !== undefined) cf.font.strikethrough = fd.strikethrough;
    }

    // Apply fill
    if (dxfData.fill) {
      const fillD = dxfData.fill as LoaderFillData;
      if (fillD.pattern_type) cf.fill.patternType = fillD.pattern_type;
      if (fillD.fg_color) cf.fill.foregroundColor = fillD.fg_color;
      if (fillD.bg_color) cf.fill.backgroundColor = fillD.bg_color;
    }

    // Apply border
    if (dxfData.border) {
      const bd = dxfData.border as LoaderBorderData;
      if (bd.left?.style) cf.border.left.lineStyle = bd.left.style;
      if (bd.left?.color) cf.border.left.color = bd.left.color;
      if (bd.right?.style) cf.border.right.lineStyle = bd.right.style;
      if (bd.right?.color) cf.border.right.color = bd.right.color;
      if (bd.top?.style) cf.border.top.lineStyle = bd.top.style;
      if (bd.top?.color) cf.border.top.color = bd.top.color;
      if (bd.bottom?.style) cf.border.bottom.lineStyle = bd.bottom.style;
      if (bd.bottom?.color) cf.border.bottom.color = bd.bottom.color;
    }

    // Apply number format
    if (dxfData.numberFormat) {
      cf.numberFormat = dxfData.numberFormat;
    }
  }
}
