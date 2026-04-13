/**
 * XML Conditional Format Saver
 *
 * Writes conditional formatting rules to XLSX worksheet XML.
 */

import { ConditionalFormat, ConditionalFormatCollection } from '../features/ConditionalFormat';

// Map internal type values to ECMA-376 cfRule type strings
const TYPE_MAP: Record<string, string> = {
  cellIs: 'cellIs',
  expression: 'expression',
  colorScale: 'colorScale',
  dataBar: 'dataBar',
  iconSet: 'iconSet',
  top10: 'top10',
  aboveAverage: 'aboveAverage',
  duplicateValues: 'duplicateValues',
  uniqueValues: 'uniqueValues',
  containsText: 'containsText',
  notContainsText: 'notContainsText',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
  containsBlanks: 'containsBlanks',
  notContainsBlanks: 'notContainsBlanks',
  containsErrors: 'containsErrors',
  notContainsErrors: 'notContainsErrors',
  timePeriod: 'timePeriod',
};

export class ConditionalFormatXmlSaver {
  private _escapeXml: (text: string) => string;

  constructor(escapeXml: (text: string) => string) {
    this._escapeXml = escapeXml;
  }

  /**
   * Formats conditional formatting XML. Returns DXF entries that need to be
   * added to the styles.xml dxfs collection.
   */
  formatConditionalFormattingXml(
    collection: ConditionalFormatCollection,
    startDxfId: number,
  ): { xml: string; dxfEntries: string[] } {
    if (collection.count === 0) return { xml: '', dxfEntries: [] };

    const dxfEntries: string[] = [];
    let dxfId = startDxfId;

    // Group by range
    const grouped = new Map<string, ConditionalFormat[]>();
    for (const cf of collection) {
      const range = cf.range ?? '';
      if (!grouped.has(range)) grouped.set(range, []);
      grouped.get(range)!.push(cf);
    }

    const lines: string[] = [];

    for (const [sqref, cfs] of grouped) {
      lines.push(`<conditionalFormatting sqref="${this._escapeXml(sqref)}">`);

      for (const cf of cfs) {
        const hasDxf = cf.hasFont() || cf.hasFill() || cf.hasBorder() || cf.numberFormat != null;
        let cfDxfId: number | null = null;

        if (hasDxf) {
          cfDxfId = dxfId++;
          dxfEntries.push(this._formatDxfEntry(cf));
        }

        lines.push(this._formatCfRuleXml(cf, cfDxfId));
      }

      lines.push('</conditionalFormatting>');
    }

    return { xml: lines.join(''), dxfEntries };
  }

  private _formatCfRuleXml(cf: ConditionalFormat, dxfId: number | null): string {
    const attrs: string[] = [];
    const ruleType = cf.type ?? 'cellIs';
    attrs.push(`type="${TYPE_MAP[ruleType] ?? ruleType}"`);

    if (dxfId !== null) {
      attrs.push(`dxfId="${dxfId}"`);
    }

    attrs.push(`priority="${cf.priority}"`);

    if (cf.stopIfTrue) {
      attrs.push('stopIfTrue="1"');
    }

    // Operator
    if (cf.operator && (ruleType === 'cellIs' || ruleType === 'expression')) {
      attrs.push(`operator="${cf.operator}"`);
    }

    // Text
    if (cf.textFormula && ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(ruleType)) {
      attrs.push(`text="${this._escapeXml(cf.textFormula)}"`);
    }

    // TimePeriod
    if (cf.dateOperator && ruleType === 'timePeriod') {
      attrs.push(`timePeriod="${cf.dateOperator}"`);
    }

    // aboveAverage
    if (ruleType === 'aboveAverage') {
      if (cf.above === false) {
        attrs.push('aboveAverage="0"');
      }
      if (cf.stdDev > 0) {
        attrs.push(`stdDev="${cf.stdDev}"`);
      }
    }

    // top10
    if (ruleType === 'top10') {
      if (cf.top === false) {
        attrs.push('bottom="1"');
      }
      if (cf.percent) {
        attrs.push('percent="1"');
      }
      attrs.push(`rank="${cf.rank}"`);
    }

    let xml = `<cfRule ${attrs.join(' ')}>`;

    // Formulas
    if (cf.formula1 != null) {
      xml += `<formula>${this._escapeXml(cf.formula1)}</formula>`;
    }
    if (cf.formula2 != null) {
      xml += `<formula>${this._escapeXml(cf.formula2)}</formula>`;
    }

    // Auto-generate text rule formulas
    if (!cf.formula1 && cf.textFormula) {
      const firstCell = this._getFirstCellFromRange(cf.range ?? 'A1');
      const textFormula = this._buildTextRuleFormula(ruleType, cf.textFormula, firstCell);
      if (textFormula) {
        xml += `<formula>${this._escapeXml(textFormula)}</formula>`;
      }
    }

    // Color scale
    if (ruleType === 'colorScale' && cf.colorScaleType) {
      xml += this._formatColorScaleXml(cf);
    }

    // Data bar
    if (ruleType === 'dataBar' && cf.barColor) {
      xml += this._formatDataBarXml(cf);
    }

    // Icon set
    if (ruleType === 'iconSet' && cf.iconSetType) {
      xml += this._formatIconSetXml(cf);
    }

    xml += '</cfRule>';
    return xml;
  }

  private _formatColorScaleXml(cf: ConditionalFormat): string {
    let xml = '<colorScale>';
    if (cf.colorScaleType === '2color') {
      xml += '<cfvo type="min"/><cfvo type="max"/>';
      xml += `<color rgb="${cf.minColor ?? 'FFF8696B'}"/>`;
      xml += `<color rgb="${cf.maxColor ?? 'FF63BE7B'}"/>`;
    } else {
      xml += '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>';
      xml += `<color rgb="${cf.minColor ?? 'FFF8696B'}"/>`;
      xml += `<color rgb="${cf.midColor ?? 'FFFFEB84'}"/>`;
      xml += `<color rgb="${cf.maxColor ?? 'FF63BE7B'}"/>`;
    }
    xml += '</colorScale>';
    return xml;
  }

  private _formatDataBarXml(cf: ConditionalFormat): string {
    let xml = '<dataBar>';
    xml += '<cfvo type="min"/><cfvo type="max"/>';
    xml += `<color rgb="${cf.barColor ?? 'FF638EC6'}"/>`;
    xml += '</dataBar>';
    return xml;
  }

  private _formatIconSetXml(cf: ConditionalFormat): string {
    const attrs: string[] = [];
    if (cf.iconSetType) attrs.push(`iconSet="${cf.iconSetType}"`);
    if (cf.reverseIcons) attrs.push('reverse="1"');
    if (cf.showIconOnly) attrs.push('showValue="0"');

    let xml = `<iconSet ${attrs.join(' ')}>`;
    // Default 3-level thresholds
    xml += '<cfvo type="percent" val="0"/>';
    xml += '<cfvo type="percent" val="33"/>';
    xml += '<cfvo type="percent" val="67"/>';
    xml += '</iconSet>';
    return xml;
  }

  private _formatDxfEntry(cf: ConditionalFormat): string {
    let xml = '<dxf>';

    if (cf.hasFont()) {
      const f = cf.font;
      xml += '<font>';
      if (f.bold) xml += '<b/>';
      if (f.italic) xml += '<i/>';
      if (f.underline) xml += '<u/>';
      if (f.strikethrough) xml += '<strike/>';
      if (f.color && f.color !== 'FF000000') {
        xml += `<color rgb="${f.color}"/>`;
      }
      if (f.name && f.name !== 'Calibri') {
        xml += `<name val="${this._escapeXml(f.name)}"/>`;
      }
      if (f.size && f.size !== 11) {
        xml += `<sz val="${f.size}"/>`;
      }
      xml += '</font>';
    }

    if (cf.hasFill()) {
      const fl = cf.fill;
      xml += '<fill><patternFill';
      if (fl.patternType !== 'none') {
        xml += ` patternType="${fl.patternType}"`;
      }
      xml += '>';
      if (fl.foregroundColor && fl.foregroundColor !== 'FFFFFFFF') {
        xml += `<fgColor rgb="${fl.foregroundColor}"/>`;
      }
      if (fl.backgroundColor && fl.backgroundColor !== 'FFFFFFFF') {
        xml += `<bgColor rgb="${fl.backgroundColor}"/>`;
      }
      xml += '</patternFill></fill>';
    }

    if (cf.hasBorder()) {
      const b = cf.border;
      xml += '<border>';
      const sides = ['left', 'right', 'top', 'bottom'] as const;
      for (const side of sides) {
        const s = b[side];
        if (s.lineStyle !== 'none') {
          xml += `<${side} style="${s.lineStyle}">`;
          xml += `<color rgb="${s.color}"/>`;
          xml += `</${side}>`;
        } else {
          xml += `<${side}/>`;
        }
      }
      xml += '</border>';
    }

    if (cf.numberFormat != null) {
      xml += `<numFmt formatCode="${this._escapeXml(cf.numberFormat)}"/>`;
    }

    xml += '</dxf>';
    return xml;
  }

  private _getFirstCellFromRange(range: string): string {
    const colonIdx = range.indexOf(':');
    if (colonIdx > 0) return range.substring(0, colonIdx);
    return range;
  }

  private _buildTextRuleFormula(ruleType: string, textValue: string, firstCell: string): string | null {
    switch (ruleType) {
      case 'containsText':
        return `NOT(ISERROR(SEARCH("${textValue}",${firstCell})))`;
      case 'notContainsText':
        return `ISERROR(SEARCH("${textValue}",${firstCell}))`;
      case 'beginsWith':
        return `LEFT(${firstCell},${textValue.length})="${textValue}"`;
      case 'endsWith':
        return `RIGHT(${firstCell},${textValue.length})="${textValue}"`;
      default:
        return null;
    }
  }
}
