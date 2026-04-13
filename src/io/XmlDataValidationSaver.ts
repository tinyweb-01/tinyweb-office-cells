/**
 * XML Data Validation Saver
 *
 * Writes data validation settings to XLSX worksheet XML.
 */

import {
  DataValidation,
  DataValidationCollection,
  DataValidationType,
  DataValidationOperator,
  DataValidationAlertStyle,
  DataValidationImeMode,
} from '../features/DataValidation';

// Maps from enum → XML string
const TYPE_TO_XML: Record<number, string> = {
  [DataValidationType.NONE]: 'none',
  [DataValidationType.WHOLE_NUMBER]: 'whole',
  [DataValidationType.DECIMAL]: 'decimal',
  [DataValidationType.LIST]: 'list',
  [DataValidationType.DATE]: 'date',
  [DataValidationType.TIME]: 'time',
  [DataValidationType.TEXT_LENGTH]: 'textLength',
  [DataValidationType.CUSTOM]: 'custom',
};

const OPERATOR_TO_XML: Record<number, string> = {
  [DataValidationOperator.BETWEEN]: 'between',
  [DataValidationOperator.NOT_BETWEEN]: 'notBetween',
  [DataValidationOperator.EQUAL]: 'equal',
  [DataValidationOperator.NOT_EQUAL]: 'notEqual',
  [DataValidationOperator.GREATER_THAN]: 'greaterThan',
  [DataValidationOperator.LESS_THAN]: 'lessThan',
  [DataValidationOperator.GREATER_THAN_OR_EQUAL]: 'greaterThanOrEqual',
  [DataValidationOperator.LESS_THAN_OR_EQUAL]: 'lessThanOrEqual',
};

const ALERT_TO_XML: Record<number, string> = {
  [DataValidationAlertStyle.STOP]: 'stop',
  [DataValidationAlertStyle.WARNING]: 'warning',
  [DataValidationAlertStyle.INFORMATION]: 'information',
};

const IME_TO_XML: Record<number, string> = {
  [DataValidationImeMode.NO_CONTROL]: 'noControl',
  [DataValidationImeMode.OFF]: 'off',
  [DataValidationImeMode.ON]: 'on',
  [DataValidationImeMode.DISABLED]: 'disabled',
  [DataValidationImeMode.HIRAGANA]: 'hiragana',
  [DataValidationImeMode.FULL_KATAKANA]: 'fullKatakana',
  [DataValidationImeMode.HALF_KATAKANA]: 'halfKatakana',
  [DataValidationImeMode.FULL_ALPHA]: 'fullAlpha',
  [DataValidationImeMode.HALF_ALPHA]: 'halfAlpha',
  [DataValidationImeMode.FULL_HANGUL]: 'fullHangul',
  [DataValidationImeMode.HALF_HANGUL]: 'halfHangul',
};

export class DataValidationXmlSaver {
  private _escapeXml: (text: string) => string;

  constructor(escapeXml: (text: string) => string) {
    this._escapeXml = escapeXml;
  }

  formatDataValidationsXml(collection: DataValidationCollection): string {
    if (collection.count === 0) return '';

    const lines: string[] = [];
    const attrs: string[] = [`count="${collection.count}"`];

    if (collection.disablePrompts) {
      attrs.push('disablePrompts="1"');
    }
    if (collection.xWindow != null) {
      attrs.push(`xWindow="${collection.xWindow}"`);
    }
    if (collection.yWindow != null) {
      attrs.push(`yWindow="${collection.yWindow}"`);
    }

    lines.push(`<dataValidations ${attrs.join(' ')}>`);

    for (const dv of collection) {
      lines.push(this._formatDataValidationXml(dv));
    }

    lines.push('</dataValidations>');
    return lines.join('');
  }

  private _formatDataValidationXml(dv: DataValidation): string {
    const attrs: string[] = [];

    // Type
    const typeStr = TYPE_TO_XML[dv.type] ?? 'none';
    if (typeStr !== 'none') {
      attrs.push(`type="${typeStr}"`);
    }

    // Operator (only emit if not default 'between' or if type needs it)
    const opStr = OPERATOR_TO_XML[dv.operator] ?? 'between';
    if (opStr !== 'between') {
      attrs.push(`operator="${opStr}"`);
    }

    // Alert style
    const alertStr = ALERT_TO_XML[dv.alertStyle] ?? 'stop';
    if (alertStr !== 'stop') {
      attrs.push(`errorStyle="${alertStr}"`);
    }

    // IME mode
    if (dv.imeMode !== DataValidationImeMode.NO_CONTROL) {
      const imeStr = IME_TO_XML[dv.imeMode];
      if (imeStr) attrs.push(`imeMode="${imeStr}"`);
    }

    // Boolean attrs
    if (dv.allowBlank) {
      attrs.push('allowBlank="1"');
    }

    // showDropDown: ECMA-376 inversion - showDropDown="1" means HIDE dropdown
    if (!dv.showDropdown) {
      attrs.push('showDropDown="1"');
    }

    if (dv.showInputMessage) {
      attrs.push('showInputMessage="1"');
    }

    if (dv.showErrorMessage) {
      attrs.push('showErrorMessage="1"');
    }

    // String attrs
    if (dv.errorTitle) {
      attrs.push(`errorTitle="${this._escapeXml(dv.errorTitle)}"`);
    }

    if (dv.errorMessage) {
      attrs.push(`error="${this._escapeXml(dv.errorMessage)}"`);
    }

    if (dv.inputTitle) {
      attrs.push(`promptTitle="${this._escapeXml(dv.inputTitle)}"`);
    }

    if (dv.inputMessage) {
      attrs.push(`prompt="${this._escapeXml(dv.inputMessage)}"`);
    }

    // sqref
    attrs.push(`sqref="${dv.sqref ?? ''}"`);

    let xml = `<dataValidation ${attrs.join(' ')}>`;

    // Formulas
    if (dv.formula1 != null) {
      xml += `<formula1>${this._escapeXml(dv.formula1)}</formula1>`;
    }
    if (dv.formula2 != null) {
      xml += `<formula2>${this._escapeXml(dv.formula2)}</formula2>`;
    }

    xml += '</dataValidation>';
    return xml;
  }
}
