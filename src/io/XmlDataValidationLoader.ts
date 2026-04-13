/**
 * XML Data Validation Loader
 *
 * Loads data validation settings from XLSX worksheet XML.
 */

import {
  DataValidation,
  DataValidationCollection,
  DataValidationType,
  DataValidationOperator,
  DataValidationAlertStyle,
  DataValidationImeMode,
} from '../features/DataValidation';

// Maps from XML string → enum
const TYPE_MAP: Record<string, DataValidationType> = {
  none: DataValidationType.NONE,
  whole: DataValidationType.WHOLE_NUMBER,
  decimal: DataValidationType.DECIMAL,
  list: DataValidationType.LIST,
  date: DataValidationType.DATE,
  time: DataValidationType.TIME,
  textLength: DataValidationType.TEXT_LENGTH,
  custom: DataValidationType.CUSTOM,
};

const OPERATOR_MAP: Record<string, DataValidationOperator> = {
  between: DataValidationOperator.BETWEEN,
  notBetween: DataValidationOperator.NOT_BETWEEN,
  equal: DataValidationOperator.EQUAL,
  notEqual: DataValidationOperator.NOT_EQUAL,
  greaterThan: DataValidationOperator.GREATER_THAN,
  lessThan: DataValidationOperator.LESS_THAN,
  greaterThanOrEqual: DataValidationOperator.GREATER_THAN_OR_EQUAL,
  lessThanOrEqual: DataValidationOperator.LESS_THAN_OR_EQUAL,
};

const ALERT_STYLE_MAP: Record<string, DataValidationAlertStyle> = {
  stop: DataValidationAlertStyle.STOP,
  warning: DataValidationAlertStyle.WARNING,
  information: DataValidationAlertStyle.INFORMATION,
};

const IME_MODE_MAP: Record<string, DataValidationImeMode> = {
  noControl: DataValidationImeMode.NO_CONTROL,
  off: DataValidationImeMode.OFF,
  on: DataValidationImeMode.ON,
  disabled: DataValidationImeMode.DISABLED,
  hiragana: DataValidationImeMode.HIRAGANA,
  fullKatakana: DataValidationImeMode.FULL_KATAKANA,
  halfKatakana: DataValidationImeMode.HALF_KATAKANA,
  fullAlpha: DataValidationImeMode.FULL_ALPHA,
  halfAlpha: DataValidationImeMode.HALF_ALPHA,
  fullHangul: DataValidationImeMode.FULL_HANGUL,
  halfHangul: DataValidationImeMode.HALF_HANGUL,
};

function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}

function attr(elem: any, name: string, defaultVal = ''): string {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : defaultVal;
}

export class DataValidationXmlLoader {
  loadDataValidations(wsRoot: any, collection: DataValidationCollection): void {
    const dvs = wsRoot?.dataValidations;
    if (!dvs) return;

    // Load collection-level attributes
    const disablePrompts = attr(dvs, 'disablePrompts', '');
    if (disablePrompts === '1' || disablePrompts === 'true') {
      collection.disablePrompts = true;
    }
    const xWindow = attr(dvs, 'xWindow', '');
    if (xWindow) collection.xWindow = parseInt(xWindow, 10);
    const yWindow = attr(dvs, 'yWindow', '');
    if (yWindow) collection.yWindow = parseInt(yWindow, 10);

    const dvElems = ensureArray(dvs.dataValidation);
    for (const dvElem of dvElems) {
      this._loadDataValidation(dvElem, collection);
    }
  }

  private _loadDataValidation(dvElem: any, collection: DataValidationCollection): void {
    const sqref = attr(dvElem, 'sqref', '');
    if (!sqref) return;

    const dv = new DataValidation(sqref);

    // Type
    const typeStr = attr(dvElem, 'type', 'none');
    dv.type = TYPE_MAP[typeStr] ?? DataValidationType.NONE;

    // Operator
    const opStr = attr(dvElem, 'operator', 'between');
    dv.operator = OPERATOR_MAP[opStr] ?? DataValidationOperator.BETWEEN;

    // Alert style
    const alertStr = attr(dvElem, 'errorStyle', 'stop');
    dv.alertStyle = ALERT_STYLE_MAP[alertStr] ?? DataValidationAlertStyle.STOP;

    // IME mode
    const imeStr = attr(dvElem, 'imeMode', '');
    if (imeStr && IME_MODE_MAP[imeStr] !== undefined) {
      dv.imeMode = IME_MODE_MAP[imeStr];
    }

    // Boolean attributes
    const allowBlank = attr(dvElem, 'allowBlank', '');
    dv.allowBlank = allowBlank === '1' || allowBlank === 'true';

    const showInputMsg = attr(dvElem, 'showInputMessage', '');
    dv.showInputMessage = showInputMsg === '1' || showInputMsg === 'true';

    const showErrorMsg = attr(dvElem, 'showErrorMessage', '');
    dv.showErrorMessage = showErrorMsg === '1' || showErrorMsg === 'true';

    // showDropDown: In ECMA-376, showDropDown="1" means HIDE the dropdown!
    const showDropDown = attr(dvElem, 'showDropDown', '');
    if (showDropDown === '1' || showDropDown === 'true') {
      dv.showDropdown = false; // inverted!
    } else {
      dv.showDropdown = true;
    }

    // String attributes
    const errorTitle = attr(dvElem, 'errorTitle', '');
    if (errorTitle) dv.errorTitle = errorTitle;

    const errorAttr = attr(dvElem, 'error', '');
    if (errorAttr) dv.errorMessage = errorAttr;

    const promptTitle = attr(dvElem, 'promptTitle', '');
    if (promptTitle) dv.inputTitle = promptTitle;

    const promptAttr = attr(dvElem, 'prompt', '');
    if (promptAttr) dv.inputMessage = promptAttr;

    // Formulas
    const formula1 = dvElem?.formula1;
    if (formula1 != null) {
      dv.formula1 = typeof formula1 === 'object' ? formula1['#text'] ?? String(formula1) : String(formula1);
    }

    const formula2 = dvElem?.formula2;
    if (formula2 != null) {
      dv.formula2 = typeof formula2 === 'object' ? formula2['#text'] ?? String(formula2) : String(formula2);
    }

    collection.addValidation(dv);
  }
}
