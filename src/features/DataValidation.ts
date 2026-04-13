/**
 * Data Validation Module
 *
 * Provides classes for Excel data validation according to ECMA-376 specification.
 * Data validation allows controlling what data can be entered into cells.
 *
 * References:
 * - ECMA-376 Part 4, Section 3.3.1.30 (dataValidation)
 */

// ==================== Enums ====================

export enum DataValidationType {
  NONE = 0,
  WHOLE_NUMBER = 1,
  DECIMAL = 2,
  LIST = 3,
  DATE = 4,
  TIME = 5,
  TEXT_LENGTH = 6,
  CUSTOM = 7,
}

export enum DataValidationOperator {
  BETWEEN = 0,
  NOT_BETWEEN = 1,
  EQUAL = 2,
  NOT_EQUAL = 3,
  GREATER_THAN = 4,
  LESS_THAN = 5,
  GREATER_THAN_OR_EQUAL = 6,
  LESS_THAN_OR_EQUAL = 7,
}

export enum DataValidationAlertStyle {
  STOP = 0,
  WARNING = 1,
  INFORMATION = 2,
}

export enum DataValidationImeMode {
  NO_CONTROL = 0,
  OFF = 1,
  ON = 2,
  DISABLED = 3,
  HIRAGANA = 4,
  FULL_KATAKANA = 5,
  HALF_KATAKANA = 6,
  FULL_ALPHA = 7,
  HALF_ALPHA = 8,
  FULL_HANGUL = 9,
  HALF_HANGUL = 10,
}

// ==================== DataValidation Class ====================

export class DataValidation {
  private _sqref: string | null;
  private _type: DataValidationType;
  private _operator: DataValidationOperator;
  private _formula1: string | null;
  private _formula2: string | null;
  private _alertStyle: DataValidationAlertStyle;
  private _showErrorMessage: boolean;
  private _errorTitle: string | null;
  private _errorMessage: string | null;
  private _showInputMessage: boolean;
  private _inputTitle: string | null;
  private _inputMessage: string | null;
  private _allowBlank: boolean;
  private _showDropdown: boolean;
  private _imeMode: DataValidationImeMode;

  constructor(sqref: string | null = null) {
    this._sqref = sqref;
    this._type = DataValidationType.NONE;
    this._operator = DataValidationOperator.BETWEEN;
    this._formula1 = null;
    this._formula2 = null;
    this._alertStyle = DataValidationAlertStyle.STOP;
    this._showErrorMessage = true;
    this._errorTitle = null;
    this._errorMessage = null;
    this._showInputMessage = false;
    this._inputTitle = null;
    this._inputMessage = null;
    this._allowBlank = true;
    this._showDropdown = true;
    this._imeMode = DataValidationImeMode.NO_CONTROL;
  }

  // ---- Cell Range ----

  get sqref(): string | null {
    return this._sqref;
  }
  set sqref(value: string | null) {
    this._sqref = value;
  }

  get ranges(): string[] {
    if (this._sqref) {
      return this._sqref.split(/\s+/);
    }
    return [];
  }

  // ---- Validation Type and Operator ----

  get type(): DataValidationType {
    return this._type;
  }
  set type(value: DataValidationType) {
    this._type = value;
  }

  get operator(): DataValidationOperator {
    return this._operator;
  }
  set operator(value: DataValidationOperator) {
    this._operator = value;
  }

  // ---- Formulas ----

  get formula1(): string | null {
    return this._formula1;
  }
  set formula1(value: string | null) {
    this._formula1 = value;
  }

  get formula2(): string | null {
    return this._formula2;
  }
  set formula2(value: string | null) {
    this._formula2 = value;
  }

  // ---- Error Alert Settings ----

  get alertStyle(): DataValidationAlertStyle {
    return this._alertStyle;
  }
  set alertStyle(value: DataValidationAlertStyle) {
    this._alertStyle = value;
  }

  get showErrorMessage(): boolean {
    return this._showErrorMessage;
  }
  set showErrorMessage(value: boolean) {
    this._showErrorMessage = value;
  }

  /** Alias for showErrorMessage */
  get showError(): boolean {
    return this._showErrorMessage;
  }
  set showError(value: boolean) {
    this._showErrorMessage = value;
  }

  get errorTitle(): string | null {
    return this._errorTitle;
  }
  set errorTitle(value: string | null) {
    this._errorTitle = value;
  }

  get errorMessage(): string | null {
    return this._errorMessage;
  }
  set errorMessage(value: string | null) {
    this._errorMessage = value;
  }

  /** Alias for errorMessage */
  get error(): string | null {
    return this._errorMessage;
  }
  set error(value: string | null) {
    this._errorMessage = value;
  }

  // ---- Input Message Settings ----

  get showInputMessage(): boolean {
    return this._showInputMessage;
  }
  set showInputMessage(value: boolean) {
    this._showInputMessage = value;
  }

  /** Alias for showInputMessage */
  get showInput(): boolean {
    return this._showInputMessage;
  }
  set showInput(value: boolean) {
    this._showInputMessage = value;
  }

  get inputTitle(): string | null {
    return this._inputTitle;
  }
  set inputTitle(value: string | null) {
    this._inputTitle = value;
  }

  /** Alias for inputTitle */
  get promptTitle(): string | null {
    return this._inputTitle;
  }
  set promptTitle(value: string | null) {
    this._inputTitle = value;
  }

  get inputMessage(): string | null {
    return this._inputMessage;
  }
  set inputMessage(value: string | null) {
    this._inputMessage = value;
  }

  /** Alias for inputMessage */
  get prompt(): string | null {
    return this._inputMessage;
  }
  set prompt(value: string | null) {
    this._inputMessage = value;
  }

  // ---- Other Settings ----

  get allowBlank(): boolean {
    return this._allowBlank;
  }
  set allowBlank(value: boolean) {
    this._allowBlank = value;
  }

  /** Alias for allowBlank */
  get ignoreBlank(): boolean {
    return this._allowBlank;
  }
  set ignoreBlank(value: boolean) {
    this._allowBlank = value;
  }

  /**
   * Whether to show the in-cell dropdown for list validations.
   * Note: In ECMA-376, showDropDown="1" means HIDE dropdown (counterintuitive).
   */
  get showDropdown(): boolean {
    return this._showDropdown;
  }
  set showDropdown(value: boolean) {
    this._showDropdown = value;
  }

  /** Alias for showDropdown */
  get inCellDropdown(): boolean {
    return this._showDropdown;
  }
  set inCellDropdown(value: boolean) {
    this._showDropdown = value;
  }

  get imeMode(): DataValidationImeMode {
    return this._imeMode;
  }
  set imeMode(value: DataValidationImeMode) {
    this._imeMode = value;
  }

  // ---- Methods ----

  add(
    validationType: DataValidationType,
    alertStyle?: DataValidationAlertStyle,
    operator?: DataValidationOperator,
    formula1?: string | null,
    formula2?: string | null,
  ): void {
    this._type = validationType;
    if (alertStyle !== undefined) this._alertStyle = alertStyle;
    if (operator !== undefined) this._operator = operator;
    if (formula1 !== undefined) this._formula1 = formula1;
    if (formula2 !== undefined) this._formula2 = formula2;
  }

  modify(options: {
    validationType?: DataValidationType;
    alertStyle?: DataValidationAlertStyle;
    operator?: DataValidationOperator;
    formula1?: string | null;
    formula2?: string | null;
  }): void {
    if (options.validationType !== undefined) this._type = options.validationType;
    if (options.alertStyle !== undefined) this._alertStyle = options.alertStyle;
    if (options.operator !== undefined) this._operator = options.operator;
    if (options.formula1 !== undefined) this._formula1 = options.formula1;
    if (options.formula2 !== undefined) this._formula2 = options.formula2;
  }

  delete(): void {
    this._type = DataValidationType.NONE;
    this._operator = DataValidationOperator.BETWEEN;
    this._formula1 = null;
    this._formula2 = null;
    this._alertStyle = DataValidationAlertStyle.STOP;
    this._showErrorMessage = true;
    this._errorTitle = null;
    this._errorMessage = null;
    this._showInputMessage = false;
    this._inputTitle = null;
    this._inputMessage = null;
    this._allowBlank = true;
    this._showDropdown = true;
    this._imeMode = DataValidationImeMode.NO_CONTROL;
  }

  copy(): DataValidation {
    const dv = new DataValidation(this._sqref);
    dv._type = this._type;
    dv._operator = this._operator;
    dv._formula1 = this._formula1;
    dv._formula2 = this._formula2;
    dv._alertStyle = this._alertStyle;
    dv._showErrorMessage = this._showErrorMessage;
    dv._errorTitle = this._errorTitle;
    dv._errorMessage = this._errorMessage;
    dv._showInputMessage = this._showInputMessage;
    dv._inputTitle = this._inputTitle;
    dv._inputMessage = this._inputMessage;
    dv._allowBlank = this._allowBlank;
    dv._showDropdown = this._showDropdown;
    dv._imeMode = this._imeMode;
    return dv;
  }
}

// ==================== DataValidationCollection ====================

export class DataValidationCollection {
  private _validations: DataValidation[] = [];
  private _disablePrompts = false;
  private _xWindow: number | null = null;
  private _yWindow: number | null = null;

  get count(): number {
    return this._validations.length;
  }

  get disablePrompts(): boolean {
    return this._disablePrompts;
  }
  set disablePrompts(value: boolean) {
    this._disablePrompts = value;
  }

  get xWindow(): number | null {
    return this._xWindow;
  }
  set xWindow(value: number | null) {
    this._xWindow = value;
  }

  get yWindow(): number | null {
    return this._yWindow;
  }
  set yWindow(value: number | null) {
    this._yWindow = value;
  }

  add(
    sqref: string,
    validationType?: DataValidationType,
    operator?: DataValidationOperator,
    formula1?: string | null,
    formula2?: string | null,
  ): DataValidation {
    const dv = new DataValidation(sqref);
    if (validationType !== undefined) dv.type = validationType;
    if (operator !== undefined) dv.operator = operator;
    if (formula1 !== undefined) dv.formula1 = formula1;
    if (formula2 !== undefined) dv.formula2 = formula2;
    this._validations.push(dv);
    return dv;
  }

  addValidation(validation: DataValidation): void {
    this._validations.push(validation);
  }

  remove(validation: DataValidation): boolean {
    const idx = this._validations.indexOf(validation);
    if (idx >= 0) {
      this._validations.splice(idx, 1);
      return true;
    }
    return false;
  }

  removeAt(index: number): void {
    if (index >= 0 && index < this._validations.length) {
      this._validations.splice(index, 1);
    }
  }

  removeByRange(sqref: string): void {
    this._validations = this._validations.filter((v) => v.sqref !== sqref);
  }

  clear(): void {
    this._validations = [];
  }

  getValidation(cellRef: string): DataValidation | null {
    for (const v of this._validations) {
      if (v.sqref && this._cellInRange(cellRef, v.sqref)) {
        return v;
      }
    }
    return null;
  }

  private _cellInRange(cellRef: string, sqref: string): boolean {
    const ranges = sqref.split(/\s+/);
    for (const range of ranges) {
      if (range === cellRef) return true;
      const parts = range.split(':');
      if (parts.length === 2) {
        const [start, end] = parts;
        const cellCol = this._colFromRef(cellRef);
        const cellRow = this._rowFromRef(cellRef);
        const startCol = this._colFromRef(start);
        const startRow = this._rowFromRef(start);
        const endCol = this._colFromRef(end);
        const endRow = this._rowFromRef(end);
        if (
          cellCol >= startCol &&
          cellCol <= endCol &&
          cellRow >= startRow &&
          cellRow <= endRow
        ) {
          return true;
        }
      }
    }
    return false;
  }

  private _colFromRef(ref: string): number {
    const match = ref.match(/^([A-Z]+)/);
    if (!match) return 0;
    let col = 0;
    for (const ch of match[1]) {
      col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return col;
  }

  private _rowFromRef(ref: string): number {
    const match = ref.match(/(\d+)$/);
    return match ? parseInt(match[1], 10) : 0;
  }

  get(index: number): DataValidation {
    if (index < 0 || index >= this._validations.length) {
      throw new Error(`Index ${index} out of range`);
    }
    return this._validations[index];
  }

  [Symbol.iterator](): Iterator<DataValidation> {
    let i = 0;
    const items = this._validations;
    return {
      next(): IteratorResult<DataValidation> {
        if (i < items.length) {
          return { value: items[i++], done: false };
        }
        return { value: undefined as any, done: true };
      },
    };
  }
}
