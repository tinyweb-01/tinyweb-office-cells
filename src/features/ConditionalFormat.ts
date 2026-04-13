/**
 * Conditional Formatting Module
 *
 * Provides classes for Excel conditional formatting according to ECMA-376 specification.
 */

import { Font, Fill, Borders, Alignment } from '../styling/Style';

// ==================== ConditionalFormat Class ====================

export class ConditionalFormat {
  // Type and identity
  private _type: string | null = null;
  private _range: string | null = null;
  private _stopIfTrue = false;
  private _priority = 0;

  // Cell value rule
  private _operator: string | null = null;
  private _formula1: string | null = null;
  private _formula2: string | null = null;

  // Text rule
  private _textOperator: string | null = null;
  private _textFormula: string | null = null;

  // Date rule
  private _dateOperator: string | null = null;
  private _dateFormula: string | null = null;

  // Duplicate/unique
  private _duplicate: boolean | null = null;

  // Top/Bottom
  private _top: boolean | null = null;
  private _percent = false;
  private _rank = 10;

  // Above/Below average
  private _above: boolean | null = null;
  private _stdDev = 0;

  // Color scale
  private _colorScaleType: string | null = null;
  private _minColor: string | null = null;
  private _midColor: string | null = null;
  private _maxColor: string | null = null;

  // Data bar
  private _barColor: string | null = null;
  private _negativeColor: string | null = null;
  private _showBorder = false;
  private _direction: string | null = null;
  private _barLength: [number, number] | null = null;

  // Icon set
  private _iconSetType: string | null = null;
  private _reverseIcons = false;
  private _showIconOnly = false;

  // Generic formula
  private _formula: string | null = null;

  // DXF styling
  private _font: Font | null = null;
  private _border: Borders | null = null;
  private _fill: Fill | null = null;
  private _alignment: Alignment | null = null;
  private _numberFormat: string | null = null;

  // Internal: dxfId for XML round-trip
  _dxfId: number | null = null;

  // ---- Type and identity ----

  get type(): string | null { return this._type; }
  set type(value: string | null) { this._type = value; }

  get range(): string | null { return this._range; }
  set range(value: string | null) { this._range = value; }

  get stopIfTrue(): boolean { return this._stopIfTrue; }
  set stopIfTrue(value: boolean) { this._stopIfTrue = value; }

  get priority(): number { return this._priority; }
  set priority(value: number) { this._priority = value; }

  // ---- Cell value rule ----

  get operator(): string | null { return this._operator; }
  set operator(value: string | null) { this._operator = value; }

  get formula1(): string | null { return this._formula1; }
  set formula1(value: string | null) { this._formula1 = value; }

  get formula2(): string | null { return this._formula2; }
  set formula2(value: string | null) { this._formula2 = value; }

  // ---- Text rule ----

  get textOperator(): string | null { return this._textOperator; }
  set textOperator(value: string | null) { this._textOperator = value; }

  get textFormula(): string | null { return this._textFormula; }
  set textFormula(value: string | null) { this._textFormula = value; }

  // ---- Date rule ----

  get dateOperator(): string | null { return this._dateOperator; }
  set dateOperator(value: string | null) { this._dateOperator = value; }

  get dateFormula(): string | null { return this._dateFormula; }
  set dateFormula(value: string | null) { this._dateFormula = value; }

  // ---- Duplicate/unique ----

  get duplicate(): boolean | null { return this._duplicate; }
  set duplicate(value: boolean | null) { this._duplicate = value; }

  // ---- Top/Bottom ----

  get top(): boolean | null { return this._top; }
  set top(value: boolean | null) { this._top = value; }

  get percent(): boolean { return this._percent; }
  set percent(value: boolean) { this._percent = value; }

  get rank(): number { return this._rank; }
  set rank(value: number) { this._rank = value; }

  // ---- Above/Below average ----

  get above(): boolean | null { return this._above; }
  set above(value: boolean | null) { this._above = value; }

  get stdDev(): number { return this._stdDev; }
  set stdDev(value: number) { this._stdDev = value; }

  // ---- Color scale ----

  get colorScaleType(): string | null { return this._colorScaleType; }
  set colorScaleType(value: string | null) { this._colorScaleType = value; }

  get minColor(): string | null { return this._minColor; }
  set minColor(value: string | null) { this._minColor = value; }

  get midColor(): string | null { return this._midColor; }
  set midColor(value: string | null) { this._midColor = value; }

  get maxColor(): string | null { return this._maxColor; }
  set maxColor(value: string | null) { this._maxColor = value; }

  // ---- Data bar ----

  get barColor(): string | null { return this._barColor; }
  set barColor(value: string | null) { this._barColor = value; }

  get negativeColor(): string | null { return this._negativeColor; }
  set negativeColor(value: string | null) { this._negativeColor = value; }

  get showBorder(): boolean { return this._showBorder; }
  set showBorder(value: boolean) { this._showBorder = value; }

  get direction(): string | null { return this._direction; }
  set direction(value: string | null) { this._direction = value; }

  get barLength(): [number, number] | null { return this._barLength; }
  set barLength(value: [number, number] | null) { this._barLength = value; }

  // ---- Icon set ----

  get iconSetType(): string | null { return this._iconSetType; }
  set iconSetType(value: string | null) { this._iconSetType = value; }

  get reverseIcons(): boolean { return this._reverseIcons; }
  set reverseIcons(value: boolean) { this._reverseIcons = value; }

  get showIconOnly(): boolean { return this._showIconOnly; }
  set showIconOnly(value: boolean) { this._showIconOnly = value; }

  // ---- Generic formula ----

  get formula(): string | null { return this._formula; }
  set formula(value: string | null) { this._formula = value; }

  // ---- DXF styling (lazily created) ----

  get font(): Font {
    if (!this._font) this._font = new Font();
    return this._font;
  }

  get border(): Borders {
    if (!this._border) this._border = new Borders();
    return this._border;
  }

  get fill(): Fill {
    if (!this._fill) this._fill = new Fill();
    return this._fill;
  }

  get alignment(): Alignment {
    if (!this._alignment) this._alignment = new Alignment();
    return this._alignment;
  }

  get numberFormat(): string | null { return this._numberFormat; }
  set numberFormat(value: string | null) { this._numberFormat = value; }

  hasFont(): boolean { return this._font !== null; }
  hasFill(): boolean { return this._fill !== null; }
  hasBorder(): boolean { return this._border !== null; }
  hasAlignment(): boolean { return this._alignment !== null; }
}

// ==================== ConditionalFormatCollection ====================

export class ConditionalFormatCollection {
  private _formats: ConditionalFormat[] = [];

  get count(): number {
    return this._formats.length;
  }

  add(): ConditionalFormat {
    const cf = new ConditionalFormat();
    cf.priority = this._formats.length + 1;
    this._formats.push(cf);
    return cf;
  }

  getByIndex(index: number): ConditionalFormat | null {
    if (index >= 0 && index < this._formats.length) {
      return this._formats[index];
    }
    return null;
  }

  getByRange(rangeStr: string): ConditionalFormat[] {
    return this._formats.filter((cf) => cf.range === rangeStr);
  }

  remove(cf: ConditionalFormat): boolean {
    const idx = this._formats.indexOf(cf);
    if (idx >= 0) {
      this._formats.splice(idx, 1);
      return true;
    }
    return false;
  }

  clear(): void {
    this._formats = [];
  }

  get(index: number): ConditionalFormat {
    if (index < 0 || index >= this._formats.length) {
      throw new Error(`Index ${index} out of range`);
    }
    return this._formats[index];
  }

  [Symbol.iterator](): Iterator<ConditionalFormat> {
    let i = 0;
    const items = this._formats;
    return {
      next(): IteratorResult<ConditionalFormat> {
        if (i < items.length) {
          return { value: items[i++], done: false };
        }
        return { value: undefined as any, done: true };
      },
    };
  }

  get length(): number {
    return this._formats.length;
  }
}
