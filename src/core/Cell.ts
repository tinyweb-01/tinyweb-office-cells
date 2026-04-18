/**
 * tinyweb-office-cells - Cell Module
 *
 * Represents a single cell in a worksheet.
 */

import { Style } from '../styling/Style';

// ─── Comment type ──────────────────────────────────────────────────────────

export interface CellComment {
  text: string;
  author: string;
  width?: number | null;
  height?: number | null;
}

// ─── Data type literals ────────────────────────────────────────────────────

export type CellDataType = 'none' | 'boolean' | 'numeric' | 'datetime' | 'string' | 'unknown';

// ─── Cell value type ───────────────────────────────────────────────────────

export type CellValue = null | undefined | string | number | boolean | Date;

// ─── Cell class ────────────────────────────────────────────────────────────

/**
 * Represents a single cell in a worksheet.
 */
export class Cell {
  private _value: CellValue;
  private _formula: string | null;
  private _style: Style;
  private _comment: CellComment | null;
  /** Internal use for saving */
  _styleIndex: number;

  constructor(value: CellValue = null, formula: string | null = null) {
    this._value = value;
    this._formula = formula;
    this._style = new Style();
    this._comment = null;
    this._styleIndex = 0;
  }

  // ── Properties ──────────────────────────────────────────────────────────

  get value(): CellValue { return this._value; }
  set value(val: CellValue) { this._value = val; }

  get formula(): string | null { return this._formula; }
  set formula(val: string | null) { this._formula = val; }

  get style(): Style { return this._style; }
  set style(val: Style) { this._style = val; }

  get comment(): CellComment | null { return this._comment; }

  // ── Data type detection ─────────────────────────────────────────────────

  get dataType(): CellDataType {
    return this.getDataType();
  }
  /** snake_case alias */
  get data_type(): CellDataType {
    return this.getDataType();
  }

  private getDataType(): CellDataType {
    if (this._value === null || this._value === undefined) return 'none';
    if (typeof this._value === 'boolean') return 'boolean';
    if (typeof this._value === 'number') return 'numeric';
    if (this._value instanceof Date) return 'datetime';
    if (typeof this._value === 'string') {
      if (this._value.toUpperCase() === 'TRUE' || this._value.toUpperCase() === 'FALSE') {
        return 'boolean';
      }
      return 'string';
    }
    return 'unknown';
  }

  // ── Cell value methods ──────────────────────────────────────────────────

  isEmpty(): boolean { return this._value === null || this._value === undefined; }
  is_empty(): boolean { return this.isEmpty(); }

  clearValue(): void { this._value = null; }
  clear_value(): void { this.clearValue(); }

  clearFormula(): void { this._formula = null; }
  clear_formula(): void { this.clearFormula(); }

  clear(): void {
    this._value = null;
    this._formula = null;
  }

  // ── Comment methods ─────────────────────────────────────────────────────

  setComment(text: string, author = 'None', width?: number | null, height?: number | null): void {
    if (author === '') author = 'None';
    this._comment = { text, author, width: width ?? null, height: height ?? null };
  }
  /** snake_case alias */
  set_comment(text: string, author = 'None', width?: number | null, height?: number | null): void {
    this.setComment(text, author, width, height);
  }
  /** .NET-style alias */
  addComment(text: string, author = 'None'): void {
    this.setComment(text, author);
  }
  add_comment(text: string, author = 'None'): void {
    this.addComment(text, author);
  }

  getComment(): CellComment | null { return this._comment; }
  get_comment(): CellComment | null { return this.getComment(); }

  clearComment(): void { this._comment = null; }
  clear_comment(): void { this.clearComment(); }
  removeComment(): void { this.clearComment(); }
  remove_comment(): void { this.clearComment(); }

  hasComment(): boolean { return this._comment !== null; }
  has_comment(): boolean { return this.hasComment(); }

  setCommentSize(width: number, height: number): void {
    if (!this._comment) throw new Error('Cell has no comment. Call setComment() first.');
    this._comment.width = width;
    this._comment.height = height;
  }
  set_comment_size(width: number, height: number): void { this.setCommentSize(width, height); }

  getCommentSize(): [number, number] | null {
    if (!this._comment) return null;
    if (this._comment.width != null && this._comment.height != null) {
      return [this._comment.width, this._comment.height];
    }
    return null;
  }
  get_comment_size(): [number, number] | null { return this.getCommentSize(); }

  // ── Style methods ───────────────────────────────────────────────────────

  applyStyle(style: Style): void { this._style = style; }
  apply_style(style: Style): void { this.applyStyle(style); }
  setStyle(style: Style): void { this.applyStyle(style); }
  set_style(style: Style): void { this.applyStyle(style); }

  getStyle(): Style { return this._style; }
  get_style(): Style { return this.getStyle(); }

  clearStyle(): void { this._style = new Style(); }
  clear_style(): void { this.clearStyle(); }

  // ── Formula methods ─────────────────────────────────────────────────────

  hasFormula(): boolean { return this._formula !== null; }
  has_formula(): boolean { return this.hasFormula(); }

  isNumericValue(): boolean { return typeof this._value === 'number'; }
  is_numeric_value(): boolean { return this.isNumericValue(); }

  isTextValue(): boolean { return typeof this._value === 'string'; }
  is_text_value(): boolean { return this.isTextValue(); }

  isBooleanValue(): boolean { return this.dataType === 'boolean'; }
  is_boolean_value(): boolean { return this.isBooleanValue(); }

  isDateTimeValue(): boolean { return this.dataType === 'datetime'; }
  is_date_time_value(): boolean { return this.isDateTimeValue(); }

  /** Sets the value of the cell (.NET compatibility). */
  putValue(value: CellValue): void { this._value = value; }
  put_value(value: CellValue): void { this.putValue(value); }

  getValue(): CellValue { return this._value; }
  get_value(): CellValue { return this.getValue(); }

  // ── String representation ───────────────────────────────────────────────

  toString(): string {
    return this._value != null ? String(this._value) : '';
  }
}
