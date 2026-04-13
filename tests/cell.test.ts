/**
 * Unit tests for src/core/Cell.ts
 */
import { Cell, CellValue } from '../src/core/Cell';
import { Style } from '../src/styling/Style';

describe('Cell', () => {
  it('defaults to null value, no formula', () => {
    const c = new Cell();
    expect(c.value).toBeNull();
    expect(c.formula).toBeNull();
    expect(c.isEmpty()).toBe(true);
    expect(c.dataType).toBe('none');
  });

  it('constructor with value and formula', () => {
    const c = new Cell(42, '=A1+B1');
    expect(c.value).toBe(42);
    expect(c.formula).toBe('=A1+B1');
    expect(c.hasFormula()).toBe(true);
    expect(c.isEmpty()).toBe(false);
  });

  // -- data type detection --
  it('detects numeric', () => {
    const c = new Cell(3.14);
    expect(c.dataType).toBe('numeric');
    expect(c.isNumericValue()).toBe(true);
    expect(c.is_numeric_value()).toBe(true);
  });

  it('detects string', () => {
    const c = new Cell('hello');
    expect(c.dataType).toBe('string');
    expect(c.isTextValue()).toBe(true);
  });

  it('detects boolean (native)', () => {
    const c = new Cell(true);
    expect(c.dataType).toBe('boolean');
    expect(c.isBooleanValue()).toBe(true);
  });

  it('detects boolean from string "TRUE"/"FALSE"', () => {
    const c = new Cell('TRUE');
    expect(c.dataType).toBe('boolean');
    const c2 = new Cell('false');
    expect(c2.dataType).toBe('boolean');
  });

  it('detects datetime', () => {
    const c = new Cell(new Date(2024, 0, 1));
    expect(c.dataType).toBe('datetime');
    expect(c.isDateTimeValue()).toBe(true);
  });

  it('snake_case data_type alias', () => {
    const c = new Cell(10);
    expect(c.data_type).toBe('numeric');
  });

  // -- value methods --
  it('putValue / getValue', () => {
    const c = new Cell();
    c.putValue(99);
    expect(c.getValue()).toBe(99);
    c.put_value('abc');
    expect(c.get_value()).toBe('abc');
  });

  it('clear / clearValue / clearFormula', () => {
    const c = new Cell(10, '=SUM(A1)');
    c.clearValue();
    expect(c.value).toBeNull();
    expect(c.formula).toBe('=SUM(A1)');
    c.clearFormula();
    expect(c.formula).toBeNull();
  });

  it('clear() resets both value and formula', () => {
    const c = new Cell(10, '=SUM(A1)');
    c.clear();
    expect(c.value).toBeNull();
    expect(c.formula).toBeNull();
  });

  // -- comments --
  it('setComment / getComment / hasComment', () => {
    const c = new Cell();
    expect(c.hasComment()).toBe(false);
    c.setComment('Note', 'Author1');
    expect(c.hasComment()).toBe(true);
    const comment = c.getComment();
    expect(comment).not.toBeNull();
    expect(comment!.text).toBe('Note');
    expect(comment!.author).toBe('Author1');
  });

  it('clearComment removes comment', () => {
    const c = new Cell();
    c.setComment('Test');
    c.clearComment();
    expect(c.hasComment()).toBe(false);
    expect(c.getComment()).toBeNull();
  });

  it('setCommentSize updates size', () => {
    const c = new Cell();
    c.setComment('Note');
    c.setCommentSize(200, 100);
    const size = c.getCommentSize();
    expect(size).toEqual([200, 100]);
  });

  it('setCommentSize throws if no comment', () => {
    const c = new Cell();
    expect(() => c.setCommentSize(10, 10)).toThrow();
  });

  it('getCommentSize returns null when no comment', () => {
    const c = new Cell();
    expect(c.getCommentSize()).toBeNull();
  });

  // -- style --
  it('has a default style', () => {
    const c = new Cell();
    expect(c.style).toBeInstanceOf(Style);
  });

  it('applyStyle / getStyle / clearStyle', () => {
    const c = new Cell();
    const s = new Style();
    s.font.bold = true;
    c.applyStyle(s);
    expect(c.getStyle().font.bold).toBe(true);
    c.clearStyle();
    expect(c.getStyle().font.bold).toBe(false);
  });

  // -- toString --
  it('toString returns value as string', () => {
    expect(new Cell(42).toString()).toBe('42');
    expect(new Cell('hello').toString()).toBe('hello');
    expect(new Cell().toString()).toBe('');
  });

  // -- snake_case aliases --
  it('snake_case comment aliases', () => {
    const c = new Cell();
    c.set_comment('Hi', 'Me');
    expect(c.has_comment()).toBe(true);
    expect(c.get_comment()!.text).toBe('Hi');
    c.clear_comment();
    expect(c.has_comment()).toBe(false);
  });

  it('snake_case value aliases', () => {
    const c = new Cell();
    expect(c.is_empty()).toBe(true);
    c.value = 5;
    c.clear_value();
    expect(c.value).toBeNull();
  });
});
