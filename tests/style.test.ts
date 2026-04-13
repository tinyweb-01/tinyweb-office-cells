/**
 * Unit tests for src/styling/Style.ts
 */
import {
  Font,
  Fill,
  Border,
  BorderType,
  Borders,
  Alignment,
  NumberFormat,
  Protection,
  Style,
} from '../src/styling/Style';

// ---------------------------------------------------------------------------
// Font
// ---------------------------------------------------------------------------
describe('Font', () => {
  it('uses sensible defaults', () => {
    const f = new Font();
    expect(f.name).toBe('Calibri');
    expect(f.size).toBe(11);
    expect(f.color).toBe('FF000000');
    expect(f.bold).toBe(false);
    expect(f.italic).toBe(false);
    expect(f.underline).toBe(false);
    expect(f.strikethrough).toBe(false);
  });

  it('accepts constructor overrides via options object', () => {
    const f = new Font({ name: 'Arial', size: 14, color: 'FFFF0000', bold: true, italic: true, underline: true, strikethrough: true });
    expect(f.name).toBe('Arial');
    expect(f.size).toBe(14);
    expect(f.color).toBe('FFFF0000');
    expect(f.bold).toBe(true);
    expect(f.italic).toBe(true);
    expect(f.underline).toBe(true);
    expect(f.strikethrough).toBe(true);
  });

  it('exposes .NET-style aliases for bold/italic', () => {
    const f = new Font();
    f.isBold = true;
    expect(f.is_bold).toBe(true);
    expect(f.bold).toBe(true);

    f.is_italic = true;
    expect(f.isItalic).toBe(true);
    expect(f.italic).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Fill
// ---------------------------------------------------------------------------
describe('Fill', () => {
  it('defaults to no fill', () => {
    const fill = new Fill();
    expect(fill.patternType).toBe('none');
    expect(fill.foregroundColor).toBe('FFFFFFFF');
    expect(fill.backgroundColor).toBe('FFFFFFFF');
  });

  it('setSolidFill sets pattern and foreground', () => {
    const fill = new Fill();
    fill.setSolidFill('FF00FF00');
    expect(fill.patternType).toBe('solid');
    expect(fill.foregroundColor).toBe('FF00FF00');
  });

  it('setPatternFill changes all properties', () => {
    const fill = new Fill();
    fill.setPatternFill('gray125', 'FF111111', 'FF222222');
    expect(fill.patternType).toBe('gray125');
    expect(fill.foregroundColor).toBe('FF111111');
    expect(fill.backgroundColor).toBe('FF222222');
  });

  it('setNoFill resets', () => {
    const fill = new Fill();
    fill.setSolidFill('FFFF0000');
    fill.setNoFill();
    expect(fill.patternType).toBe('none');
    expect(fill.foregroundColor).toBe('FFFFFFFF');
    expect(fill.backgroundColor).toBe('FFFFFFFF');
  });

  it('snake_case aliases work', () => {
    const fill = new Fill();
    fill.set_solid_fill('FFAABBCC');
    expect(fill.pattern_type).toBe('solid');
    expect(fill.foreground_color).toBe('FFAABBCC');
    fill.set_no_fill();
    expect(fill.patternType).toBe('none');
  });
});

// ---------------------------------------------------------------------------
// Border & Borders
// ---------------------------------------------------------------------------
describe('Border', () => {
  it('defaults', () => {
    const b = new Border();
    expect(b.lineStyle).toBe('none');
    expect(b.color).toBe('FF000000');
    expect(b.weight).toBe(1);
  });

  it('accepts options + snake_case alias', () => {
    const b = new Border({ lineStyle: 'thin', color: 'FFFF0000', weight: 2 });
    expect(b.line_style).toBe('thin');
    expect(b.color).toBe('FFFF0000');
    expect(b.weight).toBe(2);
  });
});

describe('Borders', () => {
  it('has all sides', () => {
    const b = new Borders();
    expect(b.left).toBeInstanceOf(Border);
    expect(b.right).toBeInstanceOf(Border);
    expect(b.top).toBeInstanceOf(Border);
    expect(b.bottom).toBeInstanceOf(Border);
    expect(b.diagonal).toBeInstanceOf(Border);
  });

  it('getByBorderType returns correct border', () => {
    const b = new Borders();
    b.left.lineStyle = 'thin';
    expect(b.getByBorderType(BorderType.LeftBorder).lineStyle).toBe('thin');
  });

  it('setBorder updates a border', () => {
    const b = new Borders();
    b.setBorder('top', 'medium', 'FF112233', 2);
    expect(b.top.lineStyle).toBe('medium');
    expect(b.top.color).toBe('FF112233');
    expect(b.top.weight).toBe(2);
  });

  it('setBorder "all" updates all four sides', () => {
    const b = new Borders();
    b.setBorder('all', 'thick', 'FFAABBCC', 3);
    for (const side of [b.top, b.bottom, b.left, b.right]) {
      expect(side.lineStyle).toBe('thick');
      expect(side.color).toBe('FFAABBCC');
      expect(side.weight).toBe(3);
    }
  });
});

// ---------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------
describe('Alignment', () => {
  it('defaults', () => {
    const a = new Alignment();
    expect(a.horizontal).toBe('general');
    expect(a.vertical).toBe('bottom');
    expect(a.wrapText).toBe(false);
    expect(a.indent).toBe(0);
    expect(a.textRotation).toBe(0);
    expect(a.shrinkToFit).toBe(false);
    expect(a.readingOrder).toBe(0);
  });

  it('options constructor + snake_case aliases', () => {
    const a = new Alignment({ horizontal: 'center', vertical: 'center', wrapText: true, indent: 2, textRotation: 45, shrinkToFit: true, readingOrder: 1 });
    expect(a.wrap_text).toBe(true);
    expect(a.text_rotation).toBe(45);
    expect(a.shrink_to_fit).toBe(true);
    expect(a.reading_order).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// NumberFormat (static utility)
// ---------------------------------------------------------------------------
describe('NumberFormat', () => {
  it('getBuiltinFormat returns known formats', () => {
    expect(NumberFormat.getBuiltinFormat(0)).toBe('General');
    expect(NumberFormat.getBuiltinFormat(1)).toBe('0');
    expect(NumberFormat.getBuiltinFormat(14)).toBe('mm-dd-yy');
  });

  it('getBuiltinFormat returns General for unknown id', () => {
    expect(NumberFormat.getBuiltinFormat(9999)).toBe('General');
  });

  it('isBuiltinFormat detects known formats', () => {
    expect(NumberFormat.isBuiltinFormat('General')).toBe(true);
    expect(NumberFormat.isBuiltinFormat('0.00')).toBe(true);
    expect(NumberFormat.isBuiltinFormat('custom-unknown')).toBe(false);
  });

  it('lookupBuiltinFormat returns id for known codes', () => {
    expect(NumberFormat.lookupBuiltinFormat('General')).toBe(0);
    expect(NumberFormat.lookupBuiltinFormat('0.00')).toBe(2);
    expect(NumberFormat.lookupBuiltinFormat('custom-unknown')).toBeNull();
  });

  it('snake_case aliases', () => {
    expect(NumberFormat.get_builtin_format(2)).toBe('0.00');
    expect(NumberFormat.is_builtin_format('0.00')).toBe(true);
    expect(NumberFormat.lookup_builtin_format('0.00')).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// Protection
// ---------------------------------------------------------------------------
describe('Protection', () => {
  it('defaults locked=true, hidden=false', () => {
    const p = new Protection();
    expect(p.locked).toBe(true);
    expect(p.hidden).toBe(false);
  });

  it('accepts options', () => {
    const p = new Protection({ locked: false, hidden: true });
    expect(p.locked).toBe(false);
    expect(p.hidden).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Style
// ---------------------------------------------------------------------------
describe('Style', () => {
  it('constructs with all sub-objects', () => {
    const s = new Style();
    expect(s.font).toBeInstanceOf(Font);
    expect(s.fill).toBeInstanceOf(Fill);
    expect(s.borders).toBeInstanceOf(Borders);
    expect(s.alignment).toBeInstanceOf(Alignment);
    expect(typeof s.numberFormat).toBe('string');
    expect(s.numberFormat).toBe('General');
    expect(s.protection).toBeInstanceOf(Protection);
  });

  it('copy() creates independent clone', () => {
    const s = new Style();
    s.font.bold = true;
    s.fill.setSolidFill('FFFF0000');
    const c = s.copy();
    expect(c.font.bold).toBe(true);
    expect(c.fill.foregroundColor).toBe('FFFF0000');

    // Mutation of copy must not affect original
    c.font.bold = false;
    expect(s.font.bold).toBe(true);
  });

  it('setFillColor sets solid fill', () => {
    const s = new Style();
    s.setFillColor('FF00FF00');
    expect(s.fill.patternType).toBe('solid');
    expect(s.fill.foregroundColor).toBe('FF00FF00');
  });

  it('setBorderColor / setBorderStyle', () => {
    const s = new Style();
    s.setBorderColor('left', 'FF112233');
    expect(s.borders.left.color).toBe('FF112233');

    s.setBorderStyle('top', 'thick');
    expect(s.borders.top.lineStyle).toBe('thick');
  });

  it('setBorder sets all properties at once', () => {
    const s = new Style();
    s.setBorder('bottom', 'dashed', 'FFAABBCC', 2);
    expect(s.borders.bottom.lineStyle).toBe('dashed');
    expect(s.borders.bottom.color).toBe('FFAABBCC');
    expect(s.borders.bottom.weight).toBe(2);
  });

  it('setDiagonalBorder', () => {
    const s = new Style();
    s.setDiagonalBorder('thin', 'FF000000', 1, true, false);
    expect(s.borders.diagonal.lineStyle).toBe('thin');
    expect(s.borders.diagonalUp).toBe(true);
    expect(s.borders.diagonalDown).toBe(false);
  });

  it('alignment convenience methods', () => {
    const s = new Style();
    s.setHorizontalAlignment('center');
    expect(s.alignment.horizontal).toBe('center');
    s.setVerticalAlignment('top');
    expect(s.alignment.vertical).toBe('top');
    s.setTextWrap(true);
    expect(s.alignment.wrapText).toBe(true);
    s.setShrinkToFit(true);
    expect(s.alignment.shrinkToFit).toBe(true);
    s.setIndent(3);
    expect(s.alignment.indent).toBe(3);
    s.setTextRotation(90);
    expect(s.alignment.textRotation).toBe(90);
    s.setReadingOrder(2);
    expect(s.alignment.readingOrder).toBe(2);
  });

  it('number format convenience', () => {
    const s = new Style();
    s.setNumberFormat('#,##0.00');
    expect(s.numberFormat).toBe('#,##0.00');
    s.setBuiltinNumberFormat(14);
    expect(s.numberFormat).toBe('mm-dd-yy');
  });

  it('protection convenience', () => {
    const s = new Style();
    s.setLocked(false);
    expect(s.protection.locked).toBe(false);
    s.setFormulaHidden(true);
    expect(s.protection.hidden).toBe(true);
  });

  it('snake_case aliases on Style', () => {
    const s = new Style();
    s.set_fill_color('FFAABB00');
    expect(s.fill.foregroundColor).toBe('FFAABB00');
    s.set_border_color('left', 'FF000001');
    expect(s.borders.left.color).toBe('FF000001');
    s.set_number_format('0.00%');
    expect(s.number_format).toBe('0.00%');
  });
});
