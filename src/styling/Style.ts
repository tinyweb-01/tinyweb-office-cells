/**
 * tinyweb-office-cells - Style Module
 *
 * Provides classes for cell styling: Font, Fill, Border, Borders,
 * Alignment, NumberFormat, Protection, and Style.
 */

// ─── Font ──────────────────────────────────────────────────────────────────

export interface FontOptions {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}

/**
 * Represents font settings for a cell or range of cells.
 */
export class Font {
  name: string;
  size: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;

  constructor(options: FontOptions = {}) {
    this.name = options.name ?? 'Calibri';
    this.size = options.size ?? 11;
    this.color = options.color ?? 'FF000000';
    this.bold = options.bold ?? false;
    this.italic = options.italic ?? false;
    this.underline = options.underline ?? false;
    this.strikethrough = options.strikethrough ?? false;
  }

  /** .NET-style alias: snake_case getter */
  get is_bold(): boolean {
    return this.bold;
  }
  set is_bold(v: boolean) {
    this.bold = v;
  }

  /** .NET-style alias: camelCase getter */
  get isBold(): boolean {
    return this.bold;
  }
  set isBold(v: boolean) {
    this.bold = v;
  }

  get is_italic(): boolean {
    return this.italic;
  }
  set is_italic(v: boolean) {
    this.italic = v;
  }

  get isItalic(): boolean {
    return this.italic;
  }
  set isItalic(v: boolean) {
    this.italic = v;
  }
}

// ─── Fill ──────────────────────────────────────────────────────────────────

export interface FillOptions {
  patternType?: string;
  foregroundColor?: string;
  backgroundColor?: string;
}

/**
 * Represents fill settings for a cell or range of cells.
 */
export class Fill {
  patternType: string;
  foregroundColor: string;
  backgroundColor: string;

  constructor(options: FillOptions = {}) {
    this.patternType = options.patternType ?? 'none';
    this.foregroundColor = options.foregroundColor ?? 'FFFFFFFF';
    this.backgroundColor = options.backgroundColor ?? 'FFFFFFFF';
  }

  // snake_case aliases
  get pattern_type(): string { return this.patternType; }
  set pattern_type(v: string) { this.patternType = v; }
  get foreground_color(): string { return this.foregroundColor; }
  set foreground_color(v: string) { this.foregroundColor = v; }
  get background_color(): string { return this.backgroundColor; }
  set background_color(v: string) { this.backgroundColor = v; }

  /** Sets a solid fill pattern with specified color. */
  setSolidFill(color: string): void {
    this.patternType = 'solid';
    this.foregroundColor = color;
    this.backgroundColor = color;
  }
  /** snake_case alias */
  set_solid_fill(color: string): void { this.setSolidFill(color); }

  /** Sets a gradient fill (simplified as solid). */
  setGradientFill(startColor: string, endColor: string): void {
    this.patternType = 'solid';
    this.foregroundColor = startColor;
    this.backgroundColor = endColor;
  }
  set_gradient_fill(startColor: string, endColor: string): void { this.setGradientFill(startColor, endColor); }

  /** Sets a pattern fill with specified pattern type and colors. */
  setPatternFill(patternType: string, fgColor = 'FFFFFFFF', bgColor = 'FFFFFFFF'): void {
    this.patternType = patternType;
    this.foregroundColor = fgColor;
    this.backgroundColor = bgColor;
  }
  set_pattern_fill(patternType: string, fgColor = 'FFFFFFFF', bgColor = 'FFFFFFFF'): void {
    this.setPatternFill(patternType, fgColor, bgColor);
  }

  /** Sets no fill (transparent background). */
  setNoFill(): void {
    this.patternType = 'none';
    this.foregroundColor = 'FFFFFFFF';
    this.backgroundColor = 'FFFFFFFF';
  }
  set_no_fill(): void { this.setNoFill(); }
}

// ─── Border ────────────────────────────────────────────────────────────────

export interface BorderOptions {
  lineStyle?: string;
  color?: string;
  weight?: number;
}

/**
 * Represents border settings for a single side of a cell.
 */
export class Border {
  lineStyle: string;
  color: string;
  weight: number;

  constructor(options: BorderOptions = {}) {
    this.lineStyle = options.lineStyle ?? 'none';
    this.color = options.color ?? 'FF000000';
    this.weight = options.weight ?? 1;
  }

  // snake_case aliases
  get line_style(): string { return this.lineStyle; }
  set line_style(v: string) { this.lineStyle = v; }
}

// ─── BorderType enum ───────────────────────────────────────────────────────

export enum BorderType {
  TopBorder = 'top',
  BottomBorder = 'bottom',
  LeftBorder = 'left',
  RightBorder = 'right',
  DiagonalUp = 'diagonalUp',
  DiagonalDown = 'diagonalDown',
}

// ─── Borders ───────────────────────────────────────────────────────────────

/**
 * Represents border settings for all sides of a cell.
 */
export class Borders {
  top: Border;
  bottom: Border;
  left: Border;
  right: Border;
  diagonal: Border;
  diagonalUp: boolean;
  diagonalDown: boolean;

  constructor() {
    this.top = new Border();
    this.bottom = new Border();
    this.left = new Border();
    this.right = new Border();
    this.diagonal = new Border();
    this.diagonalUp = false;
    this.diagonalDown = false;
  }

  // snake_case aliases
  get diagonal_up(): boolean { return this.diagonalUp; }
  set diagonal_up(v: boolean) { this.diagonalUp = v; }
  get diagonal_down(): boolean { return this.diagonalDown; }
  set diagonal_down(v: boolean) { this.diagonalDown = v; }

  /** Gets a border by BorderType. */
  getByBorderType(borderType: BorderType): Border {
    switch (borderType) {
      case BorderType.TopBorder: return this.top;
      case BorderType.BottomBorder: return this.bottom;
      case BorderType.LeftBorder: return this.left;
      case BorderType.RightBorder: return this.right;
      case BorderType.DiagonalUp:
      case BorderType.DiagonalDown:
        return this.diagonal;
      default:
        throw new Error(`Unknown border type: ${borderType}`);
    }
  }

  /** Sets border properties for a specific side or 'all'. */
  setBorder(side: string, lineStyle = 'none', color = 'FF000000', weight = 1): void {
    const apply = (b: Border): void => {
      b.lineStyle = lineStyle;
      b.color = color;
      b.weight = weight;
    };
    if (side === 'all') {
      apply(this.top);
      apply(this.bottom);
      apply(this.left);
      apply(this.right);
    } else if (side === 'top') {
      apply(this.top);
    } else if (side === 'bottom') {
      apply(this.bottom);
    } else if (side === 'left') {
      apply(this.left);
    } else if (side === 'right') {
      apply(this.right);
    }
  }
  set_border(side: string, lineStyle = 'none', color = 'FF000000', weight = 1): void {
    this.setBorder(side, lineStyle, color, weight);
  }
}

// ─── Alignment ─────────────────────────────────────────────────────────────

export interface AlignmentOptions {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  indent?: number;
  textRotation?: number;
  shrinkToFit?: boolean;
  readingOrder?: number;
  relativeIndent?: number;
}

/**
 * Represents alignment settings for a cell.
 */
export class Alignment {
  horizontal: string;
  vertical: string;
  wrapText: boolean;
  indent: number;
  textRotation: number;
  shrinkToFit: boolean;
  readingOrder: number;
  relativeIndent: number;

  constructor(options: AlignmentOptions = {}) {
    this.horizontal = options.horizontal ?? 'general';
    this.vertical = options.vertical ?? 'bottom';
    this.wrapText = options.wrapText ?? false;
    this.indent = options.indent ?? 0;
    this.textRotation = options.textRotation ?? 0;
    this.shrinkToFit = options.shrinkToFit ?? false;
    this.readingOrder = options.readingOrder ?? 0;
    this.relativeIndent = options.relativeIndent ?? 0;
  }

  // snake_case aliases
  get wrap_text(): boolean { return this.wrapText; }
  set wrap_text(v: boolean) { this.wrapText = v; }
  get text_rotation(): number { return this.textRotation; }
  set text_rotation(v: number) { this.textRotation = v; }
  get shrink_to_fit(): boolean { return this.shrinkToFit; }
  set shrink_to_fit(v: boolean) { this.shrinkToFit = v; }
  get reading_order(): number { return this.readingOrder; }
  set reading_order(v: number) { this.readingOrder = v; }
  get relative_indent(): number { return this.relativeIndent; }
  set relative_indent(v: number) { this.relativeIndent = v; }
}

// ─── NumberFormat ──────────────────────────────────────────────────────────

/**
 * Built-in number format lookup.
 */
export class NumberFormat {
  static readonly BUILTIN_FORMATS: Record<number, string> = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    5: '$#,##0_);($#,##0)',
    6: '$#,##0_);[Red]($#,##0)',
    7: '$#,##0.00_);($#,##0.00)',
    8: '$#,##0.00_);[Red]($#,##0.00)',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)',
    40: '#,##0.00_);[Red](#,##0.00)',
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mm:ss.0',
    48: '##0.0E+0',
    49: '@',
  };

  /** Gets a built-in format string by format ID. */
  static getBuiltinFormat(formatId: number): string {
    return NumberFormat.BUILTIN_FORMATS[formatId] ?? 'General';
  }
  static get_builtin_format(formatId: number): string {
    return NumberFormat.getBuiltinFormat(formatId);
  }

  /** Checks if a format code is a built-in format. */
  static isBuiltinFormat(formatCode: string): boolean {
    return Object.values(NumberFormat.BUILTIN_FORMATS).includes(formatCode);
  }
  static is_builtin_format(formatCode: string): boolean {
    return NumberFormat.isBuiltinFormat(formatCode);
  }

  /** Looks up the format ID for a built-in format code. */
  static lookupBuiltinFormat(formatCode: string): number | null {
    for (const [id, code] of Object.entries(NumberFormat.BUILTIN_FORMATS)) {
      if (code === formatCode) return Number(id);
    }
    return null;
  }
  static lookup_builtin_format(formatCode: string): number | null {
    return NumberFormat.lookupBuiltinFormat(formatCode);
  }
}

// ─── Protection ────────────────────────────────────────────────────────────

export interface ProtectionOptions {
  locked?: boolean;
  hidden?: boolean;
}

/**
 * Represents cell protection settings.
 */
export class Protection {
  locked: boolean;
  hidden: boolean;

  constructor(options: ProtectionOptions = {}) {
    this.locked = options.locked ?? true;
    this.hidden = options.hidden ?? false;
  }
}

// ─── Style ─────────────────────────────────────────────────────────────────

/**
 * Represents formatting settings for a cell or range of cells.
 */
export class Style {
  font: Font;
  fill: Fill;
  borders: Borders;
  alignment: Alignment;
  numberFormat: string;
  protection: Protection;

  constructor() {
    this.font = new Font();
    this.fill = new Fill();
    this.borders = new Borders();
    this.alignment = new Alignment();
    this.numberFormat = 'General';
    this.protection = new Protection();
  }

  // snake_case alias
  get number_format(): string { return this.numberFormat; }
  set number_format(v: string) { this.numberFormat = v; }

  /** Creates a deep copy of this Style. */
  copy(): Style {
    const s = new Style();
    s.font = new Font({
      name: this.font.name,
      size: this.font.size,
      color: this.font.color,
      bold: this.font.bold,
      italic: this.font.italic,
      underline: this.font.underline,
      strikethrough: this.font.strikethrough,
    });
    s.fill = new Fill({
      patternType: this.fill.patternType,
      foregroundColor: this.fill.foregroundColor,
      backgroundColor: this.fill.backgroundColor,
    });
    s.borders = new Borders();
    s.borders.top = new Border({ lineStyle: this.borders.top.lineStyle, color: this.borders.top.color, weight: this.borders.top.weight });
    s.borders.bottom = new Border({ lineStyle: this.borders.bottom.lineStyle, color: this.borders.bottom.color, weight: this.borders.bottom.weight });
    s.borders.left = new Border({ lineStyle: this.borders.left.lineStyle, color: this.borders.left.color, weight: this.borders.left.weight });
    s.borders.right = new Border({ lineStyle: this.borders.right.lineStyle, color: this.borders.right.color, weight: this.borders.right.weight });
    s.borders.diagonal = new Border({ lineStyle: this.borders.diagonal.lineStyle, color: this.borders.diagonal.color, weight: this.borders.diagonal.weight });
    s.borders.diagonalUp = this.borders.diagonalUp;
    s.borders.diagonalDown = this.borders.diagonalDown;
    s.alignment = new Alignment({
      horizontal: this.alignment.horizontal,
      vertical: this.alignment.vertical,
      wrapText: this.alignment.wrapText,
      indent: this.alignment.indent,
      textRotation: this.alignment.textRotation,
      shrinkToFit: this.alignment.shrinkToFit,
      readingOrder: this.alignment.readingOrder,
      relativeIndent: this.alignment.relativeIndent,
    });
    s.numberFormat = this.numberFormat;
    s.protection = new Protection({ locked: this.protection.locked, hidden: this.protection.hidden });
    return s;
  }

  // ── Convenience methods (matching Python) ─────────────────────────────

  setFillColor(color: string): void { this.fill.setSolidFill(color); }
  set_fill_color(color: string): void { this.setFillColor(color); }

  setFillPattern(patternType: string, fgColor = 'FFFFFFFF', bgColor = 'FFFFFFFF'): void {
    this.fill.setPatternFill(patternType, fgColor, bgColor);
  }
  set_fill_pattern(patternType: string, fgColor = 'FFFFFFFF', bgColor = 'FFFFFFFF'): void {
    this.setFillPattern(patternType, fgColor, bgColor);
  }

  setNoFill(): void { this.fill.setNoFill(); }
  set_no_fill(): void { this.setNoFill(); }

  setBorderColor(side: string, color: string): void {
    const sides = side === 'all' ? ['top', 'bottom', 'left', 'right'] as const : [side] as const;
    for (const s of sides) {
      const b = (this.borders as unknown as Record<string, Border>)[s];
      if (b) b.color = color;
    }
  }
  set_border_color(side: string, color: string): void { this.setBorderColor(side, color); }

  setBorderStyle(side: string, style: string): void {
    const sides = side === 'all' ? ['top', 'bottom', 'left', 'right'] as const : [side] as const;
    for (const s of sides) {
      const b = (this.borders as unknown as Record<string, Border>)[s];
      if (b) b.lineStyle = style;
    }
  }
  set_border_style(side: string, style: string): void { this.setBorderStyle(side, style); }

  setBorder(side: string, lineStyle = 'none', color = 'FF000000', weight = 1): void {
    this.borders.setBorder(side, lineStyle, color, weight);
  }
  set_border(side: string, lineStyle = 'none', color = 'FF000000', weight = 1): void {
    this.setBorder(side, lineStyle, color, weight);
  }

  setDiagonalBorder(lineStyle = 'none', color = 'FF000000', weight = 1, up = false, down = false): void {
    this.borders.diagonal.lineStyle = lineStyle;
    this.borders.diagonal.color = color;
    this.borders.diagonal.weight = weight;
    this.borders.diagonalUp = up;
    this.borders.diagonalDown = down;
  }
  set_diagonal_border(lineStyle = 'none', color = 'FF000000', weight = 1, up = false, down = false): void {
    this.setDiagonalBorder(lineStyle, color, weight, up, down);
  }

  setHorizontalAlignment(alignment: string): void {
    const valid = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
    if (!valid.includes(alignment)) throw new Error(`Invalid horizontal alignment: ${alignment}`);
    this.alignment.horizontal = alignment;
  }
  set_horizontal_alignment(alignment: string): void { this.setHorizontalAlignment(alignment); }

  setVerticalAlignment(alignment: string): void {
    const valid = ['top', 'center', 'bottom', 'justify', 'distributed'];
    if (!valid.includes(alignment)) throw new Error(`Invalid vertical alignment: ${alignment}`);
    this.alignment.vertical = alignment;
  }
  set_vertical_alignment(alignment: string): void { this.setVerticalAlignment(alignment); }

  setTextWrap(wrap = true): void { this.alignment.wrapText = wrap; }
  set_text_wrap(wrap = true): void { this.setTextWrap(wrap); }

  setShrinkToFit(shrink = true): void { this.alignment.shrinkToFit = shrink; }
  set_shrink_to_fit(shrink = true): void { this.setShrinkToFit(shrink); }

  setIndent(indent: number): void { this.alignment.indent = Math.max(0, indent); }
  set_indent(indent: number): void { this.setIndent(indent); }

  setTextRotation(rotation: number): void {
    if (rotation !== 255 && (rotation < 0 || rotation > 180)) {
      throw new Error('Text rotation must be 0-180 degrees or 255 for vertical text');
    }
    this.alignment.textRotation = rotation;
  }
  set_text_rotation(rotation: number): void { this.setTextRotation(rotation); }

  setReadingOrder(order: number): void {
    if (![0, 1, 2].includes(order)) {
      throw new Error('Reading order must be 0 (Context), 1 (Left-to-Right), or 2 (Right-to-Left)');
    }
    this.alignment.readingOrder = order;
  }
  set_reading_order(order: number): void { this.setReadingOrder(order); }

  setNumberFormat(formatCode: string): void { this.numberFormat = formatCode; }
  set_number_format(formatCode: string): void { this.setNumberFormat(formatCode); }

  setBuiltinNumberFormat(formatId: number): void {
    this.numberFormat = NumberFormat.getBuiltinFormat(formatId);
  }
  set_builtin_number_format(formatId: number): void { this.setBuiltinNumberFormat(formatId); }

  setLocked(locked = true): void { this.protection.locked = locked; }
  set_locked(locked = true): void { this.setLocked(locked); }

  setFormulaHidden(hidden = true): void { this.protection.hidden = hidden; }
  set_formula_hidden(hidden = true): void { this.setFormulaHidden(hidden); }
}
