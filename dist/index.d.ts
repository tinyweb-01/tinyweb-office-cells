import JSZip from 'jszip';

/**
 * Aspose.Cells for Node.js - Style Module
 *
 * Provides classes for cell styling: Font, Fill, Border, Borders,
 * Alignment, NumberFormat, Protection, and Style.
 *
 * Compatible with Aspose.Cells for .NET API structure.
 */
interface FontOptions {
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
declare class Font {
    name: string;
    size: number;
    color: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikethrough: boolean;
    constructor(options?: FontOptions);
    /** .NET-style alias: snake_case getter */
    get is_bold(): boolean;
    set is_bold(v: boolean);
    /** .NET-style alias: camelCase getter */
    get isBold(): boolean;
    set isBold(v: boolean);
    get is_italic(): boolean;
    set is_italic(v: boolean);
    get isItalic(): boolean;
    set isItalic(v: boolean);
}
interface FillOptions {
    patternType?: string;
    foregroundColor?: string;
    backgroundColor?: string;
}
/**
 * Represents fill settings for a cell or range of cells.
 */
declare class Fill {
    patternType: string;
    foregroundColor: string;
    backgroundColor: string;
    constructor(options?: FillOptions);
    get pattern_type(): string;
    set pattern_type(v: string);
    get foreground_color(): string;
    set foreground_color(v: string);
    get background_color(): string;
    set background_color(v: string);
    /** Sets a solid fill pattern with specified color. */
    setSolidFill(color: string): void;
    /** snake_case alias */
    set_solid_fill(color: string): void;
    /** Sets a gradient fill (simplified as solid). */
    setGradientFill(startColor: string, endColor: string): void;
    set_gradient_fill(startColor: string, endColor: string): void;
    /** Sets a pattern fill with specified pattern type and colors. */
    setPatternFill(patternType: string, fgColor?: string, bgColor?: string): void;
    set_pattern_fill(patternType: string, fgColor?: string, bgColor?: string): void;
    /** Sets no fill (transparent background). */
    setNoFill(): void;
    set_no_fill(): void;
}
interface BorderOptions {
    lineStyle?: string;
    color?: string;
    weight?: number;
}
/**
 * Represents border settings for a single side of a cell.
 */
declare class Border {
    lineStyle: string;
    color: string;
    weight: number;
    constructor(options?: BorderOptions);
    get line_style(): string;
    set line_style(v: string);
}
declare enum BorderType {
    TopBorder = "top",
    BottomBorder = "bottom",
    LeftBorder = "left",
    RightBorder = "right",
    DiagonalUp = "diagonalUp",
    DiagonalDown = "diagonalDown"
}
/**
 * Represents border settings for all sides of a cell.
 */
declare class Borders {
    top: Border;
    bottom: Border;
    left: Border;
    right: Border;
    diagonal: Border;
    diagonalUp: boolean;
    diagonalDown: boolean;
    constructor();
    get diagonal_up(): boolean;
    set diagonal_up(v: boolean);
    get diagonal_down(): boolean;
    set diagonal_down(v: boolean);
    /** Gets a border by BorderType. */
    getByBorderType(borderType: BorderType): Border;
    /** Sets border properties for a specific side or 'all'. */
    setBorder(side: string, lineStyle?: string, color?: string, weight?: number): void;
    set_border(side: string, lineStyle?: string, color?: string, weight?: number): void;
}
interface AlignmentOptions {
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
declare class Alignment {
    horizontal: string;
    vertical: string;
    wrapText: boolean;
    indent: number;
    textRotation: number;
    shrinkToFit: boolean;
    readingOrder: number;
    relativeIndent: number;
    constructor(options?: AlignmentOptions);
    get wrap_text(): boolean;
    set wrap_text(v: boolean);
    get text_rotation(): number;
    set text_rotation(v: number);
    get shrink_to_fit(): boolean;
    set shrink_to_fit(v: boolean);
    get reading_order(): number;
    set reading_order(v: number);
    get relative_indent(): number;
    set relative_indent(v: number);
}
/**
 * Built-in number format lookup.
 */
declare class NumberFormat {
    static readonly BUILTIN_FORMATS: Record<number, string>;
    /** Gets a built-in format string by format ID. */
    static getBuiltinFormat(formatId: number): string;
    static get_builtin_format(formatId: number): string;
    /** Checks if a format code is a built-in format. */
    static isBuiltinFormat(formatCode: string): boolean;
    static is_builtin_format(formatCode: string): boolean;
    /** Looks up the format ID for a built-in format code. */
    static lookupBuiltinFormat(formatCode: string): number | null;
    static lookup_builtin_format(formatCode: string): number | null;
}
interface ProtectionOptions {
    locked?: boolean;
    hidden?: boolean;
}
/**
 * Represents cell protection settings.
 */
declare class Protection {
    locked: boolean;
    hidden: boolean;
    constructor(options?: ProtectionOptions);
}
/**
 * Represents formatting settings for a cell or range of cells.
 */
declare class Style {
    font: Font;
    fill: Fill;
    borders: Borders;
    alignment: Alignment;
    numberFormat: string;
    protection: Protection;
    constructor();
    get number_format(): string;
    set number_format(v: string);
    /** Creates a deep copy of this Style. */
    copy(): Style;
    setFillColor(color: string): void;
    set_fill_color(color: string): void;
    setFillPattern(patternType: string, fgColor?: string, bgColor?: string): void;
    set_fill_pattern(patternType: string, fgColor?: string, bgColor?: string): void;
    setNoFill(): void;
    set_no_fill(): void;
    setBorderColor(side: string, color: string): void;
    set_border_color(side: string, color: string): void;
    setBorderStyle(side: string, style: string): void;
    set_border_style(side: string, style: string): void;
    setBorder(side: string, lineStyle?: string, color?: string, weight?: number): void;
    set_border(side: string, lineStyle?: string, color?: string, weight?: number): void;
    setDiagonalBorder(lineStyle?: string, color?: string, weight?: number, up?: boolean, down?: boolean): void;
    set_diagonal_border(lineStyle?: string, color?: string, weight?: number, up?: boolean, down?: boolean): void;
    setHorizontalAlignment(alignment: string): void;
    set_horizontal_alignment(alignment: string): void;
    setVerticalAlignment(alignment: string): void;
    set_vertical_alignment(alignment: string): void;
    setTextWrap(wrap?: boolean): void;
    set_text_wrap(wrap?: boolean): void;
    setShrinkToFit(shrink?: boolean): void;
    set_shrink_to_fit(shrink?: boolean): void;
    setIndent(indent: number): void;
    set_indent(indent: number): void;
    setTextRotation(rotation: number): void;
    set_text_rotation(rotation: number): void;
    setReadingOrder(order: number): void;
    set_reading_order(order: number): void;
    setNumberFormat(formatCode: string): void;
    set_number_format(formatCode: string): void;
    setBuiltinNumberFormat(formatId: number): void;
    set_builtin_number_format(formatId: number): void;
    setLocked(locked?: boolean): void;
    set_locked(locked?: boolean): void;
    setFormulaHidden(hidden?: boolean): void;
    set_formula_hidden(hidden?: boolean): void;
}

/**
 * Aspose.Cells for Node.js - Cell Module
 *
 * Represents a single cell in a worksheet.
 * Compatible with Aspose.Cells for .NET API structure.
 */

interface CellComment {
    text: string;
    author: string;
    width?: number | null;
    height?: number | null;
}
type CellDataType = 'none' | 'boolean' | 'numeric' | 'datetime' | 'string' | 'unknown';
type CellValue = null | undefined | string | number | boolean | Date;
/**
 * Represents a single cell in a worksheet.
 */
declare class Cell {
    private _value;
    private _formula;
    private _style;
    private _comment;
    /** Internal use for saving */
    _styleIndex: number;
    constructor(value?: CellValue, formula?: string | null);
    get value(): CellValue;
    set value(val: CellValue);
    get formula(): string | null;
    set formula(val: string | null);
    get style(): Style;
    set style(val: Style);
    get comment(): CellComment | null;
    get dataType(): CellDataType;
    /** snake_case alias */
    get data_type(): CellDataType;
    private getDataType;
    isEmpty(): boolean;
    is_empty(): boolean;
    clearValue(): void;
    clear_value(): void;
    clearFormula(): void;
    clear_formula(): void;
    clear(): void;
    setComment(text: string, author?: string, width?: number | null, height?: number | null): void;
    /** snake_case alias */
    set_comment(text: string, author?: string, width?: number | null, height?: number | null): void;
    /** .NET-style alias */
    addComment(text: string, author?: string): void;
    add_comment(text: string, author?: string): void;
    getComment(): CellComment | null;
    get_comment(): CellComment | null;
    clearComment(): void;
    clear_comment(): void;
    removeComment(): void;
    remove_comment(): void;
    hasComment(): boolean;
    has_comment(): boolean;
    setCommentSize(width: number, height: number): void;
    set_comment_size(width: number, height: number): void;
    getCommentSize(): [number, number] | null;
    get_comment_size(): [number, number] | null;
    applyStyle(style: Style): void;
    apply_style(style: Style): void;
    setStyle(style: Style): void;
    set_style(style: Style): void;
    getStyle(): Style;
    get_style(): Style;
    clearStyle(): void;
    clear_style(): void;
    hasFormula(): boolean;
    has_formula(): boolean;
    isNumericValue(): boolean;
    is_numeric_value(): boolean;
    isTextValue(): boolean;
    is_text_value(): boolean;
    isBooleanValue(): boolean;
    is_boolean_value(): boolean;
    isDateTimeValue(): boolean;
    is_date_time_value(): boolean;
    /** Sets the value of the cell (.NET compatibility). */
    putValue(value: CellValue): void;
    put_value(value: CellValue): void;
    getValue(): CellValue;
    get_value(): CellValue;
    toString(): string;
}

/**
 * Aspose.Cells for Node.js - Cells Module
 *
 * A1-keyed collection of cells in a worksheet.
 * Compatible with Aspose.Cells for .NET API structure.
 */

interface WorksheetLike$1 {
    _mergedCells: string[];
    _rowHeights: Record<number, number>;
    _columnWidths: Record<number, number>;
    _hiddenRows: Set<number>;
    _hiddenColumns: Set<number>;
}
/**
 * Represents a collection of cells in a worksheet.
 */
declare class Cells {
    /** Internal cell storage keyed by A1 reference. */
    _cells: Map<string, Cell>;
    /** Back-reference to owning worksheet (set at runtime). */
    _worksheet: WorksheetLike$1 | null;
    constructor(worksheet?: WorksheetLike$1 | null);
    private requireWorksheet;
    /** Normalize tuple-like [row, col] into A1 key. */
    private normalizeKey;
    /** Gets a cell by its A1 reference. Creates if not exists. */
    get(key: string | [number, number | string]): Cell;
    /** Sets a cell value by A1 reference. */
    set(key: string | [number, number | string], value: Cell | CellValue): void;
    /** Access cell by row/column (1-based). */
    cell(row: number, column: number | string): Cell;
    static columnIndexFromString(column: string): number;
    /** snake_case alias */
    static column_index_from_string(column: string): number;
    static columnLetterFromIndex(index: number): string;
    static column_letter_from_index(index: number): string;
    static coordinateFromString(coord: string): [number, number];
    static coordinate_from_string(coord: string): [number, number];
    static coordinateToString(row: number, column: number): string;
    static coordinate_to_string(row: number, column: number): string;
    /** Iterates row-by-row, yielding arrays of Cell (or values). */
    iterRows(options?: {
        minRow?: number;
        maxRow?: number;
        minCol?: number;
        maxCol?: number;
        valuesOnly?: boolean;
    }): Generator<(Cell | CellValue)[]>;
    /** snake_case alias */
    iter_rows(options?: {
        minRow?: number;
        maxRow?: number;
        minCol?: number;
        maxCol?: number;
        valuesOnly?: boolean;
    }): Generator<(Cell | CellValue)[]>;
    count(): number;
    clear(): void;
    getCellByName(cellName: string): Cell;
    get_cell_by_name(cellName: string): Cell;
    setCellByName(cellName: string, value: CellValue): void;
    set_cell_by_name(cellName: string, value: CellValue): void;
    getCell(row: number, column: number): Cell;
    get_cell(row: number, column: number): Cell;
    setCell(row: number, column: number, value: CellValue): void;
    set_cell(row: number, column: number, value: CellValue): void;
    setRowHeight(row: number, height: number): void;
    set_row_height(row: number, height: number): void;
    SetRowHeight(row: number, height: number): void;
    getRowHeight(row: number): number;
    get_row_height(row: number): number;
    GetRowHeight(row: number): number;
    setColumnWidth(column: number | string, width: number): void;
    set_column_width(column: number | string, width: number): void;
    SetColumnWidth(column: number | string, width: number): void;
    getColumnWidth(column: number | string): number;
    get_column_width(column: number | string): number;
    GetColumnWidth(column: number | string): number;
    hideRow(row: number): void;
    hide_row(row: number): void;
    unhideRow(row: number): void;
    unhide_row(row: number): void;
    isRowHidden(row: number): boolean;
    is_row_hidden(row: number): boolean;
    IsRowHidden(row: number): boolean;
    hideColumn(column: number | string): void;
    hide_column(column: number | string): void;
    unhideColumn(column: number | string): void;
    unhide_column(column: number | string): void;
    isColumnHidden(column: number | string): boolean;
    is_column_hidden(column: number | string): boolean;
    IsColumnHidden(column: number | string): boolean;
    /** Merges a rectangular range. Row/column are 0-based. */
    merge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void;
    /** .NET alias */
    Merge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void;
    /** Unmerges a previously merged rectangular range. */
    unmerge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void;
    UnMerge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void;
    mergeRange(rangeRef: string): void;
    merge_range(rangeRef: string): void;
    unmergeRange(rangeRef: string): void;
    unmerge_range(rangeRef: string): void;
    getMergedCells(): string[];
    get_merged_cells(): string[];
    getRange(startRow: number, startColumn: number, endRow: number, endColumn: number): Cell[][];
    get_range(startRow: number, startColumn: number, endRow: number, endColumn: number): Cell[][];
    hasCell(cellName: string): boolean;
    has_cell(cellName: string): boolean;
    deleteCell(cellName: string): void;
    delete_cell(cellName: string): void;
    getAllCells(): Map<string, Cell>;
    get_all_cells(): Map<string, Cell>;
    /** Track max data row (1-based). */
    get maxDataRow(): number;
    get max_data_row(): number;
    /** Track max data column (1-based). */
    get maxDataColumn(): number;
    get max_data_column(): number;
    [Symbol.iterator](): Iterator<[string, Cell]>;
    get length(): number;
}

/**
 * Data Validation Module
 *
 * Provides classes for Excel data validation according to ECMA-376 specification.
 * Data validation allows controlling what data can be entered into cells.
 *
 * References:
 * - ECMA-376 Part 4, Section 3.3.1.30 (dataValidation)
 */
declare enum DataValidationType {
    NONE = 0,
    WHOLE_NUMBER = 1,
    DECIMAL = 2,
    LIST = 3,
    DATE = 4,
    TIME = 5,
    TEXT_LENGTH = 6,
    CUSTOM = 7
}
declare enum DataValidationOperator {
    BETWEEN = 0,
    NOT_BETWEEN = 1,
    EQUAL = 2,
    NOT_EQUAL = 3,
    GREATER_THAN = 4,
    LESS_THAN = 5,
    GREATER_THAN_OR_EQUAL = 6,
    LESS_THAN_OR_EQUAL = 7
}
declare enum DataValidationAlertStyle {
    STOP = 0,
    WARNING = 1,
    INFORMATION = 2
}
declare enum DataValidationImeMode {
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
    HALF_HANGUL = 10
}
declare class DataValidation {
    private _sqref;
    private _type;
    private _operator;
    private _formula1;
    private _formula2;
    private _alertStyle;
    private _showErrorMessage;
    private _errorTitle;
    private _errorMessage;
    private _showInputMessage;
    private _inputTitle;
    private _inputMessage;
    private _allowBlank;
    private _showDropdown;
    private _imeMode;
    constructor(sqref?: string | null);
    get sqref(): string | null;
    set sqref(value: string | null);
    get ranges(): string[];
    get type(): DataValidationType;
    set type(value: DataValidationType);
    get operator(): DataValidationOperator;
    set operator(value: DataValidationOperator);
    get formula1(): string | null;
    set formula1(value: string | null);
    get formula2(): string | null;
    set formula2(value: string | null);
    get alertStyle(): DataValidationAlertStyle;
    set alertStyle(value: DataValidationAlertStyle);
    get showErrorMessage(): boolean;
    set showErrorMessage(value: boolean);
    /** Alias for showErrorMessage */
    get showError(): boolean;
    set showError(value: boolean);
    get errorTitle(): string | null;
    set errorTitle(value: string | null);
    get errorMessage(): string | null;
    set errorMessage(value: string | null);
    /** Alias for errorMessage */
    get error(): string | null;
    set error(value: string | null);
    get showInputMessage(): boolean;
    set showInputMessage(value: boolean);
    /** Alias for showInputMessage */
    get showInput(): boolean;
    set showInput(value: boolean);
    get inputTitle(): string | null;
    set inputTitle(value: string | null);
    /** Alias for inputTitle */
    get promptTitle(): string | null;
    set promptTitle(value: string | null);
    get inputMessage(): string | null;
    set inputMessage(value: string | null);
    /** Alias for inputMessage */
    get prompt(): string | null;
    set prompt(value: string | null);
    get allowBlank(): boolean;
    set allowBlank(value: boolean);
    /** Alias for allowBlank */
    get ignoreBlank(): boolean;
    set ignoreBlank(value: boolean);
    /**
     * Whether to show the in-cell dropdown for list validations.
     * Note: In ECMA-376, showDropDown="1" means HIDE dropdown (counterintuitive).
     */
    get showDropdown(): boolean;
    set showDropdown(value: boolean);
    /** Alias for showDropdown */
    get inCellDropdown(): boolean;
    set inCellDropdown(value: boolean);
    get imeMode(): DataValidationImeMode;
    set imeMode(value: DataValidationImeMode);
    add(validationType: DataValidationType, alertStyle?: DataValidationAlertStyle, operator?: DataValidationOperator, formula1?: string | null, formula2?: string | null): void;
    modify(options: {
        validationType?: DataValidationType;
        alertStyle?: DataValidationAlertStyle;
        operator?: DataValidationOperator;
        formula1?: string | null;
        formula2?: string | null;
    }): void;
    delete(): void;
    copy(): DataValidation;
}
declare class DataValidationCollection {
    private _validations;
    private _disablePrompts;
    private _xWindow;
    private _yWindow;
    get count(): number;
    get disablePrompts(): boolean;
    set disablePrompts(value: boolean);
    get xWindow(): number | null;
    set xWindow(value: number | null);
    get yWindow(): number | null;
    set yWindow(value: number | null);
    add(sqref: string, validationType?: DataValidationType, operator?: DataValidationOperator, formula1?: string | null, formula2?: string | null): DataValidation;
    addValidation(validation: DataValidation): void;
    remove(validation: DataValidation): boolean;
    removeAt(index: number): void;
    removeByRange(sqref: string): void;
    clear(): void;
    getValidation(cellRef: string): DataValidation | null;
    private _cellInRange;
    private _colFromRef;
    private _rowFromRef;
    get(index: number): DataValidation;
    [Symbol.iterator](): Iterator<DataValidation>;
}

/**
 * Conditional Formatting Module
 *
 * Provides classes for Excel conditional formatting according to ECMA-376 specification.
 */

declare class ConditionalFormat {
    private _type;
    private _range;
    private _stopIfTrue;
    private _priority;
    private _operator;
    private _formula1;
    private _formula2;
    private _textOperator;
    private _textFormula;
    private _dateOperator;
    private _dateFormula;
    private _duplicate;
    private _top;
    private _percent;
    private _rank;
    private _above;
    private _stdDev;
    private _colorScaleType;
    private _minColor;
    private _midColor;
    private _maxColor;
    private _barColor;
    private _negativeColor;
    private _showBorder;
    private _direction;
    private _barLength;
    private _iconSetType;
    private _reverseIcons;
    private _showIconOnly;
    private _formula;
    private _font;
    private _border;
    private _fill;
    private _alignment;
    private _numberFormat;
    _dxfId: number | null;
    get type(): string | null;
    set type(value: string | null);
    get range(): string | null;
    set range(value: string | null);
    get stopIfTrue(): boolean;
    set stopIfTrue(value: boolean);
    get priority(): number;
    set priority(value: number);
    get operator(): string | null;
    set operator(value: string | null);
    get formula1(): string | null;
    set formula1(value: string | null);
    get formula2(): string | null;
    set formula2(value: string | null);
    get textOperator(): string | null;
    set textOperator(value: string | null);
    get textFormula(): string | null;
    set textFormula(value: string | null);
    get dateOperator(): string | null;
    set dateOperator(value: string | null);
    get dateFormula(): string | null;
    set dateFormula(value: string | null);
    get duplicate(): boolean | null;
    set duplicate(value: boolean | null);
    get top(): boolean | null;
    set top(value: boolean | null);
    get percent(): boolean;
    set percent(value: boolean);
    get rank(): number;
    set rank(value: number);
    get above(): boolean | null;
    set above(value: boolean | null);
    get stdDev(): number;
    set stdDev(value: number);
    get colorScaleType(): string | null;
    set colorScaleType(value: string | null);
    get minColor(): string | null;
    set minColor(value: string | null);
    get midColor(): string | null;
    set midColor(value: string | null);
    get maxColor(): string | null;
    set maxColor(value: string | null);
    get barColor(): string | null;
    set barColor(value: string | null);
    get negativeColor(): string | null;
    set negativeColor(value: string | null);
    get showBorder(): boolean;
    set showBorder(value: boolean);
    get direction(): string | null;
    set direction(value: string | null);
    get barLength(): [number, number] | null;
    set barLength(value: [number, number] | null);
    get iconSetType(): string | null;
    set iconSetType(value: string | null);
    get reverseIcons(): boolean;
    set reverseIcons(value: boolean);
    get showIconOnly(): boolean;
    set showIconOnly(value: boolean);
    get formula(): string | null;
    set formula(value: string | null);
    get font(): Font;
    get border(): Borders;
    get fill(): Fill;
    get alignment(): Alignment;
    get numberFormat(): string | null;
    set numberFormat(value: string | null);
    hasFont(): boolean;
    hasFill(): boolean;
    hasBorder(): boolean;
    hasAlignment(): boolean;
}
declare class ConditionalFormatCollection {
    private _formats;
    get count(): number;
    add(): ConditionalFormat;
    getByIndex(index: number): ConditionalFormat | null;
    getByRange(rangeStr: string): ConditionalFormat[];
    remove(cf: ConditionalFormat): boolean;
    clear(): void;
    get(index: number): ConditionalFormat;
    [Symbol.iterator](): Iterator<ConditionalFormat>;
    get length(): number;
}

/**
 * Hyperlink Module
 *
 * Provides classes for Excel hyperlinks according to ECMA-376 specification.
 */
type HyperlinkType = 'external' | 'internal' | 'email' | 'unknown';
declare class Hyperlink {
    private _range;
    private _address;
    private _subAddress;
    private _textToDisplay;
    private _screenTip;
    private _deleted;
    constructor(rangeAddress: string, address?: string, subAddress?: string, textToDisplay?: string, screenTip?: string);
    get range(): string;
    set range(value: string);
    get address(): string;
    set address(value: string);
    get subAddress(): string;
    set subAddress(value: string);
    get textToDisplay(): string;
    set textToDisplay(value: string);
    get screenTip(): string;
    set screenTip(value: string);
    get type(): HyperlinkType;
    delete(): void;
    get isDeleted(): boolean;
}
declare class HyperlinkCollection {
    private _hyperlinks;
    add(rangeAddress: string, address?: string, subAddress?: string, textToDisplay?: string, screenTip?: string): Hyperlink;
    /** Add a pre-built hyperlink object */
    addHyperlink(hl: Hyperlink): void;
    delete(indexOrHyperlink?: number | Hyperlink): void;
    clear(): void;
    get count(): number;
    get(index: number): Hyperlink;
    [Symbol.iterator](): Iterator<Hyperlink>;
}

/**
 * AutoFilter Module
 *
 * Provides classes for Excel auto-filter according to ECMA-376 specification.
 */
interface CustomFilter {
    operator: string;
    value: string;
}
interface ColorFilter {
    color: string;
    cellColor: boolean;
}
interface DynamicFilter {
    type: string;
    value?: string;
}
interface Top10Filter {
    top: boolean;
    percent: boolean;
    val: number;
}
declare class FilterColumn {
    private _colId;
    private _filters;
    private _customFilters;
    private _colorFilter;
    private _dynamicFilter;
    private _top10Filter;
    private _filterButton;
    constructor(colId: number);
    get colId(): number;
    get filters(): string[];
    get customFilters(): CustomFilter[];
    get colorFilter(): ColorFilter | null;
    set colorFilter(value: ColorFilter | null);
    get dynamicFilter(): DynamicFilter | null;
    set dynamicFilter(value: DynamicFilter | null);
    get top10Filter(): Top10Filter | null;
    set top10Filter(value: Top10Filter | null);
    get filterButton(): boolean;
    set filterButton(value: boolean);
    addFilter(value: string): void;
    addCustomFilter(operator: string, value: string): void;
    clearFilters(): void;
}
interface SortState {
    ref: string;
    columnIndex: number;
    ascending: boolean;
}
declare class AutoFilter {
    private _range;
    private _filterColumns;
    private _sortState;
    get range(): string | null;
    set range(value: string | null);
    get filterColumns(): Map<number, FilterColumn>;
    get sortState(): SortState | null;
    set sortState(value: SortState | null);
    setRange(startRow: number, startCol: number, endRow: number, endCol: number): void;
    filter(colIndex: number, values: string[]): FilterColumn;
    addFilter(colIndex: number, value: string): FilterColumn;
    customFilter(colIndex: number, operator: string, value: string): FilterColumn;
    filterByColor(colIndex: number, color: string, cellColor?: boolean): FilterColumn;
    filterTop10(colIndex: number, top?: boolean, percent?: boolean, val?: number): FilterColumn;
    filterDynamic(colIndex: number, filterType: string, value?: string): FilterColumn;
    clearColumnFilter(colIndex: number): void;
    clearAllFilters(): void;
    remove(): void;
    showFilterButton(colIndex: number, show?: boolean): void;
    sort(colIndex: number, ascending?: boolean): void;
    getFilterColumn(colIndex: number): FilterColumn | null;
    hasFilter(colIndex: number): boolean;
    get hasAnyFilter(): boolean;
    private _getOrCreateColumn;
    private _colToLetter;
}

/**
 * Aspose.Cells for Node.js - Page Break Module
 *
 * Manual page break collections aligned with Aspose.Cells/.NET API.
 * Ported from Python `page_break.py`.
 *
 * @module features/PageBreak
 */
/**
 * Collection of manual horizontal page breaks (row breaks).
 */
declare class HorizontalPageBreakCollection {
    /** @internal */
    _breaks: Set<number>;
    private _normalizeRow;
    /**
     * Adds a manual horizontal page break.
     * @param rowOrCell 0-based row index or A1 cell reference.
     * @returns The 0-based row index.
     */
    add(rowOrCell: number | string): number;
    /**
     * Removes the break at zero-based collection index.
     */
    removeAt(index: number): void;
    /**
     * Removes a manual horizontal page break by row/cell.
     */
    remove(rowOrCell: number | string): void;
    /** Clears all manual horizontal page breaks. */
    clear(): void;
    get count(): number;
    /** Returns a sorted array of break rows. */
    toList(): number[];
    /** Alias */
    to_list(): number[];
    get(index: number): number;
    [Symbol.iterator](): Iterator<number>;
}
/**
 * Collection of manual vertical page breaks (column breaks).
 */
declare class VerticalPageBreakCollection {
    /** @internal */
    _breaks: Set<number>;
    private _normalizeColumn;
    /**
     * Adds a manual vertical page break.
     * @param columnOrCell 0-based column index, column letters, or A1 cell reference.
     * @returns The 0-based column index.
     */
    add(columnOrCell: number | string): number;
    /**
     * Removes a manual vertical page break by column/cell.
     */
    remove(columnOrCell: number | string): void;
    /**
     * Removes the break at zero-based collection index.
     */
    removeAt(index: number): void;
    /** Clears all manual vertical page breaks. */
    clear(): void;
    get count(): number;
    /** Returns a sorted array of break columns. */
    toList(): number[];
    /** Alias */
    to_list(): number[];
    get(index: number): number;
    [Symbol.iterator](): Iterator<number>;
}

/**
 * Aspose.Cells for Node.js - Worksheet Module
 *
 * Represents a single worksheet in an Excel workbook.
 * Compatible with Aspose.Cells for .NET API structure.
 */

declare class SheetProtection {
    sheet: boolean;
    password: string | null;
    objects: boolean;
    scenarios: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
    get format_cells(): boolean;
    set format_cells(v: boolean);
    get format_columns(): boolean;
    set format_columns(v: boolean);
    get format_rows(): boolean;
    set format_rows(v: boolean);
    get insert_columns(): boolean;
    set insert_columns(v: boolean);
    get insert_rows(): boolean;
    set insert_rows(v: boolean);
    get insert_hyperlinks(): boolean;
    set insert_hyperlinks(v: boolean);
    get delete_columns(): boolean;
    set delete_columns(v: boolean);
    get delete_rows(): boolean;
    set delete_rows(v: boolean);
    get select_locked_cells(): boolean;
    set select_locked_cells(v: boolean);
    get select_unlocked_cells(): boolean;
    set select_unlocked_cells(v: boolean);
    get auto_filter(): boolean;
    set auto_filter(v: boolean);
    get pivot_tables(): boolean;
    set pivot_tables(v: boolean);
}
declare const PROTECTION_KEYS: readonly ["protected", "password", "sheet", "objects", "scenarios", "format_cells", "format_columns", "format_rows", "insert_columns", "insert_rows", "insert_hyperlinks", "delete_columns", "delete_rows", "select_locked_cells", "select_unlocked_cells", "sort", "auto_filter", "pivot_tables"];
type ProtectionKey = typeof PROTECTION_KEYS[number];
/**
 * Dictionary-like wrapper around SheetProtection for backward compatibility.
 */
declare class SheetProtectionDictWrapper {
    private _protection;
    constructor(sheetProtection: SheetProtection);
    get(key: ProtectionKey, defaultValue?: unknown): unknown;
    getItem(key: ProtectionKey): unknown;
    setItem(key: ProtectionKey, value: unknown): void;
}
interface PageSetup {
    orientation: string | null;
    paperSize: number | null;
    scale: number | null;
    fitToWidth: number | null;
    fitToHeight: number | null;
    fitToPage: boolean;
}
interface PageMargins {
    left: number;
    right: number;
    top: number;
    bottom: number;
    header: number;
    footer: number;
}
interface FreezePane {
    row: number;
    column: number;
    freezedRows: number;
    freezedColumns: number;
}
/**
 * Represents a single worksheet in an Excel workbook.
 */
declare class Worksheet {
    private _name;
    private _cells;
    private _visible;
    private _tabColor;
    private _index;
    private _protection;
    private _pageSetup;
    private _pageMargins;
    private _freezePane;
    _mergedCells: string[];
    _rowHeights: Record<number, number>;
    _columnWidths: Record<number, number>;
    _hiddenRows: Set<number>;
    _hiddenColumns: Set<number>;
    private _printArea;
    _sourceXml: string | null;
    private _dataValidations;
    private _conditionalFormatting;
    private _hyperlinks;
    private _autoFilter;
    private _horizontalPageBreaks;
    private _verticalPageBreaks;
    _workbook: unknown;
    constructor(name?: string);
    get name(): string;
    set name(v: string);
    get cells(): Cells;
    get visible(): boolean | 'veryHidden';
    set visible(v: boolean | 'veryHidden');
    /** Alias for visible */
    get isVisible(): boolean;
    get is_visible(): boolean;
    get tabColor(): string | null;
    set tabColor(v: string | null);
    get tab_color(): string | null;
    set tab_color(v: string | null);
    get index(): number;
    set index(v: number);
    get protection(): SheetProtectionDictWrapper;
    get pageSetup(): PageSetup;
    get page_setup(): PageSetup;
    get pageMargins(): PageMargins;
    get page_margins(): PageMargins;
    get printArea(): string | null;
    set printArea(v: string | null);
    get print_area(): string | null;
    set print_area(v: string | null);
    get mergedCells(): string[];
    get merged_cells(): string[];
    get freezePane(): FreezePane | null;
    get freeze_pane(): FreezePane | null;
    get dataValidations(): DataValidationCollection;
    get data_validations(): DataValidationCollection;
    get conditionalFormatting(): ConditionalFormatCollection;
    get conditional_formatting(): ConditionalFormatCollection;
    get hyperlinks(): HyperlinkCollection;
    get autoFilter(): AutoFilter;
    get auto_filter_obj(): AutoFilter;
    get horizontalPageBreaks(): HorizontalPageBreakCollection;
    get horizontal_page_breaks(): HorizontalPageBreakCollection;
    get verticalPageBreaks(): VerticalPageBreakCollection;
    get vertical_page_breaks(): VerticalPageBreakCollection;
    rename(newName: string): void;
    setVisibility(value: boolean | 'veryHidden'): void;
    set_visibility(value: boolean | 'veryHidden'): void;
    getVisibility(): boolean | 'veryHidden';
    get_visibility(): boolean | 'veryHidden';
    setTabColor(color: string | null): void;
    set_tab_color(color: string | null): void;
    getTabColor(): string | null;
    get_tab_color(): string | null;
    clearTabColor(): void;
    clear_tab_color(): void;
    setPageOrientation(orientation: 'portrait' | 'landscape'): void;
    set_page_orientation(orientation: 'portrait' | 'landscape'): void;
    getPageOrientation(): string | null;
    get_page_orientation(): string | null;
    setPaperSize(paperSize: number): void;
    set_paper_size(paperSize: number): void;
    getPaperSize(): number | null;
    get_paper_size(): number | null;
    setPageMargins(options: Partial<PageMargins>): void;
    set_page_margins(options: Partial<PageMargins>): void;
    getPageMargins(): PageMargins;
    get_page_margins(): PageMargins;
    setFitToPages(width?: number, height?: number): void;
    set_fit_to_pages(width?: number, height?: number): void;
    setPrintScale(scale: number): void;
    set_print_scale(scale: number): void;
    setPrintArea(printArea: string): void;
    set_print_area(printArea: string): void;
    SetPrintArea(printArea: string): void;
    clearPrintArea(): void;
    clear_print_area(): void;
    ClearPrintArea(): void;
    private normalizePrintArea;
    setFreezePane(row: number, column: number, freezedRows?: number, freezedColumns?: number): void;
    set_freeze_pane(row: number, column: number, freezedRows?: number, freezedColumns?: number): void;
    clearFreezePane(): void;
    clear_freeze_pane(): void;
    isProtected(): boolean;
    is_protected(): boolean;
    protect(options?: {
        password?: string | null;
        formatCells?: boolean;
        formatColumns?: boolean;
        formatRows?: boolean;
        insertColumns?: boolean;
        insertRows?: boolean;
        deleteColumns?: boolean;
        deleteRows?: boolean;
        sort?: boolean;
        autoFilter?: boolean;
        insertHyperlinks?: boolean;
        pivotTables?: boolean;
        selectLockedCells?: boolean;
        selectUnlockedCells?: boolean;
        objects?: boolean;
        scenarios?: boolean;
    }): void;
    unprotect(_password?: string): void;
    copy(name?: string): Worksheet;
    /** Placeholder */
    delete(): void;
    /** Placeholder */
    move(_index: number): void;
    /** Placeholder */
    select(): void;
    /** Placeholder */
    activate(): void;
    /**
     * Evaluates every formula on this worksheet only.
     *
     * Delegates to the workbook-level FormulaEvaluator so that cross-sheet
     * references and defined names resolve correctly.
     */
    calculateFormula(): void;
    /** snake_case alias */
    calculate_formula(): void;
}

/**
 * Aspose.Cells for Node.js - Defined Name Module
 *
 * Represents defined names (named ranges) in the workbook.
 * Ported from Python `workbook_properties.py` DefinedName/DefinedNameCollection.
 *
 * ECMA-376 Section: 18.2.5
 *
 * @module features/DefinedName
 */
/**
 * Represents a defined name in the workbook.
 */
declare class DefinedName {
    private _name;
    private _refersTo;
    private _localSheetId;
    private _comment;
    private _description;
    private _hidden;
    constructor(name: string, refersTo: string, localSheetId?: number | null);
    /** Name of the defined name. */
    get name(): string;
    set name(v: string);
    /** Formula or range that the name refers to. */
    get refersTo(): string;
    set refersTo(v: string);
    /** Sheet index for sheet-local names (null for global names). */
    get localSheetId(): number | null;
    set localSheetId(v: number | null);
    /** Comment associated with the name. */
    get comment(): string | null;
    set comment(v: string | null);
    /** Description of the name. */
    get description(): string | null;
    set description(v: string | null);
    /** Whether the name is hidden. */
    get hidden(): boolean;
    set hidden(v: boolean);
    get refers_to(): string;
    set refers_to(v: string);
    get local_sheet_id(): number | null;
    set local_sheet_id(v: number | null);
}
/**
 * Collection of defined names in the workbook.
 */
declare class DefinedNameCollection {
    private _names;
    /**
     * Adds a defined name to the collection.
     *
     * @param nameOrDn Either a DefinedName object or a string name.
     * @param refersTo Formula or range (required if nameOrDn is a string).
     * @param localSheetId Sheet index for sheet-local names.
     * @returns The added DefinedName.
     */
    add(nameOrDn: DefinedName | string, refersTo?: string, localSheetId?: number | null): DefinedName;
    /**
     * Removes a defined name by name string.
     * @returns The removed DefinedName or null if not found.
     */
    remove(name: string): DefinedName | null;
    /**
     * Gets a defined name by index or name string.
     */
    get(key: number | string): DefinedName;
    /** Number of defined names. */
    get count(): number;
    /** Returns internal array (for iteration). */
    toArray(): DefinedName[];
    [Symbol.iterator](): Iterator<DefinedName>;
}

/**
 * Aspose.Cells for Node.js – Workbook Module
 *
 * Provides Workbook class representing an Excel workbook, along with
 * SaveFormat enum and WorkbookProtection helper.
 *
 * Phase 2: XLSX save via XmlSaver (JSZip). Load still stubbed (Phase 3).
 */

/**
 * Specifies the format for saving a workbook.
 *
 * Compatible with Aspose.Cells for .NET SaveFormat enumeration.
 */
declare enum SaveFormat {
    AUTO = "auto",
    XLSX = "xlsx",
    CSV = "csv",
    TSV = "tsv",
    MARKDOWN = "markdown",
    JSON = "json"
}
/**
 * Determine the {@link SaveFormat} from a file path's extension.
 *
 * @throws {Error} If the extension is unsupported.
 */
declare function saveFormatFromExtension(filePath: string): SaveFormat;
/** Alias for .NET naming convention */
declare const save_format_from_extension: typeof saveFormatFromExtension;
/**
 * Holds workbook-level protection settings (structure / windows lock).
 */
declare class WorkbookProtection {
    lockStructure: boolean;
    lockWindows: boolean;
    workbookPassword: string | null;
    get lock_structure(): boolean;
    set lock_structure(v: boolean);
    get lock_windows(): boolean;
    set lock_windows(v: boolean);
    get workbook_password(): string | null;
    set workbook_password(v: string | null);
}
interface WorkbookView {
    activeTab: number;
    active_tab: number;
}
declare class WorkbookProperties {
    /** Workbook protection settings */
    protection: WorkbookProtection;
    /** View settings (active tab, etc.) */
    private _view;
    /** Calculation mode */
    calcMode: string;
    /** Defined names collection */
    definedNames: DefinedNameCollection;
    constructor();
    get view(): WorkbookView;
    get calc_mode(): string;
    set calc_mode(v: string);
    get defined_names(): DefinedNameCollection;
}
declare class DocumentProperties {
    title: string;
    subject: string;
    creator: string;
    keywords: string;
    description: string;
    lastModifiedBy: string;
    category: string;
    revision: string;
    created: Date | string | null;
    modified: Date | string | null;
    /** Extended / App properties */
    application: string;
    appVersion: string;
    company: string;
    manager: string;
    /** Round-trip XML blobs */
    _coreXml: string | null;
    _appXml: string | null;
    get last_modified_by(): string;
    set last_modified_by(v: string);
}
/**
 * Represents an Excel workbook.
 *
 * Provides worksheet management, protection helpers, and (in later phases)
 * file I/O for XLSX, CSV, Markdown, and JSON.
 *
 * ```ts
 * const wb = new Workbook();
 * const ws = wb.worksheets[0];
 * ws.cells.get('A1').value = 'Hello';
 * ```
 */
declare class Workbook {
    private _worksheets;
    private _styles;
    private _sharedStrings;
    private _filePath;
    /** Workbook-level properties */
    private _properties;
    /** Document properties (lazily created) */
    private _documentProperties;
    /** @internal */ _fontStyles: Map<string, number>;
    /** @internal */ _fillStyles: Map<string, number>;
    /** @internal */ _borderStyles: Map<string, number>;
    /** @internal */ _alignmentStyles: Map<string, number>;
    /** @internal */ _protectionStyles: Map<string, number>;
    /** @internal */ _cellStyles: Map<string, number>;
    /** @internal */ _numFormats: Map<string, number>;
    /** Round-trip: raw workbook-level XML blobs */
    _sourceXml: Record<string, string>;
    /**
     * Creates a new Workbook.
     *
     * @param filePath - Reserved for Phase 2/3: path to an existing .xlsx to load.
     * @param password - Reserved for Phase 2/3: password for encrypted files.
     */
    constructor(filePath?: string | null, password?: string | null);
    /** Gets collection of worksheets in the workbook. */
    get worksheets(): Worksheet[];
    /** Gets the file path of the workbook (null if not loaded from file). */
    get filePath(): string | null;
    /** Gets workbook properties (protection, view, calc settings). */
    get properties(): WorkbookProperties;
    /** Gets document properties (title, author, etc.). Lazily created. */
    get documentProperties(): DocumentProperties;
    /** Gets the internal styles array. */
    get styles(): Style[];
    /** Gets the shared strings table. */
    get sharedStrings(): string[];
    get file_path(): string | null;
    get document_properties(): DocumentProperties;
    get shared_strings(): string[];
    /**
     * Adds a new worksheet to the workbook.
     *
     * @param name - Optional name. Auto-generated as "SheetN" if omitted.
     * @returns The newly created Worksheet.
     */
    addWorksheet(name?: string): Worksheet;
    /** Alias for {@link addWorksheet}. */
    add_worksheet(name?: string): Worksheet;
    /** Alias for {@link addWorksheet}. */
    createWorksheet(name?: string): Worksheet;
    /** Alias for {@link addWorksheet}. */
    create_worksheet(name?: string): Worksheet;
    /**
     * Gets a worksheet by index (0-based) or by name.
     *
     * @throws {RangeError} If numeric index is out of range.
     * @throws {Error} If no worksheet with the given name exists.
     */
    getWorksheet(indexOrName: number | string): Worksheet;
    /** Alias */
    get_worksheet(indexOrName: number | string): Worksheet;
    /**
     * Returns the worksheet with the given name, or `null` if not found.
     */
    getWorksheetByName(name: string): Worksheet | null;
    /** Alias */
    get_worksheet_by_name(name: string): Worksheet | null;
    /**
     * Returns the worksheet at the given 0-based index, or `null` if out of range.
     */
    getWorksheetByIndex(index: number): Worksheet | null;
    /** Alias */
    get_worksheet_by_index(index: number): Worksheet | null;
    /**
     * Removes a worksheet by index, name, or direct reference.
     *
     * @throws {RangeError} If numeric index is out of range.
     * @throws {Error} If string name is not found.
     * @throws {TypeError} If argument is none of the above.
     */
    removeWorksheet(indexOrNameOrWs: number | string | Worksheet): void;
    /** Alias */
    remove_worksheet(indexOrNameOrWs: number | string | Worksheet): void;
    /**
     * Returns the currently active worksheet.
     */
    getActiveWorksheet(): Worksheet | null;
    /** Alias */
    get_active_worksheet(): Worksheet | null;
    /**
     * Sets the active worksheet by index, name, or Worksheet reference.
     */
    setActiveWorksheet(indexOrNameOrWs: number | string | Worksheet): void;
    /** Alias */
    set_active_worksheet(indexOrNameOrWs: number | string | Worksheet): void;
    /**
     * Copies a worksheet and appends the copy to the workbook.
     *
     * @returns The new worksheet, or `null` if source could not be resolved.
     */
    copyWorksheet(indexOrNameOrWs: number | string | Worksheet): Worksheet | null;
    /** Alias */
    copy_worksheet(indexOrNameOrWs: number | string | Worksheet): Worksheet | null;
    /**
     * Protects the workbook structure/windows with an optional password.
     *
     * @param password - Optional password string.
     * @param lockStructure - Prevent adding/removing sheets (default true).
     * @param lockWindows - Prevent resizing windows (default false).
     */
    protect(password?: string | null, lockStructure?: boolean, lockWindows?: boolean): void;
    /**
     * Removes workbook structure/window protection.
     */
    unprotect(_password?: string | null): void;
    /**
     * Returns `true` if the workbook has structure or window protection enabled.
     */
    isProtected(): boolean;
    /** Alias */
    is_protected(): boolean;
    /**
     * Returns a snapshot of the current workbook protection settings.
     */
    get protection(): {
        lockStructure: boolean;
        lockWindows: boolean;
        password: string | null;
    };
    /**
     * Evaluates every formula in every worksheet of this workbook.
     *
     * Uses a lazy `require` to avoid circular-dependency issues between
     * Workbook and FormulaEvaluator.
     */
    calculateFormula(): void;
    /** snake_case alias */
    calculate_formula(): void;
    /**
     * Saves the workbook to a file.
     *
     * Currently supports XLSX format only (Phase 2). CSV, JSON, etc. will be
     * added in later phases.
     *
     * @param filePath  Destination file path.
     * @param saveFormat  Explicit format; defaults to AUTO (inferred from extension).
     * @param _options  Reserved for future use (encryption options, etc.).
     * @param _password Reserved for future use (file-level password).
     */
    save(filePath: string, saveFormat?: SaveFormat | null, _options?: unknown, _password?: string | null): Promise<void>;
    /**
     * Saves the workbook to an in-memory Buffer (XLSX format).
     *
     * @param saveFormat  Explicit format; defaults to XLSX.
     * @returns A Buffer containing the XLSX ZIP archive.
     */
    saveToBuffer(saveFormat?: SaveFormat): Promise<Buffer>;
    /** snake_case alias */
    save_to_buffer(saveFormat?: SaveFormat): Promise<Buffer>;
    /**
     * Loads a workbook from a file path (async).
     *
     * @param filePath Path to the .xlsx file.
     * @param _password Reserved for future encrypted workbook support.
     * @returns A Promise resolving to the loaded Workbook.
     */
    static load(filePath: string, _password?: string | null): Promise<Workbook>;
    /**
     * Loads a workbook from a Buffer (async).
     *
     * @param buffer Buffer containing XLSX data.
     * @returns A Promise resolving to the loaded Workbook.
     */
    static loadFromBuffer(buffer: Buffer): Promise<Workbook>;
    toString(): string;
}

/**
 * Aspose.Cells for Node.js – Formula Evaluator Module
 *
 * Provides formula evaluation functionality for cells that have formulas.
 * Ported from Python `formula_evaluator.py` with significant enhancements.
 *
 * Supports:
 * - String, numeric, and boolean literals
 * - Cell references (A1 style, absolute $A$1, sheet-qualified Sheet1!A1)
 * - Range references (A1:B5)
 * - Arithmetic operators: + - * / ^
 * - Comparison operators: = <> < > <= >=
 * - String concatenation: &
 * - Built-in functions (SUM, IF, VLOOKUP, etc.)
 * - Defined names
 * - Nested formulas with recursive evaluation
 * - Error handling (#VALUE!, #REF!, #NAME?, #DIV/0!, #N/A)
 *
 * @module features/FormulaEvaluator
 */
interface WorksheetLike {
    name: string;
    cells: CellsLike;
}
interface CellLike {
    value: any;
    formula: string | null;
    hasFormula(): boolean;
}
interface CellsLike {
    get(key: string): CellLike;
    _cells: Map<string, CellLike>;
}
interface DefinedNameLike {
    name: string;
    refersTo: string;
}
interface DefinedNameCollectionLike {
    toArray(): DefinedNameLike[];
}
interface WorkbookPropertiesLike {
    definedNames: DefinedNameCollectionLike;
}
interface WorkbookLike {
    worksheets: WorksheetLike[];
    properties: WorkbookPropertiesLike;
}
/**
 * Evaluates Excel formulas within a workbook context.
 *
 * Supports arithmetic, comparisons, string concatenation, cell references,
 * range references, defined names, and a comprehensive set of built-in functions.
 *
 * @example
 * ```ts
 * const evaluator = new FormulaEvaluator(workbook);
 * const result = evaluator.evaluate('=SUM(A1:A5)', worksheet);
 * ```
 */
declare class FormulaEvaluator {
    private _workbook;
    private _definedNamesCache;
    private _evaluatingCells;
    private _maxDepth;
    constructor(workbook: WorkbookLike);
    /**
     * Evaluate a formula and return the result.
     *
     * @param formula - The formula string (with or without leading '=').
     * @param worksheet - The worksheet context for cell references.
     * @returns The evaluated value, or an error string on failure.
     */
    evaluate(formula: string, worksheet?: WorksheetLike | null): any;
    /**
     * Evaluate all formulas in all worksheets of the workbook.
     * Sets each cell's value to the evaluated result.
     */
    evaluateAll(singleSheet?: WorksheetLike): void;
    /** snake_case alias */
    evaluate_all(singleSheet?: WorksheetLike): void;
    private _evaluateTokens;
    /**
     * Parse an expression with operator precedence (Pratt parser style).
     * Returns { result, pos } where pos is the next token index.
     */
    private _parseExpression;
    /**
     * Parse an atomic value (number, string, cell ref, function call, parenthesised expression, etc.)
     */
    private _parseAtom;
    private _getDefinedNames;
    private _resolveDefinedName;
    private _resolveCellRef;
    private _resolveSheetRef;
    /**
     * Resolves a range reference into an array of cell values.
     */
    private _resolveRange;
    /**
     * Resolves a range reference returning an array of arrays (rows x cols) for lookup functions.
     */
    private _resolveRange2D;
    private _applyArithmetic;
    private _applyComparison;
    private _toNumber;
    private _toBool;
    private _flattenArgs;
    private _flattenNumeric;
    /** Functions that need 2D arrays for their table/range argument */
    private static readonly _2D_FUNCS;
    private _callFunctionWithTokens;
    private _callFunction;
    private _funcConcatenate;
    private _funcText;
    private _funcLen;
    private _funcTrim;
    private _funcUpper;
    private _funcLower;
    private _funcLeft;
    private _funcRight;
    private _funcMid;
    private _funcSubstitute;
    private _funcRept;
    private _funcIf;
    private _funcAnd;
    private _funcOr;
    private _funcNot;
    private _funcSum;
    private _funcAverage;
    private _funcMin;
    private _funcMax;
    private _funcCount;
    private _funcCountA;
    private _funcAbs;
    private _funcRound;
    private _funcRoundUp;
    private _funcRoundDown;
    private _funcInt;
    private _funcMod;
    private _funcPower;
    private _funcSqrt;
    private _funcCeiling;
    private _funcFloor;
    private _funcVlookup;
    private _funcHlookup;
    private _funcIndex;
    private _funcMatch;
    private _valuesEqual;
    private _funcToday;
    private _funcNow;
    private _funcDate;
    private _funcYear;
    private _funcMonth;
    private _funcDay;
    private _funcIsNumber;
    private _funcIsText;
    private _funcIsBlank;
    private _funcIsError;
    private _funcIsNa;
    private _funcValue;
    private _funcChoose;
}

/**
 * Comment XML Handler
 *
 * Reads and writes Excel comments (annotations) in XLSX format.
 * Comments are stored in xl/comments{n}.xml with VML drawings for positioning.
 */

declare class CommentXmlWriter {
    /**
     * Check if worksheet has any comments
     */
    static worksheetHasComments(worksheet: Worksheet): boolean;
    /**
     * Write comments XML to the ZIP archive
     */
    writeCommentsXml(zip: JSZip, worksheet: Worksheet, sheetNum: number): void;
    /**
     * Write VML drawing XML for comment shapes
     */
    writeVmlDrawingXml(zip: JSZip, worksheet: Worksheet, sheetNum: number): void;
    private _calculateAnchor;
}
declare class CommentXmlReader {
    loadComments(zip: JSZip, worksheet: Worksheet, sheetNum: number): void;
    /**
     * Parse comments XML and apply to worksheet
     */
    parseAndApplyComments(commentsRoot: any, worksheet: Worksheet): void;
    /**
     * Parse VML drawing to get comment dimensions
     */
    parseAndApplyVmlDrawing(vmlContent: string, worksheet: Worksheet): void;
    private _ensureArray;
    private _colToLetter;
}

/**
 * XML AutoFilter Loader
 *
 * Loads auto-filter settings from XLSX worksheet XML.
 */

declare class AutoFilterXmlLoader {
    loadAutoFilter(wsRoot: any, autoFilter: AutoFilter): void;
    private _calculateSortColumnIndex;
    private _colFromRef;
}

/**
 * XML AutoFilter Saver
 *
 * Writes auto-filter settings to XLSX worksheet XML.
 */

declare class AutoFilterXmlSaver {
    formatAutoFilterXml(autoFilter: AutoFilter): string;
    private _calculateSortRefs;
    private _colFromRef;
    private _rowFromRef;
    private _colToLetter;
}

/**
 * XML Conditional Format Loader
 *
 * Loads conditional formatting rules from XLSX worksheet XML.
 */

declare class ConditionalFormatXmlLoader {
    private _dxfStyles;
    constructor(dxfStyles: any[]);
    loadConditionalFormatting(wsRoot: any, collection: ConditionalFormatCollection): void;
    private _loadCfRule;
    private _loadColorScale;
    private _loadDataBar;
    private _loadIconSet;
    private _applyDxfStyles;
    private _applyDxfData;
}

/**
 * XML Conditional Format Saver
 *
 * Writes conditional formatting rules to XLSX worksheet XML.
 */

declare class ConditionalFormatXmlSaver {
    private _escapeXml;
    constructor(escapeXml: (text: string) => string);
    /**
     * Formats conditional formatting XML. Returns DXF entries that need to be
     * added to the styles.xml dxfs collection.
     */
    formatConditionalFormattingXml(collection: ConditionalFormatCollection, startDxfId: number): {
        xml: string;
        dxfEntries: string[];
    };
    private _formatCfRuleXml;
    private _formatColorScaleXml;
    private _formatDataBarXml;
    private _formatIconSetXml;
    private _formatDxfEntry;
    private _getFirstCellFromRange;
    private _buildTextRuleFormula;
}

/**
 * XML Data Validation Loader
 *
 * Loads data validation settings from XLSX worksheet XML.
 */

declare class DataValidationXmlLoader {
    loadDataValidations(wsRoot: any, collection: DataValidationCollection): void;
    private _loadDataValidation;
}

/**
 * XML Data Validation Saver
 *
 * Writes data validation settings to XLSX worksheet XML.
 */

declare class DataValidationXmlSaver {
    private _escapeXml;
    constructor(escapeXml: (text: string) => string);
    formatDataValidationsXml(collection: DataValidationCollection): string;
    private _formatDataValidationXml;
}

/**
 * XML Hyperlink Handler
 *
 * Loads and saves hyperlinks from/to XLSX worksheet XML and .rels files.
 */

declare class HyperlinkXmlLoader {
    loadHyperlinks(wsRoot: any, collection: HyperlinkCollection, relationships: Map<string, string>): void;
    loadRelationships(relsXml: any): Map<string, string>;
    private _loadHyperlink;
}
interface HyperlinkRelationship {
    rId: string;
    target: string;
}
declare class HyperlinkXmlSaver {
    private _relIdCounter;
    formatHyperlinksXml(collection: HyperlinkCollection): string;
    private _formatHyperlinkXml;
    getHyperlinkRelationships(collection: HyperlinkCollection): HyperlinkRelationship[];
    resetRelationshipCounter(startRelId?: number): void;
}
declare class HyperlinkRelationshipWriter {
    static formatRelationshipsXml(relationships: HyperlinkRelationship[], existingRels?: string[]): string;
}

/**
 * Shared String Table for XLSX files.
 *
 * Manages deduplicated string storage per ECMA-376 §18.4 (Shared String Table).
 * Each unique string is stored once; cells reference strings by index.
 */
declare class SharedStringTable {
    /** Ordered list of unique strings */
    private _strings;
    /** Reverse lookup: string → index */
    private _stringToIndex;
    /** Total reference count (including duplicates) */
    private _count;
    /** Number of unique strings. */
    get uniqueCount(): number;
    /** Total number of string references (≥ uniqueCount). */
    get count(): number;
    /** Direct access to the strings array (read-only copy). */
    get strings(): string[];
    /**
     * Adds a string to the shared string table.
     *
     * If the string already exists, returns the existing index.
     * Always increments the total reference count.
     *
     * @param text - The string to add.
     * @returns The index of the string in the table.
     */
    addString(text: string): number;
    /**
     * Gets a string by its index.
     *
     * @param index - The 0-based index.
     * @returns The string at that index.
     * @throws {RangeError} If the index is out of range.
     */
    getString(index: number): string;
    /**
     * Serializes the shared string table to OOXML.
     *
     * Produces the content of `xl/sharedStrings.xml`.
     */
    toXml(): string;
    /**
     * Resets the table to empty state.
     */
    clear(): void;
}

/**
 * Cell Value Handler for XLSX serialisation.
 *
 * Provides type detection and value formatting for writing cell values
 * to OOXML worksheet XML, matching ECMA-376 §18.18.11 (ST_CellType).
 *
 * Cell types:
 *  - 's'   → shared string index
 *  - 'str' → inline formula string result
 *  - 'n'   → number (including dates as serial numbers)
 *  - 'b'   → boolean (0 / 1)
 *  - 'e'   → error value
 */

/** Excel cell type constants matching ST_CellType */
declare const CELL_TYPE_STRING = "s";
declare const CELL_TYPE_NUMBER = "n";
declare const CELL_TYPE_BOOLEAN = "b";
declare const CELL_TYPE_ERROR = "e";
declare class CellValueHandler {
    /**
     * Determines the OOXML cell type for a JavaScript value.
     *
     * @param value - A cell value (string, number, boolean, Date, null, etc.)
     * @returns The OOXML cell type string, or `null` for empty cells.
     */
    static getCellType(value: CellValue): string | null;
    /**
     * Formats a cell value for writing into OOXML `<v>` element.
     *
     * @param value - The cell value.
     * @param cellType - Optional pre-determined cell type; auto-detected if omitted.
     * @returns A tuple of `[formattedString, cellType]`, or `[null, null]` for empty.
     */
    static formatValueForXml(value: CellValue, cellType?: string | null): [string | null, string | null];
    /**
     * Parses an OOXML cell value string back into a JavaScript value.
     *
     * @param valueStr - The raw string from the `<v>` element.
     * @param cellType - The cell type attribute ('s', 'n', 'b', 'e', 'str').
     * @param sharedStrings - Optional shared string table for type 's'.
     * @returns The parsed JavaScript value.
     */
    static parseValueFromXml(valueStr: string | null | undefined, cellType: string | null, sharedStrings?: string[]): CellValue;
    /**
     * Converts a JavaScript Date to an Excel serial date number.
     *
     * The serial number represents days since 1899-12-30 (Excel epoch).
     * Time-of-day is represented as the fractional part.
     *
     * @param dt - The Date to convert.
     * @returns The Excel serial date number.
     */
    static dateToExcelSerial(dt: Date): number;
    /**
     * Converts an Excel serial date number to a JavaScript Date.
     *
     * @param serial - The Excel serial date number.
     * @returns A Date object.
     */
    static excelSerialToDate(serial: number): Date;
    /**
     * Checks if a string value represents an Excel error.
     */
    static isErrorValue(value: string): boolean;
    /**
     * Returns the normalized error string, or `null` if not an error.
     */
    static getErrorType(value: string): string | null;
}

/**
 * XLSX XML Saver – writes a Workbook to an OOXML ZIP (.xlsx) archive.
 *
 * Ported from Python `xml_saver.py`. Phase 2 scope:
 *   ✔ Content types, relationships, workbook.xml
 *   ✔ Worksheet XML (sheetData, dimension, cols, merged cells, freeze panes,
 *     sheet protection, page setup, page margins)
 *   ✔ Styles XML (fonts, fills, borders, alignments, protection, number formats, cellXfs)
 *   ✔ Shared strings XML
 *   ✔ Core & extended properties (docProps)
 *   ✗ Charts, pictures, shapes, tables, sparklines, conditional formatting,
 *     hyperlinks, data validations, comments, autofilter, calc chain (future phases)
 *
 * @module io/XmlSaver
 */

declare class XmlSaver {
    private _workbook;
    private _sharedStringTable;
    private _fontStyles;
    private _fillStyles;
    private _borderStyles;
    private _alignmentStyles;
    private _protectionStyles;
    private _numFormats;
    private _cellStyles;
    private _dxfEntries;
    constructor(workbook: Workbook);
    /**
     * Saves the workbook as an XLSX buffer (in-memory ZIP).
     *
     * @returns A `Buffer` containing the XLSX data.
     */
    saveToBuffer(): Promise<Buffer>;
    /**
     * Saves the workbook to a file path.
     *
     * @param filePath - Destination file path.
     */
    saveToFile(filePath: string): Promise<void>;
    private _writeContentTypes;
    private _writeRootRelationships;
    private _writeWorkbookRelationships;
    private _writeWorkbookXml;
    private _writeWorksheetXml;
    /**
     * Computes the dimension reference string (e.g. "A1:C10") for a worksheet.
     */
    private _computeDimensionRef;
    /**
     * Formats <cols> XML for column widths and hidden columns.
     */
    private _formatColsXml;
    /**
     * Formats a single cell as XML.
     */
    private _formatCellXml;
    private _writeStylesXml;
    private _formatFontXml;
    private _formatFillXml;
    private _formatBorderXml;
    private _formatAlignmentXml;
    private _formatProtectionXml;
    private _registerDefaultStyles;
    private _getOrCreateFontStyle;
    private _getOrCreateFillStyle;
    private _getOrCreateBorderStyle;
    private _getOrCreateAlignmentStyle;
    private _getOrCreateProtectionStyle;
    private _getOrCreateNumberFormatStyle;
    /**
     * Gets or creates a composite cellXfs style index for a cell.
     */
    private _getOrCreateCellStyle;
    private _writeSharedStringsXml;
    private _writeThemeXml;
    private _writeCorePropertiesXml;
    private _writeAppPropertiesXml;
}

/**
 * tinyweb-cells – XML Loader Module (Phase 3)
 *
 * Reads OOXML (.xlsx) ZIP archives and populates Workbook objects.
 * Port of the Python `xml_loader.py`.
 *
 * ECMA-376 Compliant cell value import.
 */

declare class XmlLoader {
    private _workbook;
    private _wb;
    private _parser;
    private _contentTypeOverrides;
    private _contentTypeDefaults;
    constructor(workbook: Workbook);
    /**
     * Loads workbook data from a JSZip instance.
     */
    loadWorkbook(zip: JSZip): Promise<void>;
    /**
     * Loads a workbook from a Buffer.
     */
    static loadFromBuffer(buffer: Buffer): Promise<Workbook>;
    /**
     * Loads a workbook from a file path.
     */
    static loadFromFile(filePath: string): Promise<Workbook>;
    private _loadContentTypeOverrides;
    private _loadTheme;
    private _loadWorkbookProperties;
    private _loadDefinedNames;
    private _loadWorksheetInfo;
    private _applyPrintAreasFromDefinedNames;
    private _extractPrintArea;
    private _loadSharedStrings;
    private _loadStyles;
    private _loadStylesXml;
    private _loadFonts;
    private _loadFills;
    private _extractColorAttrs;
    private _loadBorders;
    private _loadCellXfs;
    private _loadDxfStyles;
    private _loadExtraWorkbookRels;
    private _loadWorksheetsData;
    private _loadComments;
    private _loadExtraSheetRels;
    private _loadWorksheetData;
    private _loadMergedCells;
    private _loadColumnDimensions;
    private _loadRowHeights;
    private _loadFreezePane;
    /**
     * Loads manual page breaks from rowBreaks / colBreaks elements.
     * ECMA-376 Section 18.3.1.14 (colBreaks), 18.3.1.73 (rowBreaks).
     */
    private _loadPageBreaks;
    private _applyCellStyle;
    private _loadDocumentProperties;
    private _loadCoreProperties;
    private _loadAppProperties;
    private _parseDatetime;
}

export { Alignment, AutoFilter, AutoFilterXmlLoader, AutoFilterXmlSaver, Border, BorderType, Borders, CELL_TYPE_BOOLEAN, CELL_TYPE_ERROR, CELL_TYPE_NUMBER, CELL_TYPE_STRING, Cell, type CellComment, type CellDataType, type CellValue, CellValueHandler, Cells, type ColorFilter, CommentXmlReader, CommentXmlWriter, ConditionalFormat, ConditionalFormatCollection, ConditionalFormatXmlLoader, ConditionalFormatXmlSaver, type CustomFilter, DataValidation, DataValidationAlertStyle, DataValidationCollection, DataValidationImeMode, DataValidationOperator, DataValidationType, DataValidationXmlLoader, DataValidationXmlSaver, DefinedName, DefinedNameCollection, DocumentProperties, type DynamicFilter, Fill, FilterColumn, Font, FormulaEvaluator, type FreezePane, HorizontalPageBreakCollection, Hyperlink, HyperlinkCollection, type HyperlinkRelationship, HyperlinkRelationshipWriter, type HyperlinkType, HyperlinkXmlLoader, HyperlinkXmlSaver, NumberFormat, type PageMargins, type PageSetup, Protection, SaveFormat, SharedStringTable, SheetProtection, SheetProtectionDictWrapper, type SortState, Style, type Top10Filter, VerticalPageBreakCollection, Workbook, WorkbookProperties, WorkbookProtection, type WorkbookView, Worksheet, type WorksheetLike$1 as WorksheetLike, XmlLoader, XmlSaver, saveFormatFromExtension, save_format_from_extension };
