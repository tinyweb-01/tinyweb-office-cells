'use strict';

var JSZip = require('jszip');
var fastXmlParser = require('fast-xml-parser');

function _interopDefault (e) { return e && e.__esModule ? e : { default: e }; }

var JSZip__default = /*#__PURE__*/_interopDefault(JSZip);

var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __esm = (fn, res) => function __init() {
  return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;
};
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/styling/Style.ts
var Style_exports = {};
__export(Style_exports, {
  Alignment: () => exports.Alignment,
  Border: () => exports.Border,
  BorderType: () => exports.BorderType,
  Borders: () => exports.Borders,
  Fill: () => exports.Fill,
  Font: () => exports.Font,
  NumberFormat: () => exports.NumberFormat,
  Protection: () => exports.Protection,
  Style: () => exports.Style
});
exports.Font = void 0; exports.Fill = void 0; exports.Border = void 0; exports.BorderType = void 0; exports.Borders = void 0; exports.Alignment = void 0; var _NumberFormat; exports.NumberFormat = void 0; exports.Protection = void 0; exports.Style = void 0;
var init_Style = __esm({
  "src/styling/Style.ts"() {
    exports.Font = class {
      constructor(options = {}) {
        this.name = options.name ?? "Calibri";
        this.size = options.size ?? 11;
        this.color = options.color ?? "FF000000";
        this.bold = options.bold ?? false;
        this.italic = options.italic ?? false;
        this.underline = options.underline ?? false;
        this.strikethrough = options.strikethrough ?? false;
      }
      /** .NET-style alias: snake_case getter */
      get is_bold() {
        return this.bold;
      }
      set is_bold(v) {
        this.bold = v;
      }
      /** .NET-style alias: camelCase getter */
      get isBold() {
        return this.bold;
      }
      set isBold(v) {
        this.bold = v;
      }
      get is_italic() {
        return this.italic;
      }
      set is_italic(v) {
        this.italic = v;
      }
      get isItalic() {
        return this.italic;
      }
      set isItalic(v) {
        this.italic = v;
      }
    };
    exports.Fill = class {
      constructor(options = {}) {
        this.patternType = options.patternType ?? "none";
        this.foregroundColor = options.foregroundColor ?? "FFFFFFFF";
        this.backgroundColor = options.backgroundColor ?? "FFFFFFFF";
      }
      // snake_case aliases
      get pattern_type() {
        return this.patternType;
      }
      set pattern_type(v) {
        this.patternType = v;
      }
      get foreground_color() {
        return this.foregroundColor;
      }
      set foreground_color(v) {
        this.foregroundColor = v;
      }
      get background_color() {
        return this.backgroundColor;
      }
      set background_color(v) {
        this.backgroundColor = v;
      }
      /** Sets a solid fill pattern with specified color. */
      setSolidFill(color) {
        this.patternType = "solid";
        this.foregroundColor = color;
        this.backgroundColor = color;
      }
      /** snake_case alias */
      set_solid_fill(color) {
        this.setSolidFill(color);
      }
      /** Sets a gradient fill (simplified as solid). */
      setGradientFill(startColor, endColor) {
        this.patternType = "solid";
        this.foregroundColor = startColor;
        this.backgroundColor = endColor;
      }
      set_gradient_fill(startColor, endColor) {
        this.setGradientFill(startColor, endColor);
      }
      /** Sets a pattern fill with specified pattern type and colors. */
      setPatternFill(patternType, fgColor = "FFFFFFFF", bgColor = "FFFFFFFF") {
        this.patternType = patternType;
        this.foregroundColor = fgColor;
        this.backgroundColor = bgColor;
      }
      set_pattern_fill(patternType, fgColor = "FFFFFFFF", bgColor = "FFFFFFFF") {
        this.setPatternFill(patternType, fgColor, bgColor);
      }
      /** Sets no fill (transparent background). */
      setNoFill() {
        this.patternType = "none";
        this.foregroundColor = "FFFFFFFF";
        this.backgroundColor = "FFFFFFFF";
      }
      set_no_fill() {
        this.setNoFill();
      }
    };
    exports.Border = class {
      constructor(options = {}) {
        this.lineStyle = options.lineStyle ?? "none";
        this.color = options.color ?? "FF000000";
        this.weight = options.weight ?? 1;
      }
      // snake_case aliases
      get line_style() {
        return this.lineStyle;
      }
      set line_style(v) {
        this.lineStyle = v;
      }
    };
    exports.BorderType = /* @__PURE__ */ ((BorderType2) => {
      BorderType2["TopBorder"] = "top";
      BorderType2["BottomBorder"] = "bottom";
      BorderType2["LeftBorder"] = "left";
      BorderType2["RightBorder"] = "right";
      BorderType2["DiagonalUp"] = "diagonalUp";
      BorderType2["DiagonalDown"] = "diagonalDown";
      return BorderType2;
    })(exports.BorderType || {});
    exports.Borders = class {
      constructor() {
        this.top = new exports.Border();
        this.bottom = new exports.Border();
        this.left = new exports.Border();
        this.right = new exports.Border();
        this.diagonal = new exports.Border();
        this.diagonalUp = false;
        this.diagonalDown = false;
      }
      // snake_case aliases
      get diagonal_up() {
        return this.diagonalUp;
      }
      set diagonal_up(v) {
        this.diagonalUp = v;
      }
      get diagonal_down() {
        return this.diagonalDown;
      }
      set diagonal_down(v) {
        this.diagonalDown = v;
      }
      /** Gets a border by BorderType. */
      getByBorderType(borderType) {
        switch (borderType) {
          case "top" /* TopBorder */:
            return this.top;
          case "bottom" /* BottomBorder */:
            return this.bottom;
          case "left" /* LeftBorder */:
            return this.left;
          case "right" /* RightBorder */:
            return this.right;
          case "diagonalUp" /* DiagonalUp */:
          case "diagonalDown" /* DiagonalDown */:
            return this.diagonal;
          default:
            throw new Error(`Unknown border type: ${borderType}`);
        }
      }
      /** Sets border properties for a specific side or 'all'. */
      setBorder(side, lineStyle = "none", color = "FF000000", weight = 1) {
        const apply = (b) => {
          b.lineStyle = lineStyle;
          b.color = color;
          b.weight = weight;
        };
        if (side === "all") {
          apply(this.top);
          apply(this.bottom);
          apply(this.left);
          apply(this.right);
        } else if (side === "top") {
          apply(this.top);
        } else if (side === "bottom") {
          apply(this.bottom);
        } else if (side === "left") {
          apply(this.left);
        } else if (side === "right") {
          apply(this.right);
        }
      }
      set_border(side, lineStyle = "none", color = "FF000000", weight = 1) {
        this.setBorder(side, lineStyle, color, weight);
      }
    };
    exports.Alignment = class {
      constructor(options = {}) {
        this.horizontal = options.horizontal ?? "general";
        this.vertical = options.vertical ?? "bottom";
        this.wrapText = options.wrapText ?? false;
        this.indent = options.indent ?? 0;
        this.textRotation = options.textRotation ?? 0;
        this.shrinkToFit = options.shrinkToFit ?? false;
        this.readingOrder = options.readingOrder ?? 0;
        this.relativeIndent = options.relativeIndent ?? 0;
      }
      // snake_case aliases
      get wrap_text() {
        return this.wrapText;
      }
      set wrap_text(v) {
        this.wrapText = v;
      }
      get text_rotation() {
        return this.textRotation;
      }
      set text_rotation(v) {
        this.textRotation = v;
      }
      get shrink_to_fit() {
        return this.shrinkToFit;
      }
      set shrink_to_fit(v) {
        this.shrinkToFit = v;
      }
      get reading_order() {
        return this.readingOrder;
      }
      set reading_order(v) {
        this.readingOrder = v;
      }
      get relative_indent() {
        return this.relativeIndent;
      }
      set relative_indent(v) {
        this.relativeIndent = v;
      }
    };
    _NumberFormat = class _NumberFormat {
      /** Gets a built-in format string by format ID. */
      static getBuiltinFormat(formatId) {
        return _NumberFormat.BUILTIN_FORMATS[formatId] ?? "General";
      }
      static get_builtin_format(formatId) {
        return _NumberFormat.getBuiltinFormat(formatId);
      }
      /** Checks if a format code is a built-in format. */
      static isBuiltinFormat(formatCode) {
        return Object.values(_NumberFormat.BUILTIN_FORMATS).includes(formatCode);
      }
      static is_builtin_format(formatCode) {
        return _NumberFormat.isBuiltinFormat(formatCode);
      }
      /** Looks up the format ID for a built-in format code. */
      static lookupBuiltinFormat(formatCode) {
        for (const [id, code] of Object.entries(_NumberFormat.BUILTIN_FORMATS)) {
          if (code === formatCode) return Number(id);
        }
        return null;
      }
      static lookup_builtin_format(formatCode) {
        return _NumberFormat.lookupBuiltinFormat(formatCode);
      }
    };
    _NumberFormat.BUILTIN_FORMATS = {
      0: "General",
      1: "0",
      2: "0.00",
      3: "#,##0",
      4: "#,##0.00",
      5: "$#,##0_);($#,##0)",
      6: "$#,##0_);[Red]($#,##0)",
      7: "$#,##0.00_);($#,##0.00)",
      8: "$#,##0.00_);[Red]($#,##0.00)",
      9: "0%",
      10: "0.00%",
      11: "0.00E+00",
      12: "# ?/?",
      13: "# ??/??",
      14: "mm-dd-yy",
      15: "d-mmm-yy",
      16: "d-mmm",
      17: "mmm-yy",
      18: "h:mm AM/PM",
      19: "h:mm:ss AM/PM",
      20: "h:mm",
      21: "h:mm:ss",
      22: "m/d/yy h:mm",
      37: "#,##0_);(#,##0)",
      38: "#,##0_);[Red](#,##0)",
      39: "#,##0.00_);(#,##0.00)",
      40: "#,##0.00_);[Red](#,##0.00)",
      41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
      42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
      43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
      44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
      45: "mm:ss",
      46: "[h]:mm:ss",
      47: "mm:ss.0",
      48: "##0.0E+0",
      49: "@"
    };
    exports.NumberFormat = _NumberFormat;
    exports.Protection = class {
      constructor(options = {}) {
        this.locked = options.locked ?? true;
        this.hidden = options.hidden ?? false;
      }
    };
    exports.Style = class _Style {
      constructor() {
        this.font = new exports.Font();
        this.fill = new exports.Fill();
        this.borders = new exports.Borders();
        this.alignment = new exports.Alignment();
        this.numberFormat = "General";
        this.protection = new exports.Protection();
      }
      // snake_case alias
      get number_format() {
        return this.numberFormat;
      }
      set number_format(v) {
        this.numberFormat = v;
      }
      /** Creates a deep copy of this Style. */
      copy() {
        const s = new _Style();
        s.font = new exports.Font({
          name: this.font.name,
          size: this.font.size,
          color: this.font.color,
          bold: this.font.bold,
          italic: this.font.italic,
          underline: this.font.underline,
          strikethrough: this.font.strikethrough
        });
        s.fill = new exports.Fill({
          patternType: this.fill.patternType,
          foregroundColor: this.fill.foregroundColor,
          backgroundColor: this.fill.backgroundColor
        });
        s.borders = new exports.Borders();
        s.borders.top = new exports.Border({ lineStyle: this.borders.top.lineStyle, color: this.borders.top.color, weight: this.borders.top.weight });
        s.borders.bottom = new exports.Border({ lineStyle: this.borders.bottom.lineStyle, color: this.borders.bottom.color, weight: this.borders.bottom.weight });
        s.borders.left = new exports.Border({ lineStyle: this.borders.left.lineStyle, color: this.borders.left.color, weight: this.borders.left.weight });
        s.borders.right = new exports.Border({ lineStyle: this.borders.right.lineStyle, color: this.borders.right.color, weight: this.borders.right.weight });
        s.borders.diagonal = new exports.Border({ lineStyle: this.borders.diagonal.lineStyle, color: this.borders.diagonal.color, weight: this.borders.diagonal.weight });
        s.borders.diagonalUp = this.borders.diagonalUp;
        s.borders.diagonalDown = this.borders.diagonalDown;
        s.alignment = new exports.Alignment({
          horizontal: this.alignment.horizontal,
          vertical: this.alignment.vertical,
          wrapText: this.alignment.wrapText,
          indent: this.alignment.indent,
          textRotation: this.alignment.textRotation,
          shrinkToFit: this.alignment.shrinkToFit,
          readingOrder: this.alignment.readingOrder,
          relativeIndent: this.alignment.relativeIndent
        });
        s.numberFormat = this.numberFormat;
        s.protection = new exports.Protection({ locked: this.protection.locked, hidden: this.protection.hidden });
        return s;
      }
      // ── Convenience methods (matching Python) ─────────────────────────────
      setFillColor(color) {
        this.fill.setSolidFill(color);
      }
      set_fill_color(color) {
        this.setFillColor(color);
      }
      setFillPattern(patternType, fgColor = "FFFFFFFF", bgColor = "FFFFFFFF") {
        this.fill.setPatternFill(patternType, fgColor, bgColor);
      }
      set_fill_pattern(patternType, fgColor = "FFFFFFFF", bgColor = "FFFFFFFF") {
        this.setFillPattern(patternType, fgColor, bgColor);
      }
      setNoFill() {
        this.fill.setNoFill();
      }
      set_no_fill() {
        this.setNoFill();
      }
      setBorderColor(side, color) {
        const sides = side === "all" ? ["top", "bottom", "left", "right"] : [side];
        for (const s of sides) {
          const b = this.borders[s];
          if (b) b.color = color;
        }
      }
      set_border_color(side, color) {
        this.setBorderColor(side, color);
      }
      setBorderStyle(side, style) {
        const sides = side === "all" ? ["top", "bottom", "left", "right"] : [side];
        for (const s of sides) {
          const b = this.borders[s];
          if (b) b.lineStyle = style;
        }
      }
      set_border_style(side, style) {
        this.setBorderStyle(side, style);
      }
      setBorder(side, lineStyle = "none", color = "FF000000", weight = 1) {
        this.borders.setBorder(side, lineStyle, color, weight);
      }
      set_border(side, lineStyle = "none", color = "FF000000", weight = 1) {
        this.setBorder(side, lineStyle, color, weight);
      }
      setDiagonalBorder(lineStyle = "none", color = "FF000000", weight = 1, up = false, down = false) {
        this.borders.diagonal.lineStyle = lineStyle;
        this.borders.diagonal.color = color;
        this.borders.diagonal.weight = weight;
        this.borders.diagonalUp = up;
        this.borders.diagonalDown = down;
      }
      set_diagonal_border(lineStyle = "none", color = "FF000000", weight = 1, up = false, down = false) {
        this.setDiagonalBorder(lineStyle, color, weight, up, down);
      }
      setHorizontalAlignment(alignment) {
        const valid = ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
        if (!valid.includes(alignment)) throw new Error(`Invalid horizontal alignment: ${alignment}`);
        this.alignment.horizontal = alignment;
      }
      set_horizontal_alignment(alignment) {
        this.setHorizontalAlignment(alignment);
      }
      setVerticalAlignment(alignment) {
        const valid = ["top", "center", "bottom", "justify", "distributed"];
        if (!valid.includes(alignment)) throw new Error(`Invalid vertical alignment: ${alignment}`);
        this.alignment.vertical = alignment;
      }
      set_vertical_alignment(alignment) {
        this.setVerticalAlignment(alignment);
      }
      setTextWrap(wrap = true) {
        this.alignment.wrapText = wrap;
      }
      set_text_wrap(wrap = true) {
        this.setTextWrap(wrap);
      }
      setShrinkToFit(shrink = true) {
        this.alignment.shrinkToFit = shrink;
      }
      set_shrink_to_fit(shrink = true) {
        this.setShrinkToFit(shrink);
      }
      setIndent(indent) {
        this.alignment.indent = Math.max(0, indent);
      }
      set_indent(indent) {
        this.setIndent(indent);
      }
      setTextRotation(rotation) {
        if (rotation !== 255 && (rotation < 0 || rotation > 180)) {
          throw new Error("Text rotation must be 0-180 degrees or 255 for vertical text");
        }
        this.alignment.textRotation = rotation;
      }
      set_text_rotation(rotation) {
        this.setTextRotation(rotation);
      }
      setReadingOrder(order) {
        if (![0, 1, 2].includes(order)) {
          throw new Error("Reading order must be 0 (Context), 1 (Left-to-Right), or 2 (Right-to-Left)");
        }
        this.alignment.readingOrder = order;
      }
      set_reading_order(order) {
        this.setReadingOrder(order);
      }
      setNumberFormat(formatCode) {
        this.numberFormat = formatCode;
      }
      set_number_format(formatCode) {
        this.setNumberFormat(formatCode);
      }
      setBuiltinNumberFormat(formatId) {
        this.numberFormat = exports.NumberFormat.getBuiltinFormat(formatId);
      }
      set_builtin_number_format(formatId) {
        this.setBuiltinNumberFormat(formatId);
      }
      setLocked(locked = true) {
        this.protection.locked = locked;
      }
      set_locked(locked = true) {
        this.setLocked(locked);
      }
      setFormulaHidden(hidden = true) {
        this.protection.hidden = hidden;
      }
      set_formula_hidden(hidden = true) {
        this.setFormulaHidden(hidden);
      }
    };
  }
});

// src/core/Cell.ts
var Cell_exports = {};
__export(Cell_exports, {
  Cell: () => exports.Cell
});
exports.Cell = void 0;
var init_Cell = __esm({
  "src/core/Cell.ts"() {
    init_Style();
    exports.Cell = class {
      constructor(value = null, formula = null) {
        this._value = value;
        this._formula = formula;
        this._style = new exports.Style();
        this._comment = null;
        this._styleIndex = 0;
      }
      // ── Properties ──────────────────────────────────────────────────────────
      get value() {
        return this._value;
      }
      set value(val) {
        this._value = val;
      }
      get formula() {
        return this._formula;
      }
      set formula(val) {
        this._formula = val;
      }
      get style() {
        return this._style;
      }
      set style(val) {
        this._style = val;
      }
      get comment() {
        return this._comment;
      }
      // ── Data type detection ─────────────────────────────────────────────────
      get dataType() {
        return this.getDataType();
      }
      /** snake_case alias */
      get data_type() {
        return this.getDataType();
      }
      getDataType() {
        if (this._value === null || this._value === void 0) return "none";
        if (typeof this._value === "boolean") return "boolean";
        if (typeof this._value === "number") return "numeric";
        if (this._value instanceof Date) return "datetime";
        if (typeof this._value === "string") {
          if (this._value.toUpperCase() === "TRUE" || this._value.toUpperCase() === "FALSE") {
            return "boolean";
          }
          return "string";
        }
        return "unknown";
      }
      // ── Cell value methods ──────────────────────────────────────────────────
      isEmpty() {
        return this._value === null || this._value === void 0;
      }
      is_empty() {
        return this.isEmpty();
      }
      clearValue() {
        this._value = null;
      }
      clear_value() {
        this.clearValue();
      }
      clearFormula() {
        this._formula = null;
      }
      clear_formula() {
        this.clearFormula();
      }
      clear() {
        this._value = null;
        this._formula = null;
      }
      // ── Comment methods ─────────────────────────────────────────────────────
      setComment(text, author = "None", width, height) {
        if (author === "") author = "None";
        this._comment = { text, author, width: width ?? null, height: height ?? null };
      }
      /** snake_case alias */
      set_comment(text, author = "None", width, height) {
        this.setComment(text, author, width, height);
      }
      /** .NET-style alias */
      addComment(text, author = "None") {
        this.setComment(text, author);
      }
      add_comment(text, author = "None") {
        this.addComment(text, author);
      }
      getComment() {
        return this._comment;
      }
      get_comment() {
        return this.getComment();
      }
      clearComment() {
        this._comment = null;
      }
      clear_comment() {
        this.clearComment();
      }
      removeComment() {
        this.clearComment();
      }
      remove_comment() {
        this.clearComment();
      }
      hasComment() {
        return this._comment !== null;
      }
      has_comment() {
        return this.hasComment();
      }
      setCommentSize(width, height) {
        if (!this._comment) throw new Error("Cell has no comment. Call setComment() first.");
        this._comment.width = width;
        this._comment.height = height;
      }
      set_comment_size(width, height) {
        this.setCommentSize(width, height);
      }
      getCommentSize() {
        if (!this._comment) return null;
        if (this._comment.width != null && this._comment.height != null) {
          return [this._comment.width, this._comment.height];
        }
        return null;
      }
      get_comment_size() {
        return this.getCommentSize();
      }
      // ── Style methods ───────────────────────────────────────────────────────
      applyStyle(style) {
        this._style = style;
      }
      apply_style(style) {
        this.applyStyle(style);
      }
      setStyle(style) {
        this.applyStyle(style);
      }
      set_style(style) {
        this.applyStyle(style);
      }
      getStyle() {
        return this._style;
      }
      get_style() {
        return this.getStyle();
      }
      clearStyle() {
        this._style = new exports.Style();
      }
      clear_style() {
        this.clearStyle();
      }
      // ── Formula methods ─────────────────────────────────────────────────────
      hasFormula() {
        return this._formula !== null;
      }
      has_formula() {
        return this.hasFormula();
      }
      isNumericValue() {
        return typeof this._value === "number";
      }
      is_numeric_value() {
        return this.isNumericValue();
      }
      isTextValue() {
        return typeof this._value === "string";
      }
      is_text_value() {
        return this.isTextValue();
      }
      isBooleanValue() {
        return this.dataType === "boolean";
      }
      is_boolean_value() {
        return this.isBooleanValue();
      }
      isDateTimeValue() {
        return this.dataType === "datetime";
      }
      is_date_time_value() {
        return this.isDateTimeValue();
      }
      /** Sets the value of the cell (.NET compatibility). */
      putValue(value) {
        this._value = value;
      }
      put_value(value) {
        this.putValue(value);
      }
      getValue() {
        return this._value;
      }
      get_value() {
        return this.getValue();
      }
      // ── String representation ───────────────────────────────────────────────
      toString() {
        return this._value != null ? String(this._value) : "";
      }
    };
  }
});

// src/core/Cells.ts
exports.Cells = void 0;
var init_Cells = __esm({
  "src/core/Cells.ts"() {
    init_Cell();
    exports.Cells = class _Cells {
      constructor(worksheet = null) {
        this._cells = /* @__PURE__ */ new Map();
        this._worksheet = worksheet;
      }
      requireWorksheet() {
        if (!this._worksheet) throw new Error("Cells is not attached to a Worksheet");
      }
      /** Normalize tuple-like [row, col] into A1 key. */
      normalizeKey(key) {
        if (Array.isArray(key)) {
          const [row, col] = key;
          const colIdx = typeof col === "string" ? _Cells.columnIndexFromString(col) : col;
          return _Cells.coordinateToString(row, colIdx);
        }
        return key;
      }
      // ── Cell access ─────────────────────────────────────────────────────────
      /** Gets a cell by its A1 reference. Creates if not exists. */
      get(key) {
        const k = this.normalizeKey(key);
        let cell = this._cells.get(k);
        if (!cell) {
          cell = new exports.Cell();
          this._cells.set(k, cell);
        }
        return cell;
      }
      /** Sets a cell value by A1 reference. */
      set(key, value) {
        const k = this.normalizeKey(key);
        if (value instanceof exports.Cell) {
          this._cells.set(k, value);
        } else {
          const cell = this.get(k);
          cell.value = value;
        }
      }
      /** Access cell by row/column (1-based). */
      cell(row, column) {
        const colIdx = typeof column === "string" ? _Cells.columnIndexFromString(column) : column;
        const ref = `${_Cells.columnLetterFromIndex(colIdx)}${row}`;
        return this.get(ref);
      }
      // ── Coordinate conversion (static) ──────────────────────────────────────
      static columnIndexFromString(column) {
        if (!column) throw new Error("Column string cannot be empty");
        const upper = column.toUpperCase();
        let result = 0;
        for (const ch of upper) {
          if (ch < "A" || ch > "Z") throw new Error(`Invalid column character: ${ch}`);
          result = result * 26 + (ch.charCodeAt(0) - 64);
        }
        return result;
      }
      /** snake_case alias */
      static column_index_from_string(column) {
        return _Cells.columnIndexFromString(column);
      }
      static columnLetterFromIndex(index) {
        if (index < 1) throw new Error("Column index must be >= 1");
        let result = "";
        let i = index;
        while (i > 0) {
          i--;
          result = String.fromCharCode(65 + i % 26) + result;
          i = Math.floor(i / 26);
        }
        return result;
      }
      static column_letter_from_index(index) {
        return _Cells.columnLetterFromIndex(index);
      }
      static coordinateFromString(coord) {
        if (!coord) throw new Error("Coordinate cannot be empty");
        let colStr = "";
        let rowStr = "";
        for (const ch of coord) {
          if (/[a-zA-Z]/.test(ch)) colStr += ch;
          else if (/\d/.test(ch)) rowStr += ch;
          else throw new Error(`Invalid character in coordinate: ${ch}`);
        }
        if (!colStr || !rowStr) throw new Error(`Invalid coordinate format: ${coord}`);
        return [parseInt(rowStr, 10), _Cells.columnIndexFromString(colStr)];
      }
      static coordinate_from_string(coord) {
        return _Cells.coordinateFromString(coord);
      }
      static coordinateToString(row, column) {
        if (row < 1 || column < 1) throw new Error("Row and column must be >= 1");
        return `${_Cells.columnLetterFromIndex(column)}${row}`;
      }
      static coordinate_to_string(row, column) {
        return _Cells.coordinateToString(row, column);
      }
      // ── Iteration ───────────────────────────────────────────────────────────
      /** Iterates row-by-row, yielding arrays of Cell (or values). */
      *iterRows(options = {}) {
        if (this._cells.size === 0) return;
        const rows = /* @__PURE__ */ new Set();
        const cols = /* @__PURE__ */ new Set();
        for (const ref of this._cells.keys()) {
          const [r, c] = _Cells.coordinateFromString(ref);
          rows.add(r);
          cols.add(c);
        }
        const minRow = options.minRow ?? Math.min(...rows);
        const maxRow = options.maxRow ?? Math.max(...rows);
        const minCol = options.minCol ?? Math.min(...cols);
        const maxCol = options.maxCol ?? Math.max(...cols);
        for (let r = minRow; r <= maxRow; r++) {
          const rowCells = [];
          for (let c = minCol; c <= maxCol; c++) {
            const cell = this.get(_Cells.coordinateToString(r, c));
            rowCells.push(options.valuesOnly ? cell.value : cell);
          }
          yield rowCells;
        }
      }
      /** snake_case alias */
      *iter_rows(options = {}) {
        yield* this.iterRows(options);
      }
      // ── Collection methods ──────────────────────────────────────────────────
      count() {
        return this._cells.size;
      }
      clear() {
        this._cells.clear();
      }
      getCellByName(cellName) {
        return this.get(cellName);
      }
      get_cell_by_name(cellName) {
        return this.getCellByName(cellName);
      }
      setCellByName(cellName, value) {
        this.set(cellName, value);
      }
      set_cell_by_name(cellName, value) {
        this.setCellByName(cellName, value);
      }
      getCell(row, column) {
        return this.cell(row, column);
      }
      get_cell(row, column) {
        return this.getCell(row, column);
      }
      setCell(row, column, value) {
        const ref = _Cells.coordinateToString(row, column);
        this.set(ref, value);
      }
      set_cell(row, column, value) {
        this.setCell(row, column, value);
      }
      // ── Row / Column dimensions ─────────────────────────────────────────────
      setRowHeight(row, height) {
        this.requireWorksheet();
        if (row < 0) throw new Error("row must be >= 0");
        if (height <= 0) throw new Error("height must be > 0");
        this._worksheet._rowHeights[row + 1] = height;
      }
      set_row_height(row, height) {
        this.setRowHeight(row, height);
      }
      SetRowHeight(row, height) {
        this.setRowHeight(row, height);
      }
      getRowHeight(row) {
        this.requireWorksheet();
        if (row < 0) throw new Error("row must be >= 0");
        const h = this._worksheet._rowHeights[row + 1];
        return h ?? 15;
      }
      get_row_height(row) {
        return this.getRowHeight(row);
      }
      GetRowHeight(row) {
        return this.getRowHeight(row);
      }
      setColumnWidth(column, width) {
        this.requireWorksheet();
        let colIdx;
        if (typeof column === "string") {
          colIdx = _Cells.columnIndexFromString(column);
        } else {
          if (column < 0) throw new Error("column must be >= 0");
          colIdx = column + 1;
        }
        if (width <= 0) throw new Error("width must be > 0");
        this._worksheet._columnWidths[colIdx] = width;
      }
      set_column_width(column, width) {
        this.setColumnWidth(column, width);
      }
      SetColumnWidth(column, width) {
        this.setColumnWidth(column, width);
      }
      getColumnWidth(column) {
        this.requireWorksheet();
        let colIdx;
        if (typeof column === "string") {
          colIdx = _Cells.columnIndexFromString(column);
        } else {
          if (column < 0) throw new Error("column must be >= 0");
          colIdx = column + 1;
        }
        return this._worksheet._columnWidths[colIdx] ?? 8.43;
      }
      get_column_width(column) {
        return this.getColumnWidth(column);
      }
      GetColumnWidth(column) {
        return this.getColumnWidth(column);
      }
      // ── Hide / unhide rows & columns ────────────────────────────────────────
      hideRow(row) {
        this.requireWorksheet();
        if (row < 0) throw new Error("row must be >= 0");
        this._worksheet._hiddenRows.add(row + 1);
      }
      hide_row(row) {
        this.hideRow(row);
      }
      unhideRow(row) {
        this.requireWorksheet();
        if (row < 0) throw new Error("row must be >= 0");
        this._worksheet._hiddenRows.delete(row + 1);
      }
      unhide_row(row) {
        this.unhideRow(row);
      }
      isRowHidden(row) {
        this.requireWorksheet();
        if (row < 0) throw new Error("row must be >= 0");
        return this._worksheet._hiddenRows.has(row + 1);
      }
      is_row_hidden(row) {
        return this.isRowHidden(row);
      }
      IsRowHidden(row) {
        return this.isRowHidden(row);
      }
      hideColumn(column) {
        this.requireWorksheet();
        const colIdx = typeof column === "string" ? _Cells.columnIndexFromString(column) : column + 1;
        this._worksheet._hiddenColumns.add(colIdx);
      }
      hide_column(column) {
        this.hideColumn(column);
      }
      unhideColumn(column) {
        this.requireWorksheet();
        const colIdx = typeof column === "string" ? _Cells.columnIndexFromString(column) : column + 1;
        this._worksheet._hiddenColumns.delete(colIdx);
      }
      unhide_column(column) {
        this.unhideColumn(column);
      }
      isColumnHidden(column) {
        this.requireWorksheet();
        const colIdx = typeof column === "string" ? _Cells.columnIndexFromString(column) : column + 1;
        return this._worksheet._hiddenColumns.has(colIdx);
      }
      is_column_hidden(column) {
        return this.isColumnHidden(column);
      }
      IsColumnHidden(column) {
        return this.isColumnHidden(column);
      }
      // ── Merge cells ─────────────────────────────────────────────────────────
      /** Merges a rectangular range. Row/column are 0-based. */
      merge(firstRow, firstColumn, totalRows, totalColumns) {
        this.requireWorksheet();
        if (firstRow < 0) throw new Error("firstRow must be >= 0");
        if (firstColumn < 0) throw new Error("firstColumn must be >= 0");
        if (totalRows < 1) throw new Error("totalRows must be >= 1");
        if (totalColumns < 1) throw new Error("totalColumns must be >= 1");
        const startRef = _Cells.coordinateToString(firstRow + 1, firstColumn + 1);
        const endRef = _Cells.coordinateToString(firstRow + totalRows, firstColumn + totalColumns);
        const mergeRef = (startRef !== endRef ? `${startRef}:${endRef}` : startRef).toUpperCase();
        if (!this._worksheet._mergedCells.includes(mergeRef)) {
          this._worksheet._mergedCells.push(mergeRef);
        }
      }
      /** .NET alias */
      Merge(firstRow, firstColumn, totalRows, totalColumns) {
        this.merge(firstRow, firstColumn, totalRows, totalColumns);
      }
      /** Unmerges a previously merged rectangular range. */
      unmerge(firstRow, firstColumn, totalRows, totalColumns) {
        this.requireWorksheet();
        if (firstRow < 0) throw new Error("firstRow must be >= 0");
        if (firstColumn < 0) throw new Error("firstColumn must be >= 0");
        if (totalRows < 1) throw new Error("totalRows must be >= 1");
        if (totalColumns < 1) throw new Error("totalColumns must be >= 1");
        const startRef = _Cells.coordinateToString(firstRow + 1, firstColumn + 1);
        const endRef = _Cells.coordinateToString(firstRow + totalRows, firstColumn + totalColumns);
        const mergeRef = (startRef !== endRef ? `${startRef}:${endRef}` : startRef).toUpperCase();
        const idx = this._worksheet._mergedCells.indexOf(mergeRef);
        if (idx !== -1) this._worksheet._mergedCells.splice(idx, 1);
      }
      UnMerge(firstRow, firstColumn, totalRows, totalColumns) {
        this.unmerge(firstRow, firstColumn, totalRows, totalColumns);
      }
      mergeRange(rangeRef) {
        this.requireWorksheet();
        if (!rangeRef) throw new Error("rangeRef must be a non-empty string");
        const ref = rangeRef.toUpperCase();
        if (!this._worksheet._mergedCells.includes(ref)) {
          this._worksheet._mergedCells.push(ref);
        }
      }
      merge_range(rangeRef) {
        this.mergeRange(rangeRef);
      }
      unmergeRange(rangeRef) {
        this.requireWorksheet();
        if (!rangeRef) throw new Error("rangeRef must be a non-empty string");
        const ref = rangeRef.toUpperCase();
        const idx = this._worksheet._mergedCells.indexOf(ref);
        if (idx !== -1) this._worksheet._mergedCells.splice(idx, 1);
      }
      unmerge_range(rangeRef) {
        this.unmergeRange(rangeRef);
      }
      getMergedCells() {
        this.requireWorksheet();
        return [...this._worksheet._mergedCells];
      }
      get_merged_cells() {
        return this.getMergedCells();
      }
      // ── Range methods ───────────────────────────────────────────────────────
      getRange(startRow, startColumn, endRow, endColumn) {
        const result = [];
        for (let r = startRow; r <= endRow; r++) {
          const rowCells = [];
          for (let c = startColumn; c <= endColumn; c++) {
            rowCells.push(this.get(_Cells.coordinateToString(r, c)));
          }
          result.push(rowCells);
        }
        return result;
      }
      get_range(startRow, startColumn, endRow, endColumn) {
        return this.getRange(startRow, startColumn, endRow, endColumn);
      }
      // ── Utility ─────────────────────────────────────────────────────────────
      hasCell(cellName) {
        return this._cells.has(cellName);
      }
      has_cell(cellName) {
        return this.hasCell(cellName);
      }
      deleteCell(cellName) {
        this._cells.delete(cellName);
      }
      delete_cell(cellName) {
        this.deleteCell(cellName);
      }
      getAllCells() {
        return new Map(this._cells);
      }
      get_all_cells() {
        return this.getAllCells();
      }
      /** Track max data row (1-based). */
      get maxDataRow() {
        let max = 0;
        for (const ref of this._cells.keys()) {
          const [r] = _Cells.coordinateFromString(ref);
          if (r > max) max = r;
        }
        return max;
      }
      get max_data_row() {
        return this.maxDataRow;
      }
      /** Track max data column (1-based). */
      get maxDataColumn() {
        let max = 0;
        for (const ref of this._cells.keys()) {
          const [, c] = _Cells.coordinateFromString(ref);
          if (c > max) max = c;
        }
        return max;
      }
      get max_data_column() {
        return this.maxDataColumn;
      }
      // ── Iterator support ────────────────────────────────────────────────────
      [Symbol.iterator]() {
        return this._cells.entries();
      }
      get length() {
        return this._cells.size;
      }
    };
  }
});

// src/features/DataValidation.ts
exports.DataValidationType = void 0; exports.DataValidationOperator = void 0; exports.DataValidationAlertStyle = void 0; exports.DataValidationImeMode = void 0; exports.DataValidation = void 0; exports.DataValidationCollection = void 0;
var init_DataValidation = __esm({
  "src/features/DataValidation.ts"() {
    exports.DataValidationType = /* @__PURE__ */ ((DataValidationType2) => {
      DataValidationType2[DataValidationType2["NONE"] = 0] = "NONE";
      DataValidationType2[DataValidationType2["WHOLE_NUMBER"] = 1] = "WHOLE_NUMBER";
      DataValidationType2[DataValidationType2["DECIMAL"] = 2] = "DECIMAL";
      DataValidationType2[DataValidationType2["LIST"] = 3] = "LIST";
      DataValidationType2[DataValidationType2["DATE"] = 4] = "DATE";
      DataValidationType2[DataValidationType2["TIME"] = 5] = "TIME";
      DataValidationType2[DataValidationType2["TEXT_LENGTH"] = 6] = "TEXT_LENGTH";
      DataValidationType2[DataValidationType2["CUSTOM"] = 7] = "CUSTOM";
      return DataValidationType2;
    })(exports.DataValidationType || {});
    exports.DataValidationOperator = /* @__PURE__ */ ((DataValidationOperator2) => {
      DataValidationOperator2[DataValidationOperator2["BETWEEN"] = 0] = "BETWEEN";
      DataValidationOperator2[DataValidationOperator2["NOT_BETWEEN"] = 1] = "NOT_BETWEEN";
      DataValidationOperator2[DataValidationOperator2["EQUAL"] = 2] = "EQUAL";
      DataValidationOperator2[DataValidationOperator2["NOT_EQUAL"] = 3] = "NOT_EQUAL";
      DataValidationOperator2[DataValidationOperator2["GREATER_THAN"] = 4] = "GREATER_THAN";
      DataValidationOperator2[DataValidationOperator2["LESS_THAN"] = 5] = "LESS_THAN";
      DataValidationOperator2[DataValidationOperator2["GREATER_THAN_OR_EQUAL"] = 6] = "GREATER_THAN_OR_EQUAL";
      DataValidationOperator2[DataValidationOperator2["LESS_THAN_OR_EQUAL"] = 7] = "LESS_THAN_OR_EQUAL";
      return DataValidationOperator2;
    })(exports.DataValidationOperator || {});
    exports.DataValidationAlertStyle = /* @__PURE__ */ ((DataValidationAlertStyle2) => {
      DataValidationAlertStyle2[DataValidationAlertStyle2["STOP"] = 0] = "STOP";
      DataValidationAlertStyle2[DataValidationAlertStyle2["WARNING"] = 1] = "WARNING";
      DataValidationAlertStyle2[DataValidationAlertStyle2["INFORMATION"] = 2] = "INFORMATION";
      return DataValidationAlertStyle2;
    })(exports.DataValidationAlertStyle || {});
    exports.DataValidationImeMode = /* @__PURE__ */ ((DataValidationImeMode2) => {
      DataValidationImeMode2[DataValidationImeMode2["NO_CONTROL"] = 0] = "NO_CONTROL";
      DataValidationImeMode2[DataValidationImeMode2["OFF"] = 1] = "OFF";
      DataValidationImeMode2[DataValidationImeMode2["ON"] = 2] = "ON";
      DataValidationImeMode2[DataValidationImeMode2["DISABLED"] = 3] = "DISABLED";
      DataValidationImeMode2[DataValidationImeMode2["HIRAGANA"] = 4] = "HIRAGANA";
      DataValidationImeMode2[DataValidationImeMode2["FULL_KATAKANA"] = 5] = "FULL_KATAKANA";
      DataValidationImeMode2[DataValidationImeMode2["HALF_KATAKANA"] = 6] = "HALF_KATAKANA";
      DataValidationImeMode2[DataValidationImeMode2["FULL_ALPHA"] = 7] = "FULL_ALPHA";
      DataValidationImeMode2[DataValidationImeMode2["HALF_ALPHA"] = 8] = "HALF_ALPHA";
      DataValidationImeMode2[DataValidationImeMode2["FULL_HANGUL"] = 9] = "FULL_HANGUL";
      DataValidationImeMode2[DataValidationImeMode2["HALF_HANGUL"] = 10] = "HALF_HANGUL";
      return DataValidationImeMode2;
    })(exports.DataValidationImeMode || {});
    exports.DataValidation = class _DataValidation {
      constructor(sqref = null) {
        this._sqref = sqref;
        this._type = 0 /* NONE */;
        this._operator = 0 /* BETWEEN */;
        this._formula1 = null;
        this._formula2 = null;
        this._alertStyle = 0 /* STOP */;
        this._showErrorMessage = true;
        this._errorTitle = null;
        this._errorMessage = null;
        this._showInputMessage = false;
        this._inputTitle = null;
        this._inputMessage = null;
        this._allowBlank = true;
        this._showDropdown = true;
        this._imeMode = 0 /* NO_CONTROL */;
      }
      // ---- Cell Range ----
      get sqref() {
        return this._sqref;
      }
      set sqref(value) {
        this._sqref = value;
      }
      get ranges() {
        if (this._sqref) {
          return this._sqref.split(/\s+/);
        }
        return [];
      }
      // ---- Validation Type and Operator ----
      get type() {
        return this._type;
      }
      set type(value) {
        this._type = value;
      }
      get operator() {
        return this._operator;
      }
      set operator(value) {
        this._operator = value;
      }
      // ---- Formulas ----
      get formula1() {
        return this._formula1;
      }
      set formula1(value) {
        this._formula1 = value;
      }
      get formula2() {
        return this._formula2;
      }
      set formula2(value) {
        this._formula2 = value;
      }
      // ---- Error Alert Settings ----
      get alertStyle() {
        return this._alertStyle;
      }
      set alertStyle(value) {
        this._alertStyle = value;
      }
      get showErrorMessage() {
        return this._showErrorMessage;
      }
      set showErrorMessage(value) {
        this._showErrorMessage = value;
      }
      /** Alias for showErrorMessage */
      get showError() {
        return this._showErrorMessage;
      }
      set showError(value) {
        this._showErrorMessage = value;
      }
      get errorTitle() {
        return this._errorTitle;
      }
      set errorTitle(value) {
        this._errorTitle = value;
      }
      get errorMessage() {
        return this._errorMessage;
      }
      set errorMessage(value) {
        this._errorMessage = value;
      }
      /** Alias for errorMessage */
      get error() {
        return this._errorMessage;
      }
      set error(value) {
        this._errorMessage = value;
      }
      // ---- Input Message Settings ----
      get showInputMessage() {
        return this._showInputMessage;
      }
      set showInputMessage(value) {
        this._showInputMessage = value;
      }
      /** Alias for showInputMessage */
      get showInput() {
        return this._showInputMessage;
      }
      set showInput(value) {
        this._showInputMessage = value;
      }
      get inputTitle() {
        return this._inputTitle;
      }
      set inputTitle(value) {
        this._inputTitle = value;
      }
      /** Alias for inputTitle */
      get promptTitle() {
        return this._inputTitle;
      }
      set promptTitle(value) {
        this._inputTitle = value;
      }
      get inputMessage() {
        return this._inputMessage;
      }
      set inputMessage(value) {
        this._inputMessage = value;
      }
      /** Alias for inputMessage */
      get prompt() {
        return this._inputMessage;
      }
      set prompt(value) {
        this._inputMessage = value;
      }
      // ---- Other Settings ----
      get allowBlank() {
        return this._allowBlank;
      }
      set allowBlank(value) {
        this._allowBlank = value;
      }
      /** Alias for allowBlank */
      get ignoreBlank() {
        return this._allowBlank;
      }
      set ignoreBlank(value) {
        this._allowBlank = value;
      }
      /**
       * Whether to show the in-cell dropdown for list validations.
       * Note: In ECMA-376, showDropDown="1" means HIDE dropdown (counterintuitive).
       */
      get showDropdown() {
        return this._showDropdown;
      }
      set showDropdown(value) {
        this._showDropdown = value;
      }
      /** Alias for showDropdown */
      get inCellDropdown() {
        return this._showDropdown;
      }
      set inCellDropdown(value) {
        this._showDropdown = value;
      }
      get imeMode() {
        return this._imeMode;
      }
      set imeMode(value) {
        this._imeMode = value;
      }
      // ---- Methods ----
      add(validationType, alertStyle, operator, formula1, formula2) {
        this._type = validationType;
        if (alertStyle !== void 0) this._alertStyle = alertStyle;
        if (operator !== void 0) this._operator = operator;
        if (formula1 !== void 0) this._formula1 = formula1;
        if (formula2 !== void 0) this._formula2 = formula2;
      }
      modify(options) {
        if (options.validationType !== void 0) this._type = options.validationType;
        if (options.alertStyle !== void 0) this._alertStyle = options.alertStyle;
        if (options.operator !== void 0) this._operator = options.operator;
        if (options.formula1 !== void 0) this._formula1 = options.formula1;
        if (options.formula2 !== void 0) this._formula2 = options.formula2;
      }
      delete() {
        this._type = 0 /* NONE */;
        this._operator = 0 /* BETWEEN */;
        this._formula1 = null;
        this._formula2 = null;
        this._alertStyle = 0 /* STOP */;
        this._showErrorMessage = true;
        this._errorTitle = null;
        this._errorMessage = null;
        this._showInputMessage = false;
        this._inputTitle = null;
        this._inputMessage = null;
        this._allowBlank = true;
        this._showDropdown = true;
        this._imeMode = 0 /* NO_CONTROL */;
      }
      copy() {
        const dv = new _DataValidation(this._sqref);
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
    };
    exports.DataValidationCollection = class {
      constructor() {
        this._validations = [];
        this._disablePrompts = false;
        this._xWindow = null;
        this._yWindow = null;
      }
      get count() {
        return this._validations.length;
      }
      get disablePrompts() {
        return this._disablePrompts;
      }
      set disablePrompts(value) {
        this._disablePrompts = value;
      }
      get xWindow() {
        return this._xWindow;
      }
      set xWindow(value) {
        this._xWindow = value;
      }
      get yWindow() {
        return this._yWindow;
      }
      set yWindow(value) {
        this._yWindow = value;
      }
      add(sqref, validationType, operator, formula1, formula2) {
        const dv = new exports.DataValidation(sqref);
        if (validationType !== void 0) dv.type = validationType;
        if (operator !== void 0) dv.operator = operator;
        if (formula1 !== void 0) dv.formula1 = formula1;
        if (formula2 !== void 0) dv.formula2 = formula2;
        this._validations.push(dv);
        return dv;
      }
      addValidation(validation) {
        this._validations.push(validation);
      }
      remove(validation) {
        const idx = this._validations.indexOf(validation);
        if (idx >= 0) {
          this._validations.splice(idx, 1);
          return true;
        }
        return false;
      }
      removeAt(index) {
        if (index >= 0 && index < this._validations.length) {
          this._validations.splice(index, 1);
        }
      }
      removeByRange(sqref) {
        this._validations = this._validations.filter((v) => v.sqref !== sqref);
      }
      clear() {
        this._validations = [];
      }
      getValidation(cellRef) {
        for (const v of this._validations) {
          if (v.sqref && this._cellInRange(cellRef, v.sqref)) {
            return v;
          }
        }
        return null;
      }
      _cellInRange(cellRef, sqref) {
        const ranges = sqref.split(/\s+/);
        for (const range of ranges) {
          if (range === cellRef) return true;
          const parts = range.split(":");
          if (parts.length === 2) {
            const [start, end] = parts;
            const cellCol = this._colFromRef(cellRef);
            const cellRow = this._rowFromRef(cellRef);
            const startCol = this._colFromRef(start);
            const startRow = this._rowFromRef(start);
            const endCol = this._colFromRef(end);
            const endRow = this._rowFromRef(end);
            if (cellCol >= startCol && cellCol <= endCol && cellRow >= startRow && cellRow <= endRow) {
              return true;
            }
          }
        }
        return false;
      }
      _colFromRef(ref) {
        const match = ref.match(/^([A-Z]+)/);
        if (!match) return 0;
        let col = 0;
        for (const ch of match[1]) {
          col = col * 26 + (ch.charCodeAt(0) - 64);
        }
        return col;
      }
      _rowFromRef(ref) {
        const match = ref.match(/(\d+)$/);
        return match ? parseInt(match[1], 10) : 0;
      }
      get(index) {
        if (index < 0 || index >= this._validations.length) {
          throw new Error(`Index ${index} out of range`);
        }
        return this._validations[index];
      }
      [Symbol.iterator]() {
        let i = 0;
        const items = this._validations;
        return {
          next() {
            if (i < items.length) {
              return { value: items[i++], done: false };
            }
            return { value: void 0, done: true };
          }
        };
      }
    };
  }
});

// src/features/ConditionalFormat.ts
exports.ConditionalFormat = void 0; exports.ConditionalFormatCollection = void 0;
var init_ConditionalFormat = __esm({
  "src/features/ConditionalFormat.ts"() {
    init_Style();
    exports.ConditionalFormat = class {
      constructor() {
        // Type and identity
        this._type = null;
        this._range = null;
        this._stopIfTrue = false;
        this._priority = 0;
        // Cell value rule
        this._operator = null;
        this._formula1 = null;
        this._formula2 = null;
        // Text rule
        this._textOperator = null;
        this._textFormula = null;
        // Date rule
        this._dateOperator = null;
        this._dateFormula = null;
        // Duplicate/unique
        this._duplicate = null;
        // Top/Bottom
        this._top = null;
        this._percent = false;
        this._rank = 10;
        // Above/Below average
        this._above = null;
        this._stdDev = 0;
        // Color scale
        this._colorScaleType = null;
        this._minColor = null;
        this._midColor = null;
        this._maxColor = null;
        // Data bar
        this._barColor = null;
        this._negativeColor = null;
        this._showBorder = false;
        this._direction = null;
        this._barLength = null;
        // Icon set
        this._iconSetType = null;
        this._reverseIcons = false;
        this._showIconOnly = false;
        // Generic formula
        this._formula = null;
        // DXF styling
        this._font = null;
        this._border = null;
        this._fill = null;
        this._alignment = null;
        this._numberFormat = null;
        // Internal: dxfId for XML round-trip
        this._dxfId = null;
      }
      // ---- Type and identity ----
      get type() {
        return this._type;
      }
      set type(value) {
        this._type = value;
      }
      get range() {
        return this._range;
      }
      set range(value) {
        this._range = value;
      }
      get stopIfTrue() {
        return this._stopIfTrue;
      }
      set stopIfTrue(value) {
        this._stopIfTrue = value;
      }
      get priority() {
        return this._priority;
      }
      set priority(value) {
        this._priority = value;
      }
      // ---- Cell value rule ----
      get operator() {
        return this._operator;
      }
      set operator(value) {
        this._operator = value;
      }
      get formula1() {
        return this._formula1;
      }
      set formula1(value) {
        this._formula1 = value;
      }
      get formula2() {
        return this._formula2;
      }
      set formula2(value) {
        this._formula2 = value;
      }
      // ---- Text rule ----
      get textOperator() {
        return this._textOperator;
      }
      set textOperator(value) {
        this._textOperator = value;
      }
      get textFormula() {
        return this._textFormula;
      }
      set textFormula(value) {
        this._textFormula = value;
      }
      // ---- Date rule ----
      get dateOperator() {
        return this._dateOperator;
      }
      set dateOperator(value) {
        this._dateOperator = value;
      }
      get dateFormula() {
        return this._dateFormula;
      }
      set dateFormula(value) {
        this._dateFormula = value;
      }
      // ---- Duplicate/unique ----
      get duplicate() {
        return this._duplicate;
      }
      set duplicate(value) {
        this._duplicate = value;
      }
      // ---- Top/Bottom ----
      get top() {
        return this._top;
      }
      set top(value) {
        this._top = value;
      }
      get percent() {
        return this._percent;
      }
      set percent(value) {
        this._percent = value;
      }
      get rank() {
        return this._rank;
      }
      set rank(value) {
        this._rank = value;
      }
      // ---- Above/Below average ----
      get above() {
        return this._above;
      }
      set above(value) {
        this._above = value;
      }
      get stdDev() {
        return this._stdDev;
      }
      set stdDev(value) {
        this._stdDev = value;
      }
      // ---- Color scale ----
      get colorScaleType() {
        return this._colorScaleType;
      }
      set colorScaleType(value) {
        this._colorScaleType = value;
      }
      get minColor() {
        return this._minColor;
      }
      set minColor(value) {
        this._minColor = value;
      }
      get midColor() {
        return this._midColor;
      }
      set midColor(value) {
        this._midColor = value;
      }
      get maxColor() {
        return this._maxColor;
      }
      set maxColor(value) {
        this._maxColor = value;
      }
      // ---- Data bar ----
      get barColor() {
        return this._barColor;
      }
      set barColor(value) {
        this._barColor = value;
      }
      get negativeColor() {
        return this._negativeColor;
      }
      set negativeColor(value) {
        this._negativeColor = value;
      }
      get showBorder() {
        return this._showBorder;
      }
      set showBorder(value) {
        this._showBorder = value;
      }
      get direction() {
        return this._direction;
      }
      set direction(value) {
        this._direction = value;
      }
      get barLength() {
        return this._barLength;
      }
      set barLength(value) {
        this._barLength = value;
      }
      // ---- Icon set ----
      get iconSetType() {
        return this._iconSetType;
      }
      set iconSetType(value) {
        this._iconSetType = value;
      }
      get reverseIcons() {
        return this._reverseIcons;
      }
      set reverseIcons(value) {
        this._reverseIcons = value;
      }
      get showIconOnly() {
        return this._showIconOnly;
      }
      set showIconOnly(value) {
        this._showIconOnly = value;
      }
      // ---- Generic formula ----
      get formula() {
        return this._formula;
      }
      set formula(value) {
        this._formula = value;
      }
      // ---- DXF styling (lazily created) ----
      get font() {
        if (!this._font) this._font = new exports.Font();
        return this._font;
      }
      get border() {
        if (!this._border) this._border = new exports.Borders();
        return this._border;
      }
      get fill() {
        if (!this._fill) this._fill = new exports.Fill();
        return this._fill;
      }
      get alignment() {
        if (!this._alignment) this._alignment = new exports.Alignment();
        return this._alignment;
      }
      get numberFormat() {
        return this._numberFormat;
      }
      set numberFormat(value) {
        this._numberFormat = value;
      }
      hasFont() {
        return this._font !== null;
      }
      hasFill() {
        return this._fill !== null;
      }
      hasBorder() {
        return this._border !== null;
      }
      hasAlignment() {
        return this._alignment !== null;
      }
    };
    exports.ConditionalFormatCollection = class {
      constructor() {
        this._formats = [];
      }
      get count() {
        return this._formats.length;
      }
      add() {
        const cf = new exports.ConditionalFormat();
        cf.priority = this._formats.length + 1;
        this._formats.push(cf);
        return cf;
      }
      getByIndex(index) {
        if (index >= 0 && index < this._formats.length) {
          return this._formats[index];
        }
        return null;
      }
      getByRange(rangeStr) {
        return this._formats.filter((cf) => cf.range === rangeStr);
      }
      remove(cf) {
        const idx = this._formats.indexOf(cf);
        if (idx >= 0) {
          this._formats.splice(idx, 1);
          return true;
        }
        return false;
      }
      clear() {
        this._formats = [];
      }
      get(index) {
        if (index < 0 || index >= this._formats.length) {
          throw new Error(`Index ${index} out of range`);
        }
        return this._formats[index];
      }
      [Symbol.iterator]() {
        let i = 0;
        const items = this._formats;
        return {
          next() {
            if (i < items.length) {
              return { value: items[i++], done: false };
            }
            return { value: void 0, done: true };
          }
        };
      }
      get length() {
        return this._formats.length;
      }
    };
  }
});

// src/features/Hyperlink.ts
exports.Hyperlink = void 0; exports.HyperlinkCollection = void 0;
var init_Hyperlink = __esm({
  "src/features/Hyperlink.ts"() {
    exports.Hyperlink = class {
      constructor(rangeAddress, address = "", subAddress = "", textToDisplay = "", screenTip = "") {
        this._deleted = false;
        this._range = rangeAddress;
        this._address = address;
        this._subAddress = subAddress;
        this._textToDisplay = textToDisplay;
        this._screenTip = screenTip;
      }
      get range() {
        return this._range;
      }
      set range(value) {
        this._range = value;
      }
      get address() {
        return this._address;
      }
      set address(value) {
        this._address = value;
      }
      get subAddress() {
        return this._subAddress;
      }
      set subAddress(value) {
        this._subAddress = value;
      }
      get textToDisplay() {
        return this._textToDisplay;
      }
      set textToDisplay(value) {
        this._textToDisplay = value;
      }
      get screenTip() {
        return this._screenTip;
      }
      set screenTip(value) {
        this._screenTip = value;
      }
      get type() {
        if (this._subAddress && !this._address) return "internal";
        if (this._address) {
          if (this._address.startsWith("mailto:")) return "email";
          return "external";
        }
        return "unknown";
      }
      delete() {
        this._deleted = true;
      }
      get isDeleted() {
        return this._deleted;
      }
    };
    exports.HyperlinkCollection = class {
      constructor() {
        this._hyperlinks = [];
      }
      add(rangeAddress, address = "", subAddress = "", textToDisplay = "", screenTip = "") {
        if (!rangeAddress || !rangeAddress.match(/^[A-Z]+\d+(:[A-Z]+\d+)?$/i)) ;
        const hl = new exports.Hyperlink(rangeAddress, address, subAddress, textToDisplay, screenTip);
        this._hyperlinks.push(hl);
        return hl;
      }
      /** Add a pre-built hyperlink object */
      addHyperlink(hl) {
        this._hyperlinks.push(hl);
      }
      delete(indexOrHyperlink) {
        if (typeof indexOrHyperlink === "number") {
          if (indexOrHyperlink >= 0 && indexOrHyperlink < this._hyperlinks.length) {
            this._hyperlinks.splice(indexOrHyperlink, 1);
          }
        } else if (indexOrHyperlink instanceof exports.Hyperlink) {
          const idx = this._hyperlinks.indexOf(indexOrHyperlink);
          if (idx >= 0) this._hyperlinks.splice(idx, 1);
        }
      }
      clear() {
        this._hyperlinks = [];
      }
      get count() {
        return this._hyperlinks.length;
      }
      get(index) {
        if (index < 0 || index >= this._hyperlinks.length) {
          throw new Error(`Index ${index} out of range`);
        }
        return this._hyperlinks[index];
      }
      [Symbol.iterator]() {
        let i = 0;
        const items = this._hyperlinks;
        return {
          next() {
            if (i < items.length) {
              return { value: items[i++], done: false };
            }
            return { value: void 0, done: true };
          }
        };
      }
    };
  }
});

// src/features/AutoFilter.ts
exports.FilterColumn = void 0; exports.AutoFilter = void 0;
var init_AutoFilter = __esm({
  "src/features/AutoFilter.ts"() {
    exports.FilterColumn = class {
      constructor(colId) {
        this._filters = [];
        this._customFilters = [];
        this._colorFilter = null;
        this._dynamicFilter = null;
        this._top10Filter = null;
        this._filterButton = true;
        this._colId = colId;
      }
      get colId() {
        return this._colId;
      }
      get filters() {
        return this._filters;
      }
      get customFilters() {
        return this._customFilters;
      }
      get colorFilter() {
        return this._colorFilter;
      }
      set colorFilter(value) {
        this._colorFilter = value;
      }
      get dynamicFilter() {
        return this._dynamicFilter;
      }
      set dynamicFilter(value) {
        this._dynamicFilter = value;
      }
      get top10Filter() {
        return this._top10Filter;
      }
      set top10Filter(value) {
        this._top10Filter = value;
      }
      get filterButton() {
        return this._filterButton;
      }
      set filterButton(value) {
        this._filterButton = value;
      }
      addFilter(value) {
        if (!this._filters.includes(value)) {
          this._filters.push(value);
        }
      }
      addCustomFilter(operator, value) {
        this._customFilters.push({ operator, value });
      }
      clearFilters() {
        this._filters = [];
        this._customFilters = [];
        this._colorFilter = null;
        this._dynamicFilter = null;
        this._top10Filter = null;
      }
    };
    exports.AutoFilter = class {
      constructor() {
        this._range = null;
        this._filterColumns = /* @__PURE__ */ new Map();
        this._sortState = null;
      }
      get range() {
        return this._range;
      }
      set range(value) {
        this._range = value;
      }
      get filterColumns() {
        return this._filterColumns;
      }
      get sortState() {
        return this._sortState;
      }
      set sortState(value) {
        this._sortState = value;
      }
      setRange(startRow, startCol, endRow, endCol) {
        const startColLetter = this._colToLetter(startCol);
        const endColLetter = this._colToLetter(endCol);
        this._range = `${startColLetter}${startRow + 1}:${endColLetter}${endRow + 1}`;
      }
      filter(colIndex, values) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.clearFilters();
        for (const v of values) {
          fc.addFilter(v);
        }
        return fc;
      }
      addFilter(colIndex, value) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.addFilter(value);
        return fc;
      }
      customFilter(colIndex, operator, value) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.addCustomFilter(operator, value);
        return fc;
      }
      filterByColor(colIndex, color, cellColor = true) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.colorFilter = { color, cellColor };
        return fc;
      }
      filterTop10(colIndex, top = true, percent = false, val = 10) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.top10Filter = { top, percent, val };
        return fc;
      }
      filterDynamic(colIndex, filterType, value) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.dynamicFilter = { type: filterType, value };
        return fc;
      }
      clearColumnFilter(colIndex) {
        const fc = this._filterColumns.get(colIndex);
        if (fc) fc.clearFilters();
      }
      clearAllFilters() {
        for (const fc of this._filterColumns.values()) {
          fc.clearFilters();
        }
      }
      remove() {
        this._range = null;
        this._filterColumns.clear();
        this._sortState = null;
      }
      showFilterButton(colIndex, show = true) {
        const fc = this._getOrCreateColumn(colIndex);
        fc.filterButton = show;
      }
      sort(colIndex, ascending = true) {
        if (!this._range) return;
        this._sortState = {
          ref: this._range,
          columnIndex: colIndex,
          ascending
        };
      }
      getFilterColumn(colIndex) {
        return this._filterColumns.get(colIndex) ?? null;
      }
      hasFilter(colIndex) {
        const fc = this._filterColumns.get(colIndex);
        if (!fc) return false;
        return fc.filters.length > 0 || fc.customFilters.length > 0 || fc.colorFilter !== null || fc.dynamicFilter !== null || fc.top10Filter !== null;
      }
      get hasAnyFilter() {
        return this._range !== null;
      }
      _getOrCreateColumn(colIndex) {
        if (!this._filterColumns.has(colIndex)) {
          this._filterColumns.set(colIndex, new exports.FilterColumn(colIndex));
        }
        return this._filterColumns.get(colIndex);
      }
      _colToLetter(col) {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + c % 26) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      }
    };
  }
});

// src/features/PageBreak.ts
exports.HorizontalPageBreakCollection = void 0; exports.VerticalPageBreakCollection = void 0;
var init_PageBreak = __esm({
  "src/features/PageBreak.ts"() {
    init_Cells();
    exports.HorizontalPageBreakCollection = class {
      constructor() {
        /** @internal */
        this._breaks = /* @__PURE__ */ new Set();
      }
      _normalizeRow(rowOrCell) {
        if (typeof rowOrCell === "string") {
          const [row] = exports.Cells.coordinateFromString(rowOrCell);
          return row - 1;
        }
        if (rowOrCell == null || rowOrCell < 0) {
          throw new Error("row must be >= 0");
        }
        return rowOrCell;
      }
      /**
       * Adds a manual horizontal page break.
       * @param rowOrCell 0-based row index or A1 cell reference.
       * @returns The 0-based row index.
       */
      add(rowOrCell) {
        const row = this._normalizeRow(rowOrCell);
        this._breaks.add(row);
        return row;
      }
      /**
       * Removes the break at zero-based collection index.
       */
      removeAt(index) {
        const list = this.toList();
        if (index < 0 || index >= list.length) {
          throw new RangeError(`Index ${index} out of range`);
        }
        this._breaks.delete(list[index]);
      }
      /**
       * Removes a manual horizontal page break by row/cell.
       */
      remove(rowOrCell) {
        const row = this._normalizeRow(rowOrCell);
        this._breaks.delete(row);
      }
      /** Clears all manual horizontal page breaks. */
      clear() {
        this._breaks.clear();
      }
      get count() {
        return this._breaks.size;
      }
      /** Returns a sorted array of break rows. */
      toList() {
        return Array.from(this._breaks).sort((a, b) => a - b);
      }
      /** Alias */
      to_list() {
        return this.toList();
      }
      get(index) {
        return this.toList()[index];
      }
      [Symbol.iterator]() {
        return this.toList()[Symbol.iterator]();
      }
    };
    exports.VerticalPageBreakCollection = class {
      constructor() {
        /** @internal */
        this._breaks = /* @__PURE__ */ new Set();
      }
      _normalizeColumn(columnOrCell) {
        if (typeof columnOrCell === "string") {
          if (/\d/.test(columnOrCell)) {
            const [, col] = exports.Cells.coordinateFromString(columnOrCell);
            return col - 1;
          }
          return exports.Cells.columnIndexFromString(columnOrCell) - 1;
        }
        if (columnOrCell == null || columnOrCell < 0) {
          throw new Error("column must be >= 0");
        }
        return columnOrCell;
      }
      /**
       * Adds a manual vertical page break.
       * @param columnOrCell 0-based column index, column letters, or A1 cell reference.
       * @returns The 0-based column index.
       */
      add(columnOrCell) {
        const col = this._normalizeColumn(columnOrCell);
        this._breaks.add(col);
        return col;
      }
      /**
       * Removes a manual vertical page break by column/cell.
       */
      remove(columnOrCell) {
        const col = this._normalizeColumn(columnOrCell);
        this._breaks.delete(col);
      }
      /**
       * Removes the break at zero-based collection index.
       */
      removeAt(index) {
        const list = this.toList();
        if (index < 0 || index >= list.length) {
          throw new RangeError(`Index ${index} out of range`);
        }
        this._breaks.delete(list[index]);
      }
      /** Clears all manual vertical page breaks. */
      clear() {
        this._breaks.clear();
      }
      get count() {
        return this._breaks.size;
      }
      /** Returns a sorted array of break columns. */
      toList() {
        return Array.from(this._breaks).sort((a, b) => a - b);
      }
      /** Alias */
      to_list() {
        return this.toList();
      }
      get(index) {
        return this.toList()[index];
      }
      [Symbol.iterator]() {
        return this.toList()[Symbol.iterator]();
      }
    };
  }
});

// src/io/CellValueHandler.ts
function numberToString(n) {
  const s = String(n);
  return s;
}
exports.CELL_TYPE_STRING = void 0; var CELL_TYPE_INLINE_STRING; exports.CELL_TYPE_NUMBER = void 0; exports.CELL_TYPE_BOOLEAN = void 0; exports.CELL_TYPE_ERROR = void 0; var ERROR_VALUES, EXCEL_EPOCH, MS_PER_DAY; exports.CellValueHandler = void 0;
var init_CellValueHandler = __esm({
  "src/io/CellValueHandler.ts"() {
    exports.CELL_TYPE_STRING = "s";
    CELL_TYPE_INLINE_STRING = "str";
    exports.CELL_TYPE_NUMBER = "n";
    exports.CELL_TYPE_BOOLEAN = "b";
    exports.CELL_TYPE_ERROR = "e";
    ERROR_VALUES = /* @__PURE__ */ new Set([
      "#NULL!",
      "#DIV/0!",
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#NUM!",
      "#N/A",
      "#GETTING_DATA"
    ]);
    EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));
    MS_PER_DAY = 864e5;
    exports.CellValueHandler = class {
      /**
       * Determines the OOXML cell type for a JavaScript value.
       *
       * @param value - A cell value (string, number, boolean, Date, null, etc.)
       * @returns The OOXML cell type string, or `null` for empty cells.
       */
      static getCellType(value) {
        if (value === null || value === void 0) {
          return null;
        }
        if (typeof value === "boolean") {
          return exports.CELL_TYPE_BOOLEAN;
        }
        if (typeof value === "number") {
          return exports.CELL_TYPE_NUMBER;
        }
        if (value instanceof Date) {
          return exports.CELL_TYPE_NUMBER;
        }
        if (typeof value === "string") {
          if (value === "") return exports.CELL_TYPE_STRING;
          if (ERROR_VALUES.has(value.toUpperCase())) return exports.CELL_TYPE_ERROR;
          return exports.CELL_TYPE_STRING;
        }
        return exports.CELL_TYPE_STRING;
      }
      /**
       * Formats a cell value for writing into OOXML `<v>` element.
       *
       * @param value - The cell value.
       * @param cellType - Optional pre-determined cell type; auto-detected if omitted.
       * @returns A tuple of `[formattedString, cellType]`, or `[null, null]` for empty.
       */
      static formatValueForXml(value, cellType) {
        if (value === null || value === void 0) {
          return [null, null];
        }
        const resolvedType = cellType ?? this.getCellType(value);
        if (resolvedType === null) {
          return [null, null];
        }
        switch (resolvedType) {
          case exports.CELL_TYPE_BOOLEAN:
            return [value ? "1" : "0", exports.CELL_TYPE_BOOLEAN];
          case exports.CELL_TYPE_NUMBER: {
            if (value instanceof Date) {
              const serial = this.dateToExcelSerial(value);
              return [String(serial), exports.CELL_TYPE_NUMBER];
            }
            if (typeof value === "number") {
              if (!isFinite(value)) {
                return ["#NUM!", exports.CELL_TYPE_ERROR];
              }
              return [numberToString(value), exports.CELL_TYPE_NUMBER];
            }
            return [String(value), exports.CELL_TYPE_NUMBER];
          }
          case exports.CELL_TYPE_ERROR:
            return [String(value).toUpperCase(), exports.CELL_TYPE_ERROR];
          case exports.CELL_TYPE_STRING:
          case CELL_TYPE_INLINE_STRING:
          default:
            return [String(value), exports.CELL_TYPE_STRING];
        }
      }
      /**
       * Parses an OOXML cell value string back into a JavaScript value.
       *
       * @param valueStr - The raw string from the `<v>` element.
       * @param cellType - The cell type attribute ('s', 'n', 'b', 'e', 'str').
       * @param sharedStrings - Optional shared string table for type 's'.
       * @returns The parsed JavaScript value.
       */
      static parseValueFromXml(valueStr, cellType, sharedStrings) {
        if (valueStr === null || valueStr === void 0) {
          return null;
        }
        switch (cellType) {
          case exports.CELL_TYPE_BOOLEAN:
            return valueStr === "1" || valueStr.toLowerCase() === "true";
          case exports.CELL_TYPE_STRING: {
            const idx = parseInt(valueStr, 10);
            if (sharedStrings && !isNaN(idx) && idx >= 0 && idx < sharedStrings.length) {
              return sharedStrings[idx];
            }
            return valueStr;
          }
          case CELL_TYPE_INLINE_STRING:
            return valueStr;
          case exports.CELL_TYPE_ERROR:
            return valueStr;
          case exports.CELL_TYPE_NUMBER: {
            const num = parseFloat(valueStr);
            if (isNaN(num)) return valueStr;
            return num;
          }
          default: {
            const num = parseFloat(valueStr);
            if (!isNaN(num) && String(num) === valueStr.trim()) {
              return num;
            }
            return valueStr;
          }
        }
      }
      // ── Date Conversion ─────────────────────────────────────────────────────
      /**
       * Converts a JavaScript Date to an Excel serial date number.
       *
       * The serial number represents days since 1899-12-30 (Excel epoch).
       * Time-of-day is represented as the fractional part.
       *
       * @param dt - The Date to convert.
       * @returns The Excel serial date number.
       */
      static dateToExcelSerial(dt) {
        const utcMs = Date.UTC(
          dt.getFullYear(),
          dt.getMonth(),
          dt.getDate(),
          dt.getHours(),
          dt.getMinutes(),
          dt.getSeconds(),
          dt.getMilliseconds()
        ) - EXCEL_EPOCH.getTime();
        let serial = utcMs / MS_PER_DAY;
        serial = Math.round(serial * 1e10) / 1e10;
        return serial;
      }
      /**
       * Converts an Excel serial date number to a JavaScript Date.
       *
       * @param serial - The Excel serial date number.
       * @returns A Date object.
       */
      static excelSerialToDate(serial) {
        const ms = serial * MS_PER_DAY + EXCEL_EPOCH.getTime();
        return new Date(ms);
      }
      // ── Error Detection ─────────────────────────────────────────────────────
      /**
       * Checks if a string value represents an Excel error.
       */
      static isErrorValue(value) {
        return ERROR_VALUES.has(value.toUpperCase());
      }
      /**
       * Returns the normalized error string, or `null` if not an error.
       */
      static getErrorType(value) {
        const upper = value.toUpperCase();
        if (ERROR_VALUES.has(upper)) return upper;
        return null;
      }
    };
  }
});

// src/features/FormulaEvaluator.ts
var FormulaEvaluator_exports = {};
__export(FormulaEvaluator_exports, {
  FormulaEvaluator: () => exports.FormulaEvaluator
});
function isError(val) {
  if (typeof val !== "string") return false;
  return val === ERRORS.VALUE || val === ERRORS.REF || val === ERRORS.NAME || val === ERRORS.DIV0 || val === ERRORS.NA || val === ERRORS.NUM || val === ERRORS.NULL;
}
function tokenize(formula) {
  const tokens = [];
  let i = 0;
  const len = formula.length;
  while (i < len) {
    const ch = formula[i];
    if (ch === " " || ch === "	" || ch === "\n" || ch === "\r") {
      i++;
      continue;
    }
    if (ch === '"') {
      let str = "";
      i++;
      while (i < len) {
        if (formula[i] === '"') {
          if (i + 1 < len && formula[i + 1] === '"') {
            str += '"';
            i += 2;
          } else {
            break;
          }
        } else {
          str += formula[i];
          i++;
        }
      }
      i++;
      tokens.push({ type: "STRING" /* STRING */, value: str });
      continue;
    }
    if (ch === "#") {
      const remaining = formula.substring(i);
      const errMatch = remaining.match(/^(#DIV\/0!|#VALUE!|#REF!|#NAME\?|#N\/A|#NUM!|#NULL!)/i);
      if (errMatch) {
        tokens.push({ type: "ERROR" /* ERROR */, value: errMatch[1].toUpperCase() });
        i += errMatch[1].length;
        continue;
      }
      let errStr = "#";
      i++;
      while (i < len && formula[i] !== " " && formula[i] !== "," && formula[i] !== ")" && formula[i] !== "+" && formula[i] !== "-" && formula[i] !== "*" && formula[i] !== "&" && formula[i] !== "^" && formula[i] !== "=" && formula[i] !== "<" && formula[i] !== ">") {
        errStr += formula[i];
        i++;
      }
      tokens.push({ type: "ERROR" /* ERROR */, value: errStr });
      continue;
    }
    if (ch >= "0" && ch <= "9" || ch === "." && i + 1 < len && formula[i + 1] >= "0" && formula[i + 1] <= "9") {
      let num = "";
      while (i < len && (formula[i] >= "0" && formula[i] <= "9" || formula[i] === ".")) {
        num += formula[i];
        i++;
      }
      if (i < len && (formula[i] === "e" || formula[i] === "E")) {
        num += formula[i];
        i++;
        if (i < len && (formula[i] === "+" || formula[i] === "-")) {
          num += formula[i];
          i++;
        }
        while (i < len && formula[i] >= "0" && formula[i] <= "9") {
          num += formula[i];
          i++;
        }
      }
      if (i < len && formula[i] === "%") {
        const val = parseFloat(num) / 100;
        tokens.push({ type: "NUMBER" /* NUMBER */, value: String(val) });
        i++;
      } else {
        tokens.push({ type: "NUMBER" /* NUMBER */, value: num });
      }
      continue;
    }
    if (ch === "(") {
      tokens.push({ type: "LPAREN" /* LPAREN */, value: "(" });
      i++;
      continue;
    }
    if (ch === ")") {
      tokens.push({ type: "RPAREN" /* RPAREN */, value: ")" });
      i++;
      continue;
    }
    if (ch === ",") {
      tokens.push({ type: "COMMA" /* COMMA */, value: "," });
      i++;
      continue;
    }
    if (ch === "&") {
      tokens.push({ type: "CONCAT" /* CONCAT */, value: "&" });
      i++;
      continue;
    }
    if (ch === "+") {
      if (tokens.length === 0 || tokens[tokens.length - 1].type === "LPAREN" /* LPAREN */ || tokens[tokens.length - 1].type === "COMMA" /* COMMA */ || tokens[tokens.length - 1].type === "OPERATOR" /* OPERATOR */ || tokens[tokens.length - 1].type === "COMPARISON" /* COMPARISON */ || tokens[tokens.length - 1].type === "CONCAT" /* CONCAT */) {
        tokens.push({ type: "UNARY_PLUS" /* UNARY_PLUS */, value: "+" });
      } else {
        tokens.push({ type: "OPERATOR" /* OPERATOR */, value: "+" });
      }
      i++;
      continue;
    }
    if (ch === "-") {
      if (tokens.length === 0 || tokens[tokens.length - 1].type === "LPAREN" /* LPAREN */ || tokens[tokens.length - 1].type === "COMMA" /* COMMA */ || tokens[tokens.length - 1].type === "OPERATOR" /* OPERATOR */ || tokens[tokens.length - 1].type === "COMPARISON" /* COMPARISON */ || tokens[tokens.length - 1].type === "CONCAT" /* CONCAT */) {
        tokens.push({ type: "UNARY_MINUS" /* UNARY_MINUS */, value: "-" });
      } else {
        tokens.push({ type: "OPERATOR" /* OPERATOR */, value: "-" });
      }
      i++;
      continue;
    }
    if (ch === "*" || ch === "/" || ch === "^") {
      tokens.push({ type: "OPERATOR" /* OPERATOR */, value: ch });
      i++;
      continue;
    }
    if (ch === "<") {
      if (i + 1 < len && formula[i + 1] === ">") {
        tokens.push({ type: "COMPARISON" /* COMPARISON */, value: "<>" });
        i += 2;
      } else if (i + 1 < len && formula[i + 1] === "=") {
        tokens.push({ type: "COMPARISON" /* COMPARISON */, value: "<=" });
        i += 2;
      } else {
        tokens.push({ type: "COMPARISON" /* COMPARISON */, value: "<" });
        i++;
      }
      continue;
    }
    if (ch === ">") {
      if (i + 1 < len && formula[i + 1] === "=") {
        tokens.push({ type: "COMPARISON" /* COMPARISON */, value: ">=" });
        i += 2;
      } else {
        tokens.push({ type: "COMPARISON" /* COMPARISON */, value: ">" });
        i++;
      }
      continue;
    }
    if (ch === "=") {
      tokens.push({ type: "COMPARISON" /* COMPARISON */, value: "=" });
      i++;
      continue;
    }
    if (ch >= "A" && ch <= "Z" || ch >= "a" && ch <= "z" || ch === "_" || ch === "$" || ch === "'") {
      if (ch === "'") {
        let sheetName = "";
        i++;
        while (i < len && formula[i] !== "'") {
          sheetName += formula[i];
          i++;
        }
        i++;
        if (i < len && formula[i] === "!") {
          i++;
          let cellPart = "";
          while (i < len && (formula[i] >= "A" && formula[i] <= "Z" || formula[i] >= "a" && formula[i] <= "z" || formula[i] >= "0" && formula[i] <= "9" || formula[i] === "$" || formula[i] === ":")) {
            cellPart += formula[i];
            i++;
          }
          if (cellPart.includes(":")) {
            tokens.push({ type: "RANGE_REF" /* RANGE_REF */, value: `'${sheetName}'!${cellPart}` });
          } else {
            tokens.push({ type: "SHEET_REF" /* SHEET_REF */, value: `'${sheetName}'!${cellPart}` });
          }
        }
        continue;
      }
      let ident = "";
      while (i < len && (formula[i] >= "A" && formula[i] <= "Z" || formula[i] >= "a" && formula[i] <= "z" || formula[i] >= "0" && formula[i] <= "9" || formula[i] === "_" || formula[i] === "$" || formula[i] === ".")) {
        ident += formula[i];
        i++;
      }
      if (i < len && formula[i] === "!") {
        const sheetName = ident;
        i++;
        let cellPart = "";
        while (i < len && (formula[i] >= "A" && formula[i] <= "Z" || formula[i] >= "a" && formula[i] <= "z" || formula[i] >= "0" && formula[i] <= "9" || formula[i] === "$" || formula[i] === ":")) {
          cellPart += formula[i];
          i++;
        }
        if (cellPart.includes(":")) {
          tokens.push({ type: "RANGE_REF" /* RANGE_REF */, value: `${sheetName}!${cellPart}` });
        } else {
          tokens.push({ type: "SHEET_REF" /* SHEET_REF */, value: `${sheetName}!${cellPart}` });
        }
        continue;
      }
      if (i < len && formula[i] === ":") {
        const startRef = ident;
        i++;
        let endRef = "";
        while (i < len && (formula[i] >= "A" && formula[i] <= "Z" || formula[i] >= "a" && formula[i] <= "z" || formula[i] >= "0" && formula[i] <= "9" || formula[i] === "$")) {
          endRef += formula[i];
          i++;
        }
        tokens.push({ type: "RANGE_REF" /* RANGE_REF */, value: `${startRef}:${endRef}` });
        continue;
      }
      if (i < len && formula[i] === "(") {
        tokens.push({ type: "FUNCTION" /* FUNCTION */, value: ident.toUpperCase() });
        continue;
      }
      const upper = ident.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        tokens.push({ type: "BOOLEAN" /* BOOLEAN */, value: upper });
        continue;
      }
      const cleanIdent = ident.replace(/\$/g, "");
      if (/^[A-Za-z]+[0-9]+$/.test(cleanIdent)) {
        tokens.push({ type: "CELL_REF" /* CELL_REF */, value: ident });
        continue;
      }
      tokens.push({ type: "NAME" /* NAME */, value: ident });
      continue;
    }
    i++;
  }
  return tokens;
}
function precedence(op) {
  switch (op) {
    case "^":
      return 5;
    case "*":
    case "/":
      return 4;
    case "+":
    case "-":
      return 3;
    case "&":
      return 2;
    case "=":
    case "<>":
    case "<":
    case ">":
    case "<=":
    case ">=":
      return 1;
    default:
      return 0;
  }
}
function isLeftAssociative(op) {
  return op !== "^";
}
var ERRORS, _FormulaEvaluator; exports.FormulaEvaluator = void 0;
var init_FormulaEvaluator = __esm({
  "src/features/FormulaEvaluator.ts"() {
    init_Cells();
    init_CellValueHandler();
    ERRORS = {
      VALUE: "#VALUE!",
      REF: "#REF!",
      NAME: "#NAME?",
      DIV0: "#DIV/0!",
      NA: "#N/A",
      NUM: "#NUM!",
      NULL: "#NULL!"
    };
    _FormulaEvaluator = class _FormulaEvaluator {
      constructor(workbook) {
        this._definedNamesCache = null;
        this._evaluatingCells = /* @__PURE__ */ new Set();
        this._maxDepth = 100;
        this._workbook = workbook;
      }
      // ── Main API ────────────────────────────────────────────────────────────
      /**
       * Evaluate a formula and return the result.
       *
       * @param formula - The formula string (with or without leading '=').
       * @param worksheet - The worksheet context for cell references.
       * @returns The evaluated value, or an error string on failure.
       */
      evaluate(formula, worksheet) {
        if (formula == null) return null;
        let expr = formula;
        if (expr.startsWith("=")) {
          expr = expr.substring(1);
        }
        try {
          this._evaluatingCells.clear();
          const tokens = tokenize(expr);
          const result = this._evaluateTokens(tokens, worksheet ?? null);
          return result;
        } catch (e) {
          if (isError(e?.message)) return e.message;
          return ERRORS.VALUE;
        }
      }
      /**
       * Evaluate all formulas in all worksheets of the workbook.
       * Sets each cell's value to the evaluated result.
       */
      evaluateAll(singleSheet) {
        const sheets = singleSheet ? [singleSheet] : this._workbook.worksheets;
        for (const ws of sheets) {
          for (const [ref, cell] of ws.cells._cells.entries()) {
            if (cell.formula) {
              const result = this.evaluate(cell.formula, ws);
              if (result !== null && result !== void 0) {
                cell.value = result;
              }
            }
          }
        }
      }
      /** snake_case alias */
      evaluate_all(singleSheet) {
        this.evaluateAll(singleSheet);
      }
      // ── Token evaluation (recursive descent parser) ─────────────────────────
      _evaluateTokens(tokens, worksheet) {
        if (tokens.length === 0) return null;
        const { result } = this._parseExpression(tokens, 0, worksheet, 0);
        return result;
      }
      /**
       * Parse an expression with operator precedence (Pratt parser style).
       * Returns { result, pos } where pos is the next token index.
       */
      _parseExpression(tokens, pos, worksheet, minPrec) {
        let { result: left, pos: nextPos } = this._parseAtom(tokens, pos, worksheet);
        while (nextPos < tokens.length) {
          const tok = tokens[nextPos];
          if (tok.type === "OPERATOR" /* OPERATOR */) {
            const prec = precedence(tok.value);
            if (prec < minPrec) break;
            const nextMinPrec = isLeftAssociative(tok.value) ? prec + 1 : prec;
            nextPos++;
            const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, nextMinPrec);
            nextPos = rPos;
            left = this._applyArithmetic(tok.value, left, right);
            continue;
          }
          if (tok.type === "COMPARISON" /* COMPARISON */) {
            const prec = precedence(tok.value);
            if (prec < minPrec) break;
            nextPos++;
            const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, prec + 1);
            nextPos = rPos;
            left = this._applyComparison(tok.value, left, right);
            continue;
          }
          if (tok.type === "CONCAT" /* CONCAT */) {
            const prec = precedence("&");
            if (prec < minPrec) break;
            nextPos++;
            const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, prec + 1);
            nextPos = rPos;
            left = String(left ?? "") + String(right ?? "");
            continue;
          }
          break;
        }
        return { result: left, pos: nextPos };
      }
      /**
       * Parse an atomic value (number, string, cell ref, function call, parenthesised expression, etc.)
       */
      _parseAtom(tokens, pos, worksheet) {
        if (pos >= tokens.length) return { result: null, pos };
        const tok = tokens[pos];
        switch (tok.type) {
          case "NUMBER" /* NUMBER */:
            return { result: parseFloat(tok.value), pos: pos + 1 };
          case "STRING" /* STRING */:
            return { result: tok.value, pos: pos + 1 };
          case "BOOLEAN" /* BOOLEAN */:
            return { result: tok.value === "TRUE", pos: pos + 1 };
          case "ERROR" /* ERROR */:
            return { result: tok.value, pos: pos + 1 };
          case "UNARY_MINUS" /* UNARY_MINUS */: {
            const { result, pos: nPos } = this._parseAtom(tokens, pos + 1, worksheet);
            const num = this._toNumber(result);
            if (isError(num)) return { result: num, pos: nPos };
            return { result: -num, pos: nPos };
          }
          case "UNARY_PLUS" /* UNARY_PLUS */: {
            const { result, pos: nPos } = this._parseAtom(tokens, pos + 1, worksheet);
            const num = this._toNumber(result);
            if (isError(num)) return { result: num, pos: nPos };
            return { result: +num, pos: nPos };
          }
          case "LPAREN" /* LPAREN */: {
            const { result, pos: nPos } = this._parseExpression(tokens, pos + 1, worksheet, 0);
            const rPos = nPos < tokens.length && tokens[nPos].type === "RPAREN" /* RPAREN */ ? nPos + 1 : nPos;
            return { result, pos: rPos };
          }
          case "CELL_REF" /* CELL_REF */: {
            const val = this._resolveCellRef(tok.value, worksheet);
            return { result: val, pos: pos + 1 };
          }
          case "SHEET_REF" /* SHEET_REF */: {
            const val = this._resolveSheetRef(tok.value, worksheet);
            return { result: val, pos: pos + 1 };
          }
          case "RANGE_REF" /* RANGE_REF */: {
            const vals = this._resolveRange(tok.value, worksheet);
            if (Array.isArray(vals) && vals.length === 1) return { result: vals[0], pos: pos + 1 };
            return { result: vals, pos: pos + 1 };
          }
          case "FUNCTION" /* FUNCTION */: {
            const funcName = tok.value;
            pos++;
            if (pos < tokens.length && tokens[pos].type === "LPAREN" /* LPAREN */) {
              pos++;
            }
            const argTokenGroups = [];
            let currentArgTokens = [];
            let parenDepth = 0;
            while (pos < tokens.length) {
              const t = tokens[pos];
              if (t.type === "RPAREN" /* RPAREN */ && parenDepth === 0) {
                if (currentArgTokens.length > 0) {
                  argTokenGroups.push(currentArgTokens);
                }
                pos++;
                break;
              }
              if (t.type === "LPAREN" /* LPAREN */) parenDepth++;
              if (t.type === "RPAREN" /* RPAREN */) parenDepth--;
              if (t.type === "COMMA" /* COMMA */ && parenDepth === 0) {
                argTokenGroups.push(currentArgTokens);
                currentArgTokens = [];
                pos++;
                continue;
              }
              currentArgTokens.push(t);
              pos++;
            }
            if (argTokenGroups.length === 0 && currentArgTokens.length === 0) {
              return { result: this._callFunction(funcName, [], worksheet), pos };
            }
            const result = this._callFunctionWithTokens(funcName, argTokenGroups, worksheet);
            return { result, pos };
          }
          case "NAME" /* NAME */: {
            const names = this._getDefinedNames();
            const upperName = tok.value.toUpperCase();
            for (const [name, refersTo] of names.entries()) {
              if (name.toUpperCase() === upperName) {
                const resolved = this._resolveDefinedName(refersTo, worksheet);
                return { result: resolved, pos: pos + 1 };
              }
            }
            return { result: ERRORS.NAME, pos: pos + 1 };
          }
          default:
            return { result: null, pos: pos + 1 };
        }
      }
      // ── Defined Names ───────────────────────────────────────────────────────
      _getDefinedNames() {
        if (this._definedNamesCache === null) {
          this._definedNamesCache = /* @__PURE__ */ new Map();
          try {
            const names = this._workbook.properties.definedNames.toArray();
            for (const dn of names) {
              this._definedNamesCache.set(dn.name, dn.refersTo);
            }
          } catch {
          }
        }
        return this._definedNamesCache;
      }
      _resolveDefinedName(refersTo, worksheet) {
        const cleaned = refersTo.replace(/\$/g, "");
        if (cleaned.includes(":")) {
          return this._resolveRange(cleaned, worksheet);
        }
        if (cleaned.includes("!")) {
          return this._resolveSheetRef(cleaned, worksheet);
        }
        return this._resolveCellRef(cleaned, worksheet);
      }
      // ── Cell Reference Resolution ───────────────────────────────────────────
      _resolveCellRef(ref, worksheet) {
        if (!worksheet) return ERRORS.REF;
        const cleanRef = ref.replace(/\$/g, "").toUpperCase();
        const cellKey = `${worksheet.name}!${cleanRef}`;
        if (this._evaluatingCells.has(cellKey)) {
          return 0;
        }
        const cell = worksheet.cells._cells.get(cleanRef);
        if (!cell) return 0;
        if (cell.value !== null && cell.value !== void 0) {
          return cell.value;
        }
        if (cell.formula) {
          this._evaluatingCells.add(cellKey);
          try {
            const result = this.evaluate(cell.formula, worksheet);
            return result;
          } finally {
            this._evaluatingCells.delete(cellKey);
          }
        }
        return null;
      }
      _resolveSheetRef(ref, _worksheet) {
        let sheetPart;
        let cellPart;
        if (ref.includes("!")) {
          [sheetPart, cellPart] = ref.split("!", 2);
        } else {
          return ERRORS.REF;
        }
        sheetPart = sheetPart.replace(/^'|'$/g, "");
        const targetSheet = this._workbook.worksheets.find((ws) => ws.name === sheetPart);
        if (!targetSheet) return ERRORS.REF;
        const cleanCell = cellPart.replace(/\$/g, "");
        if (cleanCell.includes(":")) {
          return this._resolveRange(`${sheetPart}!${cleanCell}`, _worksheet);
        }
        return this._resolveCellRef(cleanCell, targetSheet);
      }
      /**
       * Resolves a range reference into an array of cell values.
       */
      _resolveRange(rangeRef, worksheet) {
        let sheetName = null;
        let rangePart;
        if (rangeRef.includes("!")) {
          const parts = rangeRef.split("!", 2);
          sheetName = parts[0].replace(/^'|'$/g, "");
          rangePart = parts[1];
        } else {
          rangePart = rangeRef;
        }
        const targetSheet = sheetName ? this._workbook.worksheets.find((ws) => ws.name === sheetName) : worksheet;
        if (!targetSheet) return [];
        const cleanRange = rangePart.replace(/\$/g, "");
        const [startRef, endRef] = cleanRange.split(":");
        try {
          const [startRow, startCol] = exports.Cells.coordinateFromString(startRef.toUpperCase());
          const [endRow, endCol] = exports.Cells.coordinateFromString(endRef.toUpperCase());
          const values = [];
          const minRow = Math.min(startRow, endRow);
          const maxRow = Math.max(startRow, endRow);
          const minCol = Math.min(startCol, endCol);
          const maxCol = Math.max(startCol, endCol);
          for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
              const cellRef = exports.Cells.coordinateToString(r, c);
              const val = this._resolveCellRef(cellRef, targetSheet);
              values.push(val);
            }
          }
          return values;
        } catch {
          return [];
        }
      }
      /**
       * Resolves a range reference returning an array of arrays (rows x cols) for lookup functions.
       */
      _resolveRange2D(rangeRef, worksheet) {
        let sheetName = null;
        let rangePart;
        if (rangeRef.includes("!")) {
          const parts = rangeRef.split("!", 2);
          sheetName = parts[0].replace(/^'|'$/g, "");
          rangePart = parts[1];
        } else {
          rangePart = rangeRef;
        }
        const targetSheet = sheetName ? this._workbook.worksheets.find((ws) => ws.name === sheetName) : worksheet;
        if (!targetSheet) return [];
        const cleanRange = rangePart.replace(/\$/g, "");
        const [startRef, endRef] = cleanRange.split(":");
        try {
          const [startRow, startCol] = exports.Cells.coordinateFromString(startRef.toUpperCase());
          const [endRow, endCol] = exports.Cells.coordinateFromString(endRef.toUpperCase());
          const rows = [];
          const minRow = Math.min(startRow, endRow);
          const maxRow = Math.max(startRow, endRow);
          const minCol = Math.min(startCol, endCol);
          const maxCol = Math.max(startCol, endCol);
          for (let r = minRow; r <= maxRow; r++) {
            const row = [];
            for (let c = minCol; c <= maxCol; c++) {
              const cellRef = exports.Cells.coordinateToString(r, c);
              const val = this._resolveCellRef(cellRef, targetSheet);
              row.push(val);
            }
            rows.push(row);
          }
          return rows;
        } catch {
          return [];
        }
      }
      // ── Arithmetic ──────────────────────────────────────────────────────────
      _applyArithmetic(op, left, right) {
        if (isError(left)) return left;
        if (isError(right)) return right;
        const l = this._toNumber(left);
        const r = this._toNumber(right);
        if (isError(l)) return l;
        if (isError(r)) return r;
        switch (op) {
          case "+":
            return l + r;
          case "-":
            return l - r;
          case "*":
            return l * r;
          case "/":
            if (r === 0) return ERRORS.DIV0;
            return l / r;
          case "^":
            return Math.pow(l, r);
          default:
            return ERRORS.VALUE;
        }
      }
      // ── Comparison ──────────────────────────────────────────────────────────
      _applyComparison(op, left, right) {
        if (isError(left)) return left;
        if (isError(right)) return right;
        const lNum = typeof left === "number" ? left : typeof left === "boolean" ? left ? 1 : 0 : NaN;
        const rNum = typeof right === "number" ? right : typeof right === "boolean" ? right ? 1 : 0 : NaN;
        if (!isNaN(lNum) && !isNaN(rNum)) {
          switch (op) {
            case "=":
              return lNum === rNum;
            case "<>":
              return lNum !== rNum;
            case "<":
              return lNum < rNum;
            case ">":
              return lNum > rNum;
            case "<=":
              return lNum <= rNum;
            case ">=":
              return lNum >= rNum;
          }
        }
        const lStr = String(left ?? "").toLowerCase();
        const rStr = String(right ?? "").toLowerCase();
        switch (op) {
          case "=":
            return lStr === rStr;
          case "<>":
            return lStr !== rStr;
          case "<":
            return lStr < rStr;
          case ">":
            return lStr > rStr;
          case "<=":
            return lStr <= rStr;
          case ">=":
            return lStr >= rStr;
        }
        return false;
      }
      // ── Type Conversion ─────────────────────────────────────────────────────
      _toNumber(val) {
        if (val === null || val === void 0) return 0;
        if (typeof val === "number") return val;
        if (typeof val === "boolean") return val ? 1 : 0;
        if (typeof val === "string") {
          if (isError(val)) return val;
          if (val === "") return 0;
          const num = Number(val);
          if (!isNaN(num)) return num;
          return ERRORS.VALUE;
        }
        return ERRORS.VALUE;
      }
      _toBool(val) {
        if (typeof val === "boolean") return val;
        if (typeof val === "number") return val !== 0;
        if (typeof val === "string") {
          if (val.toUpperCase() === "TRUE") return true;
          if (val.toUpperCase() === "FALSE") return false;
          return val !== "";
        }
        return val != null;
      }
      // ── Flatten helper (for range arguments) ────────────────────────────────
      _flattenArgs(args) {
        const result = [];
        for (const arg of args) {
          if (Array.isArray(arg)) {
            result.push(...this._flattenArgs(arg));
          } else {
            result.push(arg);
          }
        }
        return result;
      }
      _flattenNumeric(args) {
        const flat = this._flattenArgs(args);
        const nums = [];
        for (const v of flat) {
          if (typeof v === "number") nums.push(v);
          else if (typeof v === "boolean") nums.push(v ? 1 : 0);
          else if (typeof v === "string" && v !== "" && !isError(v)) {
            const n = Number(v);
            if (!isNaN(n)) nums.push(n);
          }
        }
        return nums;
      }
      _callFunctionWithTokens(funcName, argTokenGroups, worksheet) {
        const needs2D = _FormulaEvaluator._2D_FUNCS.has(funcName);
        const evaluatedArgs = [];
        for (let i = 0; i < argTokenGroups.length; i++) {
          const argTokens = argTokenGroups[i];
          if (argTokens.length === 1 && argTokens[0].type === "RANGE_REF" /* RANGE_REF */) {
            const isTableArg = needs2D && funcName !== "INDEX" && i === 1 || funcName === "INDEX" && i === 0;
            if (isTableArg) {
              evaluatedArgs.push(this._resolveRange2D(argTokens[0].value, worksheet));
            } else {
              evaluatedArgs.push(this._resolveRange(argTokens[0].value, worksheet));
            }
          } else {
            const { result } = this._parseExpression(argTokens, 0, worksheet, 0);
            evaluatedArgs.push(result);
          }
        }
        return this._callFunction(funcName, evaluatedArgs, worksheet);
      }
      _callFunction(funcName, args, worksheet) {
        switch (funcName) {
          // ── String Functions ──────────────────────────────────────────────
          case "CONCATENATE":
          case "CONCAT":
            return this._funcConcatenate(args);
          case "TEXT":
            return this._funcText(args);
          case "LEN":
            return this._funcLen(args);
          case "TRIM":
            return this._funcTrim(args);
          case "UPPER":
            return this._funcUpper(args);
          case "LOWER":
            return this._funcLower(args);
          case "LEFT":
            return this._funcLeft(args);
          case "RIGHT":
            return this._funcRight(args);
          case "MID":
            return this._funcMid(args);
          case "SUBSTITUTE":
            return this._funcSubstitute(args);
          case "REPT":
            return this._funcRept(args);
          // ── Logic Functions ───────────────────────────────────────────────
          case "IF":
            return this._funcIf(args);
          case "AND":
            return this._funcAnd(args);
          case "OR":
            return this._funcOr(args);
          case "NOT":
            return this._funcNot(args);
          // ── Math Functions ────────────────────────────────────────────────
          case "SUM":
            return this._funcSum(args);
          case "AVERAGE":
            return this._funcAverage(args);
          case "MIN":
            return this._funcMin(args);
          case "MAX":
            return this._funcMax(args);
          case "COUNT":
            return this._funcCount(args);
          case "COUNTA":
            return this._funcCountA(args);
          case "ABS":
            return this._funcAbs(args);
          case "ROUND":
            return this._funcRound(args);
          case "ROUNDUP":
            return this._funcRoundUp(args);
          case "ROUNDDOWN":
            return this._funcRoundDown(args);
          case "INT":
            return this._funcInt(args);
          case "MOD":
            return this._funcMod(args);
          case "POWER":
            return this._funcPower(args);
          case "SQRT":
            return this._funcSqrt(args);
          case "CEILING":
            return this._funcCeiling(args);
          case "FLOOR":
            return this._funcFloor(args);
          // ── Lookup Functions ──────────────────────────────────────────────
          case "VLOOKUP":
            return this._funcVlookup(args, worksheet);
          case "HLOOKUP":
            return this._funcHlookup(args, worksheet);
          case "INDEX":
            return this._funcIndex(args);
          case "MATCH":
            return this._funcMatch(args);
          // ── Date Functions ────────────────────────────────────────────────
          case "TODAY":
            return this._funcToday();
          case "NOW":
            return this._funcNow();
          case "DATE":
            return this._funcDate(args);
          case "YEAR":
            return this._funcYear(args);
          case "MONTH":
            return this._funcMonth(args);
          case "DAY":
            return this._funcDay(args);
          // ── Info Functions ────────────────────────────────────────────────
          case "ISNUMBER":
            return this._funcIsNumber(args);
          case "ISTEXT":
            return this._funcIsText(args);
          case "ISBLANK":
            return this._funcIsBlank(args);
          case "ISERROR":
            return this._funcIsError(args);
          case "ISNA":
            return this._funcIsNa(args);
          // ── Other Functions ───────────────────────────────────────────────
          case "VALUE":
            return this._funcValue(args);
          case "CHOOSE":
            return this._funcChoose(args);
          default:
            return ERRORS.NAME;
        }
      }
      // ── String Function Implementations ─────────────────────────────────────
      _funcConcatenate(args) {
        const flat = this._flattenArgs(args);
        return flat.map((v) => v == null ? "" : String(v)).join("");
      }
      _funcText(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const value = args[0];
        const format = String(args[1] ?? "");
        if (value == null) return "";
        if (typeof value === "number") {
          if (format === "0") return Math.round(value).toString();
          if (format === "0.00") return value.toFixed(2);
          if (/^0\.0+$/.test(format)) {
            const decimals = format.length - 2;
            return value.toFixed(decimals);
          }
        }
        return String(value);
      }
      _funcLen(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const val = args[0];
        if (val == null) return 0;
        return String(val).length;
      }
      _funcTrim(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const val = args[0];
        if (val == null) return "";
        return String(val).split(/\s+/).filter((s) => s).join(" ");
      }
      _funcUpper(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const val = args[0];
        if (val == null) return "";
        return String(val).toUpperCase();
      }
      _funcLower(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const val = args[0];
        if (val == null) return "";
        return String(val).toLowerCase();
      }
      _funcLeft(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const text = String(args[0] ?? "");
        const numChars = args.length >= 2 ? this._toNumber(args[1]) : 1;
        if (isError(numChars)) return numChars;
        return text.substring(0, numChars);
      }
      _funcRight(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const text = String(args[0] ?? "");
        const numChars = args.length >= 2 ? this._toNumber(args[1]) : 1;
        if (isError(numChars)) return numChars;
        return text.substring(text.length - numChars);
      }
      _funcMid(args) {
        if (args.length < 3) return ERRORS.VALUE;
        const text = String(args[0] ?? "");
        const startNum = this._toNumber(args[1]);
        const numChars = this._toNumber(args[2]);
        if (isError(startNum)) return startNum;
        if (isError(numChars)) return numChars;
        return text.substring(startNum - 1, startNum - 1 + numChars);
      }
      _funcSubstitute(args) {
        if (args.length < 3) return ERRORS.VALUE;
        const text = String(args[0] ?? "");
        const oldText = String(args[1] ?? "");
        const newText = String(args[2] ?? "");
        if (args.length >= 4) {
          const instanceNum = this._toNumber(args[3]);
          if (isError(instanceNum)) return instanceNum;
          let count = 0;
          let result = "";
          let searchFrom = 0;
          while (true) {
            const idx = text.indexOf(oldText, searchFrom);
            if (idx === -1) {
              result += text.substring(searchFrom);
              break;
            }
            count++;
            if (count === instanceNum) {
              result += text.substring(searchFrom, idx) + newText;
              result += text.substring(idx + oldText.length);
              break;
            }
            result += text.substring(searchFrom, idx + oldText.length);
            searchFrom = idx + oldText.length;
          }
          return result;
        }
        return text.split(oldText).join(newText);
      }
      _funcRept(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const text = String(args[0] ?? "");
        const times = this._toNumber(args[1]);
        if (isError(times)) return times;
        if (times < 0) return ERRORS.VALUE;
        return text.repeat(times);
      }
      // ── Logic Function Implementations ──────────────────────────────────────
      _funcIf(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const condition = this._toBool(args[0]);
        if (condition) {
          return args[1];
        }
        return args.length > 2 ? args[2] : false;
      }
      _funcAnd(args) {
        const flat = this._flattenArgs(args);
        if (flat.length === 0) return ERRORS.VALUE;
        for (const v of flat) {
          if (isError(v)) return v;
          if (!this._toBool(v)) return false;
        }
        return true;
      }
      _funcOr(args) {
        const flat = this._flattenArgs(args);
        if (flat.length === 0) return ERRORS.VALUE;
        for (const v of flat) {
          if (isError(v)) return v;
          if (this._toBool(v)) return true;
        }
        return false;
      }
      _funcNot(args) {
        if (args.length < 1) return ERRORS.VALUE;
        if (isError(args[0])) return args[0];
        return !this._toBool(args[0]);
      }
      // ── Math Function Implementations ───────────────────────────────────────
      _funcSum(args) {
        const nums = this._flattenNumeric(args);
        let sum = 0;
        for (const n of nums) sum += n;
        return sum;
      }
      _funcAverage(args) {
        const nums = this._flattenNumeric(args);
        if (nums.length === 0) return ERRORS.DIV0;
        let sum = 0;
        for (const n of nums) sum += n;
        return sum / nums.length;
      }
      _funcMin(args) {
        const nums = this._flattenNumeric(args);
        if (nums.length === 0) return 0;
        return Math.min(...nums);
      }
      _funcMax(args) {
        const nums = this._flattenNumeric(args);
        if (nums.length === 0) return 0;
        return Math.max(...nums);
      }
      _funcCount(args) {
        const flat = this._flattenArgs(args);
        let count = 0;
        for (const v of flat) {
          if (typeof v === "number") count++;
          else if (typeof v === "string" && v !== "" && !isError(v)) {
            const n = Number(v);
            if (!isNaN(n)) count++;
          }
        }
        return count;
      }
      _funcCountA(args) {
        const flat = this._flattenArgs(args);
        let count = 0;
        for (const v of flat) {
          if (v !== null && v !== void 0 && v !== "") count++;
        }
        return count;
      }
      _funcAbs(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        if (isError(num)) return num;
        return Math.abs(num);
      }
      _funcRound(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const digits = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(digits)) return digits;
        const factor = Math.pow(10, digits);
        return Math.round(num * factor) / factor;
      }
      _funcRoundUp(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const digits = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(digits)) return digits;
        const factor = Math.pow(10, digits);
        const n = num;
        if (n >= 0) {
          return Math.ceil(n * factor) / factor;
        } else {
          return -Math.ceil(-n * factor) / factor;
        }
      }
      _funcRoundDown(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const digits = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(digits)) return digits;
        const factor = Math.pow(10, digits);
        return Math.trunc(num * factor) / factor;
      }
      _funcInt(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        if (isError(num)) return num;
        return Math.floor(num);
      }
      _funcMod(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const divisor = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(divisor)) return divisor;
        if (divisor === 0) return ERRORS.DIV0;
        const n = num;
        const d = divisor;
        return n - d * Math.floor(n / d);
      }
      _funcPower(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const base = this._toNumber(args[0]);
        const exp = this._toNumber(args[1]);
        if (isError(base)) return base;
        if (isError(exp)) return exp;
        return Math.pow(base, exp);
      }
      _funcSqrt(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        if (isError(num)) return num;
        if (num < 0) return ERRORS.NUM;
        return Math.sqrt(num);
      }
      _funcCeiling(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const significance = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(significance)) return significance;
        if (significance === 0) return 0;
        return Math.ceil(num / significance) * significance;
      }
      _funcFloor(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const num = this._toNumber(args[0]);
        const significance = this._toNumber(args[1]);
        if (isError(num)) return num;
        if (isError(significance)) return significance;
        if (significance === 0) return ERRORS.DIV0;
        return Math.floor(num / significance) * significance;
      }
      // ── Lookup Function Implementations ─────────────────────────────────────
      _funcVlookup(args, worksheet) {
        if (args.length < 3) return ERRORS.VALUE;
        const lookupValue = args[0];
        const tableArray = args[1];
        const colIndexNum = this._toNumber(args[2]);
        if (isError(colIndexNum)) return colIndexNum;
        const rangeLookup = args.length >= 4 ? this._toBool(args[3]) : true;
        if (!Array.isArray(tableArray)) return ERRORS.VALUE;
        const is2D = tableArray.length > 0 && Array.isArray(tableArray[0]);
        let table2D;
        if (is2D) {
          table2D = tableArray;
        } else {
          table2D = tableArray.map((v) => [v]);
        }
        const colIdx = colIndexNum - 1;
        if (colIdx < 0) return ERRORS.VALUE;
        if (rangeLookup) {
          let lastMatch = -1;
          for (let r = 0; r < table2D.length; r++) {
            const cellVal = table2D[r][0];
            if (cellVal == null) continue;
            if (typeof lookupValue === "number" && typeof cellVal === "number") {
              if (cellVal <= lookupValue) lastMatch = r;
              else break;
            } else {
              const cmp = String(cellVal).toLowerCase().localeCompare(String(lookupValue).toLowerCase());
              if (cmp <= 0) lastMatch = r;
              else break;
            }
          }
          if (lastMatch === -1) return ERRORS.NA;
          if (colIdx >= table2D[lastMatch].length) return ERRORS.REF;
          return table2D[lastMatch][colIdx];
        } else {
          for (let r = 0; r < table2D.length; r++) {
            const cellVal = table2D[r][0];
            if (this._valuesEqual(cellVal, lookupValue)) {
              if (colIdx >= table2D[r].length) return ERRORS.REF;
              return table2D[r][colIdx];
            }
          }
          return ERRORS.NA;
        }
      }
      _funcHlookup(args, worksheet) {
        if (args.length < 3) return ERRORS.VALUE;
        const lookupValue = args[0];
        const tableArray = args[1];
        const rowIndexNum = this._toNumber(args[2]);
        if (isError(rowIndexNum)) return rowIndexNum;
        const rangeLookup = args.length >= 4 ? this._toBool(args[3]) : true;
        if (!Array.isArray(tableArray)) return ERRORS.VALUE;
        const is2D = tableArray.length > 0 && Array.isArray(tableArray[0]);
        let table2D;
        if (is2D) {
          table2D = tableArray;
        } else {
          table2D = [tableArray];
        }
        const rowIdx = rowIndexNum - 1;
        if (rowIdx < 0 || rowIdx >= table2D.length) return ERRORS.REF;
        if (table2D.length === 0) return ERRORS.NA;
        const firstRow = table2D[0];
        if (rangeLookup) {
          let lastMatch = -1;
          for (let c = 0; c < firstRow.length; c++) {
            const cellVal = firstRow[c];
            if (cellVal == null) continue;
            if (typeof lookupValue === "number" && typeof cellVal === "number") {
              if (cellVal <= lookupValue) lastMatch = c;
              else break;
            } else {
              const cmp = String(cellVal).toLowerCase().localeCompare(String(lookupValue).toLowerCase());
              if (cmp <= 0) lastMatch = c;
              else break;
            }
          }
          if (lastMatch === -1) return ERRORS.NA;
          return table2D[rowIdx][lastMatch];
        } else {
          for (let c = 0; c < firstRow.length; c++) {
            if (this._valuesEqual(firstRow[c], lookupValue)) {
              return table2D[rowIdx][c];
            }
          }
          return ERRORS.NA;
        }
      }
      _funcIndex(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const arr = args[0];
        const rowNum = this._toNumber(args[1]);
        if (isError(rowNum)) return rowNum;
        if (!Array.isArray(arr)) return ERRORS.REF;
        const is2D = arr.length > 0 && Array.isArray(arr[0]);
        if (is2D) {
          const table2D = arr;
          const rIdx = rowNum - 1;
          if (rIdx < 0 || rIdx >= table2D.length) return ERRORS.REF;
          if (args.length >= 3) {
            const colNum = this._toNumber(args[2]);
            if (isError(colNum)) return colNum;
            const cIdx = colNum - 1;
            if (cIdx < 0 || cIdx >= table2D[rIdx].length) return ERRORS.REF;
            return table2D[rIdx][cIdx];
          }
          if (table2D[rIdx].length === 1) return table2D[rIdx][0];
          return table2D[rIdx];
        }
        const idx = rowNum - 1;
        if (idx < 0 || idx >= arr.length) return ERRORS.REF;
        return arr[idx];
      }
      _funcMatch(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const lookupValue = args[0];
        const lookupArray = args[1];
        const matchType = args.length >= 3 ? this._toNumber(args[2]) : 1;
        if (isError(matchType)) return matchType;
        if (!Array.isArray(lookupArray)) return ERRORS.NA;
        const flat = this._flattenArgs([lookupArray]);
        if (matchType === 0) {
          for (let i = 0; i < flat.length; i++) {
            if (this._valuesEqual(flat[i], lookupValue)) return i + 1;
          }
          return ERRORS.NA;
        }
        if (matchType === 1) {
          let lastMatch = -1;
          for (let i = 0; i < flat.length; i++) {
            if (typeof lookupValue === "number" && typeof flat[i] === "number") {
              if (flat[i] <= lookupValue) lastMatch = i;
              else break;
            }
          }
          if (lastMatch === -1) return ERRORS.NA;
          return lastMatch + 1;
        }
        if (matchType === -1) {
          let lastMatch = -1;
          for (let i = 0; i < flat.length; i++) {
            if (typeof lookupValue === "number" && typeof flat[i] === "number") {
              if (flat[i] >= lookupValue) lastMatch = i;
              else break;
            }
          }
          if (lastMatch === -1) return ERRORS.NA;
          return lastMatch + 1;
        }
        return ERRORS.NA;
      }
      _valuesEqual(a, b) {
        if (a === b) return true;
        if (a == null || b == null) return false;
        if (typeof a === "number" && typeof b === "number") return a === b;
        return String(a).toLowerCase() === String(b).toLowerCase();
      }
      // ── Date Function Implementations ───────────────────────────────────────
      _funcToday() {
        const now = /* @__PURE__ */ new Date();
        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        return exports.CellValueHandler.dateToExcelSerial(today);
      }
      _funcNow() {
        return exports.CellValueHandler.dateToExcelSerial(/* @__PURE__ */ new Date());
      }
      _funcDate(args) {
        if (args.length < 3) return ERRORS.VALUE;
        const year = this._toNumber(args[0]);
        const month = this._toNumber(args[1]);
        const day = this._toNumber(args[2]);
        if (isError(year)) return year;
        if (isError(month)) return month;
        if (isError(day)) return day;
        const dt = new Date(year, month - 1, day);
        return exports.CellValueHandler.dateToExcelSerial(dt);
      }
      _funcYear(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const serial = this._toNumber(args[0]);
        if (isError(serial)) return serial;
        const dt = exports.CellValueHandler.excelSerialToDate(serial);
        return dt.getFullYear();
      }
      _funcMonth(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const serial = this._toNumber(args[0]);
        if (isError(serial)) return serial;
        const dt = exports.CellValueHandler.excelSerialToDate(serial);
        return dt.getMonth() + 1;
      }
      _funcDay(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const serial = this._toNumber(args[0]);
        if (isError(serial)) return serial;
        const dt = exports.CellValueHandler.excelSerialToDate(serial);
        return dt.getDate();
      }
      // ── Info Function Implementations ───────────────────────────────────────
      _funcIsNumber(args) {
        if (args.length < 1) return false;
        return typeof args[0] === "number";
      }
      _funcIsText(args) {
        if (args.length < 1) return false;
        return typeof args[0] === "string" && !isError(args[0]);
      }
      _funcIsBlank(args) {
        if (args.length < 1) return true;
        return args[0] === null || args[0] === void 0 || args[0] === "";
      }
      _funcIsError(args) {
        if (args.length < 1) return false;
        return isError(args[0]);
      }
      _funcIsNa(args) {
        if (args.length < 1) return false;
        return args[0] === ERRORS.NA;
      }
      // ── Other Function Implementations ──────────────────────────────────────
      _funcValue(args) {
        if (args.length < 1) return ERRORS.VALUE;
        const val = args[0];
        if (typeof val === "number") return val;
        if (typeof val === "string") {
          const num = Number(val);
          if (isNaN(num)) return ERRORS.VALUE;
          return num;
        }
        return ERRORS.VALUE;
      }
      _funcChoose(args) {
        if (args.length < 2) return ERRORS.VALUE;
        const indexNum = this._toNumber(args[0]);
        if (isError(indexNum)) return indexNum;
        const idx = indexNum;
        if (idx < 1 || idx >= args.length) return ERRORS.VALUE;
        return args[idx];
      }
    };
    // ── Function Dispatch ───────────────────────────────────────────────────
    /** Functions that need 2D arrays for their table/range argument */
    _FormulaEvaluator._2D_FUNCS = /* @__PURE__ */ new Set(["VLOOKUP", "HLOOKUP", "INDEX"]);
    exports.FormulaEvaluator = _FormulaEvaluator;
  }
});

// src/core/Worksheet.ts
exports.SheetProtection = void 0; exports.SheetProtectionDictWrapper = void 0; exports.Worksheet = void 0;
var init_Worksheet = __esm({
  "src/core/Worksheet.ts"() {
    init_Cells();
    init_Cell();
    init_DataValidation();
    init_ConditionalFormat();
    init_Hyperlink();
    init_AutoFilter();
    init_PageBreak();
    exports.SheetProtection = class {
      constructor() {
        this.sheet = false;
        this.password = null;
        this.objects = false;
        this.scenarios = false;
        this.formatCells = false;
        this.formatColumns = false;
        this.formatRows = false;
        this.insertColumns = false;
        this.insertRows = false;
        this.insertHyperlinks = false;
        this.deleteColumns = false;
        this.deleteRows = false;
        this.selectLockedCells = true;
        this.selectUnlockedCells = true;
        this.sort = false;
        this.autoFilter = false;
        this.pivotTables = false;
      }
      // snake_case aliases
      get format_cells() {
        return this.formatCells;
      }
      set format_cells(v) {
        this.formatCells = v;
      }
      get format_columns() {
        return this.formatColumns;
      }
      set format_columns(v) {
        this.formatColumns = v;
      }
      get format_rows() {
        return this.formatRows;
      }
      set format_rows(v) {
        this.formatRows = v;
      }
      get insert_columns() {
        return this.insertColumns;
      }
      set insert_columns(v) {
        this.insertColumns = v;
      }
      get insert_rows() {
        return this.insertRows;
      }
      set insert_rows(v) {
        this.insertRows = v;
      }
      get insert_hyperlinks() {
        return this.insertHyperlinks;
      }
      set insert_hyperlinks(v) {
        this.insertHyperlinks = v;
      }
      get delete_columns() {
        return this.deleteColumns;
      }
      set delete_columns(v) {
        this.deleteColumns = v;
      }
      get delete_rows() {
        return this.deleteRows;
      }
      set delete_rows(v) {
        this.deleteRows = v;
      }
      get select_locked_cells() {
        return this.selectLockedCells;
      }
      set select_locked_cells(v) {
        this.selectLockedCells = v;
      }
      get select_unlocked_cells() {
        return this.selectUnlockedCells;
      }
      set select_unlocked_cells(v) {
        this.selectUnlockedCells = v;
      }
      get auto_filter() {
        return this.autoFilter;
      }
      set auto_filter(v) {
        this.autoFilter = v;
      }
      get pivot_tables() {
        return this.pivotTables;
      }
      set pivot_tables(v) {
        this.pivotTables = v;
      }
    };
    exports.SheetProtectionDictWrapper = class {
      constructor(sheetProtection) {
        this._protection = sheetProtection;
      }
      get(key, defaultValue) {
        try {
          return this.getItem(key);
        } catch {
          return defaultValue;
        }
      }
      getItem(key) {
        switch (key) {
          case "protected":
          case "sheet":
            return this._protection.sheet;
          case "password":
            return this._protection.password;
          case "objects":
            return this._protection.objects;
          case "scenarios":
            return this._protection.scenarios;
          case "format_cells":
            return this._protection.formatCells;
          case "format_columns":
            return this._protection.formatColumns;
          case "format_rows":
            return this._protection.formatRows;
          case "insert_columns":
            return this._protection.insertColumns;
          case "insert_rows":
            return this._protection.insertRows;
          case "insert_hyperlinks":
            return this._protection.insertHyperlinks;
          case "delete_columns":
            return this._protection.deleteColumns;
          case "delete_rows":
            return this._protection.deleteRows;
          case "select_locked_cells":
            return this._protection.selectLockedCells;
          case "select_unlocked_cells":
            return this._protection.selectUnlockedCells;
          case "sort":
            return this._protection.sort;
          case "auto_filter":
            return this._protection.autoFilter;
          case "pivot_tables":
            return this._protection.pivotTables;
          default:
            throw new Error(`Unknown protection key: ${key}`);
        }
      }
      setItem(key, value) {
        switch (key) {
          case "protected":
          case "sheet":
            this._protection.sheet = value;
            break;
          case "password":
            this._protection.password = value;
            break;
          case "objects":
            this._protection.objects = value;
            break;
          case "scenarios":
            this._protection.scenarios = value;
            break;
          case "format_cells":
            this._protection.formatCells = value;
            break;
          case "format_columns":
            this._protection.formatColumns = value;
            break;
          case "format_rows":
            this._protection.formatRows = value;
            break;
          case "insert_columns":
            this._protection.insertColumns = value;
            break;
          case "insert_rows":
            this._protection.insertRows = value;
            break;
          case "insert_hyperlinks":
            this._protection.insertHyperlinks = value;
            break;
          case "delete_columns":
            this._protection.deleteColumns = value;
            break;
          case "delete_rows":
            this._protection.deleteRows = value;
            break;
          case "select_locked_cells":
            this._protection.selectLockedCells = value;
            break;
          case "select_unlocked_cells":
            this._protection.selectUnlockedCells = value;
            break;
          case "sort":
            this._protection.sort = value;
            break;
          case "auto_filter":
            this._protection.autoFilter = value;
            break;
          case "pivot_tables":
            this._protection.pivotTables = value;
            break;
          default:
            throw new Error(`Unknown protection key: ${key}`);
        }
      }
    };
    exports.Worksheet = class _Worksheet {
      constructor(name = "Sheet1") {
        this._name = name;
        this._cells = new exports.Cells(this);
        this._visible = true;
        this._tabColor = null;
        this._index = 0;
        this._protection = new exports.SheetProtection();
        this._pageSetup = {
          orientation: null,
          paperSize: null,
          scale: null,
          fitToWidth: null,
          fitToHeight: null,
          fitToPage: false
        };
        this._pageMargins = {
          left: 0.75,
          right: 0.75,
          top: 1,
          bottom: 1,
          header: 0.5,
          footer: 0.5
        };
        this._freezePane = null;
        this._mergedCells = [];
        this._rowHeights = {};
        this._columnWidths = {};
        this._hiddenRows = /* @__PURE__ */ new Set();
        this._hiddenColumns = /* @__PURE__ */ new Set();
        this._printArea = null;
        this._sourceXml = null;
        this._dataValidations = new exports.DataValidationCollection();
        this._conditionalFormatting = new exports.ConditionalFormatCollection();
        this._hyperlinks = new exports.HyperlinkCollection();
        this._autoFilter = new exports.AutoFilter();
        this._horizontalPageBreaks = new exports.HorizontalPageBreakCollection();
        this._verticalPageBreaks = new exports.VerticalPageBreakCollection();
        this._workbook = null;
      }
      // ── Properties ──────────────────────────────────────────────────────────
      get name() {
        return this._name;
      }
      set name(v) {
        this._name = v;
      }
      get cells() {
        return this._cells;
      }
      get visible() {
        return this._visible;
      }
      set visible(v) {
        this._visible = v;
      }
      /** Alias for visible */
      get isVisible() {
        return this._visible === true;
      }
      get is_visible() {
        return this.isVisible;
      }
      get tabColor() {
        return this._tabColor;
      }
      set tabColor(v) {
        this._tabColor = v;
      }
      get tab_color() {
        return this._tabColor;
      }
      set tab_color(v) {
        this._tabColor = v;
      }
      get index() {
        return this._index;
      }
      set index(v) {
        this._index = v;
      }
      get protection() {
        return new exports.SheetProtectionDictWrapper(this._protection);
      }
      get pageSetup() {
        return this._pageSetup;
      }
      get page_setup() {
        return this._pageSetup;
      }
      get pageMargins() {
        return this._pageMargins;
      }
      get page_margins() {
        return this._pageMargins;
      }
      get printArea() {
        return this._printArea;
      }
      set printArea(v) {
        this._printArea = v ? this.normalizePrintArea(v) : null;
      }
      get print_area() {
        return this._printArea;
      }
      set print_area(v) {
        this.printArea = v;
      }
      get mergedCells() {
        return [...this._mergedCells];
      }
      get merged_cells() {
        return this.mergedCells;
      }
      get freezePane() {
        return this._freezePane;
      }
      get freeze_pane() {
        return this._freezePane;
      }
      // Feature collections (Phase 4)
      get dataValidations() {
        return this._dataValidations;
      }
      get data_validations() {
        return this._dataValidations;
      }
      get conditionalFormatting() {
        return this._conditionalFormatting;
      }
      get conditional_formatting() {
        return this._conditionalFormatting;
      }
      get hyperlinks() {
        return this._hyperlinks;
      }
      get autoFilter() {
        return this._autoFilter;
      }
      get auto_filter_obj() {
        return this._autoFilter;
      }
      get horizontalPageBreaks() {
        return this._horizontalPageBreaks;
      }
      get horizontal_page_breaks() {
        return this._horizontalPageBreaks;
      }
      get verticalPageBreaks() {
        return this._verticalPageBreaks;
      }
      get vertical_page_breaks() {
        return this._verticalPageBreaks;
      }
      // ── Methods ─────────────────────────────────────────────────────────────
      rename(newName) {
        this._name = newName;
      }
      setVisibility(value) {
        if (value !== true && value !== false && value !== "veryHidden") {
          throw new Error(`Invalid visibility value: ${value}. Use true, false, or 'veryHidden'.`);
        }
        this._visible = value;
      }
      set_visibility(value) {
        this.setVisibility(value);
      }
      getVisibility() {
        return this._visible;
      }
      get_visibility() {
        return this.getVisibility();
      }
      setTabColor(color) {
        if (color !== null && color.length !== 8) {
          throw new Error(`Invalid tab color format: ${color}. Expected 8-char AARRGGBB hex string.`);
        }
        this._tabColor = color ? color.toUpperCase() : null;
      }
      set_tab_color(color) {
        this.setTabColor(color);
      }
      getTabColor() {
        return this._tabColor;
      }
      get_tab_color() {
        return this.getTabColor();
      }
      clearTabColor() {
        this._tabColor = null;
      }
      clear_tab_color() {
        this.clearTabColor();
      }
      // ── Page setup methods ──────────────────────────────────────────────────
      setPageOrientation(orientation) {
        if (orientation !== "portrait" && orientation !== "landscape") {
          throw new Error(`Invalid orientation: ${orientation}. Use 'portrait' or 'landscape'.`);
        }
        this._pageSetup.orientation = orientation;
      }
      set_page_orientation(orientation) {
        this.setPageOrientation(orientation);
      }
      getPageOrientation() {
        return this._pageSetup.orientation;
      }
      get_page_orientation() {
        return this.getPageOrientation();
      }
      setPaperSize(paperSize) {
        this._pageSetup.paperSize = paperSize;
      }
      set_paper_size(paperSize) {
        this.setPaperSize(paperSize);
      }
      getPaperSize() {
        return this._pageSetup.paperSize;
      }
      get_paper_size() {
        return this.getPaperSize();
      }
      setPageMargins(options) {
        if (options.left != null) this._pageMargins.left = options.left;
        if (options.right != null) this._pageMargins.right = options.right;
        if (options.top != null) this._pageMargins.top = options.top;
        if (options.bottom != null) this._pageMargins.bottom = options.bottom;
        if (options.header != null) this._pageMargins.header = options.header;
        if (options.footer != null) this._pageMargins.footer = options.footer;
      }
      set_page_margins(options) {
        this.setPageMargins(options);
      }
      getPageMargins() {
        return { ...this._pageMargins };
      }
      get_page_margins() {
        return this.getPageMargins();
      }
      setFitToPages(width = 1, height = 1) {
        this._pageSetup.fitToWidth = width;
        this._pageSetup.fitToHeight = height;
        this._pageSetup.fitToPage = true;
      }
      set_fit_to_pages(width = 1, height = 1) {
        this.setFitToPages(width, height);
      }
      setPrintScale(scale) {
        if (scale < 10 || scale > 400) throw new Error(`Print scale must be between 10 and 400, got ${scale}.`);
        this._pageSetup.scale = scale;
      }
      set_print_scale(scale) {
        this.setPrintScale(scale);
      }
      // ── Print area ──────────────────────────────────────────────────────────
      setPrintArea(printArea) {
        this.printArea = printArea;
      }
      set_print_area(printArea) {
        this.setPrintArea(printArea);
      }
      SetPrintArea(printArea) {
        this.setPrintArea(printArea);
      }
      clearPrintArea() {
        this._printArea = null;
      }
      clear_print_area() {
        this.clearPrintArea();
      }
      ClearPrintArea() {
        this.clearPrintArea();
      }
      normalizePrintArea(printArea) {
        if (!printArea || !printArea.trim()) throw new Error("print_area must be a non-empty string");
        const parts = [];
        for (const token of printArea.split(",")) {
          const part = token.trim().toUpperCase().replace(/\$/g, "");
          if (!part) continue;
          if (part.includes(":")) {
            const [startRef, endRef] = part.split(":");
            exports.Cells.coordinateFromString(startRef);
            exports.Cells.coordinateFromString(endRef);
            parts.push(`${startRef}:${endRef}`);
          } else {
            exports.Cells.coordinateFromString(part);
            parts.push(part);
          }
        }
        if (parts.length === 0) throw new Error("print_area must contain at least one valid range");
        return parts.join(",");
      }
      // ── Freeze pane ─────────────────────────────────────────────────────────
      setFreezePane(row, column, freezedRows, freezedColumns) {
        this._freezePane = {
          row,
          column,
          freezedRows: freezedRows ?? row,
          freezedColumns: freezedColumns ?? column
        };
      }
      set_freeze_pane(row, column, freezedRows, freezedColumns) {
        this.setFreezePane(row, column, freezedRows, freezedColumns);
      }
      clearFreezePane() {
        this._freezePane = null;
      }
      clear_freeze_pane() {
        this.clearFreezePane();
      }
      // ── Protection ──────────────────────────────────────────────────────────
      isProtected() {
        return this._protection.sheet;
      }
      is_protected() {
        return this.isProtected();
      }
      protect(options = {}) {
        this._protection.sheet = true;
        if (options.password != null) this._protection.password = options.password;
        if (options.formatCells != null) this._protection.formatCells = options.formatCells;
        if (options.formatColumns != null) this._protection.formatColumns = options.formatColumns;
        if (options.formatRows != null) this._protection.formatRows = options.formatRows;
        if (options.insertColumns != null) this._protection.insertColumns = options.insertColumns;
        if (options.insertRows != null) this._protection.insertRows = options.insertRows;
        if (options.deleteColumns != null) this._protection.deleteColumns = options.deleteColumns;
        if (options.deleteRows != null) this._protection.deleteRows = options.deleteRows;
        if (options.sort != null) this._protection.sort = options.sort;
        if (options.autoFilter != null) this._protection.autoFilter = options.autoFilter;
        if (options.insertHyperlinks != null) this._protection.insertHyperlinks = options.insertHyperlinks;
        if (options.pivotTables != null) this._protection.pivotTables = options.pivotTables;
        if (options.selectLockedCells != null) this._protection.selectLockedCells = options.selectLockedCells;
        if (options.selectUnlockedCells != null) this._protection.selectUnlockedCells = options.selectUnlockedCells;
        if (options.objects != null) this._protection.objects = options.objects;
        if (options.scenarios != null) this._protection.scenarios = options.scenarios;
      }
      unprotect(_password) {
        this._protection.sheet = false;
        this._protection.password = null;
      }
      // ── Copy / placeholder methods ──────────────────────────────────────────
      copy(name) {
        const newWs = new _Worksheet(name ?? `${this._name} (copy)`);
        for (const [ref, cell] of this._cells._cells.entries()) {
          const newCell = new exports.Cell(cell.value, cell.formula);
          if (cell.style) newCell.style = cell.style.copy();
          newWs._cells._cells.set(ref, newCell);
        }
        newWs._mergedCells = [...this._mergedCells];
        newWs._printArea = this._printArea;
        return newWs;
      }
      /** Placeholder */
      delete() {
      }
      /** Placeholder */
      move(_index) {
      }
      /** Placeholder */
      select() {
      }
      /** Placeholder */
      activate() {
      }
      /**
       * Evaluates every formula on this worksheet only.
       *
       * Delegates to the workbook-level FormulaEvaluator so that cross-sheet
       * references and defined names resolve correctly.
       */
      calculateFormula() {
        const { FormulaEvaluator: FormulaEvaluator2 } = (init_FormulaEvaluator(), __toCommonJS(FormulaEvaluator_exports));
        const evaluator = new FormulaEvaluator2(this._workbook);
        evaluator.evaluateAll(this);
      }
      /** snake_case alias */
      calculate_formula() {
        this.calculateFormula();
      }
    };
  }
});

// src/io/SharedStrings.ts
function escapeXml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
exports.SharedStringTable = void 0;
var init_SharedStrings = __esm({
  "src/io/SharedStrings.ts"() {
    exports.SharedStringTable = class {
      constructor() {
        /** Ordered list of unique strings */
        this._strings = [];
        /** Reverse lookup: string → index */
        this._stringToIndex = /* @__PURE__ */ new Map();
        /** Total reference count (including duplicates) */
        this._count = 0;
      }
      // ── Properties ──────────────────────────────────────────────────────────
      /** Number of unique strings. */
      get uniqueCount() {
        return this._strings.length;
      }
      /** Total number of string references (≥ uniqueCount). */
      get count() {
        return this._count;
      }
      /** Direct access to the strings array (read-only copy). */
      get strings() {
        return [...this._strings];
      }
      // ── Methods ─────────────────────────────────────────────────────────────
      /**
       * Adds a string to the shared string table.
       *
       * If the string already exists, returns the existing index.
       * Always increments the total reference count.
       *
       * @param text - The string to add.
       * @returns The index of the string in the table.
       */
      addString(text) {
        this._count++;
        const existing = this._stringToIndex.get(text);
        if (existing !== void 0) {
          return existing;
        }
        const index = this._strings.length;
        this._strings.push(text);
        this._stringToIndex.set(text, index);
        return index;
      }
      /**
       * Gets a string by its index.
       *
       * @param index - The 0-based index.
       * @returns The string at that index.
       * @throws {RangeError} If the index is out of range.
       */
      getString(index) {
        if (index < 0 || index >= this._strings.length) {
          throw new RangeError(
            `String index ${index} out of range (0..${this._strings.length - 1})`
          );
        }
        return this._strings[index];
      }
      /**
       * Serializes the shared string table to OOXML.
       *
       * Produces the content of `xl/sharedStrings.xml`.
       */
      toXml() {
        const parts = [];
        parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        parts.push(
          `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this._count}" uniqueCount="${this._strings.length}">`
        );
        for (const text of this._strings) {
          const needsPreserve = text !== text.trim();
          if (needsPreserve) {
            parts.push(`<si><t xml:space="preserve">${escapeXml(text)}</t></si>`);
          } else {
            parts.push(`<si><t>${escapeXml(text)}</t></si>`);
          }
        }
        parts.push("</sst>");
        return parts.join("");
      }
      /**
       * Resets the table to empty state.
       */
      clear() {
        this._strings = [];
        this._stringToIndex.clear();
        this._count = 0;
      }
    };
  }
});

// src/io/XmlDataValidationSaver.ts
var TYPE_TO_XML, OPERATOR_TO_XML, ALERT_TO_XML, IME_TO_XML; exports.DataValidationXmlSaver = void 0;
var init_XmlDataValidationSaver = __esm({
  "src/io/XmlDataValidationSaver.ts"() {
    init_DataValidation();
    TYPE_TO_XML = {
      [0 /* NONE */]: "none",
      [1 /* WHOLE_NUMBER */]: "whole",
      [2 /* DECIMAL */]: "decimal",
      [3 /* LIST */]: "list",
      [4 /* DATE */]: "date",
      [5 /* TIME */]: "time",
      [6 /* TEXT_LENGTH */]: "textLength",
      [7 /* CUSTOM */]: "custom"
    };
    OPERATOR_TO_XML = {
      [0 /* BETWEEN */]: "between",
      [1 /* NOT_BETWEEN */]: "notBetween",
      [2 /* EQUAL */]: "equal",
      [3 /* NOT_EQUAL */]: "notEqual",
      [4 /* GREATER_THAN */]: "greaterThan",
      [5 /* LESS_THAN */]: "lessThan",
      [6 /* GREATER_THAN_OR_EQUAL */]: "greaterThanOrEqual",
      [7 /* LESS_THAN_OR_EQUAL */]: "lessThanOrEqual"
    };
    ALERT_TO_XML = {
      [0 /* STOP */]: "stop",
      [1 /* WARNING */]: "warning",
      [2 /* INFORMATION */]: "information"
    };
    IME_TO_XML = {
      [0 /* NO_CONTROL */]: "noControl",
      [1 /* OFF */]: "off",
      [2 /* ON */]: "on",
      [3 /* DISABLED */]: "disabled",
      [4 /* HIRAGANA */]: "hiragana",
      [5 /* FULL_KATAKANA */]: "fullKatakana",
      [6 /* HALF_KATAKANA */]: "halfKatakana",
      [7 /* FULL_ALPHA */]: "fullAlpha",
      [8 /* HALF_ALPHA */]: "halfAlpha",
      [9 /* FULL_HANGUL */]: "fullHangul",
      [10 /* HALF_HANGUL */]: "halfHangul"
    };
    exports.DataValidationXmlSaver = class {
      constructor(escapeXml6) {
        this._escapeXml = escapeXml6;
      }
      formatDataValidationsXml(collection) {
        if (collection.count === 0) return "";
        const lines = [];
        const attrs = [`count="${collection.count}"`];
        if (collection.disablePrompts) {
          attrs.push('disablePrompts="1"');
        }
        if (collection.xWindow != null) {
          attrs.push(`xWindow="${collection.xWindow}"`);
        }
        if (collection.yWindow != null) {
          attrs.push(`yWindow="${collection.yWindow}"`);
        }
        lines.push(`<dataValidations ${attrs.join(" ")}>`);
        for (const dv of collection) {
          lines.push(this._formatDataValidationXml(dv));
        }
        lines.push("</dataValidations>");
        return lines.join("");
      }
      _formatDataValidationXml(dv) {
        const attrs = [];
        const typeStr = TYPE_TO_XML[dv.type] ?? "none";
        if (typeStr !== "none") {
          attrs.push(`type="${typeStr}"`);
        }
        const opStr = OPERATOR_TO_XML[dv.operator] ?? "between";
        if (opStr !== "between") {
          attrs.push(`operator="${opStr}"`);
        }
        const alertStr = ALERT_TO_XML[dv.alertStyle] ?? "stop";
        if (alertStr !== "stop") {
          attrs.push(`errorStyle="${alertStr}"`);
        }
        if (dv.imeMode !== 0 /* NO_CONTROL */) {
          const imeStr = IME_TO_XML[dv.imeMode];
          if (imeStr) attrs.push(`imeMode="${imeStr}"`);
        }
        if (dv.allowBlank) {
          attrs.push('allowBlank="1"');
        }
        if (!dv.showDropdown) {
          attrs.push('showDropDown="1"');
        }
        if (dv.showInputMessage) {
          attrs.push('showInputMessage="1"');
        }
        if (dv.showErrorMessage) {
          attrs.push('showErrorMessage="1"');
        }
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
        attrs.push(`sqref="${dv.sqref ?? ""}"`);
        let xml = `<dataValidation ${attrs.join(" ")}>`;
        if (dv.formula1 != null) {
          xml += `<formula1>${this._escapeXml(dv.formula1)}</formula1>`;
        }
        if (dv.formula2 != null) {
          xml += `<formula2>${this._escapeXml(dv.formula2)}</formula2>`;
        }
        xml += "</dataValidation>";
        return xml;
      }
    };
  }
});

// src/io/XmlConditionalFormatSaver.ts
var TYPE_MAP; exports.ConditionalFormatXmlSaver = void 0;
var init_XmlConditionalFormatSaver = __esm({
  "src/io/XmlConditionalFormatSaver.ts"() {
    TYPE_MAP = {
      cellIs: "cellIs",
      expression: "expression",
      colorScale: "colorScale",
      dataBar: "dataBar",
      iconSet: "iconSet",
      top10: "top10",
      aboveAverage: "aboveAverage",
      duplicateValues: "duplicateValues",
      uniqueValues: "uniqueValues",
      containsText: "containsText",
      notContainsText: "notContainsText",
      beginsWith: "beginsWith",
      endsWith: "endsWith",
      containsBlanks: "containsBlanks",
      notContainsBlanks: "notContainsBlanks",
      containsErrors: "containsErrors",
      notContainsErrors: "notContainsErrors",
      timePeriod: "timePeriod"
    };
    exports.ConditionalFormatXmlSaver = class {
      constructor(escapeXml6) {
        this._escapeXml = escapeXml6;
      }
      /**
       * Formats conditional formatting XML. Returns DXF entries that need to be
       * added to the styles.xml dxfs collection.
       */
      formatConditionalFormattingXml(collection, startDxfId) {
        if (collection.count === 0) return { xml: "", dxfEntries: [] };
        const dxfEntries = [];
        let dxfId = startDxfId;
        const grouped = /* @__PURE__ */ new Map();
        for (const cf of collection) {
          const range = cf.range ?? "";
          if (!grouped.has(range)) grouped.set(range, []);
          grouped.get(range).push(cf);
        }
        const lines = [];
        for (const [sqref, cfs] of grouped) {
          lines.push(`<conditionalFormatting sqref="${this._escapeXml(sqref)}">`);
          for (const cf of cfs) {
            const hasDxf = cf.hasFont() || cf.hasFill() || cf.hasBorder() || cf.numberFormat != null;
            let cfDxfId = null;
            if (hasDxf) {
              cfDxfId = dxfId++;
              dxfEntries.push(this._formatDxfEntry(cf));
            }
            lines.push(this._formatCfRuleXml(cf, cfDxfId));
          }
          lines.push("</conditionalFormatting>");
        }
        return { xml: lines.join(""), dxfEntries };
      }
      _formatCfRuleXml(cf, dxfId) {
        const attrs = [];
        const ruleType = cf.type ?? "cellIs";
        attrs.push(`type="${TYPE_MAP[ruleType] ?? ruleType}"`);
        if (dxfId !== null) {
          attrs.push(`dxfId="${dxfId}"`);
        }
        attrs.push(`priority="${cf.priority}"`);
        if (cf.stopIfTrue) {
          attrs.push('stopIfTrue="1"');
        }
        if (cf.operator && (ruleType === "cellIs" || ruleType === "expression")) {
          attrs.push(`operator="${cf.operator}"`);
        }
        if (cf.textFormula && ["containsText", "notContainsText", "beginsWith", "endsWith"].includes(ruleType)) {
          attrs.push(`text="${this._escapeXml(cf.textFormula)}"`);
        }
        if (cf.dateOperator && ruleType === "timePeriod") {
          attrs.push(`timePeriod="${cf.dateOperator}"`);
        }
        if (ruleType === "aboveAverage") {
          if (cf.above === false) {
            attrs.push('aboveAverage="0"');
          }
          if (cf.stdDev > 0) {
            attrs.push(`stdDev="${cf.stdDev}"`);
          }
        }
        if (ruleType === "top10") {
          if (cf.top === false) {
            attrs.push('bottom="1"');
          }
          if (cf.percent) {
            attrs.push('percent="1"');
          }
          attrs.push(`rank="${cf.rank}"`);
        }
        let xml = `<cfRule ${attrs.join(" ")}>`;
        if (cf.formula1 != null) {
          xml += `<formula>${this._escapeXml(cf.formula1)}</formula>`;
        }
        if (cf.formula2 != null) {
          xml += `<formula>${this._escapeXml(cf.formula2)}</formula>`;
        }
        if (!cf.formula1 && cf.textFormula) {
          const firstCell = this._getFirstCellFromRange(cf.range ?? "A1");
          const textFormula = this._buildTextRuleFormula(ruleType, cf.textFormula, firstCell);
          if (textFormula) {
            xml += `<formula>${this._escapeXml(textFormula)}</formula>`;
          }
        }
        if (ruleType === "colorScale" && cf.colorScaleType) {
          xml += this._formatColorScaleXml(cf);
        }
        if (ruleType === "dataBar" && cf.barColor) {
          xml += this._formatDataBarXml(cf);
        }
        if (ruleType === "iconSet" && cf.iconSetType) {
          xml += this._formatIconSetXml(cf);
        }
        xml += "</cfRule>";
        return xml;
      }
      _formatColorScaleXml(cf) {
        let xml = "<colorScale>";
        if (cf.colorScaleType === "2color") {
          xml += '<cfvo type="min"/><cfvo type="max"/>';
          xml += `<color rgb="${cf.minColor ?? "FFF8696B"}"/>`;
          xml += `<color rgb="${cf.maxColor ?? "FF63BE7B"}"/>`;
        } else {
          xml += '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>';
          xml += `<color rgb="${cf.minColor ?? "FFF8696B"}"/>`;
          xml += `<color rgb="${cf.midColor ?? "FFFFEB84"}"/>`;
          xml += `<color rgb="${cf.maxColor ?? "FF63BE7B"}"/>`;
        }
        xml += "</colorScale>";
        return xml;
      }
      _formatDataBarXml(cf) {
        let xml = "<dataBar>";
        xml += '<cfvo type="min"/><cfvo type="max"/>';
        xml += `<color rgb="${cf.barColor ?? "FF638EC6"}"/>`;
        xml += "</dataBar>";
        return xml;
      }
      _formatIconSetXml(cf) {
        const attrs = [];
        if (cf.iconSetType) attrs.push(`iconSet="${cf.iconSetType}"`);
        if (cf.reverseIcons) attrs.push('reverse="1"');
        if (cf.showIconOnly) attrs.push('showValue="0"');
        let xml = `<iconSet ${attrs.join(" ")}>`;
        xml += '<cfvo type="percent" val="0"/>';
        xml += '<cfvo type="percent" val="33"/>';
        xml += '<cfvo type="percent" val="67"/>';
        xml += "</iconSet>";
        return xml;
      }
      _formatDxfEntry(cf) {
        let xml = "<dxf>";
        if (cf.hasFont()) {
          const f = cf.font;
          xml += "<font>";
          if (f.bold) xml += "<b/>";
          if (f.italic) xml += "<i/>";
          if (f.underline) xml += "<u/>";
          if (f.strikethrough) xml += "<strike/>";
          if (f.color && f.color !== "FF000000") {
            xml += `<color rgb="${f.color}"/>`;
          }
          if (f.name && f.name !== "Calibri") {
            xml += `<name val="${this._escapeXml(f.name)}"/>`;
          }
          if (f.size && f.size !== 11) {
            xml += `<sz val="${f.size}"/>`;
          }
          xml += "</font>";
        }
        if (cf.hasFill()) {
          const fl = cf.fill;
          xml += "<fill><patternFill";
          if (fl.patternType !== "none") {
            xml += ` patternType="${fl.patternType}"`;
          }
          xml += ">";
          if (fl.foregroundColor && fl.foregroundColor !== "FFFFFFFF") {
            xml += `<fgColor rgb="${fl.foregroundColor}"/>`;
          }
          if (fl.backgroundColor && fl.backgroundColor !== "FFFFFFFF") {
            xml += `<bgColor rgb="${fl.backgroundColor}"/>`;
          }
          xml += "</patternFill></fill>";
        }
        if (cf.hasBorder()) {
          const b = cf.border;
          xml += "<border>";
          const sides = ["left", "right", "top", "bottom"];
          for (const side of sides) {
            const s = b[side];
            if (s.lineStyle !== "none") {
              xml += `<${side} style="${s.lineStyle}">`;
              xml += `<color rgb="${s.color}"/>`;
              xml += `</${side}>`;
            } else {
              xml += `<${side}/>`;
            }
          }
          xml += "</border>";
        }
        if (cf.numberFormat != null) {
          xml += `<numFmt formatCode="${this._escapeXml(cf.numberFormat)}"/>`;
        }
        xml += "</dxf>";
        return xml;
      }
      _getFirstCellFromRange(range) {
        const colonIdx = range.indexOf(":");
        if (colonIdx > 0) return range.substring(0, colonIdx);
        return range;
      }
      _buildTextRuleFormula(ruleType, textValue, firstCell) {
        switch (ruleType) {
          case "containsText":
            return `NOT(ISERROR(SEARCH("${textValue}",${firstCell})))`;
          case "notContainsText":
            return `ISERROR(SEARCH("${textValue}",${firstCell}))`;
          case "beginsWith":
            return `LEFT(${firstCell},${textValue.length})="${textValue}"`;
          case "endsWith":
            return `RIGHT(${firstCell},${textValue.length})="${textValue}"`;
          default:
            return null;
        }
      }
    };
  }
});

// src/io/XmlHyperlinkHandler.ts
function ensureArray(val) {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}
function attr(elem, name, defaultVal = "") {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : defaultVal;
}
function escapeXml2(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
exports.HyperlinkXmlLoader = void 0; exports.HyperlinkXmlSaver = void 0; exports.HyperlinkRelationshipWriter = void 0;
var init_XmlHyperlinkHandler = __esm({
  "src/io/XmlHyperlinkHandler.ts"() {
    init_Hyperlink();
    exports.HyperlinkXmlLoader = class {
      loadHyperlinks(wsRoot, collection, relationships) {
        const hlElems = ensureArray(wsRoot?.hyperlinks?.hyperlink);
        for (const hlElem of hlElems) {
          this._loadHyperlink(hlElem, collection, relationships);
        }
      }
      loadRelationships(relsXml) {
        const rels = /* @__PURE__ */ new Map();
        if (!relsXml) return rels;
        const relationships = ensureArray(relsXml?.Relationships?.Relationship);
        for (const rel of relationships) {
          const id = attr(rel, "Id", "");
          const target = attr(rel, "Target", "");
          const type = attr(rel, "Type", "");
          if (id && type.includes("hyperlink")) {
            rels.set(id, target);
          }
        }
        return rels;
      }
      _loadHyperlink(elem, collection, relationships) {
        const ref = attr(elem, "ref", "");
        if (!ref) return;
        let address = "";
        let subAddress = "";
        const display = attr(elem, "display", "");
        const tooltip = attr(elem, "tooltip", "");
        const rId = attr(elem, "r:id", "") || attr(elem, "id", "");
        if (rId && relationships.has(rId)) {
          address = relationships.get(rId);
        }
        const location = attr(elem, "location", "");
        if (location) {
          subAddress = location;
        }
        const hl = new exports.Hyperlink(ref, address, subAddress, display, tooltip);
        collection.addHyperlink(hl);
      }
    };
    exports.HyperlinkXmlSaver = class {
      constructor() {
        this._relIdCounter = 1;
      }
      formatHyperlinksXml(collection) {
        if (collection.count === 0) return "";
        const lines = ["<hyperlinks>"];
        for (const hl of collection) {
          if (hl.isDeleted) continue;
          lines.push(this._formatHyperlinkXml(hl));
        }
        lines.push("</hyperlinks>");
        return lines.join("");
      }
      _formatHyperlinkXml(hl) {
        const attrs = [`ref="${escapeXml2(hl.range)}"`];
        if (hl.type === "external" || hl.type === "email") {
          const rId = `rId${this._relIdCounter++}`;
          attrs.push(`r:id="${rId}"`);
        }
        if (hl.subAddress) {
          attrs.push(`location="${escapeXml2(hl.subAddress)}"`);
        }
        if (hl.textToDisplay) {
          attrs.push(`display="${escapeXml2(hl.textToDisplay)}"`);
        }
        if (hl.screenTip) {
          attrs.push(`tooltip="${escapeXml2(hl.screenTip)}"`);
        }
        return `<hyperlink ${attrs.join(" ")}/>`;
      }
      getHyperlinkRelationships(collection) {
        const rels = [];
        this._relIdCounter;
        let externalCount = 0;
        for (const hl of collection) {
          if (hl.isDeleted) continue;
          if (hl.type === "external" || hl.type === "email") externalCount++;
        }
        const startRelId = this._relIdCounter - externalCount;
        let rid = startRelId;
        for (const hl of collection) {
          if (hl.isDeleted) continue;
          if (hl.type === "external" || hl.type === "email") {
            rels.push({
              rId: `rId${rid++}`,
              target: hl.address
            });
          }
        }
        return rels;
      }
      resetRelationshipCounter(startRelId = 1) {
        this._relIdCounter = startRelId;
      }
    };
    exports.HyperlinkRelationshipWriter = class {
      static formatRelationshipsXml(relationships, existingRels = []) {
        const lines = [];
        for (const rel of existingRels) {
          lines.push(rel);
        }
        for (const rel of relationships) {
          lines.push(
            `<Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml2(rel.target)}" TargetMode="External"/>`
          );
        }
        return lines.join("");
      }
    };
  }
});

// src/io/XmlAutoFilterSaver.ts
function escapeXml3(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
exports.AutoFilterXmlSaver = void 0;
var init_XmlAutoFilterSaver = __esm({
  "src/io/XmlAutoFilterSaver.ts"() {
    exports.AutoFilterXmlSaver = class {
      formatAutoFilterXml(autoFilter) {
        if (!autoFilter.range) return "";
        const lines = [];
        lines.push(`<autoFilter ref="${escapeXml3(autoFilter.range)}">`);
        const sortedCols = Array.from(autoFilter.filterColumns.entries()).sort(
          (a, b) => a[0] - b[0]
        );
        for (const [, fc] of sortedCols) {
          const hasContent = fc.filters.length > 0 || fc.customFilters.length > 0 || fc.colorFilter !== null || fc.dynamicFilter !== null || fc.top10Filter !== null;
          if (!hasContent && fc.filterButton) continue;
          const fcAttrs = [`colId="${fc.colId}"`];
          if (!fc.filterButton) {
            fcAttrs.push('hiddenButton="1"');
          }
          lines.push(`<filterColumn ${fcAttrs.join(" ")}>`);
          if (fc.filters.length > 0) {
            lines.push("<filters>");
            for (const val of fc.filters) {
              lines.push(`<filter val="${escapeXml3(val)}"/>`);
            }
            lines.push("</filters>");
          }
          if (fc.customFilters.length > 0) {
            const andAttr = fc.customFilters.length > 1 ? ' and="1"' : "";
            lines.push(`<customFilters${andAttr}>`);
            for (const cf of fc.customFilters) {
              lines.push(`<customFilter operator="${escapeXml3(cf.operator)}" val="${escapeXml3(cf.value)}"/>`);
            }
            lines.push("</customFilters>");
          }
          if (fc.colorFilter) {
            const cellColorAttr = fc.colorFilter.cellColor ? "" : ' cellColor="0"';
            lines.push(`<colorFilter dxfId="${escapeXml3(fc.colorFilter.color)}"${cellColorAttr}/>`);
          }
          if (fc.dynamicFilter) {
            const valAttr = fc.dynamicFilter.value ? ` val="${escapeXml3(fc.dynamicFilter.value)}"` : "";
            lines.push(`<dynamicFilter type="${escapeXml3(fc.dynamicFilter.type)}"${valAttr}/>`);
          }
          if (fc.top10Filter) {
            const attrs = [];
            if (!fc.top10Filter.top) attrs.push('top="0"');
            if (fc.top10Filter.percent) attrs.push('percent="1"');
            attrs.push(`val="${fc.top10Filter.val}"`);
            lines.push(`<top10 ${attrs.join(" ")}/>`);
          }
          lines.push("</filterColumn>");
        }
        if (autoFilter.sortState) {
          const ss = autoFilter.sortState;
          const sortRef = this._calculateSortRefs(autoFilter.range, ss.columnIndex);
          lines.push(`<sortState ref="${escapeXml3(ss.ref)}">`);
          const descAttr = ss.ascending ? "" : ' descending="1"';
          lines.push(`<sortCondition ref="${escapeXml3(sortRef)}"${descAttr}/>`);
          lines.push("</sortState>");
        }
        lines.push("</autoFilter>");
        return lines.join("");
      }
      _calculateSortRefs(filterRange, columnIndex) {
        const parts = filterRange.split(":");
        if (parts.length !== 2) return filterRange;
        const startCol = this._colFromRef(parts[0]);
        const targetCol = startCol + columnIndex;
        const colLetter = this._colToLetter(targetCol);
        const startRow = this._rowFromRef(parts[0]);
        const endRow = this._rowFromRef(parts[1]);
        return `${colLetter}${startRow}:${colLetter}${endRow}`;
      }
      _colFromRef(ref) {
        const match = ref.match(/^([A-Z]+)/i);
        if (!match) return 0;
        let col = 0;
        for (const ch of match[1].toUpperCase()) {
          col = col * 26 + (ch.charCodeAt(0) - 64);
        }
        return col - 1;
      }
      _rowFromRef(ref) {
        const match = ref.match(/(\d+)$/);
        return match ? parseInt(match[1], 10) : 1;
      }
      _colToLetter(col) {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + c % 26) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      }
    };
  }
});

// src/io/CommentXml.ts
function escapeXml4(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function cellReferenceSortKey(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return [0, 0];
  const col = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  let colNum = 0;
  for (const ch of col) {
    colNum = colNum * 26 + (ch.charCodeAt(0) - 64);
  }
  return [row, colNum];
}
function colFromRef(ref) {
  const match = ref.match(/^([A-Z]+)/i);
  if (!match) return 0;
  let col = 0;
  for (const ch of match[1].toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return col - 1;
}
function rowFromRef(ref) {
  const match = ref.match(/(\d+)$/);
  return match ? parseInt(match[1], 10) - 1 : 0;
}
exports.CommentXmlWriter = void 0; exports.CommentXmlReader = void 0;
var init_CommentXml = __esm({
  "src/io/CommentXml.ts"() {
    exports.CommentXmlWriter = class {
      /**
       * Check if worksheet has any comments
       */
      static worksheetHasComments(worksheet) {
        const cellsObj = worksheet._cells;
        const cells = cellsObj._cells ?? cellsObj;
        for (const cell of cells.values()) {
          if (cell.comment) return true;
        }
        return false;
      }
      /**
       * Write comments XML to the ZIP archive
       */
      writeCommentsXml(zip, worksheet, sheetNum) {
        const cellsObj = worksheet._cells;
        const cells = cellsObj._cells ?? cellsObj;
        const commentCells = [];
        for (const [ref, cell] of cells.entries()) {
          if (cell.comment) {
            commentCells.push({
              ref,
              text: cell.comment.text ?? "",
              author: cell.comment.author ?? "None"
            });
          }
        }
        if (commentCells.length === 0) return;
        commentCells.sort((a, b) => {
          const ka = cellReferenceSortKey(a.ref);
          const kb = cellReferenceSortKey(b.ref);
          return ka[0] - kb[0] || ka[1] - kb[1];
        });
        const authors = [...new Set(commentCells.map((c) => c.author))];
        const lines = [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          "<authors>"
        ];
        for (const author of authors) {
          lines.push(`<author>${escapeXml4(author)}</author>`);
        }
        lines.push("</authors>");
        lines.push("<commentList>");
        for (const cc of commentCells) {
          const authorIdx = authors.indexOf(cc.author);
          lines.push(`<comment ref="${cc.ref}" authorId="${authorIdx}">`);
          lines.push("<text>");
          lines.push(`<r><t>${escapeXml4(cc.text)}</t></r>`);
          lines.push("</text>");
          lines.push("</comment>");
        }
        lines.push("</commentList>");
        lines.push("</comments>");
        zip.file(`xl/comments${sheetNum}.xml`, lines.join(""));
      }
      /**
       * Write VML drawing XML for comment shapes
       */
      writeVmlDrawingXml(zip, worksheet, sheetNum) {
        const cellsObj = worksheet._cells;
        const cells = cellsObj._cells ?? cellsObj;
        const commentCells = [];
        for (const [ref, cell] of cells.entries()) {
          if (cell.comment) {
            const width = cell.comment.width ?? 108;
            const height = cell.comment.height ?? 59;
            commentCells.push({ ref, width, height });
          }
        }
        if (commentCells.length === 0) return;
        commentCells.sort((a, b) => {
          const ka = cellReferenceSortKey(a.ref);
          const kb = cellReferenceSortKey(b.ref);
          return ka[0] - kb[0] || ka[1] - kb[1];
        });
        const lines = [
          '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
          '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
          '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
          '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
          "</v:shapetype>"
        ];
        for (let i = 0; i < commentCells.length; i++) {
          const cc = commentCells[i];
          const row = rowFromRef(cc.ref);
          const col = colFromRef(cc.ref);
          const anchor = this._calculateAnchor(row, col, cc.width, cc.height);
          lines.push(`<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" `);
          lines.push(`style="position:absolute;margin-left:0;margin-top:0;width:${cc.width}pt;height:${cc.height}pt;z-index:${i + 1};visibility:hidden" `);
          lines.push('fillcolor="#ffffe1" o:insetmode="auto">');
          lines.push('<v:fill color="#ffffe1"/>');
          lines.push('<v:shadow on="t" color="black" obscured="t"/>');
          lines.push('<v:path o:connecttype="none"/>');
          lines.push('<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>');
          lines.push('<x:ClientData ObjectType="Note">');
          lines.push("<x:MoveWithCells/>");
          lines.push("<x:SizeWithCells/>");
          lines.push(`<x:Anchor>${anchor}</x:Anchor>`);
          lines.push("<x:AutoFill>False</x:AutoFill>");
          lines.push(`<x:Row>${row}</x:Row>`);
          lines.push(`<x:Column>${col}</x:Column>`);
          lines.push("</x:ClientData>");
          lines.push("</v:shape>");
        }
        lines.push("</xml>");
        zip.file(`xl/drawings/vmlDrawing${sheetNum}.vml`, lines.join(""));
      }
      _calculateAnchor(row, col, width, height) {
        const colWidth = 64;
        const rowHeight = 20;
        const leftCol = col + 1;
        const leftOffset = 15;
        const topRow = row;
        const topOffset = 10;
        const widthPx = Math.round(width * 1.33);
        const heightPx = Math.round(height * 1.33);
        const rightCol = leftCol + Math.max(1, Math.floor(widthPx / colWidth));
        const rightOffset = widthPx % colWidth;
        const bottomRow = topRow + Math.max(1, Math.floor(heightPx / rowHeight));
        const bottomOffset = heightPx % rowHeight;
        return `${leftCol}, ${leftOffset}, ${topRow}, ${topOffset}, ${rightCol}, ${rightOffset}, ${bottomRow}, ${bottomOffset}`;
      }
    };
    exports.CommentXmlReader = class {
      loadComments(zip, worksheet, sheetNum) {
        const commentsFile = zip.file(`xl/comments${sheetNum}.xml`);
        if (!commentsFile) return;
      }
      /**
       * Parse comments XML and apply to worksheet
       */
      parseAndApplyComments(commentsRoot, worksheet) {
        if (!commentsRoot?.comments) return;
        const comments = commentsRoot.comments;
        const authors = [];
        const authorElems = this._ensureArray(comments?.authors?.author);
        for (const a of authorElems) {
          authors.push(typeof a === "string" ? a : String(a));
        }
        const commentElems = this._ensureArray(comments?.commentList?.comment);
        for (const ce of commentElems) {
          const ref = ce?.["@_ref"];
          if (!ref) continue;
          const authorId = parseInt(ce?.["@_authorId"] ?? "0", 10);
          const author = authors[authorId] ?? "None";
          let text = "";
          const textElem = ce?.text;
          if (textElem) {
            const runs = this._ensureArray(textElem.r);
            if (runs.length > 0) {
              for (const r of runs) {
                const tVal = r?.t;
                if (tVal != null) {
                  const tArr = Array.isArray(tVal) ? tVal : [tVal];
                  for (const t of tArr) {
                    if (t != null) {
                      text += typeof t === "object" ? t["#text"] ?? "" : String(t);
                    }
                  }
                }
              }
            } else if (textElem.t != null) {
              const tArr = Array.isArray(textElem.t) ? textElem.t : [textElem.t];
              for (const t of tArr) {
                if (t != null) {
                  text += typeof t === "object" ? t["#text"] ?? "" : String(t);
                }
              }
            }
          }
          if (text) {
            const cellsObj2 = worksheet._cells;
            const cells2 = cellsObj2._cells ?? cellsObj2;
            const cell = cells2.get(ref);
            if (cell) {
              cell.setComment(text, author);
            } else {
              const { Cell: Cell2 } = (init_Cell(), __toCommonJS(Cell_exports));
              const newCell = new Cell2();
              newCell.setComment(text, author);
              cells2.set(ref, newCell);
            }
          }
        }
      }
      /**
       * Parse VML drawing to get comment dimensions
       */
      parseAndApplyVmlDrawing(vmlContent, worksheet) {
        const shapeRegex = /<v:shape[^>]*>[\s\S]*?<\/v:shape>/gi;
        let match;
        while ((match = shapeRegex.exec(vmlContent)) !== null) {
          const shape = match[0];
          const rowMatch = shape.match(/<x:Row>(\d+)<\/x:Row>/i);
          const colMatch = shape.match(/<x:Column>(\d+)<\/x:Column>/i);
          if (!rowMatch || !colMatch) continue;
          const row = parseInt(rowMatch[1], 10);
          const col = parseInt(colMatch[1], 10);
          const widthMatch = shape.match(/width:\s*([\d.]+)pt/i);
          const heightMatch = shape.match(/height:\s*([\d.]+)pt/i);
          if (widthMatch && heightMatch) {
            const width = parseFloat(widthMatch[1]);
            const height = parseFloat(heightMatch[1]);
            const colLetter = this._colToLetter(col);
            const ref = `${colLetter}${row + 1}`;
            const cellsObj3 = worksheet._cells;
            const cells3 = cellsObj3._cells ?? cellsObj3;
            const cell = cells3.get(ref);
            if (cell?.comment) {
              cell.comment.width = width;
              cell.comment.height = height;
            }
          }
        }
      }
      _ensureArray(val) {
        if (val == null) return [];
        return Array.isArray(val) ? val : [val];
      }
      _colToLetter(col) {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + c % 26) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      }
    };
  }
});
function escapeXml5(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function sortedKeys(map) {
  return Array.from(map.keys()).sort((a, b) => a - b);
}
exports.XmlSaver = void 0; var DEFAULT_THEME_XML;
var init_XmlSaver = __esm({
  "src/io/XmlSaver.ts"() {
    init_Cells();
    init_SharedStrings();
    init_CellValueHandler();
    init_Style();
    init_XmlDataValidationSaver();
    init_XmlConditionalFormatSaver();
    init_XmlHyperlinkHandler();
    init_XmlAutoFilterSaver();
    init_CommentXml();
    exports.XmlSaver = class {
      constructor(workbook) {
        // Style registries: index → data
        this._fontStyles = /* @__PURE__ */ new Map();
        this._fillStyles = /* @__PURE__ */ new Map();
        this._borderStyles = /* @__PURE__ */ new Map();
        this._alignmentStyles = /* @__PURE__ */ new Map();
        this._protectionStyles = /* @__PURE__ */ new Map();
        this._numFormats = /* @__PURE__ */ new Map();
        // Composite cell style key → xf index
        this._cellStyles = /* @__PURE__ */ new Map();
        // Phase 4: DXF entries collected from conditional formatting
        this._dxfEntries = [];
        this._workbook = workbook;
        this._sharedStringTable = new exports.SharedStringTable();
      }
      // ── Public API ──────────────────────────────────────────────────────────
      /**
       * Saves the workbook as an XLSX buffer (in-memory ZIP).
       *
       * @returns A `Buffer` containing the XLSX data.
       */
      async saveToBuffer() {
        const zip = new JSZip__default.default();
        this._writeContentTypes(zip);
        this._writeRootRelationships(zip);
        this._writeWorkbookRelationships(zip);
        this._writeWorkbookXml(zip);
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          this._writeWorksheetXml(zip, this._workbook.worksheets[i], i + 1);
        }
        const commentWriter = new exports.CommentXmlWriter();
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          const ws = this._workbook.worksheets[i];
          if (exports.CommentXmlWriter.worksheetHasComments(ws)) {
            commentWriter.writeCommentsXml(zip, ws, i + 1);
            commentWriter.writeVmlDrawingXml(zip, ws, i + 1);
          }
        }
        this._writeStylesXml(zip);
        this._writeSharedStringsXml(zip);
        this._writeThemeXml(zip);
        this._writeCorePropertiesXml(zip);
        this._writeAppPropertiesXml(zip);
        const buf = await zip.generateAsync({
          type: "nodebuffer",
          compression: "DEFLATE",
          compressionOptions: { level: 6 }
        });
        return buf;
      }
      /**
       * Saves the workbook to a file path.
       *
       * @param filePath - Destination file path.
       */
      async saveToFile(filePath) {
        const fs = await import('fs');
        const path = await import('path');
        const dir = path.dirname(filePath);
        if (dir && dir !== ".") {
          fs.mkdirSync(dir, { recursive: true });
        }
        const buffer = await this.saveToBuffer();
        fs.writeFileSync(filePath, buffer);
      }
      // ══════════════════════════════════════════════════════════════════════════
      // [Content_Types].xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeContentTypes(zip) {
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        );
        lines.push(
          '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        );
        lines.push('  <Default Extension="xml" ContentType="application/xml"/>');
        lines.push(
          '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        );
        lines.push(
          '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        );
        lines.push(
          '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        );
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          lines.push(
            `  <Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
          );
        }
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          const ws = this._workbook.worksheets[i];
          if (exports.CommentXmlWriter.worksheetHasComments(ws)) {
            lines.push(
              `  <Override PartName="/xl/comments${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`
            );
            lines.push(
              `  <Override PartName="/xl/drawings/vmlDrawing${i + 1}.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`
            );
          }
        }
        lines.push(
          '  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        );
        lines.push(
          '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        );
        lines.push(
          '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        );
        lines.push("</Types>");
        zip.file("[Content_Types].xml", lines.join("\n"));
      }
      // ══════════════════════════════════════════════════════════════════════════
      // _rels/.rels
      // ══════════════════════════════════════════════════════════════════════════
      _writeRootRelationships(zip) {
        const content = [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
          '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
          '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
          "</Relationships>"
        ].join("\n");
        zip.file("_rels/.rels", content);
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/_rels/workbook.xml.rels
      // ══════════════════════════════════════════════════════════════════════════
      _writeWorkbookRelationships(zip) {
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        );
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          lines.push(
            `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
          );
        }
        lines.push(
          '  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        );
        lines.push(
          '  <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        );
        lines.push(
          '  <Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        );
        lines.push("</Relationships>");
        zip.file("xl/_rels/workbook.xml.rels", lines.join("\n"));
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/workbook.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeWorkbookXml(zip) {
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        );
        lines.push(
          '  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>'
        );
        lines.push('  <workbookPr defaultThemeVersion="124226"/>');
        const activeTab = this._workbook.properties.view.activeTab;
        lines.push("  <bookViews>");
        lines.push(
          `    <workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192" activeTab="${activeTab}"/>`
        );
        lines.push("  </bookViews>");
        lines.push("  <sheets>");
        for (let i = 0; i < this._workbook.worksheets.length; i++) {
          const ws = this._workbook.worksheets[i];
          let stateAttr = "";
          if (ws.visible === false) {
            stateAttr = ' state="hidden"';
          } else if (ws.visible === "veryHidden") {
            stateAttr = ' state="veryHidden"';
          }
          lines.push(
            `    <sheet name="${escapeXml5(ws.name)}" sheetId="${i + 1}"${stateAttr} r:id="rId${i + 1}"/>`
          );
        }
        lines.push("  </sheets>");
        const definedNames = this._workbook.properties.definedNames;
        if (definedNames.count > 0) {
          lines.push("  <definedNames>");
          for (const dn of definedNames) {
            const dnAttrs = [`name="${escapeXml5(dn.name)}"`];
            if (dn.localSheetId !== null) {
              dnAttrs.push(`localSheetId="${dn.localSheetId}"`);
            }
            if (dn.hidden) {
              dnAttrs.push('hidden="1"');
            }
            if (dn.comment) {
              dnAttrs.push(`comment="${escapeXml5(dn.comment)}"`);
            }
            lines.push(`    <definedName ${dnAttrs.join(" ")}>${escapeXml5(dn.refersTo)}</definedName>`);
          }
          lines.push("  </definedNames>");
        }
        const calcMode = this._workbook.properties.calcMode || "auto";
        lines.push(`  <calcPr calcId="145621" calcMode="${calcMode}"/>`);
        lines.push("</workbook>");
        zip.file("xl/workbook.xml", lines.join("\n"));
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/worksheets/sheetN.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeWorksheetXml(zip, worksheet, sheetNum) {
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        );
        const dimRef = this._computeDimensionRef(worksheet);
        if (dimRef) {
          lines.push(`  <dimension ref="${escapeXml5(dimRef)}"/>`);
        }
        const activeTab = this._workbook.properties.view.activeTab;
        const isActive = sheetNum - 1 === activeTab;
        lines.push("  <sheetViews>");
        let sheetViewAttrs = `workbookViewId="0"`;
        if (isActive) sheetViewAttrs += ' tabSelected="1"';
        const fp = worksheet.freezePane;
        if (fp) {
          lines.push(`    <sheetView ${sheetViewAttrs}>`);
          const topLeftCell = exports.Cells.coordinateToString(
            fp.freezedRows + 1,
            fp.freezedColumns + 1
          );
          lines.push(
            `      <pane xSplit="${fp.column}" ySplit="${fp.row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`
          );
          lines.push("    </sheetView>");
        } else {
          lines.push(`    <sheetView ${sheetViewAttrs}/>`);
        }
        lines.push("  </sheetViews>");
        lines.push('  <sheetFormatPr defaultRowHeight="15"/>');
        const colsXml = this._formatColsXml(worksheet);
        if (colsXml) {
          lines.push(colsXml);
        }
        lines.push("  <sheetData>");
        const cellsMap = worksheet.cells._cells;
        const sortedRefs = Array.from(cellsMap.keys()).sort((a, b) => {
          const [ra, ca] = exports.Cells.coordinateFromString(a);
          const [rb, cb] = exports.Cells.coordinateFromString(b);
          return ra !== rb ? ra - rb : ca - cb;
        });
        const rows = /* @__PURE__ */ new Map();
        for (const ref of sortedRefs) {
          const [row] = exports.Cells.coordinateFromString(ref);
          if (!rows.has(row)) rows.set(row, []);
          rows.get(row).push([ref, cellsMap.get(ref)]);
        }
        const rowHeights = worksheet._rowHeights ?? {};
        for (const rowNum of Object.keys(rowHeights).map(Number)) {
          if (!rows.has(rowNum)) rows.set(rowNum, []);
        }
        const hiddenRows = worksheet._hiddenRows ?? /* @__PURE__ */ new Set();
        for (const rowNum of hiddenRows) {
          if (!rows.has(rowNum)) rows.set(rowNum, []);
        }
        const sortedRowNums = Array.from(rows.keys()).sort((a, b) => a - b);
        for (const rowNum of sortedRowNums) {
          const rowAttrs = [`r="${rowNum}"`];
          const rh = rowHeights[rowNum];
          if (rh !== void 0) {
            rowAttrs.push(`ht="${rh}"`);
            rowAttrs.push('customHeight="1"');
          }
          if (hiddenRows.has(rowNum)) {
            rowAttrs.push('hidden="1"');
          }
          const cellEntries = rows.get(rowNum);
          if (cellEntries.length === 0) {
            lines.push(`    <row ${rowAttrs.join(" ")}/>`);
          } else {
            lines.push(`    <row ${rowAttrs.join(" ")}>`);
            for (const [ref, cell] of cellEntries) {
              lines.push(this._formatCellXml(ref, cell));
            }
            lines.push("    </row>");
          }
        }
        lines.push("  </sheetData>");
        const prot = worksheet._protection;
        if (prot && prot.sheet) {
          const protAttrs = ['sheet="1"'];
          if (prot.password) protAttrs.push(`password="${escapeXml5(prot.password)}"`);
          if (!prot.formatCells) protAttrs.push('formatCells="0"');
          if (!prot.formatColumns) protAttrs.push('formatColumns="0"');
          if (!prot.formatRows) protAttrs.push('formatRows="0"');
          if (!prot.insertColumns) protAttrs.push('insertColumns="0"');
          if (!prot.insertRows) protAttrs.push('insertRows="0"');
          if (!prot.deleteColumns) protAttrs.push('deleteColumns="0"');
          if (!prot.deleteRows) protAttrs.push('deleteRows="0"');
          if (!prot.sort) protAttrs.push('sort="0"');
          if (!prot.autoFilter) protAttrs.push('autoFilter="0"');
          if (!prot.insertHyperlinks) protAttrs.push('insertHyperlinks="0"');
          if (!prot.pivotTables) protAttrs.push('pivotTables="0"');
          if (prot.selectLockedCells) protAttrs.push('selectLockedCells="1"');
          if (prot.selectUnlockedCells) protAttrs.push('selectUnlockedCells="1"');
          if (prot.objects) protAttrs.push('objects="1"');
          if (prot.scenarios) protAttrs.push('scenarios="1"');
          lines.push(`  <sheetProtection ${protAttrs.join(" ")}/>`);
        }
        if (worksheet.autoFilter.range !== null) {
          const autoFilterXml = new exports.AutoFilterXmlSaver().formatAutoFilterXml(worksheet.autoFilter);
          if (autoFilterXml) lines.push(autoFilterXml);
        }
        const mergedCells = worksheet._mergedCells ?? [];
        if (mergedCells.length > 0) {
          lines.push(`  <mergeCells count="${mergedCells.length}">`);
          for (const mc of mergedCells) {
            lines.push(`    <mergeCell ref="${escapeXml5(mc)}"/>`);
          }
          lines.push("  </mergeCells>");
        }
        if (worksheet.conditionalFormatting.count > 0) {
          const cfSaver = new exports.ConditionalFormatXmlSaver(escapeXml5);
          const startDxfId = this._dxfEntries.length;
          const cfResult = cfSaver.formatConditionalFormattingXml(
            worksheet.conditionalFormatting,
            startDxfId
          );
          if (cfResult.xml) lines.push(cfResult.xml);
          this._dxfEntries.push(...cfResult.dxfEntries);
        }
        if (worksheet.dataValidations.count > 0) {
          const dvXml = new exports.DataValidationXmlSaver(escapeXml5).formatDataValidationsXml(worksheet.dataValidations);
          if (dvXml) lines.push(dvXml);
        }
        const hlSaver = new exports.HyperlinkXmlSaver();
        const nextRelId = exports.CommentXmlWriter.worksheetHasComments(worksheet) ? 3 : 1;
        if (worksheet.hyperlinks.count > 0) {
          hlSaver.resetRelationshipCounter(nextRelId);
          const hlXml = hlSaver.formatHyperlinksXml(worksheet.hyperlinks);
          if (hlXml) lines.push(hlXml);
        }
        const pm = worksheet.pageMargins;
        lines.push(
          `  <pageMargins left="${pm.left}" right="${pm.right}" top="${pm.top}" bottom="${pm.bottom}" header="${pm.header}" footer="${pm.footer}"/>`
        );
        const ps = worksheet.pageSetup;
        const psAttrs = [];
        if (ps.orientation) psAttrs.push(`orientation="${ps.orientation}"`);
        if (ps.paperSize != null) psAttrs.push(`paperSize="${ps.paperSize}"`);
        if (ps.fitToPage) {
          psAttrs.push('fitToPage="1"');
        }
        if (ps.scale != null) psAttrs.push(`scale="${ps.scale}"`);
        if (ps.fitToWidth != null) psAttrs.push(`fitToWidth="${ps.fitToWidth}"`);
        if (ps.fitToHeight != null) psAttrs.push(`fitToHeight="${ps.fitToHeight}"`);
        if (psAttrs.length > 0) {
          lines.push(`  <pageSetup ${psAttrs.join(" ")}/>`);
        }
        const hBreaks = Array.from(worksheet.horizontalPageBreaks).sort((a, b) => a - b);
        if (hBreaks.length > 0) {
          lines.push(`  <rowBreaks count="${hBreaks.length}" manualBreakCount="${hBreaks.length}">`);
          for (const row of hBreaks) {
            lines.push(`    <brk id="${row}" max="16383" man="1"/>`);
          }
          lines.push("  </rowBreaks>");
        }
        const vBreaks = Array.from(worksheet.verticalPageBreaks).sort((a, b) => a - b);
        if (vBreaks.length > 0) {
          lines.push(`  <colBreaks count="${vBreaks.length}" manualBreakCount="${vBreaks.length}">`);
          for (const col of vBreaks) {
            lines.push(`    <brk id="${col}" max="1048575" man="1"/>`);
          }
          lines.push("  </colBreaks>");
        }
        if (exports.CommentXmlWriter.worksheetHasComments(worksheet)) {
          lines.push('  <legacyDrawing r:id="rId1"/>');
        }
        lines.push("</worksheet>");
        zip.file(`xl/worksheets/sheet${sheetNum}.xml`, lines.join("\n"));
        const sheetRelsLines = [];
        if (exports.CommentXmlWriter.worksheetHasComments(worksheet)) {
          sheetRelsLines.push(
            `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${sheetNum}.vml"/>`
          );
          sheetRelsLines.push(
            `  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${sheetNum}.xml"/>`
          );
        }
        if (worksheet.hyperlinks.count > 0) {
          const hlRels = hlSaver.getHyperlinkRelationships(worksheet.hyperlinks);
          for (const rel of hlRels) {
            sheetRelsLines.push(
              `  <Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml5(rel.target)}" TargetMode="External"/>`
            );
          }
        }
        if (sheetRelsLines.length > 0) {
          const relsContent = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            ...sheetRelsLines,
            "</Relationships>"
          ].join("\n");
          zip.file(`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`, relsContent);
        }
      }
      // ── Worksheet helpers ───────────────────────────────────────────────────
      /**
       * Computes the dimension reference string (e.g. "A1:C10") for a worksheet.
       */
      _computeDimensionRef(worksheet) {
        const cellsMap = worksheet.cells._cells;
        if (cellsMap.size === 0) return "A1";
        let minRow = Infinity;
        let maxRow = -Infinity;
        let minCol = Infinity;
        let maxCol = -Infinity;
        for (const ref of cellsMap.keys()) {
          const [r, c] = exports.Cells.coordinateFromString(ref);
          if (r < minRow) minRow = r;
          if (r > maxRow) maxRow = r;
          if (c < minCol) minCol = c;
          if (c > maxCol) maxCol = c;
        }
        const merged = worksheet._mergedCells ?? [];
        for (const mc of merged) {
          if (mc.includes(":")) {
            const [startRef, endRef] = mc.split(":");
            const [sr, sc] = exports.Cells.coordinateFromString(startRef);
            const [er, ec] = exports.Cells.coordinateFromString(endRef);
            if (sr < minRow) minRow = sr;
            if (er > maxRow) maxRow = er;
            if (sc < minCol) minCol = sc;
            if (ec > maxCol) maxCol = ec;
          }
        }
        if (!isFinite(minRow)) return "A1";
        const startCell = exports.Cells.coordinateToString(minRow, minCol);
        const endCell = exports.Cells.coordinateToString(maxRow, maxCol);
        return startCell === endCell ? startCell : `${startCell}:${endCell}`;
      }
      /**
       * Formats <cols> XML for column widths and hidden columns.
       */
      _formatColsXml(worksheet) {
        const colWidths = worksheet._columnWidths ?? {};
        const hiddenCols = worksheet._hiddenColumns ?? /* @__PURE__ */ new Set();
        const allColIndices = /* @__PURE__ */ new Set();
        for (const k of Object.keys(colWidths).map(Number)) allColIndices.add(k);
        for (const k of hiddenCols) allColIndices.add(k);
        if (allColIndices.size === 0) return "";
        const sorted = Array.from(allColIndices).sort((a, b) => a - b);
        const lines = ["  <cols>"];
        for (const colIdx of sorted) {
          const attrs = [`min="${colIdx}"`, `max="${colIdx}"`];
          const w = colWidths[colIdx];
          if (w !== void 0) {
            attrs.push(`width="${w}"`);
            attrs.push('customWidth="1"');
          }
          if (hiddenCols.has(colIdx)) {
            attrs.push('hidden="1"');
          }
          lines.push(`    <col ${attrs.join(" ")}/>`);
        }
        lines.push("  </cols>");
        return lines.join("\n");
      }
      /**
       * Formats a single cell as XML.
       */
      _formatCellXml(ref, cell) {
        const styleIdx = this._getOrCreateCellStyle(cell);
        const [valueStr, cellType] = exports.CellValueHandler.formatValueForXml(cell.value);
        let actualValue = valueStr;
        let actualType = cellType;
        if (cellType === exports.CELL_TYPE_STRING && valueStr !== null) {
          const ssIdx = this._sharedStringTable.addString(valueStr);
          actualValue = String(ssIdx);
          actualType = exports.CELL_TYPE_STRING;
        }
        const attrs = [`r="${ref}"`];
        if (styleIdx > 0) attrs.push(`s="${styleIdx}"`);
        if (actualType) attrs.push(`t="${actualType}"`);
        let xml = `      <c ${attrs.join(" ")}>`;
        const formula = cell.formula;
        if (formula) {
          xml += `<f>${escapeXml5(formula)}</f>`;
        }
        if (actualValue !== null) {
          xml += `<v>${escapeXml5(actualValue)}</v>`;
        }
        xml += "</c>";
        return xml;
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/styles.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeStylesXml(zip) {
        this._registerDefaultStyles();
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        );
        const customFmts = /* @__PURE__ */ new Map();
        for (const [id, code] of this._numFormats) {
          if (id >= 164) customFmts.set(id, code);
        }
        lines.push(`  <numFmts count="${customFmts.size}">`);
        for (const [id, code] of Array.from(customFmts.entries()).sort(
          (a, b) => a[0] - b[0]
        )) {
          lines.push(
            `    <numFmt numFmtId="${id}" formatCode="${escapeXml5(code)}"/>`
          );
        }
        lines.push("  </numFmts>");
        const fontCount = this._fontStyles.size;
        lines.push(`  <fonts count="${fontCount}">`);
        for (const idx of sortedKeys(this._fontStyles)) {
          lines.push(this._formatFontXml(this._fontStyles.get(idx)));
        }
        lines.push("  </fonts>");
        const fillCount = this._fillStyles.size;
        lines.push(`  <fills count="${fillCount}">`);
        for (const idx of sortedKeys(this._fillStyles)) {
          lines.push(this._formatFillXml(this._fillStyles.get(idx)));
        }
        lines.push("  </fills>");
        const borderCount = this._borderStyles.size;
        lines.push(`  <borders count="${borderCount}">`);
        for (const idx of sortedKeys(this._borderStyles)) {
          lines.push(this._formatBorderXml(this._borderStyles.get(idx)));
        }
        lines.push("  </borders>");
        lines.push('  <cellStyleXfs count="1">');
        lines.push(
          '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        );
        lines.push("  </cellStyleXfs>");
        const xfCount = this._cellStyles.size + 1;
        lines.push(`  <cellXfs count="${xfCount}">`);
        lines.push(
          '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        );
        const xfEntries = Array.from(this._cellStyles.entries()).sort(
          (a, b) => a[1] - b[1]
        );
        for (const [keyStr, _xfIdx] of xfEntries) {
          const key = JSON.parse(keyStr);
          const [fontIdx, fillIdx, borderIdx, numFmtIdx, alignIdx, protIdx] = key;
          const applyNumFmt = numFmtIdx !== 0 ? ' applyNumberFormat="1"' : "";
          const applyProt = protIdx !== 0 ? ' applyProtection="1"' : "";
          lines.push(
            `    <xf numFmtId="${numFmtIdx}" fontId="${fontIdx}" fillId="${fillIdx}" borderId="${borderIdx}" xfId="0"${applyNumFmt}${applyProt}>`
          );
          if (alignIdx > 0) {
            const ad = this._alignmentStyles.get(alignIdx);
            if (ad) lines.push(this._formatAlignmentXml(ad));
          }
          if (protIdx > 0) {
            const pd = this._protectionStyles.get(protIdx);
            if (pd) lines.push(this._formatProtectionXml(pd));
          }
          lines.push("    </xf>");
        }
        lines.push("  </cellXfs>");
        lines.push('  <cellStyles count="1">');
        lines.push('    <cellStyle name="Normal" xfId="0" builtinId="0"/>');
        lines.push("  </cellStyles>");
        if (this._dxfEntries.length > 0) {
          lines.push(`  <dxfs count="${this._dxfEntries.length}">`);
          for (const dxfXml of this._dxfEntries) {
            lines.push(dxfXml);
          }
          lines.push("  </dxfs>");
        } else {
          lines.push('  <dxfs count="0"/>');
        }
        lines.push("</styleSheet>");
        zip.file("xl/styles.xml", lines.join("\n"));
      }
      // ── Style XML formatters ────────────────────────────────────────────────
      _formatFontXml(fd) {
        const parts = ["    <font>"];
        if (fd.bold) parts.push("      <b/>");
        if (fd.italic) parts.push("      <i/>");
        if (fd.underline) parts.push("      <u/>");
        if (fd.strikethrough) parts.push("      <strike/>");
        parts.push(`      <sz val="${fd.size}"/>`);
        parts.push(`      <color rgb="${fd.color}"/>`);
        parts.push(`      <name val="${escapeXml5(fd.name)}"/>`);
        parts.push("    </font>");
        return parts.join("\n");
      }
      _formatFillXml(fd) {
        const pt = fd.patternType;
        if (pt === "none" || pt === "gray125") {
          return [
            "    <fill>",
            `      <patternFill patternType="${pt}"/>`,
            "    </fill>"
          ].join("\n");
        }
        return [
          "    <fill>",
          `      <patternFill patternType="${pt}">`,
          `        <fgColor rgb="${fd.fgColor}"/>`,
          `        <bgColor rgb="${fd.bgColor}"/>`,
          "      </patternFill>",
          "    </fill>"
        ].join("\n");
      }
      _formatBorderXml(bd) {
        const parts = ["    <border>"];
        for (const side of ["left", "right", "top", "bottom"]) {
          const sd = bd[side];
          if (sd.style !== "none") {
            parts.push(`      <${side} style="${sd.style}">`);
            parts.push(`        <color rgb="${sd.color}"/>`);
            parts.push(`      </${side}>`);
          } else {
            parts.push(`      <${side}/>`);
          }
        }
        parts.push("      <diagonal/>");
        parts.push("    </border>");
        return parts.join("\n");
      }
      _formatAlignmentXml(ad) {
        const attrs = [];
        if (ad.horizontal !== "general") attrs.push(`horizontal="${ad.horizontal}"`);
        if (ad.vertical !== "bottom") attrs.push(`vertical="${ad.vertical}"`);
        if (ad.wrapText) attrs.push('wrapText="1"');
        if (ad.indent !== 0) attrs.push(`indent="${ad.indent}"`);
        if (ad.textRotation !== 0) attrs.push(`textRotation="${ad.textRotation}"`);
        if (ad.shrinkToFit) attrs.push('shrinkToFit="1"');
        if (ad.readingOrder !== 0) attrs.push(`readingOrder="${ad.readingOrder}"`);
        if (ad.relativeIndent !== 0)
          attrs.push(`relativeIndent="${ad.relativeIndent}"`);
        return attrs.length > 0 ? `      <alignment ${attrs.join(" ")}/>` : "      <alignment/>";
      }
      _formatProtectionXml(pd) {
        const attrs = [];
        if (!pd.locked) attrs.push('locked="0"');
        if (pd.hidden) attrs.push('hidden="1"');
        if (attrs.length === 0) return "";
        return `      <protection ${attrs.join(" ")}/>`;
      }
      // ══════════════════════════════════════════════════════════════════════════
      // Style registration
      // ══════════════════════════════════════════════════════════════════════════
      _registerDefaultStyles() {
        if (!this._fontStyles.has(0)) {
          this._fontStyles.set(0, {
            name: "Calibri",
            size: 11,
            color: "FF000000",
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false
          });
        }
        if (!this._fillStyles.has(0)) {
          this._fillStyles.set(0, {
            patternType: "none",
            fgColor: "FFFFFFFF",
            bgColor: "FFFFFFFF"
          });
        }
        if (!this._fillStyles.has(1)) {
          this._fillStyles.set(1, {
            patternType: "gray125",
            fgColor: "FFFFFFFF",
            bgColor: "FFFFFFFF"
          });
        }
        if (!this._borderStyles.has(0)) {
          this._borderStyles.set(0, {
            top: { style: "none", color: "FF000000" },
            bottom: { style: "none", color: "FF000000" },
            left: { style: "none", color: "FF000000" },
            right: { style: "none", color: "FF000000" }
          });
        }
        if (!this._alignmentStyles.has(0)) {
          this._alignmentStyles.set(0, {
            horizontal: "general",
            vertical: "bottom",
            wrapText: false,
            indent: 0,
            textRotation: 0,
            shrinkToFit: false,
            readingOrder: 0,
            relativeIndent: 0
          });
        }
        if (!this._protectionStyles.has(0)) {
          this._protectionStyles.set(0, { locked: true, hidden: false });
        }
      }
      _getOrCreateFontStyle(font) {
        for (const [idx, fd] of this._fontStyles) {
          if (fd.name === font.name && fd.size === font.size && fd.color === font.color && fd.bold === font.bold && fd.italic === font.italic && fd.underline === font.underline && fd.strikethrough === font.strikethrough) {
            return idx;
          }
        }
        const newIdx = this._fontStyles.size;
        this._fontStyles.set(newIdx, {
          name: font.name,
          size: font.size,
          color: font.color,
          bold: font.bold,
          italic: font.italic,
          underline: font.underline,
          strikethrough: font.strikethrough
        });
        return newIdx;
      }
      _getOrCreateFillStyle(fill) {
        for (const [idx, fd] of this._fillStyles) {
          if (fd.patternType === fill.patternType && fd.fgColor === fill.foregroundColor && fd.bgColor === fill.backgroundColor) {
            return idx;
          }
        }
        const newIdx = this._fillStyles.size;
        this._fillStyles.set(newIdx, {
          patternType: fill.patternType,
          fgColor: fill.foregroundColor,
          bgColor: fill.backgroundColor
        });
        return newIdx;
      }
      _getOrCreateBorderStyle(borders) {
        for (const [idx, bd] of this._borderStyles) {
          if (bd.top.style === borders.top.lineStyle && bd.top.color === borders.top.color && bd.bottom.style === borders.bottom.lineStyle && bd.bottom.color === borders.bottom.color && bd.left.style === borders.left.lineStyle && bd.left.color === borders.left.color && bd.right.style === borders.right.lineStyle && bd.right.color === borders.right.color) {
            return idx;
          }
        }
        const newIdx = this._borderStyles.size;
        this._borderStyles.set(newIdx, {
          top: { style: borders.top.lineStyle, color: borders.top.color },
          bottom: { style: borders.bottom.lineStyle, color: borders.bottom.color },
          left: { style: borders.left.lineStyle, color: borders.left.color },
          right: { style: borders.right.lineStyle, color: borders.right.color }
        });
        return newIdx;
      }
      _getOrCreateAlignmentStyle(alignment) {
        for (const [idx, ad] of this._alignmentStyles) {
          if (ad.horizontal === alignment.horizontal && ad.vertical === alignment.vertical && ad.wrapText === alignment.wrapText && ad.indent === alignment.indent && ad.textRotation === alignment.textRotation && ad.shrinkToFit === alignment.shrinkToFit && ad.readingOrder === alignment.readingOrder && ad.relativeIndent === alignment.relativeIndent) {
            return idx;
          }
        }
        const newIdx = this._alignmentStyles.size;
        this._alignmentStyles.set(newIdx, {
          horizontal: alignment.horizontal,
          vertical: alignment.vertical,
          wrapText: alignment.wrapText,
          indent: alignment.indent,
          textRotation: alignment.textRotation,
          shrinkToFit: alignment.shrinkToFit,
          readingOrder: alignment.readingOrder,
          relativeIndent: alignment.relativeIndent
        });
        return newIdx;
      }
      _getOrCreateProtectionStyle(protection) {
        for (const [idx, pd] of this._protectionStyles) {
          if (pd.locked === protection.locked && pd.hidden === protection.hidden) {
            return idx;
          }
        }
        const newIdx = this._protectionStyles.size;
        this._protectionStyles.set(newIdx, {
          locked: protection.locked,
          hidden: protection.hidden
        });
        return newIdx;
      }
      _getOrCreateNumberFormatStyle(numberFormat) {
        const builtinId = exports.NumberFormat.lookupBuiltinFormat(numberFormat);
        if (builtinId !== null) return builtinId;
        if (numberFormat === "General") return 0;
        for (const [id, code] of this._numFormats) {
          if (code === numberFormat) return id;
        }
        let newId = 164;
        for (const id of this._numFormats.keys()) {
          if (id >= newId) newId = id + 1;
        }
        this._numFormats.set(newId, numberFormat);
        return newId;
      }
      /**
       * Gets or creates a composite cellXfs style index for a cell.
       */
      _getOrCreateCellStyle(cell) {
        this._registerDefaultStyles();
        const style = cell.style;
        const fontIdx = this._getOrCreateFontStyle(style.font);
        const fillIdx = this._getOrCreateFillStyle(style.fill);
        const borderIdx = this._getOrCreateBorderStyle(style.borders);
        const numFmtIdx = this._getOrCreateNumberFormatStyle(
          style.numberFormat
        );
        const alignIdx = this._getOrCreateAlignmentStyle(style.alignment);
        const protIdx = this._getOrCreateProtectionStyle(style.protection);
        if (fontIdx === 0 && fillIdx === 0 && borderIdx === 0 && numFmtIdx === 0 && alignIdx === 0 && protIdx === 0) {
          return 0;
        }
        const key = [
          fontIdx,
          fillIdx,
          borderIdx,
          numFmtIdx,
          alignIdx,
          protIdx
        ];
        const keyStr = JSON.stringify(key);
        if (!this._cellStyles.has(keyStr)) {
          const nextIdx = this._cellStyles.size > 0 ? Math.max(...Array.from(this._cellStyles.values())) + 1 : 1;
          this._cellStyles.set(keyStr, nextIdx);
        }
        return this._cellStyles.get(keyStr);
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/sharedStrings.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeSharedStringsXml(zip) {
        zip.file("xl/sharedStrings.xml", this._sharedStringTable.toXml());
      }
      // ══════════════════════════════════════════════════════════════════════════
      // xl/theme/theme1.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeThemeXml(zip) {
        zip.file("xl/theme/theme1.xml", DEFAULT_THEME_XML);
      }
      // ══════════════════════════════════════════════════════════════════════════
      // docProps/core.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeCorePropertiesXml(zip) {
        const dp = this._workbook.documentProperties;
        const now = (/* @__PURE__ */ new Date()).toISOString().replace(/\.\d{3}Z$/, "Z");
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        );
        if (dp.title) lines.push(`  <dc:title>${escapeXml5(dp.title)}</dc:title>`);
        if (dp.subject)
          lines.push(`  <dc:subject>${escapeXml5(dp.subject)}</dc:subject>`);
        if (dp.creator)
          lines.push(`  <dc:creator>${escapeXml5(dp.creator)}</dc:creator>`);
        if (dp.keywords)
          lines.push(`  <cp:keywords>${escapeXml5(dp.keywords)}</cp:keywords>`);
        if (dp.description)
          lines.push(
            `  <dc:description>${escapeXml5(dp.description)}</dc:description>`
          );
        if (dp.lastModifiedBy)
          lines.push(
            `  <cp:lastModifiedBy>${escapeXml5(dp.lastModifiedBy)}</cp:lastModifiedBy>`
          );
        if (dp.category)
          lines.push(`  <cp:category>${escapeXml5(dp.category)}</cp:category>`);
        lines.push(
          `  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`
        );
        lines.push(
          `  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
        );
        lines.push("</cp:coreProperties>");
        zip.file("docProps/core.xml", lines.join("\n"));
      }
      // ══════════════════════════════════════════════════════════════════════════
      // docProps/app.xml
      // ══════════════════════════════════════════════════════════════════════════
      _writeAppPropertiesXml(zip) {
        const wsCount = this._workbook.worksheets.length;
        const lines = [];
        lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        lines.push(
          '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        );
        lines.push("  <Application>Microsoft Excel</Application>");
        lines.push("  <DocSecurity>0</DocSecurity>");
        lines.push("  <ScaleCrop>false</ScaleCrop>");
        lines.push("  <LinksUpToDate>false</LinksUpToDate>");
        lines.push("  <SharedDoc>false</SharedDoc>");
        lines.push("  <HeadingPairs>");
        lines.push('    <vt:vector size="2" baseType="variant">');
        lines.push("      <vt:variant>");
        lines.push("        <vt:lpstr>Worksheets</vt:lpstr>");
        lines.push("      </vt:variant>");
        lines.push("      <vt:variant>");
        lines.push(`        <vt:i4>${wsCount}</vt:i4>`);
        lines.push("      </vt:variant>");
        lines.push("    </vt:vector>");
        lines.push("  </HeadingPairs>");
        lines.push("  <TitlesOfParts>");
        lines.push(`    <vt:vector size="${wsCount}" baseType="lpstr">`);
        for (const ws of this._workbook.worksheets) {
          lines.push(`      <vt:lpstr>${escapeXml5(ws.name)}</vt:lpstr>`);
        }
        lines.push("    </vt:vector>");
        lines.push("  </TitlesOfParts>");
        lines.push("</Properties>");
        zip.file("docProps/app.xml", lines.join("\n"));
      }
    };
    DEFAULT_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
  }
});

// src/features/DefinedName.ts
exports.DefinedName = void 0; exports.DefinedNameCollection = void 0;
var init_DefinedName = __esm({
  "src/features/DefinedName.ts"() {
    exports.DefinedName = class {
      constructor(name, refersTo, localSheetId = null) {
        this._comment = null;
        this._description = null;
        this._hidden = false;
        this._name = name;
        this._refersTo = refersTo;
        this._localSheetId = localSheetId;
      }
      /** Name of the defined name. */
      get name() {
        return this._name;
      }
      set name(v) {
        this._name = v;
      }
      /** Formula or range that the name refers to. */
      get refersTo() {
        return this._refersTo;
      }
      set refersTo(v) {
        this._refersTo = v;
      }
      /** Sheet index for sheet-local names (null for global names). */
      get localSheetId() {
        return this._localSheetId;
      }
      set localSheetId(v) {
        this._localSheetId = v;
      }
      /** Comment associated with the name. */
      get comment() {
        return this._comment;
      }
      set comment(v) {
        this._comment = v;
      }
      /** Description of the name. */
      get description() {
        return this._description;
      }
      set description(v) {
        this._description = v;
      }
      /** Whether the name is hidden. */
      get hidden() {
        return this._hidden;
      }
      set hidden(v) {
        this._hidden = v;
      }
      // snake_case aliases
      get refers_to() {
        return this._refersTo;
      }
      set refers_to(v) {
        this._refersTo = v;
      }
      get local_sheet_id() {
        return this._localSheetId;
      }
      set local_sheet_id(v) {
        this._localSheetId = v;
      }
    };
    exports.DefinedNameCollection = class {
      constructor() {
        this._names = [];
      }
      /**
       * Adds a defined name to the collection.
       *
       * @param nameOrDn Either a DefinedName object or a string name.
       * @param refersTo Formula or range (required if nameOrDn is a string).
       * @param localSheetId Sheet index for sheet-local names.
       * @returns The added DefinedName.
       */
      add(nameOrDn, refersTo, localSheetId) {
        if (nameOrDn instanceof exports.DefinedName) {
          this._names.push(nameOrDn);
          return nameOrDn;
        }
        const dn = new exports.DefinedName(nameOrDn, refersTo ?? "", localSheetId ?? null);
        this._names.push(dn);
        return dn;
      }
      /**
       * Removes a defined name by name string.
       * @returns The removed DefinedName or null if not found.
       */
      remove(name) {
        const idx = this._names.findIndex((dn) => dn.name === name);
        if (idx >= 0) {
          return this._names.splice(idx, 1)[0];
        }
        return null;
      }
      /**
       * Gets a defined name by index or name string.
       */
      get(key) {
        if (typeof key === "number") {
          if (key < 0 || key >= this._names.length) {
            throw new RangeError(`Index ${key} out of range`);
          }
          return this._names[key];
        }
        const found = this._names.find((dn) => dn.name === key);
        if (!found) {
          throw new Error(`Defined name '${key}' not found`);
        }
        return found;
      }
      /** Number of defined names. */
      get count() {
        return this._names.length;
      }
      /** Returns internal array (for iteration). */
      toArray() {
        return [...this._names];
      }
      [Symbol.iterator]() {
        return this._names[Symbol.iterator]();
      }
    };
  }
});

// src/io/XmlAutoFilterLoader.ts
function ensureArray2(val) {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}
function attr2(elem, name, defaultVal = "") {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : defaultVal;
}
function attrOpt(elem, name) {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : null;
}
exports.AutoFilterXmlLoader = void 0;
var init_XmlAutoFilterLoader = __esm({
  "src/io/XmlAutoFilterLoader.ts"() {
    init_AutoFilter();
    exports.AutoFilterXmlLoader = class {
      loadAutoFilter(wsRoot, autoFilter) {
        const afElem = wsRoot?.autoFilter;
        if (!afElem) return;
        const ref = attr2(afElem, "ref", "");
        if (!ref) return;
        autoFilter.range = ref;
        const fcElems = ensureArray2(afElem.filterColumn);
        for (const fcElem of fcElems) {
          const colId = parseInt(attr2(fcElem, "colId", "0"), 10);
          const fc = autoFilter.getFilterColumn(colId) ?? new exports.FilterColumn(colId);
          const hiddenButton = attr2(fcElem, "hiddenButton", "");
          if (hiddenButton === "1" || hiddenButton === "true") {
            fc.filterButton = false;
          }
          const filtersElem = fcElem.filters;
          if (filtersElem) {
            const filterElems = ensureArray2(filtersElem.filter);
            for (const fe of filterElems) {
              const val = attr2(fe, "val", "");
              if (val) fc.addFilter(val);
            }
          }
          const customFiltersElem = fcElem.customFilters;
          if (customFiltersElem) {
            const customElems = ensureArray2(customFiltersElem.customFilter);
            for (const ce of customElems) {
              const op = attr2(ce, "operator", "equal");
              const val = attr2(ce, "val", "");
              fc.addCustomFilter(op, val);
            }
          }
          const colorFilterElem = fcElem.colorFilter;
          if (colorFilterElem) {
            const color = attr2(colorFilterElem, "dxfId", "");
            const cellColor = attr2(colorFilterElem, "cellColor", "1") !== "0";
            fc.colorFilter = { color, cellColor };
          }
          const dynamicElem = fcElem.dynamicFilter;
          if (dynamicElem) {
            const type = attr2(dynamicElem, "type", "");
            const val = attrOpt(dynamicElem, "val");
            fc.dynamicFilter = { type, value: val ?? void 0 };
          }
          const top10Elem = fcElem.top10;
          if (top10Elem) {
            const top = attr2(top10Elem, "top", "1") !== "0";
            const percent = attr2(top10Elem, "percent", "0") === "1";
            const val = parseInt(attr2(top10Elem, "val", "10"), 10);
            fc.top10Filter = { top, percent, val };
          }
          if (!autoFilter.getFilterColumn(colId)) {
            autoFilter.filterColumns.set(colId, fc);
          } else {
            autoFilter.filterColumns.set(colId, fc);
          }
        }
        const sortStateElem = afElem.sortState;
        if (sortStateElem) {
          const sortRef = attr2(sortStateElem, "ref", "");
          const sortConditions = ensureArray2(sortStateElem.sortCondition);
          if (sortConditions.length > 0) {
            const condRef = attr2(sortConditions[0], "ref", "");
            const descending = attr2(sortConditions[0], "descending", "");
            const colIndex = this._calculateSortColumnIndex(ref, condRef);
            autoFilter.sortState = {
              ref: sortRef || ref,
              columnIndex: colIndex,
              ascending: descending !== "1"
            };
          }
        }
      }
      _calculateSortColumnIndex(filterRange, sortCondRef) {
        if (!filterRange || !sortCondRef) return 0;
        const filterCol = this._colFromRef(filterRange.split(":")[0]);
        const sortCol = this._colFromRef(sortCondRef.split(":")[0]);
        return sortCol - filterCol;
      }
      _colFromRef(ref) {
        const match = ref.match(/^([A-Z]+)/i);
        if (!match) return 0;
        let col = 0;
        for (const ch of match[1].toUpperCase()) {
          col = col * 26 + (ch.charCodeAt(0) - 64);
        }
        return col - 1;
      }
    };
  }
});

// src/io/XmlConditionalFormatLoader.ts
function ensureArray3(val) {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}
function attr3(elem, name, defaultVal = "") {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : defaultVal;
}
function attrOpt2(elem, name) {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : null;
}
exports.ConditionalFormatXmlLoader = void 0;
var init_XmlConditionalFormatLoader = __esm({
  "src/io/XmlConditionalFormatLoader.ts"() {
    exports.ConditionalFormatXmlLoader = class {
      constructor(dxfStyles) {
        this._dxfStyles = dxfStyles;
      }
      loadConditionalFormatting(wsRoot, collection) {
        const cfElems = ensureArray3(wsRoot?.conditionalFormatting);
        for (const cfElem of cfElems) {
          const sqref = attr3(cfElem, "sqref", "");
          if (!sqref) continue;
          const rules = ensureArray3(cfElem.cfRule);
          for (const rule of rules) {
            const cf = collection.add();
            cf.range = sqref;
            this._loadCfRule(cf, rule);
          }
        }
        this._applyDxfStyles(collection);
      }
      _loadCfRule(cf, rule) {
        const ruleType = attr3(rule, "type", "");
        cf.type = ruleType || null;
        cf.priority = parseInt(attr3(rule, "priority", "0"), 10);
        const stopIfTrue = attr3(rule, "stopIfTrue", "");
        cf.stopIfTrue = stopIfTrue === "1" || stopIfTrue === "true";
        const dxfId = attrOpt2(rule, "dxfId");
        if (dxfId !== null) {
          cf._dxfId = parseInt(dxfId, 10);
        }
        const op = attrOpt2(rule, "operator");
        if (op) cf.operator = op;
        const text = attrOpt2(rule, "text");
        if (text) cf.textFormula = text;
        const timePeriod = attrOpt2(rule, "timePeriod");
        if (timePeriod) cf.dateOperator = timePeriod;
        const aboveAvg = attrOpt2(rule, "aboveAverage");
        if (ruleType === "aboveAverage") {
          cf.above = aboveAvg !== "0";
        }
        const stdDev = attrOpt2(rule, "stdDev");
        if (stdDev) cf.stdDev = parseInt(stdDev, 10);
        if (ruleType === "top10") {
          const bottom = attrOpt2(rule, "bottom");
          cf.top = bottom !== "1";
          const percent = attrOpt2(rule, "percent");
          cf.percent = percent === "1";
          const rank = attrOpt2(rule, "rank");
          if (rank) cf.rank = parseInt(rank, 10);
        }
        if (ruleType === "duplicateValues") cf.duplicate = true;
        if (ruleType === "uniqueValues") cf.duplicate = false;
        const formulas = ensureArray3(rule.formula);
        if (formulas.length > 0) {
          cf.formula1 = String(formulas[0]);
        }
        if (formulas.length > 1) {
          cf.formula2 = String(formulas[1]);
        }
        if (rule.colorScale) {
          this._loadColorScale(cf, rule.colorScale);
        }
        if (rule.dataBar) {
          this._loadDataBar(cf, rule.dataBar);
        }
        if (rule.iconSet) {
          this._loadIconSet(cf, rule.iconSet);
        }
        if (ruleType === "containsText") cf.textOperator = "containsText";
        else if (ruleType === "notContainsText") cf.textOperator = "notContains";
        else if (ruleType === "beginsWith") cf.textOperator = "beginsWith";
        else if (ruleType === "endsWith") cf.textOperator = "endsWith";
        else if (ruleType === "containsBlanks") cf.textOperator = "containsBlanks";
        else if (ruleType === "notContainsBlanks") cf.textOperator = "notContainsBlanks";
        else if (ruleType === "containsErrors") cf.textOperator = "containsErrors";
        else if (ruleType === "notContainsErrors") cf.textOperator = "notContainsErrors";
      }
      _loadColorScale(cf, elem) {
        const colors = ensureArray3(elem.color);
        if (colors.length === 2) {
          cf.colorScaleType = "2color";
          cf.minColor = attrOpt2(colors[0], "rgb") ?? attrOpt2(colors[0], "theme");
          cf.maxColor = attrOpt2(colors[1], "rgb") ?? attrOpt2(colors[1], "theme");
        } else if (colors.length >= 3) {
          cf.colorScaleType = "3color";
          cf.minColor = attrOpt2(colors[0], "rgb") ?? attrOpt2(colors[0], "theme");
          cf.midColor = attrOpt2(colors[1], "rgb") ?? attrOpt2(colors[1], "theme");
          cf.maxColor = attrOpt2(colors[2], "rgb") ?? attrOpt2(colors[2], "theme");
        }
      }
      _loadDataBar(cf, elem) {
        const colors = ensureArray3(elem.color);
        if (colors.length > 0) {
          cf.barColor = attrOpt2(colors[0], "rgb");
        }
        if (colors.length > 1) {
          cf.negativeColor = attrOpt2(colors[1], "rgb");
        }
        const showValue = attrOpt2(elem, "showValue");
        if (showValue === "0") cf.showIconOnly = true;
      }
      _loadIconSet(cf, elem) {
        const iconSet = attrOpt2(elem, "iconSet");
        if (iconSet) cf.iconSetType = iconSet;
        const reverse = attrOpt2(elem, "reverse");
        if (reverse === "1") cf.reverseIcons = true;
        const showValue = attrOpt2(elem, "showValue");
        if (showValue === "0") cf.showIconOnly = true;
      }
      _applyDxfStyles(collection) {
        for (const cf of collection) {
          if (cf._dxfId !== null && cf._dxfId >= 0 && cf._dxfId < this._dxfStyles.length) {
            const dxfData = this._dxfStyles[cf._dxfId];
            this._applyDxfData(cf, dxfData);
          }
        }
      }
      _applyDxfData(cf, dxfData) {
        if (!dxfData) return;
        if (dxfData.font) {
          const fd = dxfData.font;
          if (fd.name) cf.font.name = fd.name;
          if (fd.size) cf.font.size = fd.size;
          if (fd.color) cf.font.color = fd.color;
          if (fd.bold !== void 0) cf.font.bold = fd.bold;
          if (fd.italic !== void 0) cf.font.italic = fd.italic;
          if (fd.underline !== void 0) cf.font.underline = fd.underline;
          if (fd.strikethrough !== void 0) cf.font.strikethrough = fd.strikethrough;
        }
        if (dxfData.fill) {
          const fillD = dxfData.fill;
          if (fillD.pattern_type) cf.fill.patternType = fillD.pattern_type;
          if (fillD.fg_color) cf.fill.foregroundColor = fillD.fg_color;
          if (fillD.bg_color) cf.fill.backgroundColor = fillD.bg_color;
        }
        if (dxfData.border) {
          const bd = dxfData.border;
          if (bd.left?.style) cf.border.left.lineStyle = bd.left.style;
          if (bd.left?.color) cf.border.left.color = bd.left.color;
          if (bd.right?.style) cf.border.right.lineStyle = bd.right.style;
          if (bd.right?.color) cf.border.right.color = bd.right.color;
          if (bd.top?.style) cf.border.top.lineStyle = bd.top.style;
          if (bd.top?.color) cf.border.top.color = bd.top.color;
          if (bd.bottom?.style) cf.border.bottom.lineStyle = bd.bottom.style;
          if (bd.bottom?.color) cf.border.bottom.color = bd.bottom.color;
        }
        if (dxfData.numberFormat) {
          cf.numberFormat = dxfData.numberFormat;
        }
      }
    };
  }
});

// src/io/XmlDataValidationLoader.ts
function ensureArray4(val) {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}
function attr4(elem, name, defaultVal = "") {
  const v = elem?.["@_" + name];
  return v != null ? String(v) : defaultVal;
}
var TYPE_MAP2, OPERATOR_MAP, ALERT_STYLE_MAP, IME_MODE_MAP; exports.DataValidationXmlLoader = void 0;
var init_XmlDataValidationLoader = __esm({
  "src/io/XmlDataValidationLoader.ts"() {
    init_DataValidation();
    TYPE_MAP2 = {
      none: 0 /* NONE */,
      whole: 1 /* WHOLE_NUMBER */,
      decimal: 2 /* DECIMAL */,
      list: 3 /* LIST */,
      date: 4 /* DATE */,
      time: 5 /* TIME */,
      textLength: 6 /* TEXT_LENGTH */,
      custom: 7 /* CUSTOM */
    };
    OPERATOR_MAP = {
      between: 0 /* BETWEEN */,
      notBetween: 1 /* NOT_BETWEEN */,
      equal: 2 /* EQUAL */,
      notEqual: 3 /* NOT_EQUAL */,
      greaterThan: 4 /* GREATER_THAN */,
      lessThan: 5 /* LESS_THAN */,
      greaterThanOrEqual: 6 /* GREATER_THAN_OR_EQUAL */,
      lessThanOrEqual: 7 /* LESS_THAN_OR_EQUAL */
    };
    ALERT_STYLE_MAP = {
      stop: 0 /* STOP */,
      warning: 1 /* WARNING */,
      information: 2 /* INFORMATION */
    };
    IME_MODE_MAP = {
      noControl: 0 /* NO_CONTROL */,
      off: 1 /* OFF */,
      on: 2 /* ON */,
      disabled: 3 /* DISABLED */,
      hiragana: 4 /* HIRAGANA */,
      fullKatakana: 5 /* FULL_KATAKANA */,
      halfKatakana: 6 /* HALF_KATAKANA */,
      fullAlpha: 7 /* FULL_ALPHA */,
      halfAlpha: 8 /* HALF_ALPHA */,
      fullHangul: 9 /* FULL_HANGUL */,
      halfHangul: 10 /* HALF_HANGUL */
    };
    exports.DataValidationXmlLoader = class {
      loadDataValidations(wsRoot, collection) {
        const dvs = wsRoot?.dataValidations;
        if (!dvs) return;
        const disablePrompts = attr4(dvs, "disablePrompts", "");
        if (disablePrompts === "1" || disablePrompts === "true") {
          collection.disablePrompts = true;
        }
        const xWindow = attr4(dvs, "xWindow", "");
        if (xWindow) collection.xWindow = parseInt(xWindow, 10);
        const yWindow = attr4(dvs, "yWindow", "");
        if (yWindow) collection.yWindow = parseInt(yWindow, 10);
        const dvElems = ensureArray4(dvs.dataValidation);
        for (const dvElem of dvElems) {
          this._loadDataValidation(dvElem, collection);
        }
      }
      _loadDataValidation(dvElem, collection) {
        const sqref = attr4(dvElem, "sqref", "");
        if (!sqref) return;
        const dv = new exports.DataValidation(sqref);
        const typeStr = attr4(dvElem, "type", "none");
        dv.type = TYPE_MAP2[typeStr] ?? 0 /* NONE */;
        const opStr = attr4(dvElem, "operator", "between");
        dv.operator = OPERATOR_MAP[opStr] ?? 0 /* BETWEEN */;
        const alertStr = attr4(dvElem, "errorStyle", "stop");
        dv.alertStyle = ALERT_STYLE_MAP[alertStr] ?? 0 /* STOP */;
        const imeStr = attr4(dvElem, "imeMode", "");
        if (imeStr && IME_MODE_MAP[imeStr] !== void 0) {
          dv.imeMode = IME_MODE_MAP[imeStr];
        }
        const allowBlank = attr4(dvElem, "allowBlank", "");
        dv.allowBlank = allowBlank === "1" || allowBlank === "true";
        const showInputMsg = attr4(dvElem, "showInputMessage", "");
        dv.showInputMessage = showInputMsg === "1" || showInputMsg === "true";
        const showErrorMsg = attr4(dvElem, "showErrorMessage", "");
        dv.showErrorMessage = showErrorMsg === "1" || showErrorMsg === "true";
        const showDropDown = attr4(dvElem, "showDropDown", "");
        if (showDropDown === "1" || showDropDown === "true") {
          dv.showDropdown = false;
        } else {
          dv.showDropdown = true;
        }
        const errorTitle = attr4(dvElem, "errorTitle", "");
        if (errorTitle) dv.errorTitle = errorTitle;
        const errorAttr = attr4(dvElem, "error", "");
        if (errorAttr) dv.errorMessage = errorAttr;
        const promptTitle = attr4(dvElem, "promptTitle", "");
        if (promptTitle) dv.inputTitle = promptTitle;
        const promptAttr = attr4(dvElem, "prompt", "");
        if (promptAttr) dv.inputMessage = promptAttr;
        const formula1 = dvElem?.formula1;
        if (formula1 != null) {
          dv.formula1 = typeof formula1 === "object" ? formula1["#text"] ?? String(formula1) : String(formula1);
        }
        const formula2 = dvElem?.formula2;
        if (formula2 != null) {
          dv.formula2 = typeof formula2 === "object" ? formula2["#text"] ?? String(formula2) : String(formula2);
        }
        collection.addValidation(dv);
      }
    };
  }
});

// src/io/XmlLoader.ts
var XmlLoader_exports = {};
__export(XmlLoader_exports, {
  XmlLoader: () => exports.XmlLoader
});
function createParser() {
  return new fastXmlParser.XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    textNodeName: "#text",
    isArray: (name) => {
      const arrayElements = /* @__PURE__ */ new Set([
        "sheet",
        "row",
        "c",
        "si",
        "t",
        "r",
        "font",
        "fill",
        "border",
        "xf",
        "numFmt",
        "mergeCell",
        "col",
        "Relationship",
        "Override",
        "Default",
        "dxf",
        "definedName"
      ]);
      return arrayElements.has(name);
    },
    removeNSPrefix: true,
    parseTagValue: false,
    trimValues: false
  });
}
function attr5(elem, name, defaultVal = "") {
  if (!elem) return defaultVal;
  const key = `@_${name}`;
  return elem[key] !== void 0 ? String(elem[key]) : defaultVal;
}
function attrOpt3(elem, name) {
  if (!elem) return null;
  const key = `@_${name}`;
  return elem[key] !== void 0 ? String(elem[key]) : null;
}
function ensureArray5(val) {
  if (val === void 0 || val === null) return [];
  return Array.isArray(val) ? val : [val];
}
function extractRootExtraAttrs(xmlText, tagName) {
  const regex = new RegExp(`<${tagName}\\b([^>]*)>`, "s");
  const match = regex.exec(xmlText);
  if (!match) return "";
  let attrs = match[1];
  attrs = attrs.replace(/\s+xmlns="[^"]*"/g, "");
  attrs = attrs.replace(/\s+xmlns:r="[^"]*"/g, "");
  return attrs.trim();
}
function extractRawElement(xmlText, tagName) {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const patterns = [
    new RegExp(`(<${escaped}\\b[^>]*/>)`, "s"),
    new RegExp(`(<${escaped}\\b[^>]*>.*?</${escaped}>)`, "s")
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(xmlText);
    if (match) return match[1];
  }
  return null;
}
var BUILTIN_FORMATS; exports.XmlLoader = void 0;
var init_XmlLoader = __esm({
  "src/io/XmlLoader.ts"() {
    init_Workbook();
    init_Worksheet();
    init_Cell();
    init_CellValueHandler();
    init_XmlAutoFilterLoader();
    init_XmlConditionalFormatLoader();
    init_XmlDataValidationLoader();
    init_XmlHyperlinkHandler();
    init_CommentXml();
    init_DefinedName();
    BUILTIN_FORMATS = {
      0: "General",
      1: "0",
      2: "0.00",
      3: "#,##0",
      4: "#,##0.00",
      5: "$#,##0_);($#,##0)",
      6: "$#,##0_);[Red]($#,##0)",
      7: "$#,##0.00_);($#,##0.00)",
      8: "$#,##0.00_);[Red]($#,##0.00)",
      9: "0%",
      10: "0.00%",
      11: "0.00E+00",
      12: "# ?/?",
      13: "# ??/??",
      14: "m/d/yyyy",
      15: "d-mmm-yy",
      16: "d-mmm",
      17: "mmm-yy",
      18: "h:mm AM/PM",
      19: "h:mm:ss AM/PM",
      20: "h:mm",
      21: "h:mm:ss",
      22: "m/d/yy h:mm",
      37: "#,##0_);(#,##0)",
      38: "#,##0_);[Red](#,##0)",
      39: "#,##0.00_);(#,##0.00)",
      40: "#,##0.00_);[Red](#,##0.00)",
      45: "mm:ss",
      46: "[h]:mm:ss",
      47: "mm:ss.0",
      48: "##0.0E+0",
      49: "@"
    };
    exports.XmlLoader = class _XmlLoader {
      constructor(workbook) {
        // Content-type tracking
        this._contentTypeOverrides = {};
        this._contentTypeDefaults = {};
        this._workbook = workbook;
        this._wb = workbook;
        this._parser = createParser();
        this._wb._loaderFontStyles = /* @__PURE__ */ new Map();
        this._wb._loaderFillStyles = /* @__PURE__ */ new Map();
        this._wb._loaderBorderStyles = /* @__PURE__ */ new Map();
        this._wb._loaderAlignmentStyles = /* @__PURE__ */ new Map();
        this._wb._loaderProtectionStyles = /* @__PURE__ */ new Map();
        this._wb._loaderNumFormats = /* @__PURE__ */ new Map();
        this._wb._loaderCellStyles = /* @__PURE__ */ new Map();
        this._wb._loaderCellXfByIndex = /* @__PURE__ */ new Map();
        this._wb._sourceStylesXml = null;
        this._wb._sourceCellXfsCount = 0;
        this._wb._themeXml = null;
        this._wb._contentTypeDefaults = {};
        this._wb._contentTypeOverrides = {};
        this._wb._sourceExtraWorkbookRels = [];
        this._wb._sourceCalcChainBytes = null;
        this._wb._sourceCalcChainRel = null;
        this._wb._dxfStyles = [];
      }
      // ── Public API ──────────────────────────────────────────────────────────
      /**
       * Loads workbook data from a JSZip instance.
       */
      async loadWorkbook(zip) {
        const wbXmlBuf = await zip.file("xl/workbook.xml")?.async("nodebuffer");
        if (!wbXmlBuf) throw new Error("xl/workbook.xml not found in XLSX archive");
        const wbXmlText = wbXmlBuf.toString("utf-8");
        const wbRoot = this._parser.parse(wbXmlText);
        const workbookNode = wbRoot.workbook || wbRoot;
        this._wb._sourceXml["workbook_root_extra_attrs"] = extractRootExtraAttrs(wbXmlText, "workbook");
        const altContent = extractRawElement(wbXmlText, "mc:AlternateContent");
        if (altContent) this._wb._sourceXml["workbook_alt_content"] = altContent;
        const revPtr = extractRawElement(wbXmlText, "xr:revisionPtr");
        if (revPtr) this._wb._sourceXml["workbook_revision_ptr"] = revPtr;
        const extLst = extractRawElement(wbXmlText, "extLst");
        if (extLst) this._wb._sourceXml["workbook_extlst"] = extLst;
        this._loadWorkbookProperties(workbookNode);
        await this._loadDocumentProperties(zip);
        await this._loadContentTypeOverrides(zip);
        await this._loadTheme(zip);
        this._loadWorksheetInfo(workbookNode);
        this._applyPrintAreasFromDefinedNames(workbookNode);
        await this._loadExtraWorkbookRels(zip);
        await this._loadSharedStrings(zip);
        await this._loadStyles(zip);
        await this._loadWorksheetsData(zip);
      }
      /**
       * Loads a workbook from a Buffer.
       */
      static async loadFromBuffer(buffer) {
        const zip = await JSZip__default.default.loadAsync(buffer);
        const workbook = Object.create(exports.Workbook.prototype);
        const wb = workbook;
        wb._worksheets = [];
        wb._styles = [];
        wb._sharedStrings = [];
        workbook._filePath = null;
        workbook._properties = new (await Promise.resolve().then(() => (init_Workbook(), Workbook_exports))).WorkbookProperties();
        workbook._documentProperties = null;
        wb._sourceXml = {};
        workbook._fontStyles = /* @__PURE__ */ new Map();
        workbook._fillStyles = /* @__PURE__ */ new Map();
        workbook._borderStyles = /* @__PURE__ */ new Map();
        workbook._alignmentStyles = /* @__PURE__ */ new Map();
        workbook._protectionStyles = /* @__PURE__ */ new Map();
        workbook._cellStyles = /* @__PURE__ */ new Map();
        workbook._numFormats = /* @__PURE__ */ new Map();
        const { Style: Style3 } = await Promise.resolve().then(() => (init_Style(), Style_exports));
        wb._styles.push(new Style3());
        const loader = new _XmlLoader(workbook);
        await loader.loadWorkbook(zip);
        return workbook;
      }
      /**
       * Loads a workbook from a file path.
       */
      static async loadFromFile(filePath) {
        const fs = await import('fs');
        const buffer = fs.readFileSync(filePath);
        const workbook = await _XmlLoader.loadFromBuffer(buffer);
        workbook._filePath = filePath;
        return workbook;
      }
      // ── Content Types ───────────────────────────────────────────────────────
      async _loadContentTypeOverrides(zip) {
        this._contentTypeDefaults = {};
        this._contentTypeOverrides = {};
        this._wb._contentTypeDefaults = {};
        this._wb._contentTypeOverrides = {};
        const file = zip.file("[Content_Types].xml");
        if (!file) return;
        try {
          const buf = await file.async("nodebuffer");
          const root = this._parser.parse(buf.toString("utf-8"));
          const types = root.Types || root;
          for (const def of ensureArray5(types.Default)) {
            const ext = attrOpt3(def, "Extension");
            const ct = attrOpt3(def, "ContentType");
            if (ext && ct) {
              this._contentTypeDefaults[ext.toLowerCase()] = ct;
              this._wb._contentTypeDefaults[ext.toLowerCase()] = ct;
            }
          }
          for (const ov of ensureArray5(types.Override)) {
            const partName = attrOpt3(ov, "PartName");
            const ct = attrOpt3(ov, "ContentType");
            if (partName && ct) {
              this._contentTypeOverrides[partName] = ct;
              this._wb._contentTypeOverrides[partName] = ct;
            }
          }
        } catch {
        }
      }
      // ── Theme ───────────────────────────────────────────────────────────────
      async _loadTheme(zip) {
        const file = zip.file("xl/theme/theme1.xml");
        if (file) {
          this._wb._themeXml = await file.async("nodebuffer");
        } else {
          this._wb._themeXml = null;
        }
      }
      // ── Workbook Properties ─────────────────────────────────────────────────
      _loadWorkbookProperties(workbookNode) {
        const props = this._wb._properties;
        const fv = workbookNode.fileVersion;
        if (fv && props.file_version) {
          props.file_version.appName = attr5(fv, "appName", "xl");
          props.file_version.lastEdited = attr5(fv, "lastEdited", "5");
          props.file_version.lowestEdited = attr5(fv, "lowestEdited", "5");
          props.file_version.rupBuild = attr5(fv, "rupBuild", "9302");
        }
        const wbPr = workbookNode.workbookPr;
        if (wbPr && props.workbook_pr) {
          const dtvStr = attrOpt3(wbPr, "defaultThemeVersion");
          if (dtvStr) props.workbook_pr.defaultThemeVersion = dtvStr;
        }
        const bookViews = workbookNode.bookViews;
        if (bookViews) {
          const wbView = bookViews.workbookView;
          const view = Array.isArray(wbView) ? wbView[0] : wbView;
          if (view) {
            const tab = attrOpt3(view, "activeTab");
            if (tab !== null) props.view.activeTab = parseInt(tab, 10) || 0;
          }
        }
        const calcPr = workbookNode.calcPr;
        if (calcPr && props.calculation !== void 0) {
          const calcMode = attrOpt3(calcPr, "calcMode");
          if (calcMode) props.calcMode = calcMode;
        }
        this._loadDefinedNames(workbookNode);
      }
      _loadDefinedNames(workbookNode) {
        const dnContainer = workbookNode.definedNames;
        if (!dnContainer) return;
        const dnElems = ensureArray5(dnContainer.definedName);
        const collection = this._wb._properties.definedNames;
        for (const dnElem of dnElems) {
          const name = dnElem?.["@_name"] ?? "";
          if (!name) continue;
          const refersTo = typeof dnElem === "string" ? dnElem : dnElem?.["#text"] ?? "";
          let localSheetId = null;
          const lsId = dnElem?.["@_localSheetId"];
          if (lsId != null) {
            localSheetId = parseInt(String(lsId), 10);
          }
          const dn = new exports.DefinedName(name, String(refersTo), localSheetId);
          dn.hidden = dnElem?.["@_hidden"] === "1";
          const comment = dnElem?.["@_comment"];
          if (comment) dn.comment = String(comment);
          collection._names.push(dn);
        }
      }
      // ── Worksheet Info ──────────────────────────────────────────────────────
      _loadWorksheetInfo(workbookNode) {
        const sheetsNode = workbookNode.sheets;
        if (!sheetsNode) return;
        const sheets = ensureArray5(sheetsNode.sheet);
        for (const sheet of sheets) {
          const sheetName = attr5(sheet, "name", "Sheet");
          const ws = new exports.Worksheet(sheetName);
          ws._workbook = this._workbook;
          const state = attrOpt3(sheet, "state");
          if (state === "hidden") {
            ws._visible = false;
          } else if (state === "veryHidden") {
            ws._visible = "veryHidden";
          }
          this._wb._worksheets.push(ws);
        }
      }
      // ── Print Areas from Defined Names ──────────────────────────────────────
      _applyPrintAreasFromDefinedNames(workbookNode) {
        const dnNode = workbookNode.definedNames;
        if (!dnNode) return;
        const names = ensureArray5(dnNode.definedName);
        for (const dn of names) {
          const name = attr5(dn, "name");
          if (name !== "_xlnm.Print_Area") continue;
          const sheetIdStr = attrOpt3(dn, "localSheetId");
          if (sheetIdStr === null) continue;
          const sheetId = parseInt(sheetIdStr, 10);
          if (isNaN(sheetId) || sheetId < 0 || sheetId >= this._wb._worksheets.length) continue;
          const refersTo = typeof dn === "object" && dn["#text"] ? String(dn["#text"]) : "";
          if (!refersTo) continue;
          const ws = this._wb._worksheets[sheetId];
          const printArea = this._extractPrintArea(refersTo);
          if (printArea) {
            ws._printArea = printArea;
          }
        }
      }
      _extractPrintArea(refersTo) {
        if (!refersTo) return null;
        const parts = [];
        for (let token of refersTo.split(",")) {
          token = token.trim();
          if (!token) continue;
          let addr = token;
          if (token.includes("!")) {
            addr = token.split("!")[1];
          }
          addr = addr.replace(/\$/g, "").trim().toUpperCase();
          parts.push(addr);
        }
        return parts.length > 0 ? parts.join(",") : null;
      }
      // ── Shared Strings ──────────────────────────────────────────────────────
      async _loadSharedStrings(zip) {
        const file = zip.file("xl/sharedStrings.xml");
        if (!file) {
          this._wb._sharedStrings = [];
          return;
        }
        try {
          const buf = await file.async("nodebuffer");
          const root = this._parser.parse(buf.toString("utf-8"));
          const sst = root.sst || root;
          this._wb._sharedStrings = [];
          const siList = ensureArray5(sst.si);
          for (const si of siList) {
            if (si.t !== void 0) {
              const tArr = ensureArray5(si.t);
              const texts = [];
              for (const t of tArr) {
                if (typeof t === "string") {
                  texts.push(t);
                } else if (t && t["#text"] !== void 0) {
                  texts.push(String(t["#text"]));
                } else if (typeof t === "number") {
                  texts.push(String(t));
                } else {
                  texts.push("");
                }
              }
              this._wb._sharedStrings.push(texts.join(""));
              continue;
            }
            if (si.r !== void 0) {
              const runs = ensureArray5(si.r);
              const texts = [];
              for (const run of runs) {
                if (run.t !== void 0) {
                  const tArr = ensureArray5(run.t);
                  for (const t of tArr) {
                    if (typeof t === "string") {
                      texts.push(t);
                    } else if (t && t["#text"] !== void 0) {
                      texts.push(String(t["#text"]));
                    } else if (typeof t === "number") {
                      texts.push(String(t));
                    } else {
                      texts.push("");
                    }
                  }
                }
              }
              this._wb._sharedStrings.push(texts.join(""));
              continue;
            }
            this._wb._sharedStrings.push("");
          }
        } catch {
          this._wb._sharedStrings = [];
        }
      }
      // ── Styles ──────────────────────────────────────────────────────────────
      async _loadStyles(zip) {
        const file = zip.file("xl/styles.xml");
        if (!file) return;
        try {
          const buf = await file.async("nodebuffer");
          this._wb._sourceStylesXml = buf;
          const root = this._parser.parse(buf.toString("utf-8"));
          const stylesRoot = root.styleSheet || root;
          const cellXfsElem = stylesRoot.cellXfs;
          if (cellXfsElem) {
            this._wb._sourceCellXfsCount = parseInt(attr5(cellXfsElem, "count", "0"), 10);
          }
          this._loadStylesXml(stylesRoot);
          this._loadDxfStyles(stylesRoot);
        } catch {
        }
      }
      _loadStylesXml(stylesRoot) {
        for (const [id, code] of Object.entries(BUILTIN_FORMATS)) {
          this._wb._loaderNumFormats.set(Number(id), code);
        }
        const numFmtsNode = stylesRoot.numFmts;
        if (numFmtsNode) {
          const fmts = ensureArray5(numFmtsNode.numFmt);
          for (const fmt of fmts) {
            const id = parseInt(attr5(fmt, "numFmtId", "0"), 10);
            const code = attr5(fmt, "formatCode", "General");
            this._wb._loaderNumFormats.set(id, code);
          }
        }
        this._loadFonts(stylesRoot);
        this._loadFills(stylesRoot);
        this._loadBorders(stylesRoot);
        this._loadCellXfs(stylesRoot);
      }
      // ── Fonts ───────────────────────────────────────────────────────────────
      _loadFonts(stylesRoot) {
        const fontsNode = stylesRoot.fonts;
        if (!fontsNode) return;
        const fonts = ensureArray5(fontsNode.font);
        for (let i = 0; i < fonts.length; i++) {
          const fontElem = fonts[i];
          const szElem = fontElem.sz;
          const colorElem = fontElem.color;
          const nameElem = fontElem.name;
          const familyElem = fontElem.family;
          const charsetElem = fontElem.charset;
          const schemeElem = fontElem.scheme;
          const bElem = fontElem.b;
          const iElem = fontElem.i;
          const uElem = fontElem.u;
          const strikeElem = fontElem.strike;
          let colorType = null;
          let colorValue = null;
          let colorTint = null;
          if (colorElem) {
            if (attrOpt3(colorElem, "rgb") !== null) {
              colorType = "rgb";
              colorValue = attr5(colorElem, "rgb");
            } else if (attrOpt3(colorElem, "theme") !== null) {
              colorType = "theme";
              colorValue = attr5(colorElem, "theme");
            } else if (attrOpt3(colorElem, "indexed") !== null) {
              colorType = "indexed";
              colorValue = attr5(colorElem, "indexed");
            } else if (attrOpt3(colorElem, "auto") !== null) {
              colorType = "auto";
              colorValue = attr5(colorElem, "auto");
            }
            if (attrOpt3(colorElem, "tint") !== null) {
              colorTint = attr5(colorElem, "tint");
            }
          }
          const fontData = {
            name: nameElem ? attr5(nameElem, "val", "Calibri") : "Calibri",
            size: szElem ? parseFloat(attr5(szElem, "val", "11")) : 11,
            color: colorType === "rgb" && colorValue ? colorValue : "FF000000",
            color_type: colorType,
            color_value: colorValue,
            color_tint: colorTint,
            family: familyElem ? attr5(familyElem, "val") : null,
            charset: charsetElem ? attr5(charsetElem, "val") : null,
            scheme: schemeElem ? attr5(schemeElem, "val") : null,
            bold: bElem !== void 0,
            italic: iElem !== void 0,
            underline: uElem !== void 0,
            strikethrough: strikeElem !== void 0
          };
          this._wb._loaderFontStyles.set(i, fontData);
          if (i === 0 && this._wb._styles.length > 0) {
            const defaultFont = this._wb._styles[0].font;
            defaultFont.name = fontData.name;
            defaultFont.size = fontData.size;
            defaultFont.bold = fontData.bold;
            defaultFont.italic = fontData.italic;
            defaultFont.underline = fontData.underline;
            defaultFont.strikethrough = fontData.strikethrough;
            if (fontData.color_type === "rgb" && fontData.color_value) {
              defaultFont.color = fontData.color_value;
            }
          }
        }
      }
      // ── Fills ───────────────────────────────────────────────────────────────
      _loadFills(stylesRoot) {
        const fillsNode = stylesRoot.fills;
        if (!fillsNode) return;
        const fills = ensureArray5(fillsNode.fill);
        for (let i = 0; i < fills.length; i++) {
          const fillElem = fills[i];
          const patternElem = fillElem.patternFill;
          const fgColorElem = patternElem ? patternElem.fgColor : null;
          const bgColorElem = patternElem ? patternElem.bgColor : null;
          const [fgType, fgValue, fgTint] = this._extractColorAttrs(fgColorElem);
          const [bgType, bgValue, bgTint] = this._extractColorAttrs(bgColorElem);
          const fgRgb = fgType === "rgb" && fgValue ? fgValue : "FFFFFFFF";
          const bgRgb = bgType === "rgb" && bgValue ? bgValue : "FFFFFFFF";
          const fillData = {
            pattern_type: patternElem ? attr5(patternElem, "patternType", "none") : "none",
            fg_color: fgRgb,
            bg_color: bgRgb,
            fg_color_type: fgType,
            fg_color_value: fgValue,
            fg_color_tint: fgTint,
            bg_color_type: bgType,
            bg_color_value: bgValue,
            bg_color_tint: bgTint
          };
          this._wb._loaderFillStyles.set(i, fillData);
        }
      }
      _extractColorAttrs(colorElem) {
        if (!colorElem) return ["rgb", "FFFFFFFF", null];
        const tintStr = attrOpt3(colorElem, "tint");
        const tint = tintStr !== null ? parseFloat(tintStr) : null;
        if (attrOpt3(colorElem, "rgb") !== null) {
          return ["rgb", attr5(colorElem, "rgb"), tint];
        }
        if (attrOpt3(colorElem, "theme") !== null) {
          return ["theme", attr5(colorElem, "theme"), tint];
        }
        if (attrOpt3(colorElem, "indexed") !== null) {
          return ["indexed", attr5(colorElem, "indexed"), tint];
        }
        if (attrOpt3(colorElem, "auto") !== null) {
          return ["auto", attr5(colorElem, "auto"), tint];
        }
        return ["rgb", "FFFFFFFF", tint];
      }
      // ── Borders ─────────────────────────────────────────────────────────────
      _loadBorders(stylesRoot) {
        const bordersNode = stylesRoot.borders;
        if (!bordersNode) return;
        const borders = ensureArray5(bordersNode.border);
        for (let i = 0; i < borders.length; i++) {
          if (i === 0) continue;
          const borderElem = borders[i];
          const loadSide = (side) => {
            const sideElem = borderElem[side];
            if (!sideElem) return { style: "none", color: "FF000000" };
            const style = attr5(sideElem, "style", "none");
            const colorElem = sideElem.color;
            const color = colorElem ? attr5(colorElem, "rgb", "FF000000") : "FF000000";
            return { style, color };
          };
          const borderData = {
            top: loadSide("top"),
            bottom: loadSide("bottom"),
            left: loadSide("left"),
            right: loadSide("right")
          };
          this._wb._loaderBorderStyles.set(i, borderData);
        }
      }
      // ── Cell XFs ────────────────────────────────────────────────────────────
      _loadCellXfs(stylesRoot) {
        if (!this._wb._loaderAlignmentStyles.has(0)) {
          this._wb._loaderAlignmentStyles.set(0, {
            horizontal: "general",
            vertical: "bottom",
            wrap_text: false,
            indent: 0,
            text_rotation: 0,
            shrink_to_fit: false,
            reading_order: 0,
            relative_indent: 0
          });
        }
        if (!this._wb._loaderProtectionStyles.has(0)) {
          this._wb._loaderProtectionStyles.set(0, {
            locked: true,
            hidden: false
          });
        }
        const cellXfsNode = stylesRoot.cellXfs;
        if (!cellXfsNode) return;
        const xfs = ensureArray5(cellXfsNode.xf);
        for (let i = 0; i < xfs.length; i++) {
          const xfElem = xfs[i];
          const fontIdx = parseInt(attr5(xfElem, "fontId", "0"), 10);
          const fillIdx = parseInt(attr5(xfElem, "fillId", "0"), 10);
          const borderIdx = parseInt(attr5(xfElem, "borderId", "0"), 10);
          const numFmtIdx = parseInt(attr5(xfElem, "numFmtId", "0"), 10);
          let alignmentIdx = 0;
          const alignElem = xfElem.alignment;
          if (alignElem) {
            const horizontal = attr5(alignElem, "horizontal", "general");
            const vertical = attr5(alignElem, "vertical", "bottom");
            const textRotation = parseInt(attr5(alignElem, "textRotation", "0"), 10);
            const wrapText = attr5(alignElem, "wrapText") === "1";
            const shrinkToFit = attr5(alignElem, "shrinkToFit") === "1";
            const indent = parseInt(attr5(alignElem, "indent", "0"), 10);
            const readingOrder = parseInt(attr5(alignElem, "readingOrder", "0"), 10);
            const relativeIndent = parseInt(attr5(alignElem, "relativeIndent", "0"), 10);
            let found = false;
            for (const [idx, ad] of this._wb._loaderAlignmentStyles) {
              if (ad.horizontal === horizontal && ad.vertical === vertical && ad.wrap_text === wrapText && ad.indent === indent && ad.text_rotation === textRotation && ad.shrink_to_fit === shrinkToFit && ad.reading_order === readingOrder && ad.relative_indent === relativeIndent) {
                alignmentIdx = idx;
                found = true;
                break;
              }
            }
            if (!found) {
              alignmentIdx = this._wb._loaderAlignmentStyles.size;
              this._wb._loaderAlignmentStyles.set(alignmentIdx, {
                horizontal,
                vertical,
                wrap_text: wrapText,
                indent,
                text_rotation: textRotation,
                shrink_to_fit: shrinkToFit,
                reading_order: readingOrder,
                relative_indent: relativeIndent
              });
            }
          }
          let protectionIdx = 0;
          const protElem = xfElem.protection;
          if (protElem) {
            const locked = attr5(protElem, "locked", "1") === "1";
            const hidden = attr5(protElem, "hidden", "0") === "1";
            let found = false;
            for (const [idx, pd] of this._wb._loaderProtectionStyles) {
              if (pd.locked === locked && pd.hidden === hidden) {
                protectionIdx = idx;
                found = true;
                break;
              }
            }
            if (!found && !(locked && !hidden)) {
              protectionIdx = this._wb._loaderProtectionStyles.size;
              this._wb._loaderProtectionStyles.set(protectionIdx, {
                locked,
                hidden
              });
            }
          }
          const cellStyleKey = [
            fontIdx,
            fillIdx,
            borderIdx,
            numFmtIdx,
            alignmentIdx,
            protectionIdx
          ];
          const keyStr = JSON.stringify(cellStyleKey);
          this._wb._loaderCellStyles.set(keyStr, i);
          this._wb._loaderCellXfByIndex.set(i, cellStyleKey);
        }
      }
      // ── DXF Styles (conditional formatting) ─────────────────────────────────
      _loadDxfStyles(stylesRoot) {
        this._wb._dxfStyles = [];
        const dxfsNode = stylesRoot.dxfs;
        if (!dxfsNode) return;
        const dxfs = ensureArray5(dxfsNode.dxf);
        for (const dxfElem of dxfs) {
          const dxfData = {};
          const fontElem = dxfElem.font;
          if (fontElem) {
            const fontData = {};
            if (fontElem.b !== void 0) {
              fontData.bold = attr5(fontElem.b, "val", "1") !== "0";
            }
            if (fontElem.i !== void 0) {
              fontData.italic = attr5(fontElem.i, "val", "1") !== "0";
            }
            if (fontElem.u !== void 0) {
              fontData.underline = true;
            }
            if (fontElem.strike !== void 0) {
              fontData.strikethrough = true;
            }
            if (fontElem.color) {
              fontData.color = attr5(fontElem.color, "rgb", "FF000000");
            }
            if (Object.keys(fontData).length > 0) {
              dxfData.font = fontData;
            }
          }
          const fillElem = dxfElem.fill;
          if (fillElem) {
            const patternElem = fillElem.patternFill;
            if (patternElem) {
              const fillData = {
                pattern_type: attr5(patternElem, "patternType", "solid")
              };
              if (patternElem.fgColor) {
                fillData.fg_color = attr5(patternElem.fgColor, "rgb", "FFFFFFFF");
              }
              if (patternElem.bgColor) {
                fillData.bg_color = attr5(patternElem.bgColor, "rgb", "FFFFFFFF");
              }
              dxfData.fill = fillData;
            }
          }
          const borderElem = dxfElem.border;
          if (borderElem) {
            for (const side of ["left", "right", "top", "bottom"]) {
              const sideElem = borderElem[side];
              if (sideElem) {
                const style = attr5(sideElem, "style", "thin");
                let color = "FF000000";
                if (sideElem.color) {
                  color = attr5(sideElem.color, "rgb", "FF000000");
                }
                dxfData.border = { style, color };
                break;
              }
            }
          }
          this._wb._dxfStyles.push(dxfData);
        }
      }
      // ── Extra Workbook Rels ─────────────────────────────────────────────────
      async _loadExtraWorkbookRels(zip) {
        const file = zip.file("xl/_rels/workbook.xml.rels");
        if (!file) return;
        const buf = await file.async("nodebuffer");
        const root = this._parser.parse(buf.toString("utf-8"));
        const relsNode = root.Relationships || root;
        const NATIVE_TYPES = /* @__PURE__ */ new Set([
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        ]);
        const extraRels = [];
        const rels = ensureArray5(relsNode.Relationship);
        for (const rel of rels) {
          const relType = attr5(rel, "Type");
          const target = attr5(rel, "Target");
          const relId = attr5(rel, "Id");
          const targetMode = attr5(rel, "TargetMode");
          if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain") {
            const partPath2 = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
            try {
              const partFile = zip.file(partPath2);
              if (partFile) {
                this._wb._sourceCalcChainBytes = await partFile.async("nodebuffer");
                this._wb._sourceCalcChainRel = {
                  rel_id: relId,
                  target,
                  part_path: partPath2,
                  content_type: this._contentTypeOverrides[`/${partPath2}`] || "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                };
              }
            } catch {
            }
            continue;
          }
          if (NATIVE_TYPES.has(relType)) continue;
          if (targetMode === "External") continue;
          const partPath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
          let partBytes = null;
          try {
            const partFile = zip.file(partPath);
            if (partFile) {
              partBytes = await partFile.async("nodebuffer");
            }
          } catch {
          }
          const lastSlash = partPath.lastIndexOf("/");
          const dirPart = lastSlash >= 0 ? partPath.substring(0, lastSlash) : "";
          const filePart = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
          const partRelsPath = `${dirPart}/_rels/${filePart}.rels`;
          let partRelsBytes = null;
          try {
            const rlsFile = zip.file(partRelsPath);
            if (rlsFile) {
              partRelsBytes = await rlsFile.async("nodebuffer");
            }
          } catch {
          }
          const contentType = this._contentTypeOverrides[`/${partPath}`] || null;
          extraRels.push({
            rel_id: relId,
            rel_type: relType,
            target,
            part_path: partPath,
            part_bytes: partBytes,
            part_rels_path: partRelsPath,
            part_rels_bytes: partRelsBytes,
            content_type: contentType
          });
        }
        if (extraRels.length > 0) {
          this._wb._sourceExtraWorkbookRels = extraRels;
        }
      }
      // ── Worksheets Data ─────────────────────────────────────────────────────
      async _loadWorksheetsData(zip) {
        const worksheetsAndRoots = [];
        for (let i = 0; i < this._wb._worksheets.length; i++) {
          const ws = this._wb._worksheets[i];
          const file = zip.file(`xl/worksheets/sheet${i + 1}.xml`);
          if (!file) {
            worksheetsAndRoots.push({ idx: i, ws, root: null });
            continue;
          }
          try {
            const buf = await file.async("nodebuffer");
            const xmlText = buf.toString("utf-8");
            const parsed = this._parser.parse(xmlText);
            const wsRoot = parsed.worksheet || parsed;
            ws._sourceXml = xmlText;
            const extraAttrs = extractRootExtraAttrs(xmlText, "worksheet");
            if (extraAttrs) {
              ws._source_root_extra_attrs = extraAttrs;
            }
            const sheetPr = extractRawElement(xmlText, "sheetPr");
            if (sheetPr) ws._source_sheet_pr_xml = sheetPr;
            const phoneticPr = extractRawElement(xmlText, "phoneticPr");
            if (phoneticPr) ws._source_phonetic_pr_xml = phoneticPr;
            const dimensionXml = extractRawElement(xmlText, "dimension");
            if (dimensionXml) {
              const m = /ref="([^"]+)"/.exec(dimensionXml);
              if (m) ws._source_dimension_ref = m[1];
            }
            await this._loadExtraSheetRels(zip, ws, i + 1);
            worksheetsAndRoots.push({ idx: i, ws, root: wsRoot });
          } catch {
            worksheetsAndRoots.push({ idx: i, ws, root: null });
          }
        }
        const claimedCommentsPaths = /* @__PURE__ */ new Map();
        for (const { ws } of worksheetsAndRoots) {
          for (const r of ws._source_extra_sheet_rels ?? []) {
            if ((r.rel_type ?? "").toLowerCase().includes("comments")) {
              claimedCommentsPaths.set(r.part_path, ws);
            }
          }
        }
        const sheetRelsMap = /* @__PURE__ */ new Map();
        for (const { idx } of worksheetsAndRoots) {
          const relsPath = `xl/worksheets/_rels/sheet${idx + 1}.xml.rels`;
          const relsFile = zip.file(relsPath);
          if (relsFile) {
            try {
              const relsBuf = await relsFile.async("nodebuffer");
              const relsRoot = this._parser.parse(relsBuf.toString("utf-8"));
              sheetRelsMap.set(idx, relsRoot);
            } catch {
            }
          }
        }
        const commentReader = new exports.CommentXmlReader();
        const hlLoader = new exports.HyperlinkXmlLoader();
        for (const { idx, ws, root } of worksheetsAndRoots) {
          if (root === null) continue;
          try {
            this._loadWorksheetData(ws, root);
            const defaultCommentsPath = `xl/comments${idx + 1}.xml`;
            const commentsOwner = claimedCommentsPaths.get(defaultCommentsPath);
            if (commentsOwner == null || commentsOwner === ws) {
              await this._loadComments(zip, ws, idx + 1, commentReader);
            }
            const relsRoot = sheetRelsMap.get(idx);
            const relationships = relsRoot ? hlLoader.loadRelationships(relsRoot) : /* @__PURE__ */ new Map();
            hlLoader.loadHyperlinks(root, ws.hyperlinks, relationships);
          } catch {
          }
        }
      }
      async _loadComments(zip, worksheet, sheetNum, reader) {
        const commentsFile = zip.file(`xl/comments${sheetNum}.xml`);
        if (!commentsFile) return;
        try {
          const buf = await commentsFile.async("nodebuffer");
          const xmlText = buf.toString("utf-8");
          const parsed = this._parser.parse(xmlText);
          reader.parseAndApplyComments(parsed, worksheet);
        } catch {
        }
        const vmlFile = zip.file(`xl/drawings/vmlDrawing${sheetNum}.vml`);
        if (vmlFile) {
          try {
            const vmlBuf = await vmlFile.async("nodebuffer");
            const vmlContent = vmlBuf.toString("utf-8");
            reader.parseAndApplyVmlDrawing(vmlContent, worksheet);
          } catch {
          }
        }
      }
      // ── Extra Sheet Rels ────────────────────────────────────────────────────
      async _loadExtraSheetRels(zip, worksheet, sheetNum) {
        const relsPath = `xl/worksheets/_rels/sheet${sheetNum}.xml.rels`;
        const file = zip.file(relsPath);
        if (!file) return;
        const buf = await file.async("nodebuffer");
        const root = this._parser.parse(buf.toString("utf-8"));
        const relsNode = root.Relationships || root;
        const PRESERVED_TYPES = /* @__PURE__ */ new Set([
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        ]);
        const extraRels = [];
        const rels = ensureArray5(relsNode.Relationship);
        for (const rel of rels) {
          const relType = attr5(rel, "Type");
          const target = attr5(rel, "Target");
          const targetMode = attr5(rel, "TargetMode");
          const relId = attr5(rel, "Id");
          if (!PRESERVED_TYPES.has(relType)) continue;
          if (targetMode === "External") continue;
          let partPath;
          if (target.startsWith("../")) {
            partPath = "xl/" + target.slice(3);
          } else if (target.startsWith("/")) {
            partPath = target.slice(1);
          } else {
            partPath = "xl/worksheets/" + target;
          }
          let partBytes = null;
          try {
            const partFile = zip.file(partPath);
            if (partFile) {
              partBytes = await partFile.async("nodebuffer");
            }
          } catch {
          }
          extraRels.push({
            rel_id: relId,
            rel_type: relType,
            target,
            target_mode: targetMode,
            part_path: partPath,
            part_bytes: partBytes
          });
        }
        if (extraRels.length > 0) {
          worksheet._source_extra_sheet_rels = extraRels;
        }
      }
      // ── Worksheet Data (cells, merges, columns, rows) ───────────────────────
      _loadWorksheetData(worksheet, wsRoot) {
        this._loadColumnDimensions(worksheet, wsRoot);
        this._loadRowHeights(worksheet, wsRoot);
        this._loadMergedCells(worksheet, wsRoot);
        this._loadFreezePane(worksheet, wsRoot);
        new exports.AutoFilterXmlLoader().loadAutoFilter(wsRoot, worksheet.autoFilter);
        new exports.ConditionalFormatXmlLoader(this._wb._dxfStyles).loadConditionalFormatting(
          wsRoot,
          worksheet.conditionalFormatting
        );
        new exports.DataValidationXmlLoader().loadDataValidations(wsRoot, worksheet.dataValidations);
        this._loadPageBreaks(worksheet, wsRoot);
        const sharedStrings = this._wb._sharedStrings;
        const sheetData = wsRoot.sheetData;
        if (!sheetData) return;
        const rows = ensureArray5(sheetData.row);
        for (const rowElem of rows) {
          const cells = ensureArray5(rowElem.c);
          for (const cellElem of cells) {
            const cellRef = attr5(cellElem, "r");
            if (!cellRef) continue;
            const cellType = attr5(cellElem, "t", "n");
            let formula = null;
            if (cellElem.f !== void 0) {
              const fValue = cellElem.f;
              if (typeof fValue === "string") {
                formula = fValue;
              } else if (fValue && fValue["#text"] !== void 0) {
                formula = String(fValue["#text"]);
              } else if (typeof fValue === "object") {
                formula = null;
              }
              if (formula !== null && !formula.startsWith("=")) {
                formula = "=" + formula;
              }
            }
            const sAttr = attrOpt3(cellElem, "s");
            const styleIdx = sAttr !== null ? parseInt(sAttr, 10) : 0;
            let value = null;
            if (cellElem.v !== void 0) {
              let valueStr = null;
              if (typeof cellElem.v === "string") {
                valueStr = cellElem.v;
              } else if (typeof cellElem.v === "number") {
                valueStr = String(cellElem.v);
              } else if (cellElem.v && cellElem.v["#text"] !== void 0) {
                valueStr = String(cellElem.v["#text"]);
              }
              if (valueStr !== null) {
                value = exports.CellValueHandler.parseValueFromXml(
                  valueStr,
                  cellType,
                  sharedStrings
                );
              }
            }
            const cell = new exports.Cell(value, formula);
            this._applyCellStyle(cell, styleIdx);
            worksheet.cells.set(cellRef, cell);
          }
        }
      }
      // ── Merged Cells ────────────────────────────────────────────────────────
      _loadMergedCells(worksheet, wsRoot) {
        const mergeCellsNode = wsRoot.mergeCells;
        if (!mergeCellsNode) {
          worksheet._mergedCells = [];
          return;
        }
        const merged = [];
        const cells = ensureArray5(mergeCellsNode.mergeCell);
        for (const mc of cells) {
          const ref = attrOpt3(mc, "ref");
          if (ref) {
            merged.push(ref.toUpperCase());
          }
        }
        worksheet._mergedCells = merged;
      }
      // ── Column Dimensions ───────────────────────────────────────────────────
      _loadColumnDimensions(worksheet, wsRoot) {
        const colsNode = wsRoot.cols;
        if (!colsNode) return;
        const wsAny = worksheet;
        if (!wsAny._columnWidths) wsAny._columnWidths = {};
        if (!wsAny._hiddenColumns) wsAny._hiddenColumns = /* @__PURE__ */ new Set();
        const cols = ensureArray5(colsNode.col);
        for (const colElem of cols) {
          const minVal = attrOpt3(colElem, "min");
          const maxVal = attrOpt3(colElem, "max");
          const widthVal = attrOpt3(colElem, "width");
          const hiddenVal = attrOpt3(colElem, "hidden");
          if (minVal === null || maxVal === null) continue;
          const minCol = Math.floor(parseFloat(minVal));
          const maxCol = Math.floor(parseFloat(maxVal));
          if (minCol < 1 || maxCol < minCol) continue;
          let width = null;
          if (widthVal !== null) {
            width = parseFloat(widthVal);
            if (isNaN(width) || width <= 0) width = null;
          }
          for (let colIdx = minCol; colIdx <= maxCol; colIdx++) {
            if (width !== null) {
              wsAny._columnWidths[colIdx] = width;
            }
            if (hiddenVal === "1" || hiddenVal === "true" || hiddenVal === "True") {
              wsAny._hiddenColumns.add(colIdx);
            }
          }
        }
      }
      // ── Row Heights ─────────────────────────────────────────────────────────
      _loadRowHeights(worksheet, wsRoot) {
        const wsAny = worksheet;
        if (!wsAny._rowHeights) wsAny._rowHeights = {};
        if (!wsAny._hiddenRows) wsAny._hiddenRows = /* @__PURE__ */ new Set();
        const sheetData = wsRoot.sheetData;
        if (!sheetData) return;
        const rows = ensureArray5(sheetData.row);
        for (const rowElem of rows) {
          const ht = attrOpt3(rowElem, "ht");
          const hiddenVal = attrOpt3(rowElem, "hidden");
          const rowNum = attrOpt3(rowElem, "r");
          if (ht === null && !(hiddenVal === "1" || hiddenVal === "true" || hiddenVal === "True")) {
            continue;
          }
          if (rowNum === null) continue;
          const rowIdx = parseInt(rowNum, 10);
          if (isNaN(rowIdx) || rowIdx < 1) continue;
          if (ht !== null) {
            const height = parseFloat(ht);
            if (!isNaN(height) && height > 0) {
              wsAny._rowHeights[rowIdx] = height;
            }
          }
          if (hiddenVal === "1" || hiddenVal === "true" || hiddenVal === "True") {
            wsAny._hiddenRows.add(rowIdx);
          }
        }
      }
      // ── Freeze Pane ─────────────────────────────────────────────────────────
      _loadFreezePane(worksheet, wsRoot) {
        const sheetViews = wsRoot.sheetViews;
        if (!sheetViews) return;
        const views = ensureArray5(sheetViews.sheetView || sheetViews);
        if (views.length === 0) return;
        let view = views[0];
        if (view.sheetView) {
          view = Array.isArray(view.sheetView) ? view.sheetView[0] : view.sheetView;
        }
        const pane = view.pane;
        if (!pane) return;
        const state = attrOpt3(pane, "state");
        if (state !== "frozen") return;
        const xSplit = parseInt(attr5(pane, "xSplit", "0"), 10);
        const ySplit = parseInt(attr5(pane, "ySplit", "0"), 10);
        if (xSplit > 0 || ySplit > 0) {
          worksheet.setFreezePane(ySplit, xSplit, ySplit, xSplit);
        }
      }
      /**
       * Loads manual page breaks from rowBreaks / colBreaks elements.
       * ECMA-376 Section 18.3.1.14 (colBreaks), 18.3.1.73 (rowBreaks).
       */
      _loadPageBreaks(worksheet, wsRoot) {
        const rowBreaks = wsRoot.rowBreaks;
        if (rowBreaks) {
          const brks = ensureArray5(rowBreaks.brk);
          for (const brk of brks) {
            const id = brk?.["@_id"];
            if (id != null) {
              worksheet.horizontalPageBreaks.add(parseInt(String(id), 10));
            }
          }
        }
        const colBreaks = wsRoot.colBreaks;
        if (colBreaks) {
          const brks = ensureArray5(colBreaks.brk);
          for (const brk of brks) {
            const id = brk?.["@_id"];
            if (id != null) {
              worksheet.verticalPageBreaks.add(parseInt(String(id), 10));
            }
          }
        }
      }
      // ── Apply Cell Style ────────────────────────────────────────────────────
      _applyCellStyle(cell, styleIdx) {
        let cellStyleKey = this._wb._loaderCellXfByIndex.get(styleIdx);
        if (!cellStyleKey) {
          for (const [keyStr, idx] of this._wb._loaderCellStyles) {
            if (idx === styleIdx) {
              cellStyleKey = JSON.parse(keyStr);
              break;
            }
          }
        }
        if (!cellStyleKey) return;
        cell._sourceStyleIdx = styleIdx;
        const [fontKey, fillKey, borderKey, numFmtKey, alignmentKey, protectionKey] = cellStyleKey;
        const fontData = this._wb._loaderFontStyles.get(fontKey);
        if (fontData) {
          cell.style.font.name = fontData.name;
          cell.style.font.size = fontData.size;
          cell.style.font.color = fontData.color;
          cell.style.font.bold = fontData.bold;
          cell.style.font.italic = fontData.italic;
          cell.style.font.underline = fontData.underline;
          cell.style.font.strikethrough = fontData.strikethrough;
        }
        const fillData = this._wb._loaderFillStyles.get(fillKey);
        if (fillData) {
          cell.style.fill.patternType = fillData.pattern_type;
          cell.style.fill.foregroundColor = fillData.fg_color;
          cell.style.fill.backgroundColor = fillData.bg_color;
        }
        const borderData = this._wb._loaderBorderStyles.get(borderKey);
        if (borderData) {
          cell.style.borders.top.lineStyle = borderData.top.style;
          cell.style.borders.top.color = borderData.top.color;
          cell.style.borders.bottom.lineStyle = borderData.bottom.style;
          cell.style.borders.bottom.color = borderData.bottom.color;
          cell.style.borders.left.lineStyle = borderData.left.style;
          cell.style.borders.left.color = borderData.left.color;
          cell.style.borders.right.lineStyle = borderData.right.style;
          cell.style.borders.right.color = borderData.right.color;
        }
        const numFmtCode = this._wb._loaderNumFormats.get(numFmtKey);
        if (numFmtCode !== void 0) {
          cell.style.numberFormat = numFmtCode;
        }
        const alignData = this._wb._loaderAlignmentStyles.get(alignmentKey);
        if (alignData) {
          cell.style.alignment.horizontal = alignData.horizontal;
          cell.style.alignment.vertical = alignData.vertical;
          cell.style.alignment.wrapText = alignData.wrap_text;
          cell.style.alignment.indent = alignData.indent;
          cell.style.alignment.textRotation = alignData.text_rotation;
          cell.style.alignment.shrinkToFit = alignData.shrink_to_fit;
          cell.style.alignment.readingOrder = alignData.reading_order;
          cell.style.alignment.relativeIndent = alignData.relative_indent;
        }
        const protData = this._wb._loaderProtectionStyles.get(protectionKey);
        if (protData) {
          cell.style.protection.locked = protData.locked;
          cell.style.protection.hidden = protData.hidden;
        }
      }
      // ── Document Properties ─────────────────────────────────────────────────
      async _loadDocumentProperties(zip) {
        await this._loadCoreProperties(zip);
        await this._loadAppProperties(zip);
      }
      async _loadCoreProperties(zip) {
        const file = zip.file("docProps/core.xml");
        if (!file) return;
        try {
          const buf = await file.async("nodebuffer");
          const coreParser = new fastXmlParser.XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            textNodeName: "#text",
            removeNSPrefix: true,
            parseTagValue: false,
            trimValues: false
          });
          const root = coreParser.parse(buf.toString("utf-8"));
          const coreProps = root.coreProperties || root["cp:coreProperties"] || root;
          const docProps = this._workbook.documentProperties;
          if (coreProps.title) {
            docProps.title = typeof coreProps.title === "string" ? coreProps.title : coreProps.title["#text"] || "";
          }
          if (coreProps.subject) {
            docProps.subject = typeof coreProps.subject === "string" ? coreProps.subject : coreProps.subject["#text"] || "";
          }
          if (coreProps.creator) {
            docProps.creator = typeof coreProps.creator === "string" ? coreProps.creator : coreProps.creator["#text"] || "";
          }
          if (coreProps.description) {
            docProps.description = typeof coreProps.description === "string" ? coreProps.description : coreProps.description["#text"] || "";
          }
          if (coreProps.keywords) {
            docProps.keywords = typeof coreProps.keywords === "string" ? coreProps.keywords : coreProps.keywords["#text"] || "";
          }
          if (coreProps.lastModifiedBy) {
            docProps.lastModifiedBy = typeof coreProps.lastModifiedBy === "string" ? coreProps.lastModifiedBy : coreProps.lastModifiedBy["#text"] || "";
          }
          if (coreProps.revision) {
            docProps.revision = typeof coreProps.revision === "string" ? coreProps.revision : coreProps.revision["#text"] || "";
          }
          if (coreProps.category) {
            docProps.category = typeof coreProps.category === "string" ? coreProps.category : coreProps.category["#text"] || "";
          }
          if (coreProps.created) {
            const dateStr = typeof coreProps.created === "string" ? coreProps.created : coreProps.created["#text"] || "";
            if (dateStr) docProps.created = this._parseDatetime(dateStr);
          }
          if (coreProps.modified) {
            const dateStr = typeof coreProps.modified === "string" ? coreProps.modified : coreProps.modified["#text"] || "";
            if (dateStr) docProps.modified = this._parseDatetime(dateStr);
          }
        } catch {
        }
      }
      async _loadAppProperties(zip) {
        const file = zip.file("docProps/app.xml");
        if (!file) return;
        try {
          const buf = await file.async("nodebuffer");
          const appParser = new fastXmlParser.XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            textNodeName: "#text",
            removeNSPrefix: true,
            parseTagValue: false,
            trimValues: false
          });
          const root = appParser.parse(buf.toString("utf-8"));
          const appProps = root.Properties || root;
          const docProps = this._workbook.documentProperties;
          if (appProps.Application) {
            docProps.application = typeof appProps.Application === "string" ? appProps.Application : appProps.Application["#text"] || "";
          }
          if (appProps.AppVersion) {
            docProps.appVersion = typeof appProps.AppVersion === "string" ? appProps.AppVersion : appProps.AppVersion["#text"] || "";
          }
          if (appProps.Company) {
            docProps.company = typeof appProps.Company === "string" ? appProps.Company : appProps.Company["#text"] || "";
          }
          if (appProps.Manager) {
            docProps.manager = typeof appProps.Manager === "string" ? appProps.Manager : appProps.Manager["#text"] || "";
          }
        } catch {
        }
      }
      _parseDatetime(dateStr) {
        try {
          const d = new Date(dateStr);
          if (isNaN(d.getTime())) return dateStr;
          return d;
        } catch {
          return dateStr;
        }
      }
    };
  }
});

// src/core/Workbook.ts
var Workbook_exports = {};
__export(Workbook_exports, {
  DocumentProperties: () => exports.DocumentProperties,
  SaveFormat: () => exports.SaveFormat,
  Workbook: () => exports.Workbook,
  WorkbookProperties: () => exports.WorkbookProperties,
  WorkbookProtection: () => exports.WorkbookProtection,
  saveFormatFromExtension: () => saveFormatFromExtension,
  save_format_from_extension: () => exports.save_format_from_extension
});
function saveFormatFromExtension(filePath) {
  const dotIdx = filePath.lastIndexOf(".");
  if (dotIdx === -1) {
    throw new Error(`Unsupported file extension: (none)`);
  }
  const ext = filePath.slice(dotIdx).toLowerCase();
  const fmt = EXTENSION_FORMAT_MAP[ext];
  if (fmt === void 0) {
    throw new Error(`Unsupported file extension: ${ext}`);
  }
  return fmt;
}
exports.SaveFormat = void 0; var EXTENSION_FORMAT_MAP; exports.save_format_from_extension = void 0; exports.WorkbookProtection = void 0; exports.WorkbookProperties = void 0; exports.DocumentProperties = void 0; exports.Workbook = void 0;
var init_Workbook = __esm({
  "src/core/Workbook.ts"() {
    init_Worksheet();
    init_Style();
    init_XmlSaver();
    init_DefinedName();
    exports.SaveFormat = /* @__PURE__ */ ((SaveFormat2) => {
      SaveFormat2["AUTO"] = "auto";
      SaveFormat2["XLSX"] = "xlsx";
      SaveFormat2["CSV"] = "csv";
      SaveFormat2["TSV"] = "tsv";
      SaveFormat2["MARKDOWN"] = "markdown";
      SaveFormat2["JSON"] = "json";
      return SaveFormat2;
    })(exports.SaveFormat || {});
    EXTENSION_FORMAT_MAP = {
      ".xlsx": "xlsx" /* XLSX */,
      ".xlsm": "xlsx" /* XLSX */,
      ".csv": "csv" /* CSV */,
      ".tsv": "tsv" /* TSV */,
      ".md": "markdown" /* MARKDOWN */,
      ".markdown": "markdown" /* MARKDOWN */,
      ".json": "json" /* JSON */
    };
    exports.save_format_from_extension = saveFormatFromExtension;
    exports.WorkbookProtection = class {
      constructor() {
        this.lockStructure = false;
        this.lockWindows = false;
        this.workbookPassword = null;
      }
      /* .NET aliases */
      get lock_structure() {
        return this.lockStructure;
      }
      set lock_structure(v) {
        this.lockStructure = v;
      }
      get lock_windows() {
        return this.lockWindows;
      }
      set lock_windows(v) {
        this.lockWindows = v;
      }
      get workbook_password() {
        return this.workbookPassword;
      }
      set workbook_password(v) {
        this.workbookPassword = v;
      }
    };
    exports.WorkbookProperties = class {
      constructor() {
        /** Workbook protection settings */
        this.protection = new exports.WorkbookProtection();
        /** Calculation mode */
        this.calcMode = "auto";
        /** Defined names collection */
        this.definedNames = new exports.DefinedNameCollection();
        const viewObj = { activeTab: 0 };
        Object.defineProperty(viewObj, "active_tab", {
          get() {
            return this.activeTab;
          },
          set(v) {
            this.activeTab = v;
          },
          enumerable: true,
          configurable: true
        });
        this._view = viewObj;
      }
      get view() {
        return this._view;
      }
      /* .NET aliases */
      get calc_mode() {
        return this.calcMode;
      }
      set calc_mode(v) {
        this.calcMode = v;
      }
      get defined_names() {
        return this.definedNames;
      }
    };
    exports.DocumentProperties = class {
      constructor() {
        this.title = "";
        this.subject = "";
        this.creator = "";
        this.keywords = "";
        this.description = "";
        this.lastModifiedBy = "";
        this.category = "";
        this.revision = "";
        this.created = null;
        this.modified = null;
        /** Extended / App properties */
        this.application = "";
        this.appVersion = "";
        this.company = "";
        this.manager = "";
        /** Round-trip XML blobs */
        this._coreXml = null;
        this._appXml = null;
      }
      /* .NET aliases */
      get last_modified_by() {
        return this.lastModifiedBy;
      }
      set last_modified_by(v) {
        this.lastModifiedBy = v;
      }
    };
    exports.Workbook = class {
      // -------------------------------------------------------------------------
      // Constructor
      // -------------------------------------------------------------------------
      /**
       * Creates a new Workbook.
       *
       * @param filePath - Reserved for Phase 2/3: path to an existing .xlsx to load.
       * @param password - Reserved for Phase 2/3: password for encrypted files.
       */
      constructor(filePath, password) {
        // -- internal state -------------------------------------------------------
        this._worksheets = [];
        this._styles = [];
        this._sharedStrings = [];
        this._filePath = null;
        /** Document properties (lazily created) */
        this._documentProperties = null;
        // Style management maps (used by XML saver in Phase 2/3)
        /** @internal */
        this._fontStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._fillStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._borderStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._alignmentStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._protectionStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._cellStyles = /* @__PURE__ */ new Map();
        /** @internal */
        this._numFormats = /* @__PURE__ */ new Map();
        /** Round-trip: raw workbook-level XML blobs */
        this._sourceXml = {};
        this._filePath = filePath ?? null;
        this._properties = new exports.WorkbookProperties();
        this._styles.push(new exports.Style());
        const ws = new exports.Worksheet("Sheet1");
        ws._workbook = this;
        this._worksheets.push(ws);
      }
      // -------------------------------------------------------------------------
      // Properties
      // -------------------------------------------------------------------------
      /** Gets collection of worksheets in the workbook. */
      get worksheets() {
        return this._worksheets;
      }
      /** Gets the file path of the workbook (null if not loaded from file). */
      get filePath() {
        return this._filePath;
      }
      /** Gets workbook properties (protection, view, calc settings). */
      get properties() {
        return this._properties;
      }
      /** Gets document properties (title, author, etc.). Lazily created. */
      get documentProperties() {
        if (this._documentProperties === null) {
          this._documentProperties = new exports.DocumentProperties();
        }
        return this._documentProperties;
      }
      /** Gets the internal styles array. */
      get styles() {
        return this._styles;
      }
      /** Gets the shared strings table. */
      get sharedStrings() {
        return this._sharedStrings;
      }
      // .NET aliases
      get file_path() {
        return this.filePath;
      }
      get document_properties() {
        return this.documentProperties;
      }
      get shared_strings() {
        return this.sharedStrings;
      }
      // -------------------------------------------------------------------------
      // Worksheet management
      // -------------------------------------------------------------------------
      /**
       * Adds a new worksheet to the workbook.
       *
       * @param name - Optional name. Auto-generated as "SheetN" if omitted.
       * @returns The newly created Worksheet.
       */
      addWorksheet(name) {
        if (name === void 0 || name === null) {
          const existing = new Set(this._worksheets.map((ws2) => ws2.name));
          let i = 1;
          while (existing.has(`Sheet${i}`)) {
            i++;
          }
          name = `Sheet${i}`;
        }
        const ws = new exports.Worksheet(name);
        ws._workbook = this;
        this._worksheets.push(ws);
        return ws;
      }
      /** Alias for {@link addWorksheet}. */
      add_worksheet(name) {
        return this.addWorksheet(name);
      }
      /** Alias for {@link addWorksheet}. */
      createWorksheet(name) {
        return this.addWorksheet(name);
      }
      /** Alias for {@link addWorksheet}. */
      create_worksheet(name) {
        return this.addWorksheet(name);
      }
      /**
       * Gets a worksheet by index (0-based) or by name.
       *
       * @throws {RangeError} If numeric index is out of range.
       * @throws {Error} If no worksheet with the given name exists.
       */
      getWorksheet(indexOrName) {
        if (typeof indexOrName === "number") {
          if (indexOrName >= 0 && indexOrName < this._worksheets.length) {
            return this._worksheets[indexOrName];
          }
          throw new RangeError(`Worksheet index ${indexOrName} out of range`);
        }
        for (const ws of this._worksheets) {
          if (ws.name === indexOrName) {
            return ws;
          }
        }
        throw new Error(`Worksheet '${indexOrName}' not found`);
      }
      /** Alias */
      get_worksheet(indexOrName) {
        return this.getWorksheet(indexOrName);
      }
      /**
       * Returns the worksheet with the given name, or `null` if not found.
       */
      getWorksheetByName(name) {
        for (const ws of this._worksheets) {
          if (ws.name === name) {
            return ws;
          }
        }
        return null;
      }
      /** Alias */
      get_worksheet_by_name(name) {
        return this.getWorksheetByName(name);
      }
      /**
       * Returns the worksheet at the given 0-based index, or `null` if out of range.
       */
      getWorksheetByIndex(index) {
        if (Number.isInteger(index) && index >= 0 && index < this._worksheets.length) {
          return this._worksheets[index];
        }
        return null;
      }
      /** Alias */
      get_worksheet_by_index(index) {
        return this.getWorksheetByIndex(index);
      }
      /**
       * Removes a worksheet by index, name, or direct reference.
       *
       * @throws {RangeError} If numeric index is out of range.
       * @throws {Error} If string name is not found.
       * @throws {TypeError} If argument is none of the above.
       */
      removeWorksheet(indexOrNameOrWs) {
        if (typeof indexOrNameOrWs === "number") {
          if (indexOrNameOrWs >= 0 && indexOrNameOrWs < this._worksheets.length) {
            this._worksheets.splice(indexOrNameOrWs, 1);
            return;
          }
          throw new RangeError(`Worksheet index ${indexOrNameOrWs} out of range`);
        }
        if (typeof indexOrNameOrWs === "string") {
          const idx2 = this._worksheets.findIndex((ws) => ws.name === indexOrNameOrWs);
          if (idx2 !== -1) {
            this._worksheets.splice(idx2, 1);
            return;
          }
          throw new Error(`Worksheet '${indexOrNameOrWs}' not found`);
        }
        const idx = this._worksheets.findIndex((ws) => ws === indexOrNameOrWs);
        if (idx !== -1) {
          this._worksheets.splice(idx, 1);
          return;
        }
        throw new TypeError("removeWorksheet: argument must be number, string, or Worksheet");
      }
      /** Alias */
      remove_worksheet(indexOrNameOrWs) {
        this.removeWorksheet(indexOrNameOrWs);
      }
      /**
       * Returns the currently active worksheet.
       */
      getActiveWorksheet() {
        if (this._worksheets.length === 0) return null;
        let idx = this._properties.view.activeTab;
        idx = Math.max(0, Math.min(idx, this._worksheets.length - 1));
        return this._worksheets[idx];
      }
      /** Alias */
      get_active_worksheet() {
        return this.getActiveWorksheet();
      }
      /**
       * Sets the active worksheet by index, name, or Worksheet reference.
       */
      setActiveWorksheet(indexOrNameOrWs) {
        let idx;
        if (typeof indexOrNameOrWs === "number") {
          if (indexOrNameOrWs >= 0 && indexOrNameOrWs < this._worksheets.length) {
            idx = indexOrNameOrWs;
          }
        } else if (typeof indexOrNameOrWs === "string") {
          idx = this._worksheets.findIndex((ws) => ws.name === indexOrNameOrWs);
          if (idx === -1) idx = void 0;
        } else {
          idx = this._worksheets.findIndex((ws) => ws === indexOrNameOrWs);
          if (idx === -1) idx = void 0;
        }
        if (idx !== void 0) {
          this._properties.view.activeTab = idx;
        }
      }
      /** Alias */
      set_active_worksheet(indexOrNameOrWs) {
        this.setActiveWorksheet(indexOrNameOrWs);
      }
      /**
       * Copies a worksheet and appends the copy to the workbook.
       *
       * @returns The new worksheet, or `null` if source could not be resolved.
       */
      copyWorksheet(indexOrNameOrWs) {
        let ws;
        if (typeof indexOrNameOrWs === "number") {
          ws = this.getWorksheetByIndex(indexOrNameOrWs);
        } else if (typeof indexOrNameOrWs === "string") {
          ws = this.getWorksheetByName(indexOrNameOrWs);
        } else if (indexOrNameOrWs instanceof exports.Worksheet) {
          ws = indexOrNameOrWs;
        }
        if (!ws) return null;
        const base = ws.name;
        const existing = new Set(this._worksheets.map((w) => w.name));
        let copyName = `${base} (copy)`;
        if (existing.has(copyName)) {
          let i = 1;
          while (existing.has(`${base} (copy${i})`)) {
            i++;
          }
          copyName = `${base} (copy${i})`;
        }
        const newWs = ws.copy(copyName);
        newWs._workbook = this;
        this._worksheets.push(newWs);
        return newWs;
      }
      /** Alias */
      copy_worksheet(indexOrNameOrWs) {
        return this.copyWorksheet(indexOrNameOrWs);
      }
      // -------------------------------------------------------------------------
      // Protection
      // -------------------------------------------------------------------------
      /**
       * Protects the workbook structure/windows with an optional password.
       *
       * @param password - Optional password string.
       * @param lockStructure - Prevent adding/removing sheets (default true).
       * @param lockWindows - Prevent resizing windows (default false).
       */
      protect(password, lockStructure = true, lockWindows = false) {
        const prot = this._properties.protection;
        prot.lockStructure = lockStructure;
        prot.lockWindows = lockWindows;
        prot.workbookPassword = password ?? null;
      }
      /**
       * Removes workbook structure/window protection.
       */
      unprotect(_password) {
        const prot = this._properties.protection;
        prot.lockStructure = false;
        prot.lockWindows = false;
        prot.workbookPassword = null;
      }
      /**
       * Returns `true` if the workbook has structure or window protection enabled.
       */
      isProtected() {
        const prot = this._properties.protection;
        return prot.lockStructure || prot.lockWindows;
      }
      /** Alias */
      is_protected() {
        return this.isProtected();
      }
      /**
       * Returns a snapshot of the current workbook protection settings.
       */
      get protection() {
        const prot = this._properties.protection;
        return {
          lockStructure: prot.lockStructure,
          lockWindows: prot.lockWindows,
          password: prot.workbookPassword
        };
      }
      // -------------------------------------------------------------------------
      // Formula calculation (Phase 5)
      // -------------------------------------------------------------------------
      /**
       * Evaluates every formula in every worksheet of this workbook.
       *
       * Uses a lazy `require` to avoid circular-dependency issues between
       * Workbook and FormulaEvaluator.
       */
      calculateFormula() {
        const { FormulaEvaluator: FormulaEvaluator2 } = (init_FormulaEvaluator(), __toCommonJS(FormulaEvaluator_exports));
        const evaluator = new FormulaEvaluator2(this);
        evaluator.evaluateAll();
      }
      /** snake_case alias */
      calculate_formula() {
        this.calculateFormula();
      }
      // -------------------------------------------------------------------------
      // File I/O stubs (Phase 2/3)
      // -------------------------------------------------------------------------
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
      async save(filePath, saveFormat, _options, _password) {
        const fmt = saveFormat && saveFormat !== "auto" /* AUTO */ ? saveFormat : saveFormatFromExtension(filePath);
        if (fmt !== "xlsx" /* XLSX */) {
          throw new Error(
            `Workbook.save(): format '${fmt}' is not yet supported. Only XLSX is implemented.`
          );
        }
        const saver = new exports.XmlSaver(this);
        await saver.saveToFile(filePath);
      }
      /**
       * Saves the workbook to an in-memory Buffer (XLSX format).
       *
       * @param saveFormat  Explicit format; defaults to XLSX.
       * @returns A Buffer containing the XLSX ZIP archive.
       */
      async saveToBuffer(saveFormat = "xlsx" /* XLSX */) {
        if (saveFormat !== "xlsx" /* XLSX */ && saveFormat !== "auto" /* AUTO */) {
          throw new Error(
            `Workbook.saveToBuffer(): format '${saveFormat}' is not yet supported. Only XLSX is implemented.`
          );
        }
        const saver = new exports.XmlSaver(this);
        return saver.saveToBuffer();
      }
      /** snake_case alias */
      async save_to_buffer(saveFormat = "xlsx" /* XLSX */) {
        return this.saveToBuffer(saveFormat);
      }
      /**
       * Loads a workbook from a file path (async).
       *
       * @param filePath Path to the .xlsx file.
       * @param _password Reserved for future encrypted workbook support.
       * @returns A Promise resolving to the loaded Workbook.
       */
      static async load(filePath, _password) {
        const { XmlLoader: XmlLoader2 } = await Promise.resolve().then(() => (init_XmlLoader(), XmlLoader_exports));
        return XmlLoader2.loadFromFile(filePath);
      }
      /**
       * Loads a workbook from a Buffer (async).
       *
       * @param buffer Buffer containing XLSX data.
       * @returns A Promise resolving to the loaded Workbook.
       */
      static async loadFromBuffer(buffer) {
        const { XmlLoader: XmlLoader2 } = await Promise.resolve().then(() => (init_XmlLoader(), XmlLoader_exports));
        return XmlLoader2.loadFromBuffer(buffer);
      }
      // -------------------------------------------------------------------------
      // String representation
      // -------------------------------------------------------------------------
      toString() {
        return `Workbook(worksheets=${this._worksheets.length})`;
      }
    };
  }
});

// src/index.ts
init_Style();
init_Cell();
init_Cells();
init_Worksheet();
init_Workbook();
init_DataValidation();
init_ConditionalFormat();
init_Hyperlink();
init_AutoFilter();
init_PageBreak();
init_DefinedName();
init_FormulaEvaluator();
init_CommentXml();
init_XmlAutoFilterLoader();
init_XmlAutoFilterSaver();
init_XmlConditionalFormatLoader();
init_XmlConditionalFormatSaver();
init_XmlDataValidationLoader();
init_XmlDataValidationSaver();
init_XmlHyperlinkHandler();
init_SharedStrings();
init_CellValueHandler();
init_XmlSaver();
init_XmlLoader();

exports.saveFormatFromExtension = saveFormatFromExtension;
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map