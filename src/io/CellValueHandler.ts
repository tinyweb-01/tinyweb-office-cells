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

import type { CellValue } from '../core/Cell';

// ── Constants ───────────────────────────────────────────────────────────────

/** Excel cell type constants matching ST_CellType */
export const CELL_TYPE_STRING = 's';
export const CELL_TYPE_INLINE_STRING = 'str';
export const CELL_TYPE_NUMBER = 'n';
export const CELL_TYPE_BOOLEAN = 'b';
export const CELL_TYPE_ERROR = 'e';

/** Standard Excel error values */
export const ERROR_VALUES = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
]);

/**
 * Excel date epoch: 1899-12-30T00:00:00.000Z
 *
 * Excel uses a serial date system where day 1 = 1900-01-01.
 * Due to the intentional Lotus 1-2-3 leap year bug (1900 is treated as a
 * leap year), the epoch is effectively 1899-12-30.
 */
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
const MS_PER_DAY = 86_400_000;

// ── Public API ──────────────────────────────────────────────────────────────

export class CellValueHandler {
  /**
   * Determines the OOXML cell type for a JavaScript value.
   *
   * @param value - A cell value (string, number, boolean, Date, null, etc.)
   * @returns The OOXML cell type string, or `null` for empty cells.
   */
  static getCellType(value: CellValue): string | null {
    if (value === null || value === undefined) {
      return null;
    }

    if (typeof value === 'boolean') {
      return CELL_TYPE_BOOLEAN;
    }

    if (typeof value === 'number') {
      return CELL_TYPE_NUMBER;
    }

    if (value instanceof Date) {
      return CELL_TYPE_NUMBER; // Dates become serial numbers
    }

    if (typeof value === 'string') {
      if (value === '') return CELL_TYPE_STRING;
      if (ERROR_VALUES.has(value.toUpperCase())) return CELL_TYPE_ERROR;
      return CELL_TYPE_STRING;
    }

    // Fallback: convert to string
    return CELL_TYPE_STRING;
  }

  /**
   * Formats a cell value for writing into OOXML `<v>` element.
   *
   * @param value - The cell value.
   * @param cellType - Optional pre-determined cell type; auto-detected if omitted.
   * @returns A tuple of `[formattedString, cellType]`, or `[null, null]` for empty.
   */
  static formatValueForXml(
    value: CellValue,
    cellType?: string | null,
  ): [string | null, string | null] {
    if (value === null || value === undefined) {
      return [null, null];
    }

    const resolvedType = cellType ?? this.getCellType(value);
    if (resolvedType === null) {
      return [null, null];
    }

    switch (resolvedType) {
      case CELL_TYPE_BOOLEAN:
        return [value ? '1' : '0', CELL_TYPE_BOOLEAN];

      case CELL_TYPE_NUMBER: {
        if (value instanceof Date) {
          const serial = this.dateToExcelSerial(value);
          return [String(serial), CELL_TYPE_NUMBER];
        }
        if (typeof value === 'number') {
          if (!isFinite(value)) {
            return ['#NUM!', CELL_TYPE_ERROR];
          }
          // Use full precision; avoid scientific notation for moderate values
          return [numberToString(value), CELL_TYPE_NUMBER];
        }
        // Numeric string fallthrough
        return [String(value), CELL_TYPE_NUMBER];
      }

      case CELL_TYPE_ERROR:
        return [String(value).toUpperCase(), CELL_TYPE_ERROR];

      case CELL_TYPE_STRING:
      case CELL_TYPE_INLINE_STRING:
      default:
        return [String(value), CELL_TYPE_STRING];
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
  static parseValueFromXml(
    valueStr: string | null | undefined,
    cellType: string | null,
    sharedStrings?: string[],
  ): CellValue {
    if (valueStr === null || valueStr === undefined) {
      return null;
    }

    switch (cellType) {
      case CELL_TYPE_BOOLEAN:
        return valueStr === '1' || valueStr.toLowerCase() === 'true';

      case CELL_TYPE_STRING: {
        // Shared string index
        const idx = parseInt(valueStr, 10);
        if (sharedStrings && !isNaN(idx) && idx >= 0 && idx < sharedStrings.length) {
          return sharedStrings[idx];
        }
        return valueStr;
      }

      case CELL_TYPE_INLINE_STRING:
        return valueStr;

      case CELL_TYPE_ERROR:
        return valueStr;

      case CELL_TYPE_NUMBER: {
        const num = parseFloat(valueStr);
        if (isNaN(num)) return valueStr;
        return num;
      }

      default: {
        // No type specified – try number, then fall back to string
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
  static dateToExcelSerial(dt: Date): number {
    const utcMs =
      Date.UTC(
        dt.getFullYear(),
        dt.getMonth(),
        dt.getDate(),
        dt.getHours(),
        dt.getMinutes(),
        dt.getSeconds(),
        dt.getMilliseconds(),
      ) - EXCEL_EPOCH.getTime();

    let serial = utcMs / MS_PER_DAY;

    // Round to avoid floating-point noise (max 10 decimal digits)
    serial = Math.round(serial * 1e10) / 1e10;

    return serial;
  }

  /**
   * Converts an Excel serial date number to a JavaScript Date.
   *
   * @param serial - The Excel serial date number.
   * @returns A Date object.
   */
  static excelSerialToDate(serial: number): Date {
    const ms = serial * MS_PER_DAY + EXCEL_EPOCH.getTime();
    return new Date(ms);
  }

  // ── Error Detection ─────────────────────────────────────────────────────

  /**
   * Checks if a string value represents an Excel error.
   */
  static isErrorValue(value: string): boolean {
    return ERROR_VALUES.has(value.toUpperCase());
  }

  /**
   * Returns the normalized error string, or `null` if not an error.
   */
  static getErrorType(value: string): string | null {
    const upper = value.toUpperCase();
    if (ERROR_VALUES.has(upper)) return upper;
    return null;
  }
}

// ── Internal Helpers ────────────────────────────────────────────────────────

/**
 * Converts a number to a string suitable for OOXML `<v>` elements.
 * Avoids unnecessary scientific notation for moderate values while
 * preserving full precision.
 */
function numberToString(n: number): string {
  // toPrecision(15) matches Excel's 15-digit precision
  // but we prefer plain decimal notation
  const s = String(n);
  // JavaScript's String(n) already does a good job for most values
  return s;
}
