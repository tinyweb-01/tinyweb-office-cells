/**
 * tinyweb-office-cells - Page Break Module
 *
 * Manual page break collections.
 *
 * @module features/PageBreak
 */

import { Cells } from '../core/Cells';

// ─── HorizontalPageBreakCollection ────────────────────────────────────────

/**
 * Collection of manual horizontal page breaks (row breaks).
 */
export class HorizontalPageBreakCollection {
  /** @internal */
  _breaks: Set<number> = new Set();

  private _normalizeRow(rowOrCell: number | string): number {
    if (typeof rowOrCell === 'string') {
      const [row] = Cells.coordinateFromString(rowOrCell);
      return row - 1; // convert 1-based to 0-based
    }
    if (rowOrCell == null || rowOrCell < 0) {
      throw new Error('row must be >= 0');
    }
    return rowOrCell;
  }

  /**
   * Adds a manual horizontal page break.
   * @param rowOrCell 0-based row index or A1 cell reference.
   * @returns The 0-based row index.
   */
  add(rowOrCell: number | string): number {
    const row = this._normalizeRow(rowOrCell);
    this._breaks.add(row);
    return row;
  }

  /**
   * Removes the break at zero-based collection index.
   */
  removeAt(index: number): void {
    const list = this.toList();
    if (index < 0 || index >= list.length) {
      throw new RangeError(`Index ${index} out of range`);
    }
    this._breaks.delete(list[index]);
  }

  /**
   * Removes a manual horizontal page break by row/cell.
   */
  remove(rowOrCell: number | string): void {
    const row = this._normalizeRow(rowOrCell);
    this._breaks.delete(row);
  }

  /** Clears all manual horizontal page breaks. */
  clear(): void {
    this._breaks.clear();
  }

  get count(): number {
    return this._breaks.size;
  }

  /** Returns a sorted array of break rows. */
  toList(): number[] {
    return Array.from(this._breaks).sort((a, b) => a - b);
  }

  /** Alias */
  to_list(): number[] { return this.toList(); }

  get(index: number): number {
    return this.toList()[index];
  }

  [Symbol.iterator](): Iterator<number> {
    return this.toList()[Symbol.iterator]();
  }
}

// ─── VerticalPageBreakCollection ──────────────────────────────────────────

/**
 * Collection of manual vertical page breaks (column breaks).
 */
export class VerticalPageBreakCollection {
  /** @internal */
  _breaks: Set<number> = new Set();

  private _normalizeColumn(columnOrCell: number | string): number {
    if (typeof columnOrCell === 'string') {
      // Check if it contains digits -> cell reference like "B5"
      if (/\d/.test(columnOrCell)) {
        const [, col] = Cells.coordinateFromString(columnOrCell);
        return col - 1; // convert 1-based to 0-based
      }
      // Pure column letters like "B"
      return Cells.columnIndexFromString(columnOrCell) - 1;
    }
    if (columnOrCell == null || columnOrCell < 0) {
      throw new Error('column must be >= 0');
    }
    return columnOrCell;
  }

  /**
   * Adds a manual vertical page break.
   * @param columnOrCell 0-based column index, column letters, or A1 cell reference.
   * @returns The 0-based column index.
   */
  add(columnOrCell: number | string): number {
    const col = this._normalizeColumn(columnOrCell);
    this._breaks.add(col);
    return col;
  }

  /**
   * Removes a manual vertical page break by column/cell.
   */
  remove(columnOrCell: number | string): void {
    const col = this._normalizeColumn(columnOrCell);
    this._breaks.delete(col);
  }

  /**
   * Removes the break at zero-based collection index.
   */
  removeAt(index: number): void {
    const list = this.toList();
    if (index < 0 || index >= list.length) {
      throw new RangeError(`Index ${index} out of range`);
    }
    this._breaks.delete(list[index]);
  }

  /** Clears all manual vertical page breaks. */
  clear(): void {
    this._breaks.clear();
  }

  get count(): number {
    return this._breaks.size;
  }

  /** Returns a sorted array of break columns. */
  toList(): number[] {
    return Array.from(this._breaks).sort((a, b) => a - b);
  }

  /** Alias */
  to_list(): number[] { return this.toList(); }

  get(index: number): number {
    return this.toList()[index];
  }

  [Symbol.iterator](): Iterator<number> {
    return this.toList()[Symbol.iterator]();
  }
}
