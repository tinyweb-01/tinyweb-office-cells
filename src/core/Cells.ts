/**
 * tinyweb-office-cells - Cells Module
 *
 * A1-keyed collection of cells in a worksheet.
 */

import { Cell, CellValue } from './Cell';

// Forward-declare to avoid circular import at value level
// Worksheet is only used for type-checking; the actual reference
// is injected at runtime via the constructor.
export interface WorksheetLike {
  _mergedCells: string[];
  _rowHeights: Record<number, number>;
  _columnWidths: Record<number, number>;
  _hiddenRows: Set<number>;
  _hiddenColumns: Set<number>;
}

/**
 * Represents a collection of cells in a worksheet.
 */
export class Cells {
  /** Internal cell storage keyed by A1 reference. */
  _cells: Map<string, Cell>;
  /** Back-reference to owning worksheet (set at runtime). */
  _worksheet: WorksheetLike | null;

  constructor(worksheet: WorksheetLike | null = null) {
    this._cells = new Map();
    this._worksheet = worksheet;
  }

  private requireWorksheet(): void {
    if (!this._worksheet) throw new Error('Cells is not attached to a Worksheet');
  }

  /** Normalize tuple-like [row, col] into A1 key. */
  private normalizeKey(key: string | [number, number | string]): string {
    if (Array.isArray(key)) {
      const [row, col] = key;
      const colIdx = typeof col === 'string' ? Cells.columnIndexFromString(col) : col;
      return Cells.coordinateToString(row, colIdx);
    }
    return key;
  }

  // ── Cell access ─────────────────────────────────────────────────────────

  /** Gets a cell by its A1 reference. Creates if not exists. */
  get(key: string | [number, number | string]): Cell {
    const k = this.normalizeKey(key);
    let cell = this._cells.get(k);
    if (!cell) {
      cell = new Cell();
      this._cells.set(k, cell);
    }
    return cell;
  }

  /** Sets a cell value by A1 reference. */
  set(key: string | [number, number | string], value: Cell | CellValue): void {
    const k = this.normalizeKey(key);
    if (value instanceof Cell) {
      this._cells.set(k, value);
    } else {
      const cell = this.get(k);
      cell.value = value;
    }
  }

  /** Access cell by row/column (1-based). */
  cell(row: number, column: number | string): Cell {
    const colIdx = typeof column === 'string' ? Cells.columnIndexFromString(column) : column;
    const ref = `${Cells.columnLetterFromIndex(colIdx)}${row}`;
    return this.get(ref);
  }

  // ── Coordinate conversion (static) ──────────────────────────────────────

  static columnIndexFromString(column: string): number {
    if (!column) throw new Error('Column string cannot be empty');
    const upper = column.toUpperCase();
    let result = 0;
    for (const ch of upper) {
      if (ch < 'A' || ch > 'Z') throw new Error(`Invalid column character: ${ch}`);
      result = result * 26 + (ch.charCodeAt(0) - 64);
    }
    return result;
  }
  /** snake_case alias */
  static column_index_from_string(column: string): number { return Cells.columnIndexFromString(column); }

  static columnLetterFromIndex(index: number): string {
    if (index < 1) throw new Error('Column index must be >= 1');
    let result = '';
    let i = index;
    while (i > 0) {
      i--;
      result = String.fromCharCode(65 + (i % 26)) + result;
      i = Math.floor(i / 26);
    }
    return result;
  }
  static column_letter_from_index(index: number): string { return Cells.columnLetterFromIndex(index); }

  static coordinateFromString(coord: string): [number, number] {
    if (!coord) throw new Error('Coordinate cannot be empty');
    let colStr = '';
    let rowStr = '';
    for (const ch of coord) {
      if (/[a-zA-Z]/.test(ch)) colStr += ch;
      else if (/\d/.test(ch)) rowStr += ch;
      else throw new Error(`Invalid character in coordinate: ${ch}`);
    }
    if (!colStr || !rowStr) throw new Error(`Invalid coordinate format: ${coord}`);
    return [parseInt(rowStr, 10), Cells.columnIndexFromString(colStr)];
  }
  static coordinate_from_string(coord: string): [number, number] { return Cells.coordinateFromString(coord); }

  static coordinateToString(row: number, column: number): string {
    if (row < 1 || column < 1) throw new Error('Row and column must be >= 1');
    return `${Cells.columnLetterFromIndex(column)}${row}`;
  }
  static coordinate_to_string(row: number, column: number): string { return Cells.coordinateToString(row, column); }

  // ── Iteration ───────────────────────────────────────────────────────────

  /** Iterates row-by-row, yielding arrays of Cell (or values). */
  *iterRows(options: { minRow?: number; maxRow?: number; minCol?: number; maxCol?: number; valuesOnly?: boolean } = {}): Generator<(Cell | CellValue)[]> {
    if (this._cells.size === 0) return;
    const rows = new Set<number>();
    const cols = new Set<number>();
    for (const ref of this._cells.keys()) {
      const [r, c] = Cells.coordinateFromString(ref);
      rows.add(r);
      cols.add(c);
    }
    const minRow = options.minRow ?? Math.min(...rows);
    const maxRow = options.maxRow ?? Math.max(...rows);
    const minCol = options.minCol ?? Math.min(...cols);
    const maxCol = options.maxCol ?? Math.max(...cols);

    for (let r = minRow; r <= maxRow; r++) {
      const rowCells: (Cell | CellValue)[] = [];
      for (let c = minCol; c <= maxCol; c++) {
        const cell = this.get(Cells.coordinateToString(r, c));
        rowCells.push(options.valuesOnly ? cell.value : cell);
      }
      yield rowCells;
    }
  }
  /** snake_case alias */
  *iter_rows(options: { minRow?: number; maxRow?: number; minCol?: number; maxCol?: number; valuesOnly?: boolean } = {}): Generator<(Cell | CellValue)[]> {
    yield* this.iterRows(options);
  }

  // ── Collection methods ──────────────────────────────────────────────────

  count(): number { return this._cells.size; }

  clear(): void { this._cells.clear(); }

  getCellByName(cellName: string): Cell { return this.get(cellName); }
  get_cell_by_name(cellName: string): Cell { return this.getCellByName(cellName); }

  setCellByName(cellName: string, value: CellValue): void { this.set(cellName, value); }
  set_cell_by_name(cellName: string, value: CellValue): void { this.setCellByName(cellName, value); }

  getCell(row: number, column: number): Cell { return this.cell(row, column); }
  get_cell(row: number, column: number): Cell { return this.getCell(row, column); }

  setCell(row: number, column: number, value: CellValue): void {
    const ref = Cells.coordinateToString(row, column);
    this.set(ref, value);
  }
  set_cell(row: number, column: number, value: CellValue): void { this.setCell(row, column, value); }

  // ── Row / Column dimensions ─────────────────────────────────────────────

  setRowHeight(row: number, height: number): void {
    this.requireWorksheet();
    if (row < 0) throw new Error('row must be >= 0');
    if (height <= 0) throw new Error('height must be > 0');
    this._worksheet!._rowHeights[row + 1] = height;
  }
  set_row_height(row: number, height: number): void { this.setRowHeight(row, height); }
  SetRowHeight(row: number, height: number): void { this.setRowHeight(row, height); }

  getRowHeight(row: number): number {
    this.requireWorksheet();
    if (row < 0) throw new Error('row must be >= 0');
    const h = this._worksheet!._rowHeights[row + 1];
    return h ?? 15; // Default Excel row height
  }
  get_row_height(row: number): number { return this.getRowHeight(row); }
  GetRowHeight(row: number): number { return this.getRowHeight(row); }

  setColumnWidth(column: number | string, width: number): void {
    this.requireWorksheet();
    let colIdx: number;
    if (typeof column === 'string') {
      colIdx = Cells.columnIndexFromString(column);
    } else {
      if (column < 0) throw new Error('column must be >= 0');
      colIdx = column + 1;
    }
    if (width <= 0) throw new Error('width must be > 0');
    this._worksheet!._columnWidths[colIdx] = width;
  }
  set_column_width(column: number | string, width: number): void { this.setColumnWidth(column, width); }
  SetColumnWidth(column: number | string, width: number): void { this.setColumnWidth(column, width); }

  getColumnWidth(column: number | string): number {
    this.requireWorksheet();
    let colIdx: number;
    if (typeof column === 'string') {
      colIdx = Cells.columnIndexFromString(column);
    } else {
      if (column < 0) throw new Error('column must be >= 0');
      colIdx = column + 1;
    }
    return this._worksheet!._columnWidths[colIdx] ?? 8.43; // Default Excel column width
  }
  get_column_width(column: number | string): number { return this.getColumnWidth(column); }
  GetColumnWidth(column: number | string): number { return this.getColumnWidth(column); }

  // ── Hide / unhide rows & columns ────────────────────────────────────────

  hideRow(row: number): void {
    this.requireWorksheet();
    if (row < 0) throw new Error('row must be >= 0');
    this._worksheet!._hiddenRows.add(row + 1);
  }
  hide_row(row: number): void { this.hideRow(row); }

  unhideRow(row: number): void {
    this.requireWorksheet();
    if (row < 0) throw new Error('row must be >= 0');
    this._worksheet!._hiddenRows.delete(row + 1);
  }
  unhide_row(row: number): void { this.unhideRow(row); }

  isRowHidden(row: number): boolean {
    this.requireWorksheet();
    if (row < 0) throw new Error('row must be >= 0');
    return this._worksheet!._hiddenRows.has(row + 1);
  }
  is_row_hidden(row: number): boolean { return this.isRowHidden(row); }
  IsRowHidden(row: number): boolean { return this.isRowHidden(row); }

  hideColumn(column: number | string): void {
    this.requireWorksheet();
    const colIdx = typeof column === 'string' ? Cells.columnIndexFromString(column) : column + 1;
    this._worksheet!._hiddenColumns.add(colIdx);
  }
  hide_column(column: number | string): void { this.hideColumn(column); }

  unhideColumn(column: number | string): void {
    this.requireWorksheet();
    const colIdx = typeof column === 'string' ? Cells.columnIndexFromString(column) : column + 1;
    this._worksheet!._hiddenColumns.delete(colIdx);
  }
  unhide_column(column: number | string): void { this.unhideColumn(column); }

  isColumnHidden(column: number | string): boolean {
    this.requireWorksheet();
    const colIdx = typeof column === 'string' ? Cells.columnIndexFromString(column) : column + 1;
    return this._worksheet!._hiddenColumns.has(colIdx);
  }
  is_column_hidden(column: number | string): boolean { return this.isColumnHidden(column); }
  IsColumnHidden(column: number | string): boolean { return this.isColumnHidden(column); }

  // ── Merge cells ─────────────────────────────────────────────────────────

  /** Merges a rectangular range. Row/column are 0-based. */
  merge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    this.requireWorksheet();
    if (firstRow < 0) throw new Error('firstRow must be >= 0');
    if (firstColumn < 0) throw new Error('firstColumn must be >= 0');
    if (totalRows < 1) throw new Error('totalRows must be >= 1');
    if (totalColumns < 1) throw new Error('totalColumns must be >= 1');

    const startRef = Cells.coordinateToString(firstRow + 1, firstColumn + 1);
    const endRef = Cells.coordinateToString(firstRow + totalRows, firstColumn + totalColumns);
    const mergeRef = (startRef !== endRef ? `${startRef}:${endRef}` : startRef).toUpperCase();

    if (!this._worksheet!._mergedCells.includes(mergeRef)) {
      this._worksheet!._mergedCells.push(mergeRef);
    }
  }
  /** .NET alias */
  Merge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    this.merge(firstRow, firstColumn, totalRows, totalColumns);
  }

  /** Unmerges a previously merged rectangular range. */
  unmerge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    this.requireWorksheet();
    if (firstRow < 0) throw new Error('firstRow must be >= 0');
    if (firstColumn < 0) throw new Error('firstColumn must be >= 0');
    if (totalRows < 1) throw new Error('totalRows must be >= 1');
    if (totalColumns < 1) throw new Error('totalColumns must be >= 1');

    const startRef = Cells.coordinateToString(firstRow + 1, firstColumn + 1);
    const endRef = Cells.coordinateToString(firstRow + totalRows, firstColumn + totalColumns);
    const mergeRef = (startRef !== endRef ? `${startRef}:${endRef}` : startRef).toUpperCase();

    const idx = this._worksheet!._mergedCells.indexOf(mergeRef);
    if (idx !== -1) this._worksheet!._mergedCells.splice(idx, 1);
  }
  UnMerge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    this.unmerge(firstRow, firstColumn, totalRows, totalColumns);
  }

  mergeRange(rangeRef: string): void {
    this.requireWorksheet();
    if (!rangeRef) throw new Error('rangeRef must be a non-empty string');
    const ref = rangeRef.toUpperCase();
    if (!this._worksheet!._mergedCells.includes(ref)) {
      this._worksheet!._mergedCells.push(ref);
    }
  }
  merge_range(rangeRef: string): void { this.mergeRange(rangeRef); }

  unmergeRange(rangeRef: string): void {
    this.requireWorksheet();
    if (!rangeRef) throw new Error('rangeRef must be a non-empty string');
    const ref = rangeRef.toUpperCase();
    const idx = this._worksheet!._mergedCells.indexOf(ref);
    if (idx !== -1) this._worksheet!._mergedCells.splice(idx, 1);
  }
  unmerge_range(rangeRef: string): void { this.unmergeRange(rangeRef); }

  getMergedCells(): string[] {
    this.requireWorksheet();
    return [...this._worksheet!._mergedCells];
  }
  get_merged_cells(): string[] { return this.getMergedCells(); }

  // ── Range methods ───────────────────────────────────────────────────────

  getRange(startRow: number, startColumn: number, endRow: number, endColumn: number): Cell[][] {
    const result: Cell[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const rowCells: Cell[] = [];
      for (let c = startColumn; c <= endColumn; c++) {
        rowCells.push(this.get(Cells.coordinateToString(r, c)));
      }
      result.push(rowCells);
    }
    return result;
  }
  get_range(startRow: number, startColumn: number, endRow: number, endColumn: number): Cell[][] {
    return this.getRange(startRow, startColumn, endRow, endColumn);
  }

  // ── Utility ─────────────────────────────────────────────────────────────

  hasCell(cellName: string): boolean { return this._cells.has(cellName); }
  has_cell(cellName: string): boolean { return this.hasCell(cellName); }

  deleteCell(cellName: string): void { this._cells.delete(cellName); }
  delete_cell(cellName: string): void { this.deleteCell(cellName); }

  getAllCells(): Map<string, Cell> { return new Map(this._cells); }
  get_all_cells(): Map<string, Cell> { return this.getAllCells(); }

  /** Track max data row (1-based). */
  get maxDataRow(): number {
    let max = 0;
    for (const ref of this._cells.keys()) {
      const [r] = Cells.coordinateFromString(ref);
      if (r > max) max = r;
    }
    return max;
  }
  get max_data_row(): number { return this.maxDataRow; }

  /** Track max data column (1-based). */
  get maxDataColumn(): number {
    let max = 0;
    for (const ref of this._cells.keys()) {
      const [, c] = Cells.coordinateFromString(ref);
      if (c > max) max = c;
    }
    return max;
  }
  get max_data_column(): number { return this.maxDataColumn; }

  // ── Iterator support ────────────────────────────────────────────────────

  [Symbol.iterator](): Iterator<[string, Cell]> {
    return this._cells.entries();
  }

  get length(): number { return this._cells.size; }
}
