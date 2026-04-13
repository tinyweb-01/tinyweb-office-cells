/**
 * AutoFilter Module
 *
 * Provides classes for Excel auto-filter according to ECMA-376 specification.
 */

// ==================== FilterColumn ====================

export interface CustomFilter {
  operator: string;
  value: string;
}

export interface ColorFilter {
  color: string;
  cellColor: boolean;
}

export interface DynamicFilter {
  type: string;
  value?: string;
}

export interface Top10Filter {
  top: boolean;
  percent: boolean;
  val: number;
}

export class FilterColumn {
  private _colId: number;
  private _filters: string[] = [];
  private _customFilters: CustomFilter[] = [];
  private _colorFilter: ColorFilter | null = null;
  private _dynamicFilter: DynamicFilter | null = null;
  private _top10Filter: Top10Filter | null = null;
  private _filterButton = true;

  constructor(colId: number) {
    this._colId = colId;
  }

  get colId(): number { return this._colId; }

  get filters(): string[] { return this._filters; }
  get customFilters(): CustomFilter[] { return this._customFilters; }

  get colorFilter(): ColorFilter | null { return this._colorFilter; }
  set colorFilter(value: ColorFilter | null) { this._colorFilter = value; }

  get dynamicFilter(): DynamicFilter | null { return this._dynamicFilter; }
  set dynamicFilter(value: DynamicFilter | null) { this._dynamicFilter = value; }

  get top10Filter(): Top10Filter | null { return this._top10Filter; }
  set top10Filter(value: Top10Filter | null) { this._top10Filter = value; }

  get filterButton(): boolean { return this._filterButton; }
  set filterButton(value: boolean) { this._filterButton = value; }

  addFilter(value: string): void {
    if (!this._filters.includes(value)) {
      this._filters.push(value);
    }
  }

  addCustomFilter(operator: string, value: string): void {
    this._customFilters.push({ operator, value });
  }

  clearFilters(): void {
    this._filters = [];
    this._customFilters = [];
    this._colorFilter = null;
    this._dynamicFilter = null;
    this._top10Filter = null;
  }
}

// ==================== SortState ====================

export interface SortState {
  ref: string;
  columnIndex: number;
  ascending: boolean;
}

// ==================== AutoFilter ====================

export class AutoFilter {
  private _range: string | null = null;
  private _filterColumns: Map<number, FilterColumn> = new Map();
  private _sortState: SortState | null = null;

  get range(): string | null { return this._range; }
  set range(value: string | null) { this._range = value; }

  get filterColumns(): Map<number, FilterColumn> { return this._filterColumns; }

  get sortState(): SortState | null { return this._sortState; }
  set sortState(value: SortState | null) { this._sortState = value; }

  setRange(startRow: number, startCol: number, endRow: number, endCol: number): void {
    const startColLetter = this._colToLetter(startCol);
    const endColLetter = this._colToLetter(endCol);
    this._range = `${startColLetter}${startRow + 1}:${endColLetter}${endRow + 1}`;
  }

  filter(colIndex: number, values: string[]): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.clearFilters();
    for (const v of values) {
      fc.addFilter(v);
    }
    return fc;
  }

  addFilter(colIndex: number, value: string): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.addFilter(value);
    return fc;
  }

  customFilter(colIndex: number, operator: string, value: string): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.addCustomFilter(operator, value);
    return fc;
  }

  filterByColor(colIndex: number, color: string, cellColor = true): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.colorFilter = { color, cellColor };
    return fc;
  }

  filterTop10(colIndex: number, top = true, percent = false, val = 10): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.top10Filter = { top, percent, val };
    return fc;
  }

  filterDynamic(colIndex: number, filterType: string, value?: string): FilterColumn {
    const fc = this._getOrCreateColumn(colIndex);
    fc.dynamicFilter = { type: filterType, value };
    return fc;
  }

  clearColumnFilter(colIndex: number): void {
    const fc = this._filterColumns.get(colIndex);
    if (fc) fc.clearFilters();
  }

  clearAllFilters(): void {
    for (const fc of this._filterColumns.values()) {
      fc.clearFilters();
    }
  }

  remove(): void {
    this._range = null;
    this._filterColumns.clear();
    this._sortState = null;
  }

  showFilterButton(colIndex: number, show = true): void {
    const fc = this._getOrCreateColumn(colIndex);
    fc.filterButton = show;
  }

  sort(colIndex: number, ascending = true): void {
    if (!this._range) return;
    this._sortState = {
      ref: this._range,
      columnIndex: colIndex,
      ascending,
    };
  }

  getFilterColumn(colIndex: number): FilterColumn | null {
    return this._filterColumns.get(colIndex) ?? null;
  }

  hasFilter(colIndex: number): boolean {
    const fc = this._filterColumns.get(colIndex);
    if (!fc) return false;
    return (
      fc.filters.length > 0 ||
      fc.customFilters.length > 0 ||
      fc.colorFilter !== null ||
      fc.dynamicFilter !== null ||
      fc.top10Filter !== null
    );
  }

  get hasAnyFilter(): boolean {
    return this._range !== null;
  }

  private _getOrCreateColumn(colIndex: number): FilterColumn {
    if (!this._filterColumns.has(colIndex)) {
      this._filterColumns.set(colIndex, new FilterColumn(colIndex));
    }
    return this._filterColumns.get(colIndex)!;
  }

  private _colToLetter(col: number): string {
    let result = '';
    let c = col;
    while (c >= 0) {
      result = String.fromCharCode(65 + (c % 26)) + result;
      c = Math.floor(c / 26) - 1;
    }
    return result;
  }
}
