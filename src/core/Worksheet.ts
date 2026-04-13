/**
 * Aspose.Cells for Node.js - Worksheet Module
 *
 * Represents a single worksheet in an Excel workbook.
 * Compatible with Aspose.Cells for .NET API structure.
 */

import { Cells } from './Cells';
import { Cell } from './Cell';
import { DataValidationCollection } from '../features/DataValidation';
import { ConditionalFormatCollection } from '../features/ConditionalFormat';
import { HyperlinkCollection } from '../features/Hyperlink';
import { AutoFilter } from '../features/AutoFilter';
import { HorizontalPageBreakCollection, VerticalPageBreakCollection } from '../features/PageBreak';

// ─── SheetProtection ──────────────────────────────────────────────────────

export class SheetProtection {
  sheet = false;
  password: string | null = null;
  objects = false;
  scenarios = false;
  formatCells = false;
  formatColumns = false;
  formatRows = false;
  insertColumns = false;
  insertRows = false;
  insertHyperlinks = false;
  deleteColumns = false;
  deleteRows = false;
  selectLockedCells = true;
  selectUnlockedCells = true;
  sort = false;
  autoFilter = false;
  pivotTables = false;

  // snake_case aliases
  get format_cells(): boolean { return this.formatCells; }
  set format_cells(v: boolean) { this.formatCells = v; }
  get format_columns(): boolean { return this.formatColumns; }
  set format_columns(v: boolean) { this.formatColumns = v; }
  get format_rows(): boolean { return this.formatRows; }
  set format_rows(v: boolean) { this.formatRows = v; }
  get insert_columns(): boolean { return this.insertColumns; }
  set insert_columns(v: boolean) { this.insertColumns = v; }
  get insert_rows(): boolean { return this.insertRows; }
  set insert_rows(v: boolean) { this.insertRows = v; }
  get insert_hyperlinks(): boolean { return this.insertHyperlinks; }
  set insert_hyperlinks(v: boolean) { this.insertHyperlinks = v; }
  get delete_columns(): boolean { return this.deleteColumns; }
  set delete_columns(v: boolean) { this.deleteColumns = v; }
  get delete_rows(): boolean { return this.deleteRows; }
  set delete_rows(v: boolean) { this.deleteRows = v; }
  get select_locked_cells(): boolean { return this.selectLockedCells; }
  set select_locked_cells(v: boolean) { this.selectLockedCells = v; }
  get select_unlocked_cells(): boolean { return this.selectUnlockedCells; }
  set select_unlocked_cells(v: boolean) { this.selectUnlockedCells = v; }
  get auto_filter(): boolean { return this.autoFilter; }
  set auto_filter(v: boolean) { this.autoFilter = v; }
  get pivot_tables(): boolean { return this.pivotTables; }
  set pivot_tables(v: boolean) { this.pivotTables = v; }
}

// ─── SheetProtectionDictWrapper ───────────────────────────────────────────

const PROTECTION_KEYS = [
  'protected', 'password', 'sheet', 'objects', 'scenarios',
  'format_cells', 'format_columns', 'format_rows',
  'insert_columns', 'insert_rows', 'insert_hyperlinks',
  'delete_columns', 'delete_rows',
  'select_locked_cells', 'select_unlocked_cells',
  'sort', 'auto_filter', 'pivot_tables',
] as const;

type ProtectionKey = typeof PROTECTION_KEYS[number];

/**
 * Dictionary-like wrapper around SheetProtection for backward compatibility.
 */
export class SheetProtectionDictWrapper {
  private _protection: SheetProtection;

  constructor(sheetProtection: SheetProtection) {
    this._protection = sheetProtection;
  }

  get(key: ProtectionKey, defaultValue?: unknown): unknown {
    try {
      return this.getItem(key);
    } catch {
      return defaultValue;
    }
  }

  getItem(key: ProtectionKey): unknown {
    switch (key) {
      case 'protected':
      case 'sheet': return this._protection.sheet;
      case 'password': return this._protection.password;
      case 'objects': return this._protection.objects;
      case 'scenarios': return this._protection.scenarios;
      case 'format_cells': return this._protection.formatCells;
      case 'format_columns': return this._protection.formatColumns;
      case 'format_rows': return this._protection.formatRows;
      case 'insert_columns': return this._protection.insertColumns;
      case 'insert_rows': return this._protection.insertRows;
      case 'insert_hyperlinks': return this._protection.insertHyperlinks;
      case 'delete_columns': return this._protection.deleteColumns;
      case 'delete_rows': return this._protection.deleteRows;
      case 'select_locked_cells': return this._protection.selectLockedCells;
      case 'select_unlocked_cells': return this._protection.selectUnlockedCells;
      case 'sort': return this._protection.sort;
      case 'auto_filter': return this._protection.autoFilter;
      case 'pivot_tables': return this._protection.pivotTables;
      default: throw new Error(`Unknown protection key: ${key}`);
    }
  }

  setItem(key: ProtectionKey, value: unknown): void {
    switch (key) {
      case 'protected':
      case 'sheet': this._protection.sheet = value as boolean; break;
      case 'password': this._protection.password = value as string | null; break;
      case 'objects': this._protection.objects = value as boolean; break;
      case 'scenarios': this._protection.scenarios = value as boolean; break;
      case 'format_cells': this._protection.formatCells = value as boolean; break;
      case 'format_columns': this._protection.formatColumns = value as boolean; break;
      case 'format_rows': this._protection.formatRows = value as boolean; break;
      case 'insert_columns': this._protection.insertColumns = value as boolean; break;
      case 'insert_rows': this._protection.insertRows = value as boolean; break;
      case 'insert_hyperlinks': this._protection.insertHyperlinks = value as boolean; break;
      case 'delete_columns': this._protection.deleteColumns = value as boolean; break;
      case 'delete_rows': this._protection.deleteRows = value as boolean; break;
      case 'select_locked_cells': this._protection.selectLockedCells = value as boolean; break;
      case 'select_unlocked_cells': this._protection.selectUnlockedCells = value as boolean; break;
      case 'sort': this._protection.sort = value as boolean; break;
      case 'auto_filter': this._protection.autoFilter = value as boolean; break;
      case 'pivot_tables': this._protection.pivotTables = value as boolean; break;
      default: throw new Error(`Unknown protection key: ${key}`);
    }
  }
}

// ─── PageSetup ────────────────────────────────────────────────────────────

export interface PageSetup {
  orientation: string | null;
  paperSize: number | null;
  scale: number | null;
  fitToWidth: number | null;
  fitToHeight: number | null;
  fitToPage: boolean;
}

export interface PageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

// ─── FreezePane ───────────────────────────────────────────────────────────

export interface FreezePane {
  row: number;
  column: number;
  freezedRows: number;
  freezedColumns: number;
}

// ─── Worksheet ────────────────────────────────────────────────────────────

/**
 * Represents a single worksheet in an Excel workbook.
 */
export class Worksheet {
  private _name: string;
  private _cells: Cells;
  private _visible: boolean | 'veryHidden';
  private _tabColor: string | null;
  private _index: number;

  // Sheet protection
  private _protection: SheetProtection;

  // Page setup
  private _pageSetup: PageSetup;
  private _pageMargins: PageMargins;

  // Freeze panes
  private _freezePane: FreezePane | null;

  // Merge tracking
  _mergedCells: string[];

  // Row/Column dimensions (exposed for Cells to access)
  _rowHeights: Record<number, number>;
  _columnWidths: Record<number, number>;
  _hiddenRows: Set<number>;
  _hiddenColumns: Set<number>;

  // Print area
  private _printArea: string | null;

  // Round-trip fidelity: raw XML preservation
  _sourceXml: string | null;

  // Feature collections (Phase 4)
  private _dataValidations: DataValidationCollection;
  private _conditionalFormatting: ConditionalFormatCollection;
  private _hyperlinks: HyperlinkCollection;
  private _autoFilter: AutoFilter;
  private _horizontalPageBreaks: HorizontalPageBreakCollection;
  private _verticalPageBreaks: VerticalPageBreakCollection;

  // Workbook back-reference (set by Workbook)
  _workbook: unknown;

  constructor(name = 'Sheet1') {
    this._name = name;
    this._cells = new Cells(this);
    this._visible = true;
    this._tabColor = null;
    this._index = 0;
    this._protection = new SheetProtection();
    this._pageSetup = {
      orientation: null,
      paperSize: null,
      scale: null,
      fitToWidth: null,
      fitToHeight: null,
      fitToPage: false,
    };
    this._pageMargins = {
      left: 0.75,
      right: 0.75,
      top: 1.0,
      bottom: 1.0,
      header: 0.5,
      footer: 0.5,
    };
    this._freezePane = null;
    this._mergedCells = [];
    this._rowHeights = {};
    this._columnWidths = {};
    this._hiddenRows = new Set();
    this._hiddenColumns = new Set();
    this._printArea = null;
    this._sourceXml = null;
    this._dataValidations = new DataValidationCollection();
    this._conditionalFormatting = new ConditionalFormatCollection();
    this._hyperlinks = new HyperlinkCollection();
    this._autoFilter = new AutoFilter();
    this._horizontalPageBreaks = new HorizontalPageBreakCollection();
    this._verticalPageBreaks = new VerticalPageBreakCollection();
    this._workbook = null;
  }

  // ── Properties ──────────────────────────────────────────────────────────

  get name(): string { return this._name; }
  set name(v: string) { this._name = v; }

  get cells(): Cells { return this._cells; }

  get visible(): boolean | 'veryHidden' { return this._visible; }
  set visible(v: boolean | 'veryHidden') { this._visible = v; }

  /** Alias for visible */
  get isVisible(): boolean { return this._visible === true; }
  get is_visible(): boolean { return this.isVisible; }

  get tabColor(): string | null { return this._tabColor; }
  set tabColor(v: string | null) { this._tabColor = v; }
  get tab_color(): string | null { return this._tabColor; }
  set tab_color(v: string | null) { this._tabColor = v; }

  get index(): number { return this._index; }
  set index(v: number) { this._index = v; }

  get protection(): SheetProtectionDictWrapper {
    return new SheetProtectionDictWrapper(this._protection);
  }

  get pageSetup(): PageSetup { return this._pageSetup; }
  get page_setup(): PageSetup { return this._pageSetup; }

  get pageMargins(): PageMargins { return this._pageMargins; }
  get page_margins(): PageMargins { return this._pageMargins; }

  get printArea(): string | null { return this._printArea; }
  set printArea(v: string | null) {
    this._printArea = v ? this.normalizePrintArea(v) : null;
  }
  get print_area(): string | null { return this._printArea; }
  set print_area(v: string | null) { this.printArea = v; }

  get mergedCells(): string[] { return [...this._mergedCells]; }
  get merged_cells(): string[] { return this.mergedCells; }

  get freezePane(): FreezePane | null { return this._freezePane; }
  get freeze_pane(): FreezePane | null { return this._freezePane; }

  // Feature collections (Phase 4)
  get dataValidations(): DataValidationCollection { return this._dataValidations; }
  get data_validations(): DataValidationCollection { return this._dataValidations; }

  get conditionalFormatting(): ConditionalFormatCollection { return this._conditionalFormatting; }
  get conditional_formatting(): ConditionalFormatCollection { return this._conditionalFormatting; }

  get hyperlinks(): HyperlinkCollection { return this._hyperlinks; }

  get autoFilter(): AutoFilter { return this._autoFilter; }
  get auto_filter_obj(): AutoFilter { return this._autoFilter; }

  get horizontalPageBreaks(): HorizontalPageBreakCollection { return this._horizontalPageBreaks; }
  get horizontal_page_breaks(): HorizontalPageBreakCollection { return this._horizontalPageBreaks; }

  get verticalPageBreaks(): VerticalPageBreakCollection { return this._verticalPageBreaks; }
  get vertical_page_breaks(): VerticalPageBreakCollection { return this._verticalPageBreaks; }

  // ── Methods ─────────────────────────────────────────────────────────────

  rename(newName: string): void { this._name = newName; }

  setVisibility(value: boolean | 'veryHidden'): void {
    if (value !== true && value !== false && value !== 'veryHidden') {
      throw new Error(`Invalid visibility value: ${value}. Use true, false, or 'veryHidden'.`);
    }
    this._visible = value;
  }
  set_visibility(value: boolean | 'veryHidden'): void { this.setVisibility(value); }

  getVisibility(): boolean | 'veryHidden' { return this._visible; }
  get_visibility(): boolean | 'veryHidden' { return this.getVisibility(); }

  setTabColor(color: string | null): void {
    if (color !== null && color.length !== 8) {
      throw new Error(`Invalid tab color format: ${color}. Expected 8-char AARRGGBB hex string.`);
    }
    this._tabColor = color ? color.toUpperCase() : null;
  }
  set_tab_color(color: string | null): void { this.setTabColor(color); }

  getTabColor(): string | null { return this._tabColor; }
  get_tab_color(): string | null { return this.getTabColor(); }

  clearTabColor(): void { this._tabColor = null; }
  clear_tab_color(): void { this.clearTabColor(); }

  // ── Page setup methods ──────────────────────────────────────────────────

  setPageOrientation(orientation: 'portrait' | 'landscape'): void {
    if (orientation !== 'portrait' && orientation !== 'landscape') {
      throw new Error(`Invalid orientation: ${orientation}. Use 'portrait' or 'landscape'.`);
    }
    this._pageSetup.orientation = orientation;
  }
  set_page_orientation(orientation: 'portrait' | 'landscape'): void { this.setPageOrientation(orientation); }

  getPageOrientation(): string | null { return this._pageSetup.orientation; }
  get_page_orientation(): string | null { return this.getPageOrientation(); }

  setPaperSize(paperSize: number): void { this._pageSetup.paperSize = paperSize; }
  set_paper_size(paperSize: number): void { this.setPaperSize(paperSize); }

  getPaperSize(): number | null { return this._pageSetup.paperSize; }
  get_paper_size(): number | null { return this.getPaperSize(); }

  setPageMargins(options: Partial<PageMargins>): void {
    if (options.left != null) this._pageMargins.left = options.left;
    if (options.right != null) this._pageMargins.right = options.right;
    if (options.top != null) this._pageMargins.top = options.top;
    if (options.bottom != null) this._pageMargins.bottom = options.bottom;
    if (options.header != null) this._pageMargins.header = options.header;
    if (options.footer != null) this._pageMargins.footer = options.footer;
  }
  set_page_margins(options: Partial<PageMargins>): void { this.setPageMargins(options); }

  getPageMargins(): PageMargins { return { ...this._pageMargins }; }
  get_page_margins(): PageMargins { return this.getPageMargins(); }

  setFitToPages(width = 1, height = 1): void {
    this._pageSetup.fitToWidth = width;
    this._pageSetup.fitToHeight = height;
    this._pageSetup.fitToPage = true;
  }
  set_fit_to_pages(width = 1, height = 1): void { this.setFitToPages(width, height); }

  setPrintScale(scale: number): void {
    if (scale < 10 || scale > 400) throw new Error(`Print scale must be between 10 and 400, got ${scale}.`);
    this._pageSetup.scale = scale;
  }
  set_print_scale(scale: number): void { this.setPrintScale(scale); }

  // ── Print area ──────────────────────────────────────────────────────────

  setPrintArea(printArea: string): void { this.printArea = printArea; }
  set_print_area(printArea: string): void { this.setPrintArea(printArea); }
  SetPrintArea(printArea: string): void { this.setPrintArea(printArea); }

  clearPrintArea(): void { this._printArea = null; }
  clear_print_area(): void { this.clearPrintArea(); }
  ClearPrintArea(): void { this.clearPrintArea(); }

  private normalizePrintArea(printArea: string): string {
    if (!printArea || !printArea.trim()) throw new Error('print_area must be a non-empty string');
    const parts: string[] = [];
    for (const token of printArea.split(',')) {
      const part = token.trim().toUpperCase().replace(/\$/g, '');
      if (!part) continue;
      if (part.includes(':')) {
        const [startRef, endRef] = part.split(':');
        Cells.coordinateFromString(startRef);
        Cells.coordinateFromString(endRef);
        parts.push(`${startRef}:${endRef}`);
      } else {
        Cells.coordinateFromString(part);
        parts.push(part);
      }
    }
    if (parts.length === 0) throw new Error('print_area must contain at least one valid range');
    return parts.join(',');
  }

  // ── Freeze pane ─────────────────────────────────────────────────────────

  setFreezePane(row: number, column: number, freezedRows?: number, freezedColumns?: number): void {
    this._freezePane = {
      row,
      column,
      freezedRows: freezedRows ?? row,
      freezedColumns: freezedColumns ?? column,
    };
  }
  set_freeze_pane(row: number, column: number, freezedRows?: number, freezedColumns?: number): void {
    this.setFreezePane(row, column, freezedRows, freezedColumns);
  }

  clearFreezePane(): void { this._freezePane = null; }
  clear_freeze_pane(): void { this.clearFreezePane(); }

  // ── Protection ──────────────────────────────────────────────────────────

  isProtected(): boolean { return this._protection.sheet; }
  is_protected(): boolean { return this.isProtected(); }

  protect(options: {
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
  } = {}): void {
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

  unprotect(_password?: string): void {
    this._protection.sheet = false;
    this._protection.password = null;
  }

  // ── Copy / placeholder methods ──────────────────────────────────────────

  copy(name?: string): Worksheet {
    const newWs = new Worksheet(name ?? `${this._name} (copy)`);
    for (const [ref, cell] of this._cells._cells.entries()) {
      const newCell = new Cell(cell.value, cell.formula);
      if (cell.style) newCell.style = cell.style.copy();
      newWs._cells._cells.set(ref, newCell);
    }
    newWs._mergedCells = [...this._mergedCells];
    newWs._printArea = this._printArea;
    return newWs;
  }

  /** Placeholder */
  delete(): void { /* handled by Workbook */ }
  /** Placeholder */
  move(_index: number): void { /* handled by Workbook */ }
  /** Placeholder */
  select(): void { /* handled by Workbook */ }
  /** Placeholder */
  activate(): void { /* handled by Workbook */ }
  /**
   * Evaluates every formula on this worksheet only.
   *
   * Delegates to the workbook-level FormulaEvaluator so that cross-sheet
   * references and defined names resolve correctly.
   */
  calculateFormula(): void {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { FormulaEvaluator } = require('../features/FormulaEvaluator');
    const evaluator = new FormulaEvaluator(this._workbook);
    evaluator.evaluateAll(this);
  }

  /** snake_case alias */
  calculate_formula(): void { this.calculateFormula(); }
}
