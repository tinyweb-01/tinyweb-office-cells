/**
 * Aspose.Cells for Node.js – Workbook Module
 *
 * Provides Workbook class representing an Excel workbook, along with
 * SaveFormat enum and WorkbookProtection helper.
 *
 * Phase 2: XLSX save via XmlSaver (JSZip). Load still stubbed (Phase 3).
 */

import { Worksheet } from './Worksheet';
import { Style } from '../styling/Style';
import { XmlSaver } from '../io/XmlSaver';
import { DefinedNameCollection } from '../features/DefinedName';

// ---------------------------------------------------------------------------
// SaveFormat enum
// ---------------------------------------------------------------------------

/**
 * Specifies the format for saving a workbook.
 *
 * Compatible with Aspose.Cells for .NET SaveFormat enumeration.
 */
export enum SaveFormat {
  AUTO = 'auto',
  XLSX = 'xlsx',
  CSV = 'csv',
  TSV = 'tsv',
  MARKDOWN = 'markdown',
  JSON = 'json',
}

const EXTENSION_FORMAT_MAP: Record<string, SaveFormat> = {
  '.xlsx': SaveFormat.XLSX,
  '.xlsm': SaveFormat.XLSX,
  '.csv': SaveFormat.CSV,
  '.tsv': SaveFormat.TSV,
  '.md': SaveFormat.MARKDOWN,
  '.markdown': SaveFormat.MARKDOWN,
  '.json': SaveFormat.JSON,
};

/**
 * Determine the {@link SaveFormat} from a file path's extension.
 *
 * @throws {Error} If the extension is unsupported.
 */
export function saveFormatFromExtension(filePath: string): SaveFormat {
  const dotIdx = filePath.lastIndexOf('.');
  if (dotIdx === -1) {
    throw new Error(`Unsupported file extension: (none)`);
  }
  const ext = filePath.slice(dotIdx).toLowerCase();
  const fmt = EXTENSION_FORMAT_MAP[ext];
  if (fmt === undefined) {
    throw new Error(`Unsupported file extension: ${ext}`);
  }
  return fmt;
}

/** Alias for .NET naming convention */
export const save_format_from_extension = saveFormatFromExtension;

// ---------------------------------------------------------------------------
// WorkbookProtection
// ---------------------------------------------------------------------------

/**
 * Holds workbook-level protection settings (structure / windows lock).
 */
export class WorkbookProtection {
  lockStructure: boolean = false;
  lockWindows: boolean = false;
  workbookPassword: string | null = null;

  /* .NET aliases */
  get lock_structure(): boolean { return this.lockStructure; }
  set lock_structure(v: boolean) { this.lockStructure = v; }

  get lock_windows(): boolean { return this.lockWindows; }
  set lock_windows(v: boolean) { this.lockWindows = v; }

  get workbook_password(): string | null { return this.workbookPassword; }
  set workbook_password(v: string | null) { this.workbookPassword = v; }
}

// ---------------------------------------------------------------------------
// WorkbookProperties (lightweight stand-in until Phase 3)
// ---------------------------------------------------------------------------

export interface WorkbookView {
  activeTab: number;
  /* .NET alias */
  active_tab: number;
}

export class WorkbookProperties {
  /** Workbook protection settings */
  protection: WorkbookProtection = new WorkbookProtection();

  /** View settings (active tab, etc.) */
  private _view: WorkbookView;

  /** Calculation mode */
  calcMode: string = 'auto';

  /** Defined names collection */
  definedNames: DefinedNameCollection = new DefinedNameCollection();

  constructor() {
    const viewObj = { activeTab: 0 } as WorkbookView;
    // Implement active_tab as alias
    Object.defineProperty(viewObj, 'active_tab', {
      get() { return this.activeTab; },
      set(v: number) { this.activeTab = v; },
      enumerable: true,
      configurable: true,
    });
    this._view = viewObj;
  }

  get view(): WorkbookView { return this._view; }

  /* .NET aliases */
  get calc_mode(): string { return this.calcMode; }
  set calc_mode(v: string) { this.calcMode = v; }

  get defined_names(): DefinedNameCollection { return this.definedNames; }
}

// ---------------------------------------------------------------------------
// DocumentProperties (stub – Phase 3)
// ---------------------------------------------------------------------------

export class DocumentProperties {
  title: string = '';
  subject: string = '';
  creator: string = '';
  keywords: string = '';
  description: string = '';
  lastModifiedBy: string = '';
  category: string = '';
  revision: string = '';
  created: Date | string | null = null;
  modified: Date | string | null = null;

  /** Extended / App properties */
  application: string = '';
  appVersion: string = '';
  company: string = '';
  manager: string = '';

  /** Round-trip XML blobs */
  _coreXml: string | null = null;
  _appXml: string | null = null;

  /* .NET aliases */
  get last_modified_by(): string { return this.lastModifiedBy; }
  set last_modified_by(v: string) { this.lastModifiedBy = v; }
}

// ---------------------------------------------------------------------------
// Workbook
// ---------------------------------------------------------------------------

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
export class Workbook {
  // -- internal state -------------------------------------------------------
  private _worksheets: Worksheet[] = [];
  private _styles: Style[] = [];
  private _sharedStrings: string[] = [];
  private _filePath: string | null = null;

  /** Workbook-level properties */
  private _properties: WorkbookProperties;

  /** Document properties (lazily created) */
  private _documentProperties: DocumentProperties | null = null;

  // Style management maps (used by XML saver in Phase 2/3)
  /** @internal */ _fontStyles: Map<string, number> = new Map();
  /** @internal */ _fillStyles: Map<string, number> = new Map();
  /** @internal */ _borderStyles: Map<string, number> = new Map();
  /** @internal */ _alignmentStyles: Map<string, number> = new Map();
  /** @internal */ _protectionStyles: Map<string, number> = new Map();
  /** @internal */ _cellStyles: Map<string, number> = new Map();
  /** @internal */ _numFormats: Map<string, number> = new Map();

  /** Round-trip: raw workbook-level XML blobs */
  _sourceXml: Record<string, string> = {};

  // -------------------------------------------------------------------------
  // Constructor
  // -------------------------------------------------------------------------

  /**
   * Creates a new Workbook.
   *
   * @param filePath - Reserved for Phase 2/3: path to an existing .xlsx to load.
   * @param password - Reserved for Phase 2/3: password for encrypted files.
   */
  constructor(filePath?: string | null, password?: string | null) {
    this._filePath = filePath ?? null;
    this._properties = new WorkbookProperties();

    // Default style
    this._styles.push(new Style());

    // TODO Phase 2/3: if filePath provided, call _load(filePath, password)
    // For now, always create a default empty workbook.
    const ws = new Worksheet('Sheet1');
    (ws as unknown as { _workbook: Workbook })._workbook = this;
    this._worksheets.push(ws);
  }

  // -------------------------------------------------------------------------
  // Properties
  // -------------------------------------------------------------------------

  /** Gets collection of worksheets in the workbook. */
  get worksheets(): Worksheet[] {
    return this._worksheets;
  }

  /** Gets the file path of the workbook (null if not loaded from file). */
  get filePath(): string | null {
    return this._filePath;
  }

  /** Gets workbook properties (protection, view, calc settings). */
  get properties(): WorkbookProperties {
    return this._properties;
  }

  /** Gets document properties (title, author, etc.). Lazily created. */
  get documentProperties(): DocumentProperties {
    if (this._documentProperties === null) {
      this._documentProperties = new DocumentProperties();
    }
    return this._documentProperties;
  }

  /** Gets the internal styles array. */
  get styles(): Style[] {
    return this._styles;
  }

  /** Gets the shared strings table. */
  get sharedStrings(): string[] {
    return this._sharedStrings;
  }

  // .NET aliases
  get file_path(): string | null { return this.filePath; }
  get document_properties(): DocumentProperties { return this.documentProperties; }
  get shared_strings(): string[] { return this.sharedStrings; }

  // -------------------------------------------------------------------------
  // Worksheet management
  // -------------------------------------------------------------------------

  /**
   * Adds a new worksheet to the workbook.
   *
   * @param name - Optional name. Auto-generated as "SheetN" if omitted.
   * @returns The newly created Worksheet.
   */
  addWorksheet(name?: string): Worksheet {
    if (name === undefined || name === null) {
      const existing = new Set(this._worksheets.map(ws => ws.name));
      let i = 1;
      while (existing.has(`Sheet${i}`)) {
        i++;
      }
      name = `Sheet${i}`;
    }

    const ws = new Worksheet(name);
    (ws as unknown as { _workbook: Workbook })._workbook = this;
    this._worksheets.push(ws);
    return ws;
  }

  /** Alias for {@link addWorksheet}. */
  add_worksheet(name?: string): Worksheet { return this.addWorksheet(name); }

  /** Alias for {@link addWorksheet}. */
  createWorksheet(name?: string): Worksheet { return this.addWorksheet(name); }

  /** Alias for {@link addWorksheet}. */
  create_worksheet(name?: string): Worksheet { return this.addWorksheet(name); }

  /**
   * Gets a worksheet by index (0-based) or by name.
   *
   * @throws {RangeError} If numeric index is out of range.
   * @throws {Error} If no worksheet with the given name exists.
   */
  getWorksheet(indexOrName: number | string): Worksheet {
    if (typeof indexOrName === 'number') {
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
  get_worksheet(indexOrName: number | string): Worksheet { return this.getWorksheet(indexOrName); }

  /**
   * Returns the worksheet with the given name, or `null` if not found.
   */
  getWorksheetByName(name: string): Worksheet | null {
    for (const ws of this._worksheets) {
      if (ws.name === name) {
        return ws;
      }
    }
    return null;
  }

  /** Alias */
  get_worksheet_by_name(name: string): Worksheet | null { return this.getWorksheetByName(name); }

  /**
   * Returns the worksheet at the given 0-based index, or `null` if out of range.
   */
  getWorksheetByIndex(index: number): Worksheet | null {
    if (Number.isInteger(index) && index >= 0 && index < this._worksheets.length) {
      return this._worksheets[index];
    }
    return null;
  }

  /** Alias */
  get_worksheet_by_index(index: number): Worksheet | null { return this.getWorksheetByIndex(index); }

  /**
   * Removes a worksheet by index, name, or direct reference.
   *
   * @throws {RangeError} If numeric index is out of range.
   * @throws {Error} If string name is not found.
   * @throws {TypeError} If argument is none of the above.
   */
  removeWorksheet(indexOrNameOrWs: number | string | Worksheet): void {
    if (typeof indexOrNameOrWs === 'number') {
      if (indexOrNameOrWs >= 0 && indexOrNameOrWs < this._worksheets.length) {
        this._worksheets.splice(indexOrNameOrWs, 1);
        return;
      }
      throw new RangeError(`Worksheet index ${indexOrNameOrWs} out of range`);
    }
    if (typeof indexOrNameOrWs === 'string') {
      const idx = this._worksheets.findIndex(ws => ws.name === indexOrNameOrWs);
      if (idx !== -1) {
        this._worksheets.splice(idx, 1);
        return;
      }
      throw new Error(`Worksheet '${indexOrNameOrWs}' not found`);
    }
    // Treat as Worksheet object reference
    const idx = this._worksheets.findIndex(ws => ws === indexOrNameOrWs);
    if (idx !== -1) {
      this._worksheets.splice(idx, 1);
      return;
    }
    throw new TypeError('removeWorksheet: argument must be number, string, or Worksheet');
  }

  /** Alias */
  remove_worksheet(indexOrNameOrWs: number | string | Worksheet): void {
    this.removeWorksheet(indexOrNameOrWs);
  }

  /**
   * Returns the currently active worksheet.
   */
  getActiveWorksheet(): Worksheet | null {
    if (this._worksheets.length === 0) return null;
    let idx = this._properties.view.activeTab;
    idx = Math.max(0, Math.min(idx, this._worksheets.length - 1));
    return this._worksheets[idx];
  }

  /** Alias */
  get_active_worksheet(): Worksheet | null { return this.getActiveWorksheet(); }

  /**
   * Sets the active worksheet by index, name, or Worksheet reference.
   */
  setActiveWorksheet(indexOrNameOrWs: number | string | Worksheet): void {
    let idx: number | undefined;
    if (typeof indexOrNameOrWs === 'number') {
      if (indexOrNameOrWs >= 0 && indexOrNameOrWs < this._worksheets.length) {
        idx = indexOrNameOrWs;
      }
    } else if (typeof indexOrNameOrWs === 'string') {
      idx = this._worksheets.findIndex(ws => ws.name === indexOrNameOrWs);
      if (idx === -1) idx = undefined;
    } else {
      idx = this._worksheets.findIndex(ws => ws === indexOrNameOrWs);
      if (idx === -1) idx = undefined;
    }
    if (idx !== undefined) {
      this._properties.view.activeTab = idx;
    }
  }

  /** Alias */
  set_active_worksheet(indexOrNameOrWs: number | string | Worksheet): void {
    this.setActiveWorksheet(indexOrNameOrWs);
  }

  /**
   * Copies a worksheet and appends the copy to the workbook.
   *
   * @returns The new worksheet, or `null` if source could not be resolved.
   */
  copyWorksheet(indexOrNameOrWs: number | string | Worksheet): Worksheet | null {
    let ws: Worksheet | null | undefined;
    if (typeof indexOrNameOrWs === 'number') {
      ws = this.getWorksheetByIndex(indexOrNameOrWs);
    } else if (typeof indexOrNameOrWs === 'string') {
      ws = this.getWorksheetByName(indexOrNameOrWs);
    } else if (indexOrNameOrWs instanceof Worksheet) {
      ws = indexOrNameOrWs;
    }
    if (!ws) return null;

    // Generate unique copy name
    const base = ws.name;
    const existing = new Set(this._worksheets.map(w => w.name));
    let copyName = `${base} (copy)`;
    if (existing.has(copyName)) {
      let i = 1;
      while (existing.has(`${base} (copy${i})`)) {
        i++;
      }
      copyName = `${base} (copy${i})`;
    }

    const newWs = ws.copy(copyName);
    (newWs as unknown as { _workbook: Workbook })._workbook = this;
    this._worksheets.push(newWs);
    return newWs;
  }

  /** Alias */
  copy_worksheet(indexOrNameOrWs: number | string | Worksheet): Worksheet | null {
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
  protect(password?: string | null, lockStructure: boolean = true, lockWindows: boolean = false): void {
    const prot = this._properties.protection;
    prot.lockStructure = lockStructure;
    prot.lockWindows = lockWindows;
    // In Phase 3, hash the password. For now store plain text.
    prot.workbookPassword = password ?? null;
  }

  /**
   * Removes workbook structure/window protection.
   */
  unprotect(_password?: string | null): void {
    const prot = this._properties.protection;
    prot.lockStructure = false;
    prot.lockWindows = false;
    prot.workbookPassword = null;
  }

  /**
   * Returns `true` if the workbook has structure or window protection enabled.
   */
  isProtected(): boolean {
    const prot = this._properties.protection;
    return prot.lockStructure || prot.lockWindows;
  }

  /** Alias */
  is_protected(): boolean { return this.isProtected(); }

  /**
   * Returns a snapshot of the current workbook protection settings.
   */
  get protection(): { lockStructure: boolean; lockWindows: boolean; password: string | null } {
    const prot = this._properties.protection;
    return {
      lockStructure: prot.lockStructure,
      lockWindows: prot.lockWindows,
      password: prot.workbookPassword,
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
  calculateFormula(): void {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { FormulaEvaluator } = require('../features/FormulaEvaluator');
    const evaluator = new FormulaEvaluator(this);
    evaluator.evaluateAll();
  }

  /** snake_case alias */
  calculate_formula(): void { this.calculateFormula(); }

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
  async save(
    filePath: string,
    saveFormat?: SaveFormat | null,
    _options?: unknown,
    _password?: string | null,
  ): Promise<void> {
    const fmt = saveFormat && saveFormat !== SaveFormat.AUTO
      ? saveFormat
      : saveFormatFromExtension(filePath);

    if (fmt !== SaveFormat.XLSX) {
      throw new Error(
        `Workbook.save(): format '${fmt}' is not yet supported. Only XLSX is implemented.`,
      );
    }

    const saver = new XmlSaver(this);
    await saver.saveToFile(filePath);
  }

  /**
   * Saves the workbook to an in-memory Buffer (XLSX format).
   *
   * @param saveFormat  Explicit format; defaults to XLSX.
   * @returns A Buffer containing the XLSX ZIP archive.
   */
  async saveToBuffer(
    saveFormat: SaveFormat = SaveFormat.XLSX,
  ): Promise<Buffer> {
    if (saveFormat !== SaveFormat.XLSX && saveFormat !== SaveFormat.AUTO) {
      throw new Error(
        `Workbook.saveToBuffer(): format '${saveFormat}' is not yet supported. Only XLSX is implemented.`,
      );
    }

    const saver = new XmlSaver(this);
    return saver.saveToBuffer();
  }

  /** snake_case alias */
  async save_to_buffer(
    saveFormat: SaveFormat = SaveFormat.XLSX,
  ): Promise<Buffer> {
    return this.saveToBuffer(saveFormat);
  }

  /**
   * Loads a workbook from a file path (async).
   *
   * @param filePath Path to the .xlsx file.
   * @param _password Reserved for future encrypted workbook support.
   * @returns A Promise resolving to the loaded Workbook.
   */
  static async load(
    filePath: string,
    _password?: string | null,
  ): Promise<Workbook> {
    // Lazy import to avoid circular dependency at module load time
    const { XmlLoader } = await import('../io/XmlLoader');
    return XmlLoader.loadFromFile(filePath);
  }

  /**
   * Loads a workbook from a Buffer (async).
   *
   * @param buffer Buffer containing XLSX data.
   * @returns A Promise resolving to the loaded Workbook.
   */
  static async loadFromBuffer(
    buffer: Buffer,
  ): Promise<Workbook> {
    const { XmlLoader } = await import('../io/XmlLoader');
    return XmlLoader.loadFromBuffer(buffer);
  }

  // -------------------------------------------------------------------------
  // String representation
  // -------------------------------------------------------------------------

  toString(): string {
    return `Workbook(worksheets=${this._worksheets.length})`;
  }
}
