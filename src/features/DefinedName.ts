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

// ─── DefinedName ──────────────────────────────────────────────────────────

/**
 * Represents a defined name in the workbook.
 */
export class DefinedName {
  private _name: string;
  private _refersTo: string;
  private _localSheetId: number | null;
  private _comment: string | null = null;
  private _description: string | null = null;
  private _hidden = false;

  constructor(name: string, refersTo: string, localSheetId: number | null = null) {
    this._name = name;
    this._refersTo = refersTo;
    this._localSheetId = localSheetId;
  }

  /** Name of the defined name. */
  get name(): string { return this._name; }
  set name(v: string) { this._name = v; }

  /** Formula or range that the name refers to. */
  get refersTo(): string { return this._refersTo; }
  set refersTo(v: string) { this._refersTo = v; }

  /** Sheet index for sheet-local names (null for global names). */
  get localSheetId(): number | null { return this._localSheetId; }
  set localSheetId(v: number | null) { this._localSheetId = v; }

  /** Comment associated with the name. */
  get comment(): string | null { return this._comment; }
  set comment(v: string | null) { this._comment = v; }

  /** Description of the name. */
  get description(): string | null { return this._description; }
  set description(v: string | null) { this._description = v; }

  /** Whether the name is hidden. */
  get hidden(): boolean { return this._hidden; }
  set hidden(v: boolean) { this._hidden = v; }

  // snake_case aliases
  get refers_to(): string { return this._refersTo; }
  set refers_to(v: string) { this._refersTo = v; }
  get local_sheet_id(): number | null { return this._localSheetId; }
  set local_sheet_id(v: number | null) { this._localSheetId = v; }
}

// ─── DefinedNameCollection ────────────────────────────────────────────────

/**
 * Collection of defined names in the workbook.
 */
export class DefinedNameCollection {
  private _names: DefinedName[] = [];

  /**
   * Adds a defined name to the collection.
   *
   * @param nameOrDn Either a DefinedName object or a string name.
   * @param refersTo Formula or range (required if nameOrDn is a string).
   * @param localSheetId Sheet index for sheet-local names.
   * @returns The added DefinedName.
   */
  add(nameOrDn: DefinedName | string, refersTo?: string, localSheetId?: number | null): DefinedName {
    if (nameOrDn instanceof DefinedName) {
      this._names.push(nameOrDn);
      return nameOrDn;
    }
    const dn = new DefinedName(nameOrDn, refersTo ?? '', localSheetId ?? null);
    this._names.push(dn);
    return dn;
  }

  /**
   * Removes a defined name by name string.
   * @returns The removed DefinedName or null if not found.
   */
  remove(name: string): DefinedName | null {
    const idx = this._names.findIndex(dn => dn.name === name);
    if (idx >= 0) {
      return this._names.splice(idx, 1)[0];
    }
    return null;
  }

  /**
   * Gets a defined name by index or name string.
   */
  get(key: number | string): DefinedName {
    if (typeof key === 'number') {
      if (key < 0 || key >= this._names.length) {
        throw new RangeError(`Index ${key} out of range`);
      }
      return this._names[key];
    }
    const found = this._names.find(dn => dn.name === key);
    if (!found) {
      throw new Error(`Defined name '${key}' not found`);
    }
    return found;
  }

  /** Number of defined names. */
  get count(): number {
    return this._names.length;
  }

  /** Returns internal array (for iteration). */
  toArray(): DefinedName[] {
    return [...this._names];
  }

  [Symbol.iterator](): Iterator<DefinedName> {
    return this._names[Symbol.iterator]();
  }
}
