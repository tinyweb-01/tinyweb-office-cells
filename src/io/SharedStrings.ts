/**
 * Shared String Table for XLSX files.
 *
 * Manages deduplicated string storage per ECMA-376 §18.4 (Shared String Table).
 * Each unique string is stored once; cells reference strings by index.
 */

export class SharedStringTable {
  /** Ordered list of unique strings */
  private _strings: string[] = [];

  /** Reverse lookup: string → index */
  private _stringToIndex: Map<string, number> = new Map();

  /** Total reference count (including duplicates) */
  private _count: number = 0;

  // ── Properties ──────────────────────────────────────────────────────────

  /** Number of unique strings. */
  get uniqueCount(): number {
    return this._strings.length;
  }

  /** Total number of string references (≥ uniqueCount). */
  get count(): number {
    return this._count;
  }

  /** Direct access to the strings array (read-only copy). */
  get strings(): string[] {
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
  addString(text: string): number {
    this._count++;

    const existing = this._stringToIndex.get(text);
    if (existing !== undefined) {
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
  getString(index: number): string {
    if (index < 0 || index >= this._strings.length) {
      throw new RangeError(
        `String index ${index} out of range (0..${this._strings.length - 1})`,
      );
    }
    return this._strings[index];
  }

  /**
   * Serializes the shared string table to OOXML.
   *
   * Produces the content of `xl/sharedStrings.xml`.
   */
  toXml(): string {
    const parts: string[] = [];
    parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    parts.push(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
        ` count="${this._count}" uniqueCount="${this._strings.length}">`,
    );

    for (const text of this._strings) {
      // Preserve leading/trailing whitespace with xml:space="preserve"
      const needsPreserve = text !== text.trim();
      if (needsPreserve) {
        parts.push(`<si><t xml:space="preserve">${escapeXml(text)}</t></si>`);
      } else {
        parts.push(`<si><t>${escapeXml(text)}</t></si>`);
      }
    }

    parts.push('</sst>');
    return parts.join('');
  }

  /**
   * Resets the table to empty state.
   */
  clear(): void {
    this._strings = [];
    this._stringToIndex.clear();
    this._count = 0;
  }
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Escapes special XML characters in text content.
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
