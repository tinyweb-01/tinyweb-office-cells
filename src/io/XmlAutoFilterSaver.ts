/**
 * XML AutoFilter Saver
 *
 * Writes auto-filter settings to XLSX worksheet XML.
 */

import { AutoFilter } from '../features/AutoFilter';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export class AutoFilterXmlSaver {
  formatAutoFilterXml(autoFilter: AutoFilter): string {
    if (!autoFilter.range) return '';

    const lines: string[] = [];
    lines.push(`<autoFilter ref="${escapeXml(autoFilter.range)}">`);

    // Sort filter columns by column index
    const sortedCols = Array.from(autoFilter.filterColumns.entries()).sort(
      (a, b) => a[0] - b[0],
    );

    for (const [, fc] of sortedCols) {
      const hasContent =
        fc.filters.length > 0 ||
        fc.customFilters.length > 0 ||
        fc.colorFilter !== null ||
        fc.dynamicFilter !== null ||
        fc.top10Filter !== null;

      if (!hasContent && fc.filterButton) continue;

      const fcAttrs: string[] = [`colId="${fc.colId}"`];
      if (!fc.filterButton) {
        fcAttrs.push('hiddenButton="1"');
      }

      lines.push(`<filterColumn ${fcAttrs.join(' ')}>`);

      // Value filters
      if (fc.filters.length > 0) {
        lines.push('<filters>');
        for (const val of fc.filters) {
          lines.push(`<filter val="${escapeXml(val)}"/>`);
        }
        lines.push('</filters>');
      }

      // Custom filters
      if (fc.customFilters.length > 0) {
        const andAttr = fc.customFilters.length > 1 ? ' and="1"' : '';
        lines.push(`<customFilters${andAttr}>`);
        for (const cf of fc.customFilters) {
          lines.push(`<customFilter operator="${escapeXml(cf.operator)}" val="${escapeXml(cf.value)}"/>`);
        }
        lines.push('</customFilters>');
      }

      // Color filter
      if (fc.colorFilter) {
        const cellColorAttr = fc.colorFilter.cellColor ? '' : ' cellColor="0"';
        lines.push(`<colorFilter dxfId="${escapeXml(fc.colorFilter.color)}"${cellColorAttr}/>`);
      }

      // Dynamic filter
      if (fc.dynamicFilter) {
        const valAttr = fc.dynamicFilter.value ? ` val="${escapeXml(fc.dynamicFilter.value)}"` : '';
        lines.push(`<dynamicFilter type="${escapeXml(fc.dynamicFilter.type)}"${valAttr}/>`);
      }

      // Top10 filter
      if (fc.top10Filter) {
        const attrs: string[] = [];
        if (!fc.top10Filter.top) attrs.push('top="0"');
        if (fc.top10Filter.percent) attrs.push('percent="1"');
        attrs.push(`val="${fc.top10Filter.val}"`);
        lines.push(`<top10 ${attrs.join(' ')}/>`);
      }

      lines.push('</filterColumn>');
    }

    // Sort state
    if (autoFilter.sortState) {
      const ss = autoFilter.sortState;
      const sortRef = this._calculateSortRefs(autoFilter.range, ss.columnIndex);
      lines.push(`<sortState ref="${escapeXml(ss.ref)}">`);
      const descAttr = ss.ascending ? '' : ' descending="1"';
      lines.push(`<sortCondition ref="${escapeXml(sortRef)}"${descAttr}/>`);
      lines.push('</sortState>');
    }

    lines.push('</autoFilter>');
    return lines.join('');
  }

  private _calculateSortRefs(filterRange: string, columnIndex: number): string {
    const parts = filterRange.split(':');
    if (parts.length !== 2) return filterRange;

    const startCol = this._colFromRef(parts[0]);
    const targetCol = startCol + columnIndex;
    const colLetter = this._colToLetter(targetCol);

    const startRow = this._rowFromRef(parts[0]);
    const endRow = this._rowFromRef(parts[1]);

    return `${colLetter}${startRow}:${colLetter}${endRow}`;
  }

  private _colFromRef(ref: string): number {
    const match = ref.match(/^([A-Z]+)/i);
    if (!match) return 0;
    let col = 0;
    for (const ch of match[1].toUpperCase()) {
      col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return col - 1;
  }

  private _rowFromRef(ref: string): number {
    const match = ref.match(/(\d+)$/);
    return match ? parseInt(match[1], 10) : 1;
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
