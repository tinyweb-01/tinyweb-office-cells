/**
 * XML AutoFilter Loader
 *
 * Loads auto-filter settings from XLSX worksheet XML.
 */

import { AutoFilter, FilterColumn } from '../features/AutoFilter';

function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}

function attr(elem: any, name: string, defaultVal = ''): string {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : defaultVal;
}

function attrOpt(elem: any, name: string): string | null {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : null;
}

export class AutoFilterXmlLoader {
  loadAutoFilter(wsRoot: any, autoFilter: AutoFilter): void {
    const afElem = wsRoot?.autoFilter;
    if (!afElem) return;

    const ref = attr(afElem, 'ref', '');
    if (!ref) return;

    autoFilter.range = ref;

    // Load filter columns
    const fcElems = ensureArray(afElem.filterColumn);
    for (const fcElem of fcElems) {
      const colId = parseInt(attr(fcElem, 'colId', '0'), 10);
      const fc = autoFilter.getFilterColumn(colId) ?? new FilterColumn(colId);

      // hiddenButton means HIDE the dropdown button
      const hiddenButton = attr(fcElem, 'hiddenButton', '');
      if (hiddenButton === '1' || hiddenButton === 'true') {
        fc.filterButton = false;
      }

      // Filters (value list)
      const filtersElem = fcElem.filters;
      if (filtersElem) {
        const filterElems = ensureArray(filtersElem.filter);
        for (const fe of filterElems) {
          const val = attr(fe, 'val', '');
          if (val) fc.addFilter(val);
        }
      }

      // Custom filters
      const customFiltersElem = fcElem.customFilters;
      if (customFiltersElem) {
        const customElems = ensureArray(customFiltersElem.customFilter);
        for (const ce of customElems) {
          const op = attr(ce, 'operator', 'equal');
          const val = attr(ce, 'val', '');
          fc.addCustomFilter(op, val);
        }
      }

      // Color filter
      const colorFilterElem = fcElem.colorFilter;
      if (colorFilterElem) {
        const color = attr(colorFilterElem, 'dxfId', '');
        const cellColor = attr(colorFilterElem, 'cellColor', '1') !== '0';
        fc.colorFilter = { color, cellColor };
      }

      // Dynamic filter
      const dynamicElem = fcElem.dynamicFilter;
      if (dynamicElem) {
        const type = attr(dynamicElem, 'type', '');
        const val = attrOpt(dynamicElem, 'val');
        fc.dynamicFilter = { type, value: val ?? undefined };
      }

      // Top10 filter
      const top10Elem = fcElem.top10;
      if (top10Elem) {
        const top = attr(top10Elem, 'top', '1') !== '0';
        const percent = attr(top10Elem, 'percent', '0') === '1';
        const val = parseInt(attr(top10Elem, 'val', '10'), 10);
        fc.top10Filter = { top, percent, val };
      }

      // Add to autoFilter if not already there
      if (!autoFilter.getFilterColumn(colId)) {
        autoFilter.filterColumns.set(colId, fc);
      } else {
        // Merge loaded data
        autoFilter.filterColumns.set(colId, fc);
      }
    }

    // Load sort state
    const sortStateElem = afElem.sortState;
    if (sortStateElem) {
      const sortRef = attr(sortStateElem, 'ref', '');
      const sortConditions = ensureArray(sortStateElem.sortCondition);
      if (sortConditions.length > 0) {
        const condRef = attr(sortConditions[0], 'ref', '');
        const descending = attr(sortConditions[0], 'descending', '');
        const colIndex = this._calculateSortColumnIndex(ref, condRef);
        autoFilter.sortState = {
          ref: sortRef || ref,
          columnIndex: colIndex,
          ascending: descending !== '1',
        };
      }
    }
  }

  private _calculateSortColumnIndex(filterRange: string, sortCondRef: string): number {
    if (!filterRange || !sortCondRef) return 0;

    const filterCol = this._colFromRef(filterRange.split(':')[0]);
    const sortCol = this._colFromRef(sortCondRef.split(':')[0]);
    return sortCol - filterCol;
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
}
