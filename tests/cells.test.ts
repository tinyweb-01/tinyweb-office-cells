/**
 * Unit tests for src/core/Cells.ts
 */
import { Cells, WorksheetLike } from '../src/core/Cells';
import { Cell } from '../src/core/Cell';

/** Helper: create a Cells instance backed by a minimal WorksheetLike. */
function makeCells(): Cells {
  const wsLike: WorksheetLike = {
    _mergedCells: [],
    _rowHeights: {},
    _columnWidths: {},
    _hiddenRows: new Set(),
    _hiddenColumns: new Set(),
  };
  return new Cells(wsLike);
}

// ---------------------------------------------------------------------------
// Coordinate conversion (static)
// ---------------------------------------------------------------------------
describe('Cells – coordinate helpers', () => {
  it('columnIndexFromString A=1, Z=26, AA=27', () => {
    expect(Cells.columnIndexFromString('A')).toBe(1);
    expect(Cells.columnIndexFromString('Z')).toBe(26);
    expect(Cells.columnIndexFromString('AA')).toBe(27);
    expect(Cells.columnIndexFromString('AZ')).toBe(52);
  });

  it('columnLetterFromIndex 1=A, 26=Z, 27=AA', () => {
    expect(Cells.columnLetterFromIndex(1)).toBe('A');
    expect(Cells.columnLetterFromIndex(26)).toBe('Z');
    expect(Cells.columnLetterFromIndex(27)).toBe('AA');
  });

  it('coordinateFromString parses A1 → [1,1], B3 → [3,2]', () => {
    expect(Cells.coordinateFromString('A1')).toEqual([1, 1]);
    expect(Cells.coordinateFromString('B3')).toEqual([3, 2]);
    expect(Cells.coordinateFromString('AA10')).toEqual([10, 27]);
  });

  it('coordinateToString converts row/col to A1', () => {
    expect(Cells.coordinateToString(1, 1)).toBe('A1');
    expect(Cells.coordinateToString(3, 2)).toBe('B3');
    expect(Cells.coordinateToString(10, 27)).toBe('AA10');
  });

  it('roundtrip from/to string', () => {
    const ref = 'C5';
    const [r, c] = Cells.coordinateFromString(ref);
    expect(Cells.coordinateToString(r, c)).toBe(ref);
  });

  it('snake_case aliases', () => {
    expect(Cells.column_index_from_string('B')).toBe(2);
    expect(Cells.column_letter_from_index(2)).toBe('B');
    expect(Cells.coordinate_from_string('C3')).toEqual([3, 3]);
    expect(Cells.coordinate_to_string(3, 3)).toBe('C3');
  });
});

// ---------------------------------------------------------------------------
// Cell access
// ---------------------------------------------------------------------------
describe('Cells – cell access', () => {
  it('get() auto-creates cells', () => {
    const cells = makeCells();
    const c = cells.get('A1');
    expect(c).toBeInstanceOf(Cell);
    expect(c.value).toBeNull();
  });

  it('set() with value creates/updates cell', () => {
    const cells = makeCells();
    cells.set('B2', 42);
    expect(cells.get('B2').value).toBe(42);
  });

  it('set() with Cell instance replaces', () => {
    const cells = makeCells();
    const c = new Cell('hello');
    cells.set('C1', c);
    expect(cells.get('C1').value).toBe('hello');
  });

  it('cell(row, col) by number (1-based)', () => {
    const cells = makeCells();
    cells.cell(1, 1).value = 'A1';
    cells.cell(2, 3).value = 'C2';
    expect(cells.get('A1').value).toBe('A1');
    expect(cells.get('C2').value).toBe('C2');
  });

  it('cell(row, col) by letter', () => {
    const cells = makeCells();
    cells.cell(5, 'D').value = 99;
    expect(cells.get('D5').value).toBe(99);
  });

  it('get with tuple key [row, col]', () => {
    const cells = makeCells();
    cells.set([1, 1], 'tuple');
    expect(cells.get([1, 1]).value).toBe('tuple');
  });
});

// ---------------------------------------------------------------------------
// Collection methods
// ---------------------------------------------------------------------------
describe('Cells – collection', () => {
  it('count / clear', () => {
    const cells = makeCells();
    cells.set('A1', 1);
    cells.set('B1', 2);
    expect(cells.count()).toBe(2);
    cells.clear();
    expect(cells.count()).toBe(0);
  });

  it('hasCell / deleteCell', () => {
    const cells = makeCells();
    cells.set('A1', 'x');
    expect(cells.hasCell('A1')).toBe(true);
    cells.deleteCell('A1');
    expect(cells.hasCell('A1')).toBe(false);
  });

  it('maxDataRow / maxDataColumn', () => {
    const cells = makeCells();
    cells.set('C5', 1);
    cells.set('A1', 2);
    expect(cells.maxDataRow).toBe(5);
    expect(cells.maxDataColumn).toBe(3);
  });

  it('length mirrors count', () => {
    const cells = makeCells();
    cells.set('A1', 1);
    expect(cells.length).toBe(1);
  });

  it('Symbol.iterator yields [ref, Cell] pairs', () => {
    const cells = makeCells();
    cells.set('A1', 10);
    cells.set('B2', 20);
    const entries = [...cells];
    expect(entries.length).toBe(2);
    expect(entries[0][0]).toBe('A1');
    expect(entries[0][1].value).toBe(10);
  });

  it('getAllCells returns copy of internal map', () => {
    const cells = makeCells();
    cells.set('A1', 1);
    const all = cells.getAllCells();
    expect(all.size).toBe(1);
    // mutating returned map should not affect original
    all.delete('A1');
    expect(cells.count()).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Named access aliases
// ---------------------------------------------------------------------------
describe('Cells – named access', () => {
  it('getCellByName / setCellByName', () => {
    const cells = makeCells();
    cells.setCellByName('D4', 'val');
    expect(cells.getCellByName('D4').value).toBe('val');
  });

  it('getCell / setCell (row, col)', () => {
    const cells = makeCells();
    cells.setCell(2, 3, 'rc');
    expect(cells.getCell(2, 3).value).toBe('rc');
  });
});

// ---------------------------------------------------------------------------
// Row/Column dimensions
// ---------------------------------------------------------------------------
describe('Cells – row/column dimensions', () => {
  it('setRowHeight / getRowHeight (0-based API)', () => {
    const cells = makeCells();
    cells.setRowHeight(0, 30);
    expect(cells.getRowHeight(0)).toBe(30);
    // Default height for unset row
    expect(cells.getRowHeight(5)).toBe(15);
  });

  it('setColumnWidth / getColumnWidth (0-based numeric)', () => {
    const cells = makeCells();
    cells.setColumnWidth(0, 20);
    expect(cells.getColumnWidth(0)).toBe(20);
    // Default width for unset column
    expect(cells.getColumnWidth(9)).toBe(8.43);
  });

  it('setColumnWidth / getColumnWidth (string letter)', () => {
    const cells = makeCells();
    cells.setColumnWidth('B', 15);
    expect(cells.getColumnWidth('B')).toBe(15);
  });

  it('rejects invalid row/height', () => {
    const cells = makeCells();
    expect(() => cells.setRowHeight(-1, 10)).toThrow();
    expect(() => cells.setRowHeight(0, 0)).toThrow();
    expect(() => cells.setColumnWidth(-1, 10)).toThrow();
  });
});

// ---------------------------------------------------------------------------
// Hide / unhide
// ---------------------------------------------------------------------------
describe('Cells – hide/unhide', () => {
  it('hideRow / unhideRow / isRowHidden', () => {
    const cells = makeCells();
    expect(cells.isRowHidden(0)).toBe(false);
    cells.hideRow(0);
    expect(cells.isRowHidden(0)).toBe(true);
    cells.unhideRow(0);
    expect(cells.isRowHidden(0)).toBe(false);
  });

  it('hideColumn / unhideColumn / isColumnHidden (number)', () => {
    const cells = makeCells();
    cells.hideColumn(2);
    expect(cells.isColumnHidden(2)).toBe(true);
    cells.unhideColumn(2);
    expect(cells.isColumnHidden(2)).toBe(false);
  });

  it('hideColumn / isColumnHidden (string)', () => {
    const cells = makeCells();
    cells.hideColumn('C');
    expect(cells.isColumnHidden('C')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Merge
// ---------------------------------------------------------------------------
describe('Cells – merge', () => {
  it('merge(0-based) adds to mergedCells', () => {
    const cells = makeCells();
    cells.merge(0, 0, 2, 3); // rows 0-1, cols 0-2 → A1:C2
    const merged = cells.getMergedCells();
    expect(merged).toContain('A1:C2');
  });

  it('unmerge removes from mergedCells', () => {
    const cells = makeCells();
    cells.merge(0, 0, 2, 3);
    cells.unmerge(0, 0, 2, 3);
    expect(cells.getMergedCells()).toHaveLength(0);
  });

  it('mergeRange / unmergeRange by A1 ref', () => {
    const cells = makeCells();
    cells.mergeRange('A1:D5');
    expect(cells.getMergedCells()).toContain('A1:D5');
    cells.unmergeRange('A1:D5');
    expect(cells.getMergedCells()).toHaveLength(0);
  });

  it('duplicate merge is ignored', () => {
    const cells = makeCells();
    cells.merge(0, 0, 1, 1);
    cells.merge(0, 0, 1, 1);
    expect(cells.getMergedCells()).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Range
// ---------------------------------------------------------------------------
describe('Cells – getRange', () => {
  it('returns 2D array of cells', () => {
    const cells = makeCells();
    cells.set('A1', 1);
    cells.set('B1', 2);
    cells.set('A2', 3);
    cells.set('B2', 4);
    const range = cells.getRange(1, 1, 2, 2);
    expect(range.length).toBe(2);
    expect(range[0].length).toBe(2);
    expect(range[0][0].value).toBe(1);
    expect(range[1][1].value).toBe(4);
  });
});

// ---------------------------------------------------------------------------
// iterRows
// ---------------------------------------------------------------------------
describe('Cells – iterRows', () => {
  it('iterates rows in order', () => {
    const cells = makeCells();
    cells.set('A1', 'a1');
    cells.set('B1', 'b1');
    cells.set('A2', 'a2');
    const rows = [...cells.iterRows()];
    expect(rows.length).toBe(2); // rows 1..2
    expect(rows[0].length).toBe(2); // cols A..B
    // First row, first cell
    expect((rows[0][0] as Cell).value).toBe('a1');
  });

  it('valuesOnly yields values', () => {
    const cells = makeCells();
    cells.set('A1', 10);
    cells.set('B1', 20);
    const rows = [...cells.iterRows({ valuesOnly: true })];
    expect(rows[0]).toEqual([10, 20]);
  });
});
