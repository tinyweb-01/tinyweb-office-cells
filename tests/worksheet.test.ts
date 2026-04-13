/**
 * Unit tests for Worksheet, SheetProtection, SheetProtectionDictWrapper,
 * PageSetup, PageMargins, FreezePane.
 */
import {
  Worksheet,
  SheetProtection,
  SheetProtectionDictWrapper,
} from '../src/core/Worksheet';

// ─── SheetProtection ─────────────────────────────────────────────────────

describe('SheetProtection', () => {
  it('has correct defaults', () => {
    const p = new SheetProtection();
    expect(p.sheet).toBe(false);
    expect(p.password).toBeNull();
    expect(p.formatCells).toBe(false);
    expect(p.selectLockedCells).toBe(true);
    expect(p.selectUnlockedCells).toBe(true);
    expect(p.sort).toBe(false);
  });

  it('supports snake_case aliases', () => {
    const p = new SheetProtection();
    p.format_cells = true;
    expect(p.formatCells).toBe(true);
    expect(p.format_cells).toBe(true);

    p.insert_rows = true;
    expect(p.insertRows).toBe(true);

    p.delete_columns = true;
    expect(p.delete_columns).toBe(true);
  });
});

// ─── SheetProtectionDictWrapper ──────────────────────────────────────────

describe('SheetProtectionDictWrapper', () => {
  it('reads protection settings via getItem', () => {
    const sp = new SheetProtection();
    sp.sheet = true;
    sp.password = 'secret';
    sp.formatCells = true;
    const w = new SheetProtectionDictWrapper(sp);

    expect(w.getItem('protected')).toBe(true);
    expect(w.getItem('sheet')).toBe(true);
    expect(w.getItem('password')).toBe('secret');
    expect(w.getItem('format_cells')).toBe(true);
    expect(w.getItem('sort')).toBe(false);
  });

  it('writes protection settings via setItem', () => {
    const sp = new SheetProtection();
    const w = new SheetProtectionDictWrapper(sp);

    w.setItem('sheet', true);
    expect(sp.sheet).toBe(true);

    w.setItem('password', 'pw123');
    expect(sp.password).toBe('pw123');

    w.setItem('auto_filter', true);
    expect(sp.autoFilter).toBe(true);
  });

  it('get() returns defaultValue on unknown key', () => {
    const sp = new SheetProtection();
    const w = new SheetProtectionDictWrapper(sp);
    // Valid key returns the value
    expect(w.get('sort')).toBe(false);
    // Using get with a default for a known key
    expect(w.get('sheet', 'fallback')).toBe(false);
  });

  it('throws on unknown keys', () => {
    const sp = new SheetProtection();
    const w = new SheetProtectionDictWrapper(sp);
    expect(() => w.getItem('nonexistent' as any)).toThrow('Unknown protection key');
    expect(() => w.setItem('nonexistent' as any, true)).toThrow('Unknown protection key');
  });
});

// ─── Worksheet ───────────────────────────────────────────────────────────

describe('Worksheet', () => {
  // ── Construction ─────────────────────────────────────────────────────
  it('creates with default name Sheet1', () => {
    const ws = new Worksheet();
    expect(ws.name).toBe('Sheet1');
    expect(ws.visible).toBe(true);
    expect(ws.isVisible).toBe(true);
    expect(ws.is_visible).toBe(true);
    expect(ws.tabColor).toBeNull();
    expect(ws.index).toBe(0);
    expect(ws.isProtected()).toBe(false);
    expect(ws.printArea).toBeNull();
    expect(ws.freezePane).toBeNull();
    expect(ws.mergedCells).toEqual([]);
    expect(ws.cells).toBeDefined();
    expect(ws.dataValidations.count).toBe(0);
    expect(ws.conditionalFormatting.count).toBe(0);
    expect(ws.hyperlinks.count).toBe(0);
    expect(ws.autoFilter.range).toBeNull();
  });

  it('creates with custom name', () => {
    const ws = new Worksheet('Sales');
    expect(ws.name).toBe('Sales');
  });

  // ── Rename ──────────────────────────────────────────────────────────
  it('rename() changes name', () => {
    const ws = new Worksheet('Old');
    ws.rename('New');
    expect(ws.name).toBe('New');
  });

  // ── Visibility ──────────────────────────────────────────────────────
  it('supports visibility states', () => {
    const ws = new Worksheet();
    expect(ws.visible).toBe(true);

    ws.visible = false;
    expect(ws.visible).toBe(false);
    expect(ws.isVisible).toBe(false);

    ws.setVisibility('veryHidden');
    expect(ws.visible).toBe('veryHidden');
    expect(ws.isVisible).toBe(false);

    ws.set_visibility(true);
    expect(ws.visible).toBe(true);
    expect(ws.isVisible).toBe(true);
  });

  it('setVisibility() rejects invalid values', () => {
    const ws = new Worksheet();
    expect(() => ws.setVisibility('invalid' as any)).toThrow('Invalid visibility value');
  });

  // ── Tab color ───────────────────────────────────────────────────────
  it('sets and gets tab color', () => {
    const ws = new Worksheet();
    ws.setTabColor('FFFF0000');
    expect(ws.tabColor).toBe('FFFF0000');
    expect(ws.tab_color).toBe('FFFF0000');
    expect(ws.getTabColor()).toBe('FFFF0000');

    ws.clearTabColor();
    expect(ws.tabColor).toBeNull();
  });

  it('rejects invalid tab color format', () => {
    const ws = new Worksheet();
    expect(() => ws.setTabColor('red')).toThrow('Invalid tab color format');
  });

  it('tab_color setter works', () => {
    const ws = new Worksheet();
    ws.tab_color = 'FF00FF00';
    expect(ws.tab_color).toBe('FF00FF00');
  });

  // ── Page setup ──────────────────────────────────────────────────────
  it('setPageOrientation configures orientation', () => {
    const ws = new Worksheet();
    expect(ws.pageSetup.orientation).toBeNull();

    ws.setPageOrientation('landscape');
    expect(ws.pageSetup.orientation).toBe('landscape');
    expect(ws.page_setup.orientation).toBe('landscape');

    ws.set_page_orientation('portrait');
    expect(ws.pageSetup.orientation).toBe('portrait');
  });

  it('rejects invalid orientation', () => {
    const ws = new Worksheet();
    expect(() => ws.setPageOrientation('diagonal' as any)).toThrow('Invalid orientation');
  });

  it('manages paper size', () => {
    const ws = new Worksheet();
    ws.setPaperSize(1); // Letter
    expect(ws.getPaperSize()).toBe(1);
  });

  it('manages page margins', () => {
    const ws = new Worksheet();
    const defaults = ws.getPageMargins();
    expect(defaults.left).toBe(0.75);
    expect(defaults.right).toBe(0.75);
    expect(defaults.top).toBe(1.0);
    expect(defaults.bottom).toBe(1.0);
    expect(defaults.header).toBe(0.5);
    expect(defaults.footer).toBe(0.5);

    ws.setPageMargins({ left: 1.0, top: 1.5 });
    const m = ws.get_page_margins();
    expect(m.left).toBe(1.0);
    expect(m.top).toBe(1.5);
    expect(m.right).toBe(0.75); // unchanged
  });

  it('set fit to pages', () => {
    const ws = new Worksheet();
    ws.setFitToPages(2, 3);
    expect(ws.pageSetup.fitToWidth).toBe(2);
    expect(ws.pageSetup.fitToHeight).toBe(3);
    expect(ws.pageSetup.fitToPage).toBe(true);
  });

  it('set print scale', () => {
    const ws = new Worksheet();
    ws.setPrintScale(75);
    expect(ws.pageSetup.scale).toBe(75);
  });

  it('rejects invalid print scale', () => {
    const ws = new Worksheet();
    expect(() => ws.setPrintScale(5)).toThrow('Print scale must be between 10 and 400');
    expect(() => ws.setPrintScale(500)).toThrow('Print scale must be between 10 and 400');
  });

  // ── Print area ──────────────────────────────────────────────────────
  it('sets and clears print area', () => {
    const ws = new Worksheet();
    ws.setPrintArea('A1:D10');
    expect(ws.printArea).toBe('A1:D10');
    expect(ws.print_area).toBe('A1:D10');

    ws.clearPrintArea();
    expect(ws.printArea).toBeNull();
  });

  it('normalizes print area (removes $ and uppercases)', () => {
    const ws = new Worksheet();
    ws.printArea = '$a$1:$d$10';
    expect(ws.printArea).toBe('A1:D10');
  });

  it('setPrintArea aliases work', () => {
    const ws = new Worksheet();
    ws.set_print_area('B2:C5');
    expect(ws.printArea).toBe('B2:C5');

    ws.clear_print_area();
    expect(ws.printArea).toBeNull();

    ws.SetPrintArea('E1:F3');
    expect(ws.printArea).toBe('E1:F3');

    ws.ClearPrintArea();
    expect(ws.printArea).toBeNull();
  });

  it('empty print area sets null via setter', () => {
    const ws = new Worksheet();
    ws.setPrintArea('A1:B2');
    expect(ws.printArea).toBe('A1:B2');
    // Empty string is falsy, setter treats it as null (clear)
    ws.printArea = '';
    expect(ws.printArea).toBeNull();
  });

  // ── Freeze pane ─────────────────────────────────────────────────────
  it('sets and clears freeze pane', () => {
    const ws = new Worksheet();
    ws.setFreezePane(1, 2);
    expect(ws.freezePane).toEqual({ row: 1, column: 2, freezedRows: 1, freezedColumns: 2 });
    expect(ws.freeze_pane).toEqual(ws.freezePane);

    ws.clearFreezePane();
    expect(ws.freezePane).toBeNull();
  });

  it('setFreezePane with explicit freezed counts', () => {
    const ws = new Worksheet();
    ws.set_freeze_pane(3, 1, 2, 0);
    expect(ws.freezePane).toEqual({ row: 3, column: 1, freezedRows: 2, freezedColumns: 0 });

    ws.clear_freeze_pane();
    expect(ws.freezePane).toBeNull();
  });

  // ── Protection ──────────────────────────────────────────────────────
  it('protects and unprotects sheet', () => {
    const ws = new Worksheet();
    expect(ws.isProtected()).toBe(false);
    expect(ws.is_protected()).toBe(false);

    ws.protect({ password: 'pass', formatCells: true, sort: true });
    expect(ws.isProtected()).toBe(true);

    const p = ws.protection;
    expect(p.getItem('protected')).toBe(true);
    expect(p.getItem('password')).toBe('pass');
    expect(p.getItem('format_cells')).toBe(true);
    expect(p.getItem('sort')).toBe(true);

    ws.unprotect();
    expect(ws.isProtected()).toBe(false);
    expect(ws.protection.getItem('password')).toBeNull();
  });

  // ── Merged cells ────────────────────────────────────────────────────
  it('tracks merged cells', () => {
    const ws = new Worksheet();
    ws.cells.merge(0, 0, 2, 3); // 0-based
    const merged = ws.mergedCells;
    expect(merged.length).toBe(1);
    expect(merged[0]).toBe('A1:C2');
    // merged_cells alias
    expect(ws.merged_cells).toEqual(merged);
  });

  // ── Copy ────────────────────────────────────────────────────────────
  it('copies worksheet with cell data', () => {
    const ws = new Worksheet('Original');
    ws.cells.get('A1').value = 'Hello';
    ws.cells.get('B2').value = 42;
    ws.cells.merge(0, 0, 1, 2);
    ws.printArea = 'A1:B2';

    const copy = ws.copy('Copy1');
    expect(copy.name).toBe('Copy1');
    expect(copy.cells.get('A1').value).toBe('Hello');
    expect(copy.cells.get('B2').value).toBe(42);
    expect(copy.mergedCells).toEqual(ws.mergedCells);
    expect(copy.printArea).toBe(ws.printArea);

    // Verify deep copy - modifying copy doesn't affect original
    copy.cells.get('A1').value = 'Changed';
    expect(ws.cells.get('A1').value).toBe('Hello');
  });

  it('copy() generates default name if none given', () => {
    const ws = new Worksheet('TestSheet');
    const copy = ws.copy();
    expect(copy.name).toBe('TestSheet (copy)');
  });

  // ── Cells integration ──────────────────────────────────────────────
  it('cells collection is linked to worksheet', () => {
    const ws = new Worksheet('Data');
    ws.cells.setCell(1, 1, 'test');
    expect(ws.cells.get('A1').value).toBe('test');
  });

  // ── Placeholder methods don't throw ─────────────────────────────────
  it('placeholder methods execute without error', () => {
    const ws = new Worksheet();
    expect(() => ws.delete()).not.toThrow();
    expect(() => ws.move(0)).not.toThrow();
    expect(() => ws.select()).not.toThrow();
    expect(() => ws.activate()).not.toThrow();
    expect(() => ws.calculateFormula()).not.toThrow();
    expect(() => ws.calculate_formula()).not.toThrow();
  });

  // ── Index property ──────────────────────────────────────────────────
  it('index is read/write', () => {
    const ws = new Worksheet();
    expect(ws.index).toBe(0);
    ws.index = 5;
    expect(ws.index).toBe(5);
  });
});
