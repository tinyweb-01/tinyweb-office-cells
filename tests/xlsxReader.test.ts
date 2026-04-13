/**
 * Phase 3 – XLSX Reader round-trip tests
 *
 * Pattern: create → save → load → compare
 *
 * These tests verify that the XmlLoader correctly reads back what XmlSaver writes.
 */

import { Workbook } from '../src/core/Workbook';
import { XmlLoader } from '../src/io/XmlLoader';

// Helper: create → save → load round-trip
async function roundTrip(setup: (wb: Workbook) => void): Promise<Workbook> {
  const wb = new Workbook();
  setup(wb);
  const buf = await wb.saveToBuffer();
  return XmlLoader.loadFromBuffer(buf);
}

describe('XLSX Reader – round-trip tests', () => {
  // ── 1. Basic cell values ──────────────────────────────────────────────

  it('RT-01: round-trips string, number, boolean, null cell values', async () => {
    const wb2 = await roundTrip((wb) => {
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 'Hello';
      ws.cells.get('B1').value = 42;
      ws.cells.get('C1').value = true;
      ws.cells.get('D1').value = null;
    });

    const ws = wb2.worksheets[0];
    expect(ws.cells.get('A1').value).toBe('Hello');
    expect(ws.cells.get('B1').value).toBe(42);
    expect(ws.cells.get('C1').value).toBe(true);
    // D1 was null → should not be present or be null
    expect(ws.cells.get('D1').value).toBeNull();
  });

  // ── 2. Multiple worksheets ────────────────────────────────────────────

  it('RT-02: round-trips multiple worksheets with names', async () => {
    const wb2 = await roundTrip((wb) => {
      wb.worksheets[0].name = 'Data';
      const ws2 = wb.addWorksheet('Summary');
      wb.worksheets[0].cells.get('A1').value = 'data-sheet';
      ws2.cells.get('A1').value = 'summary-sheet';
    });

    expect(wb2.worksheets.length).toBe(2);
    expect(wb2.worksheets[0].name).toBe('Data');
    expect(wb2.worksheets[1].name).toBe('Summary');
    expect(wb2.worksheets[0].cells.get('A1').value).toBe('data-sheet');
    expect(wb2.worksheets[1].cells.get('A1').value).toBe('summary-sheet');
  });

  // ── 3. Font styling ──────────────────────────────────────────────────

  it('RT-03: round-trips font bold, italic, color, size', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 'styled';
      cell.style.font.bold = true;
      cell.style.font.italic = true;
      cell.style.font.color = 'FFFF0000';
      cell.style.font.size = 14;
      cell.style.font.name = 'Arial';
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe('styled');
    expect(cell.style.font.bold).toBe(true);
    expect(cell.style.font.italic).toBe(true);
    expect(cell.style.font.color).toBe('FFFF0000');
    expect(cell.style.font.size).toBe(14);
    expect(cell.style.font.name).toBe('Arial');
  });

  // ── 4. Fill styling ──────────────────────────────────────────────────

  it('RT-04: round-trips solid fill color', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 'filled';
      cell.style.fill.setSolidFill('FF00FF00');
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe('filled');
    expect(cell.style.fill.patternType).toBe('solid');
    expect(cell.style.fill.foregroundColor).toBe('FF00FF00');
  });

  // ── 5. Border styling ────────────────────────────────────────────────

  it('RT-05: round-trips border styles', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 'bordered';
      cell.style.borders.setBorder('top', 'thin', 'FF000000');
      cell.style.borders.setBorder('bottom', 'double', 'FFFF0000');
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe('bordered');
    expect(cell.style.borders.top.lineStyle).toBe('thin');
    expect(cell.style.borders.top.color).toBe('FF000000');
    expect(cell.style.borders.bottom.lineStyle).toBe('double');
    expect(cell.style.borders.bottom.color).toBe('FFFF0000');
  });

  // ── 6. Number format ─────────────────────────────────────────────────

  it('RT-06: round-trips number format', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 1234.56;
      cell.style.numberFormat = '#,##0.00';
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe(1234.56);
    expect(cell.style.numberFormat).toBe('#,##0.00');
  });

  // ── 7. Alignment ─────────────────────────────────────────────────────

  it('RT-07: round-trips alignment properties', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 'aligned';
      cell.style.alignment.horizontal = 'center';
      cell.style.alignment.vertical = 'top';
      cell.style.alignment.wrapText = true;
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe('aligned');
    expect(cell.style.alignment.horizontal).toBe('center');
    expect(cell.style.alignment.vertical).toBe('top');
    expect(cell.style.alignment.wrapText).toBe(true);
  });

  // ── 8. Merged cells ──────────────────────────────────────────────────

  it('RT-08: round-trips merged cells', async () => {
    const wb2 = await roundTrip((wb) => {
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 'merged';
      ws.cells.merge(0, 0, 2, 3); // A1:C2
    });

    const ws = wb2.worksheets[0];
    expect(ws.cells.get('A1').value).toBe('merged');
    const merged = ws.cells.getMergedCells();
    expect(merged.length).toBeGreaterThanOrEqual(1);
    expect(merged[0]).toBe('A1:C2');
  });

  // ── 9. Column widths & row heights ────────────────────────────────────

  it('RT-09: round-trips column widths and row heights', async () => {
    const wb2 = await roundTrip((wb) => {
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 'wide';
      ws.cells.setColumnWidth(1, 25);  // Column A (1-based)
      ws.cells.setRowHeight(1, 30);    // Row 1 (1-based)
    });

    const ws = wb2.worksheets[0];
    expect(ws.cells.get('A1').value).toBe('wide');
    expect(ws.cells.getColumnWidth(1)).toBeCloseTo(25, 0);
    expect(ws.cells.getRowHeight(1)).toBeCloseTo(30, 0);
  });

  // ── 10. Freeze pane ──────────────────────────────────────────────────

  it('RT-10: round-trips freeze pane', async () => {
    const wb2 = await roundTrip((wb) => {
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 'frozen';
      ws.setFreezePane(1, 2); // Freeze row 1, columns A-B
    });

    const ws = wb2.worksheets[0];
    expect(ws.cells.get('A1').value).toBe('frozen');
    const fp = (ws as any)._freezePane;
    expect(fp).toBeTruthy();
    expect(fp.row).toBe(1);
    expect(fp.column).toBe(2);
  });

  // ── 11. Document properties ───────────────────────────────────────────

  it('RT-11: round-trips document properties', async () => {
    const wb2 = await roundTrip((wb) => {
      wb.documentProperties.title = 'Test Title';
      wb.documentProperties.creator = 'Test Author';
      wb.documentProperties.subject = 'Test Subject';
      wb.documentProperties.keywords = 'test, xlsx, reader';
      wb.documentProperties.description = 'A test description';
      wb.documentProperties.category = 'Testing';
      wb.worksheets[0].cells.get('A1').value = 'props';
    });

    expect(wb2.documentProperties.title).toBe('Test Title');
    expect(wb2.documentProperties.creator).toBe('Test Author');
    expect(wb2.documentProperties.subject).toBe('Test Subject');
    expect(wb2.documentProperties.keywords).toBe('test, xlsx, reader');
    expect(wb2.documentProperties.description).toBe('A test description');
    expect(wb2.documentProperties.category).toBe('Testing');
    expect(wb2.worksheets[0].cells.get('A1').value).toBe('props');
  });

  // ── Bonus: Workbook.load static ───────────────────────────────────────

  it('Workbook.loadFromBuffer works as static entry point', async () => {
    const wb = new Workbook();
    wb.worksheets[0].cells.get('A1').value = 'hello';
    const buf = await wb.saveToBuffer();
    const wb2 = await Workbook.loadFromBuffer(buf);
    expect(wb2.worksheets[0].cells.get('A1').value).toBe('hello');
  });

  // ── Formula ───────────────────────────────────────────────────────────

  it('RT-12: round-trips formulas', async () => {
    const wb2 = await roundTrip((wb) => {
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 10;
      ws.cells.get('A2').value = 20;
      ws.cells.get('A3').formula = '=SUM(A1:A2)';
    });

    const ws = wb2.worksheets[0];
    expect(ws).toBeTruthy();
    expect(ws.cells.get('A1').value).toBe(10);
    expect(ws.cells.get('A2').value).toBe(20);
    expect(ws.cells.get('A3').formula).toBe('=SUM(A1:A2)');
  });

  // ── Protection ────────────────────────────────────────────────────────

  it('RT-13: round-trips cell protection locked/hidden', async () => {
    const wb2 = await roundTrip((wb) => {
      const cell = wb.worksheets[0].cells.get('A1');
      cell.value = 'protected';
      cell.style.protection.locked = false;
      cell.style.protection.hidden = true;
    });

    const cell = wb2.worksheets[0].cells.get('A1');
    expect(cell.value).toBe('protected');
    expect(cell.style.protection.locked).toBe(false);
    expect(cell.style.protection.hidden).toBe(true);
  });
});
