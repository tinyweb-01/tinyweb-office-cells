/**
 * Unit tests for rendering module: worksheetToHtml, worksheetToPng
 * TDD — tests written first, implementation follows.
 */
import { Worksheet } from '../src/core/Worksheet';
import { worksheetToHtml, worksheetToPng } from '../src/rendering';

// ─── Helper: build a small worksheet with data + styles ─────────────────

function buildTestSheet(): Worksheet {
  const ws = new Worksheet('TestSheet');
  // Header row (bold, blue fill, white font)
  const h1 = ws.cells.get('A1');
  h1.value = 'Name';
  h1.style.font.bold = true;
  h1.style.font.color = 'FFFFFFFF';
  h1.style.fill.foregroundColor = 'FF205EB5';
  h1.style.fill.patternType = 'solid';

  const h2 = ws.cells.get('B1');
  h2.value = 'Score';
  h2.style.font.bold = true;
  h2.style.font.color = 'FFFFFFFF';
  h2.style.fill.foregroundColor = 'FF205EB5';
  h2.style.fill.patternType = 'solid';

  // Data rows
  const a2 = ws.cells.get('A2');
  a2.value = 'Alice';

  const b2 = ws.cells.get('B2');
  b2.value = 95;

  const a3 = ws.cells.get('A3');
  a3.value = 'Bob';
  a3.style.font.strikethrough = true;

  const b3 = ws.cells.get('B3');
  b3.value = 80;
  b3.style.font.color = 'FFFF0000'; // red

  return ws;
}

function buildMergedSheet(): Worksheet {
  const ws = new Worksheet('MergedSheet');
  ws.cells.get('A1').value = 'Merged Title';
  // merge(firstRow, firstCol, totalRows, totalCols) — 0-based
  ws.cells.merge(0, 0, 1, 3); // A1:C1
  ws.cells.get('A2').value = 'Col1';
  ws.cells.get('B2').value = 'Col2';
  ws.cells.get('C2').value = 'Col3';
  return ws;
}

// ─── worksheetToHtml ────────────────────────────────────────────────────

describe('worksheetToHtml', () => {
  it('returns a valid HTML string with <table>', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html).toContain('<table');
    expect(html).toContain('</table>');
  });

  it('includes cell values in the output', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html).toContain('Name');
    expect(html).toContain('Score');
    expect(html).toContain('Alice');
    expect(html).toContain('95');
    expect(html).toContain('Bob');
    expect(html).toContain('80');
  });

  it('applies bold font style', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html).toMatch(/font-weight:\s*bold/);
  });

  it('applies fill background color', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html.toLowerCase()).toContain('background-color');
    expect(html.toLowerCase()).toContain('205eb5');
  });

  it('applies font color', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    // White font in header
    expect(html.toLowerCase()).toContain('#ffffff');
    // Red font for B3
    expect(html.toLowerCase()).toContain('#ff0000');
  });

  it('applies strikethrough', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html).toContain('line-through');
  });

  it('handles merged cells with colspan/rowspan', () => {
    const ws = buildMergedSheet();
    const html = worksheetToHtml(ws);
    expect(html).toContain('colspan="3"');
    expect(html).toContain('Merged Title');
  });

  it('returns complete HTML document when fullPage option is true', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws, { fullPage: true });
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('<html');
    expect(html).toContain('</html>');
  });

  it('returns only table fragment by default', () => {
    const ws = buildTestSheet();
    const html = worksheetToHtml(ws);
    expect(html).not.toContain('<!DOCTYPE html>');
  });

  it('handles empty worksheet', () => {
    const ws = new Worksheet('Empty');
    const html = worksheetToHtml(ws);
    expect(html).toContain('<table');
    expect(html).toContain('</table>');
  });

  it('applies column widths when set', () => {
    const ws = buildTestSheet();
    ws._columnWidths[1] = 20;
    ws._columnWidths[2] = 15;
    const html = worksheetToHtml(ws);
    expect(html).toContain('width:');
  });

  it('applies row heights when set', () => {
    const ws = buildTestSheet();
    ws._rowHeights[1] = 30;
    const html = worksheetToHtml(ws);
    expect(html).toContain('height:');
  });

  it('applies font size', () => {
    const ws = new Worksheet('FontSize');
    const c = ws.cells.get('A1');
    c.value = 'Big';
    c.style.font.size = 18;
    const html = worksheetToHtml(ws);
    expect(html).toContain('font-size:');
  });

  it('applies italic font', () => {
    const ws = new Worksheet('Italic');
    const c = ws.cells.get('A1');
    c.value = 'Slanted';
    c.style.font.italic = true;
    const html = worksheetToHtml(ws);
    expect(html).toContain('font-style:italic');
  });

  it('applies border styles', () => {
    const ws = new Worksheet('Borders');
    const c = ws.cells.get('A1');
    c.value = 'Boxed';
    c.style.borders.top.lineStyle = 'thin';
    c.style.borders.bottom.lineStyle = 'thin';
    const html = worksheetToHtml(ws);
    expect(html).toContain('border');
  });
});

// ─── worksheetToPng ─────────────────────────────────────────────────────

describe('worksheetToPng', () => {
  it('returns a PNG buffer', async () => {
    const ws = buildTestSheet();
    const png = await worksheetToPng(ws);
    expect(Buffer.isBuffer(png)).toBe(true);
    // PNG magic bytes: 0x89 0x50 0x4E 0x47
    expect(png[0]).toBe(0x89);
    expect(png[1]).toBe(0x50);
    expect(png[2]).toBe(0x4E);
    expect(png[3]).toBe(0x47);
  }, 30000);

  it('accepts viewport width option', async () => {
    const ws = buildTestSheet();
    const png = await worksheetToPng(ws, { viewportWidth: 800 });
    expect(Buffer.isBuffer(png)).toBe(true);
    expect(png.length).toBeGreaterThan(0);
  }, 30000);

  it('handles empty worksheet', async () => {
    const ws = new Worksheet('Empty');
    const png = await worksheetToPng(ws);
    expect(Buffer.isBuffer(png)).toBe(true);
    expect(png.length).toBeGreaterThan(0);
  }, 30000);

  it('handles merged cells in PNG', async () => {
    const ws = buildMergedSheet();
    const png = await worksheetToPng(ws);
    expect(Buffer.isBuffer(png)).toBe(true);
    expect(png[0]).toBe(0x89);
  }, 30000);
});

// ─── Range-based rendering tests ────────────────────────────────────────

describe('worksheetToHtml with range option', () => {
  function build3x3Sheet(): Worksheet {
    const ws = new Worksheet('Range');
    // 3x3 grid: A1..C3
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 3; c++) {
        const ref = String.fromCharCode(64 + c) + r; // A1, B1, ...
        ws.cells.get(ref).value = `${ref}`;
      }
    }
    return ws;
  }

  it('renders only cells within the specified range', () => {
    const ws = build3x3Sheet();
    const html = worksheetToHtml(ws, { range: 'B2:C3' });
    expect(html).toContain('B2');
    expect(html).toContain('C3');
    expect(html).not.toContain('A1');
    expect(html).not.toContain('A2');
    expect(html).not.toContain('B1');
  });

  it('renders correct number of rows and columns for range', () => {
    const ws = build3x3Sheet();
    const html = worksheetToHtml(ws, { range: 'A1:B2' });
    // Should have 2 rows, 2 cells each
    const trCount = (html.match(/<tr/g) || []).length;
    const tdCount = (html.match(/<td/g) || []).length;
    expect(trCount).toBe(2);
    expect(tdCount).toBe(4);
  });

  it('single-cell range renders one cell', () => {
    const ws = build3x3Sheet();
    const html = worksheetToHtml(ws, { range: 'B2:B2' });
    const tdCount = (html.match(/<td/g) || []).length;
    expect(tdCount).toBe(1);
    expect(html).toContain('B2');
  });

  it('falls back to full sheet when range is invalid', () => {
    const ws = build3x3Sheet();
    const htmlFull = worksheetToHtml(ws);
    const htmlBad = worksheetToHtml(ws, { range: 'INVALID' });
    expect(htmlBad).toBe(htmlFull);
  });

  it('range works with fullPage option', () => {
    const ws = build3x3Sheet();
    const html = worksheetToHtml(ws, { range: 'A1:A1', fullPage: true });
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('A1');
    expect(html).not.toContain('B1');
  });
});
