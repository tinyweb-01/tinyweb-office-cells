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
  
  // ─── Border-clipping regression tests ───────────────────────────────────
  //
  // Bug: with `border-collapse:collapse`, the table's outer borders extend
  // ~1px outside the sum of column widths. When the wrapper <div> is sized
  // to exactly `totalTableWidth`, puppeteer's `boundingBox`-based screenshot
  // crops the right (and bottom) borders. Fix: wrap the table in an
  // inline-block container with a small padding/buffer so the collapsed
  // outer borders are never clipped.
  
  describe('worksheetToHtml — outer border clipping', () => {
    function buildBorderedTable(): Worksheet {
      const ws = new Worksheet('Bordered');
      // 2x3 table with thin borders on every cell
      for (let r = 1; r <= 2; r++) {
        for (let c = 1; c <= 3; c++) {
          const ref = String.fromCharCode(64 + c) + r; // A1,B1,...
          const cell = ws.cells.get(ref);
          cell.value = ref;
          cell.style.borders.top.lineStyle = 'thin';
          cell.style.borders.bottom.lineStyle = 'thin';
          cell.style.borders.left.lineStyle = 'thin';
          cell.style.borders.right.lineStyle = 'thin';
        }
      }
      return ws;
    }
  
    it('wraps the table in an inline-block container so outer borders are not clipped', () => {
      const ws = buildBorderedTable();
      const html = worksheetToHtml(ws);
      // The table must be wrapped in a <div> (so puppeteer's boundingBox
      // includes the wrapper's padding / buffer rather than the table's
      // exact column-sum width).
      expect(html.startsWith('<div')).toBe(true);
      expect(html).toMatch(/display:inline-block/);
    });
  
    it('wrapper width/padding is strictly greater than the sum of column widths', () => {
      const ws = buildBorderedTable();
      const html = worksheetToHtml(ws);
  
      // Extract table inner width from `width:NNNpx` inside the <table ...> tag
      const tableMatch = html.match(/<table[^>]*width:(\d+)px/);
      expect(tableMatch).not.toBeNull();
      const tableWidth = parseInt(tableMatch![1], 10);
  
      // The wrapper must reserve space for the right-edge border:
      //   either via explicit width  > tableWidth
      //   or via padding-right > 0.
      const wrapperWidthMatch = html.match(/<div[^>]*width:(\d+)px/);
      const padRightMatch = html.match(/<div[^>]*padding-right:(\d+)px/);
  
      const wrapperWidth = wrapperWidthMatch ? parseInt(wrapperWidthMatch[1], 10) : 0;
      const padRight = padRightMatch ? parseInt(padRightMatch[1], 10) : 0;
  
      const reserved = wrapperWidth > 0 ? wrapperWidth - tableWidth : padRight;
      expect(reserved).toBeGreaterThanOrEqual(1);
    });
  
    it('still wraps even when there are no drawings (regression)', () => {
      const ws = buildBorderedTable();
      expect((ws as any)._drawings).toBeFalsy();
      const html = worksheetToHtml(ws);
      expect(html.startsWith('<div')).toBe(true);
      expect(html).toContain('<table');
    });
  
    it('fullPage output still embeds the wrapped table', () => {
      const ws = buildBorderedTable();
      const html = worksheetToHtml(ws, { fullPage: true });
      expect(html).toContain('<!DOCTYPE html>');
      // body contains the wrapper div
      expect(html).toMatch(/<body>[\s\S]*<div[\s\S]*<table/);
    });
  
    it('mirrors right neighbour border-left as own border-right (collapsed-edge fix)', () => {
      // Real bug from sample.xlsx 表紙 sheet: only "left" border was set on
      // each cell; the rightmost column had no neighbour to provide the
      // right edge, leaving the table visually open on the right.
      const ws = new Worksheet('MirrorBorder');
      // 1-row, 2-col table; only cell A1 has border-left, only cell B1 has
      // border-left (= shared edge between A1 and B1). Neither has border-right.
      const a = ws.cells.get('A1');
      a.value = 'A';
      a.style.borders.left.lineStyle = 'thin';
      const b = ws.cells.get('B1');
      b.value = 'B';
      b.style.borders.left.lineStyle = 'thin';
      // Note: nothing has border-right defined at all.
  
      const html = worksheetToHtml(ws);
      // After the fix, cell A1 should mirror B1's border-left into its own
      // border-right, so the rendered HTML must contain a `border-right`.
      expect(html).toMatch(/border-right:\s*1px/);
    });

    it('uses inner-cell border-right for merged ranges (sample.xlsx 表紙 bug)', () => {
      // Real bug: a merged range like F19:J19 (colspan=5) had the right
      // border on cell J19 (the last cell *inside* the merge), not on the
      // anchor F19. The renderer used to read borders only from the anchor
      // cell, so border-right was lost — leaving the rightmost cell of the
      // metadata table visually open on the right.
      const ws = new Worksheet('MergedRightBorder');
      ws._mergedCells.push('A1:E1');
      const a = ws.cells.get('A1');
      a.value = 'merged value';
      a.style.borders.left.lineStyle = 'thin';
      a.style.borders.top.lineStyle = 'thin';
      a.style.borders.bottom.lineStyle = 'thin';
      // anchor cell deliberately has NO right border.
      const e = ws.cells.get('E1');
      e.style.borders.right.lineStyle = 'thin';

      const html = worksheetToHtml(ws);
      // The merged <td colspan="5"> for "merged value" must carry border-right
      // sourced from E1.
      const mergedTd = html.match(/<td[^>]*colspan="5"[^>]*>merged value<\/td>/);
      expect(mergedTd).not.toBeNull();
      expect(mergedTd![0]).toMatch(/border-right:\s*1px/);
    });

    it('rendered PNG is wide enough to include the right border (no clipping)', async () => {
      const ws = buildBorderedTable();
      const png = await worksheetToPng(ws, { viewportWidth: 400 });
      expect(Buffer.isBuffer(png)).toBe(true);
  
      // PNG width is stored as a 4-byte big-endian integer at offset 16.
      const pngWidth =
        (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
  
      // Compute expected minimum width: 3 columns × default width (~59px)
      // + body padding (8 left + 8 right) + buffer (≥1) + viewport padding (20).
      // We don't need an exact value — just assert the screenshot is wider
      // than the bare column sum, which proves the wrapper buffer survived
      // through to the rasterized output.
      const html = worksheetToHtml(ws);
      const tableMatch = html.match(/<table[^>]*width:(\d+)px/);
      const tableWidth = tableMatch ? parseInt(tableMatch[1], 10) : 0;
  
      expect(pngWidth).toBeGreaterThan(tableWidth);
    }, 30000);
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
