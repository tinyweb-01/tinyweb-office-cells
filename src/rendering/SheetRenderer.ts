/**
 * SheetRenderer — converts a Worksheet into an HTML table string
 * with full styling (fill, font, borders, merges, column widths, row heights).
 */

import { Worksheet } from '../core/Worksheet';
import { Cells } from '../core/Cells';
import { Cell } from '../core/Cell';

// ─── Types ──────────────────────────────────────────────────────────────

export interface HtmlRenderOptions {
  /** Wrap table in full HTML document (default: false) */
  fullPage?: boolean;
  /** Default font family (default: 'Calibri, Arial, sans-serif') */
  defaultFont?: string;
  /** Render only cells within this A1 range, e.g. "A1:D10" */
  range?: string;
}

export interface PngRenderOptions extends HtmlRenderOptions {
  /** Viewport width in pixels (default: 1200) */
  viewportWidth?: number;
}

// ─── Merge helpers ──────────────────────────────────────────────────────

interface MergeInfo {
  rowSpan: number;
  colSpan: number;
}

function parseMergedCells(mergedCells: string[]): {
  origins: Map<string, MergeInfo>;  // "row,col" → span info
  hidden: Set<string>;              // "row,col" cells that are covered
} {
  const origins = new Map<string, MergeInfo>();
  const hidden = new Set<string>();

  for (const range of mergedCells) {
    // Range format: "A1:C3" or similar from _mergedCells
    const parts = range.split(':');
    if (parts.length !== 2) continue;
    const tl = parseA1(parts[0]);
    const br = parseA1(parts[1]);
    if (!tl || !br) continue;

    const rowSpan = br.row - tl.row + 1;
    const colSpan = br.col - tl.col + 1;
    origins.set(`${tl.row},${tl.col}`, { rowSpan, colSpan });

    for (let r = tl.row; r <= br.row; r++) {
      for (let c = tl.col; c <= br.col; c++) {
        if (r === tl.row && c === tl.col) continue;
        hidden.add(`${r},${c}`);
      }
    }
  }
  return { origins, hidden };
}

function parseA1(ref: string): { row: number; col: number } | null {
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  const col = colFromLetter(m[1].toUpperCase());
  const row = parseInt(m[2], 10);
  return { row, col };
}

function colFromLetter(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function colToLetter(col: number): string {
  let s = '';
  let c = col;
  while (c > 0) {
    const rem = ((c - 1) % 26);
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

// ─── Color helpers ──────────────────────────────────────────────────────

function argbToHex(argb: string): string | null {
  if (!argb || argb.length < 6) return null;
  const hex = argb.slice(-6);
  return '#' + hex;
}

function isDefaultFill(color: string): boolean {
  return !color || color === 'FFFFFFFF' || color === '00000000';
}

function shouldRenderFill(patternType: string, foregroundColor: string): boolean {
  // Skip no fill
  if (!patternType || patternType === 'none') return false;
  // Skip gray125 (Excel default fill index 1 — not a visible fill)
  if (patternType === 'gray125') return false;
  // Skip default/unset foreground colors
  if (isDefaultFill(foregroundColor)) return false;
  // Skip black foreground — usually an unset/default value, not intentional black fill
  if (foregroundColor === 'FF000000') return false;
  return true;
}

function isDefaultFontColor(color: string): boolean {
  return !color || color === 'FF000000';
}

// ─── Border style mapping ───────────────────────────────────────────────

function borderToCss(lineStyle: string, color?: string): string | null {
  if (!lineStyle || lineStyle === 'none') return null;
  const widthMap: Record<string, string> = {
    thin: '1px', medium: '2px', thick: '3px',
    hair: '1px', dotted: '1px', dashed: '1px',
    dashDot: '1px', dashDotDot: '1px',
    mediumDashed: '2px', mediumDashDot: '2px', mediumDashDotDot: '2px',
    slantDashDot: '2px', double: '3px',
  };
  const styleMap: Record<string, string> = {
    thin: 'solid', medium: 'solid', thick: 'solid',
    hair: 'solid', dotted: 'dotted', dashed: 'dashed',
    dashDot: 'dashed', dashDotDot: 'dashed',
    mediumDashed: 'dashed', mediumDashDot: 'dashed', mediumDashDotDot: 'dashed',
    slantDashDot: 'dashed', double: 'double',
  };
  const w = widthMap[lineStyle] || '1px';
  const s = styleMap[lineStyle] || 'solid';
  const c = color ? argbToHex(color) || '#000000' : '#000000';
  return `${w} ${s} ${c}`;
}

// ─── Main render function ───────────────────────────────────────────────

export function worksheetToHtml(ws: Worksheet, options: HtmlRenderOptions = {}): string {
  const cells = ws.cells;
  const cellMap = cells._cells;
  const { origins, hidden } = parseMergedCells(ws._mergedCells || []);

  // Parse optional range bounds
  let minRow = 1, minCol = 1, maxRow = 0, maxCol = 0;
  if (options.range) {
    const rangeParts = options.range.split(':');
    if (rangeParts.length === 2) {
      const tl = parseA1(rangeParts[0]);
      const br = parseA1(rangeParts[1]);
      if (tl && br) {
        minRow = tl.row; minCol = tl.col;
        maxRow = br.row; maxCol = br.col;
      }
    }
  }

  // Auto-detect bounds if no range or invalid range
  if (maxRow === 0) {
    for (const ref of cellMap.keys()) {
      const p = parseA1(ref);
      if (p) {
        if (p.row > maxRow) maxRow = p.row;
        if (p.col > maxCol) maxCol = p.col;
      }
    }
    // Also check merged cell spans
    for (const range of (ws._mergedCells || [])) {
      const parts = range.split(':');
      if (parts.length === 2) {
        const br = parseA1(parts[1]);
        if (br) {
          if (br.row > maxRow) maxRow = br.row;
          if (br.col > maxCol) maxCol = br.col;
        }
      }
    }
  }

  // Build colgroup for column widths
  let colgroup = '';
  const colWidths = ws._columnWidths || {};
  if (Object.keys(colWidths).length > 0) {
    colgroup = '<colgroup>';
    for (let c = minCol; c <= maxCol; c++) {
      const w = colWidths[c];
      if (w) {
        // Excel width units ≈ 7px per unit
        colgroup += `<col style="width:${Math.round(w * 7)}px">`;
      } else {
        colgroup += '<col>';
      }
    }
    colgroup += '</colgroup>';
  }

  // Build rows
  const rowHeights = ws._rowHeights || {};
  let tbody = '';
  for (let r = minRow; r <= maxRow; r++) {
    const rh = rowHeights[r];
    const rowStyle = rh ? ` style="height:${Math.round(rh * 1.33)}px"` : '';
    tbody += `<tr${rowStyle}>`;

    for (let c = minCol; c <= maxCol; c++) {
      const key = `${r},${c}`;
      if (hidden.has(key)) continue;

      const ref = `${colToLetter(c)}${r}`;
      const cell = cellMap.get(ref);
      const merge = origins.get(key);

      const styles: string[] = [];
      let spanAttr = '';

      // Merge spans
      if (merge) {
        if (merge.colSpan > 1) spanAttr += ` colspan="${merge.colSpan}"`;
        if (merge.rowSpan > 1) spanAttr += ` rowspan="${merge.rowSpan}"`;
      }

      if (cell) {
        const style = cell.style;
        const font = style.font;
        const fill = style.fill;
        const borders = style.borders;

        // Font styles
        if (font.bold) styles.push('font-weight:bold');
        if (font.italic) styles.push('font-style:italic');
        if (font.strikethrough) styles.push('text-decoration:line-through');
        if (font.size && font.size !== 11) styles.push(`font-size:${font.size}pt`);
        if (!isDefaultFontColor(font.color)) {
          const hex = argbToHex(font.color);
          if (hex) styles.push(`color:${hex}`);
        }

        // Fill
        if (shouldRenderFill(fill.patternType, fill.foregroundColor)) {
          const hex = '#' + fill.foregroundColor.slice(-6);
          styles.push(`background-color:${hex}`);
        }

        // Borders
        if (borders) {
          const bt = borderToCss(borders.top?.lineStyle, borders.top?.color);
          const bb = borderToCss(borders.bottom?.lineStyle, borders.bottom?.color);
          const bl = borderToCss(borders.left?.lineStyle, borders.left?.color);
          const br2 = borderToCss(borders.right?.lineStyle, borders.right?.color);
          if (bt) styles.push(`border-top:${bt}`);
          if (bb) styles.push(`border-bottom:${bb}`);
          if (bl) styles.push(`border-left:${bl}`);
          if (br2) styles.push(`border-right:${br2}`);
        }
      }

      const styleAttr = styles.length ? ` style="${styles.join(';')}"` : '';
      const value = cell?.value != null ? escapeHtml(String(cell.value)) : '';
      tbody += `<td${spanAttr}${styleAttr}>${value}</td>`;
    }
    tbody += '</tr>';
  }

  const tableHtml = `<table style="border-collapse:collapse;font-family:${options.defaultFont || 'Calibri,Arial,sans-serif'}">${colgroup}${tbody}</table>`;

  if (options.fullPage) {
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{margin:0;padding:8px;background:#fff}td{padding:2px 6px;vertical-align:middle;background-color:#fff}</style></head><body>${tableHtml}</body></html>`;
  }

  return tableHtml;
}

// ─── HTML escape ────────────────────────────────────────────────────────

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
