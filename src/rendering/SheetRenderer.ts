/**
 * SheetRenderer — converts a Worksheet into an HTML table string
 * with full styling (fill, font, borders, merges, column widths, row heights).
 *
 * Also overlays DrawingML pictures and shapes (parsed by DrawingsLoader)
 * on top of the cell grid, and applies number-format conversion for
 * dates / times / numbers.
 */

import { Worksheet } from '../core/Worksheet';
import type {
  DrawingItem,
  DrawingPicture,
  DrawingShape,
  DrawingAnchor,
  ShapeParagraph,
} from '../io/DrawingsLoader';

// ─── Types ──────────────────────────────────────────────────────────────

export interface HtmlRenderOptions {
  /** Wrap table in full HTML document (default: false) */
  fullPage?: boolean;
  /** Default font family (default: 'Calibri, Arial, sans-serif') */
  defaultFont?: string;
  /** Render only cells within this A1 range, e.g. "A1:D10" */
  range?: string;
  /** If true, omit the drawing overlay layer (pictures + shapes). Default: false */
  skipDrawings?: boolean;
}

export interface PngRenderOptions extends HtmlRenderOptions {
  /** Viewport width in pixels (default: 1200) */
  viewportWidth?: number;
}

// ─── Constants ──────────────────────────────────────────────────────────

const PX_PER_UNIT = 7;          // px per Excel column-width unit (must match colgroup)
const DEFAULT_COL_WIDTH = 8.43; // Excel default
const DEFAULT_ROW_HEIGHT_PT = 15;
const PT_TO_PX = 1.33;
const EMU_PER_PX = 9525;
const CELL_PADDING_LR_PX = 6;   // matches td padding 2px 6px
const CELL_PADDING_TB_PX = 2;

// ─── Merge helpers ──────────────────────────────────────────────────────

interface MergeInfo {
  rowSpan: number;
  colSpan: number;
}

function parseMergedCells(mergedCells: string[]): {
  origins: Map<string, MergeInfo>;
  hidden: Set<string>;
} {
  const origins = new Map<string, MergeInfo>();
  const hidden = new Set<string>();

  for (const range of mergedCells) {
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
  if (!patternType || patternType === 'none') return false;
  if (patternType === 'gray125') return false;
  if (isDefaultFill(foregroundColor)) return false;
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

// ─── Number-format engine ───────────────────────────────────────────────

/**
 * Strip Excel format-section noise: color tags ([Red]), conditions ([>0]),
 * locale tags ([$-409]), backslash escapes, quoted literals, and trailing
 * underscores (column padding). Returns a sanitized format string still
 * preserving the date/time tokens.
 */
function sanitizeFormat(fmt: string): string {
  let s = fmt;
  // Remove [Color], [Locale], [Conditions]
  s = s.replace(/\[[^\]]*\]/g, '');
  // Remove escaped chars: \X -> X (drop the backslash)
  s = s.replace(/\\(.)/g, '$1');
  // Remove quoted literals "..."
  s = s.replace(/"[^"]*"/g, '');
  // Underscore + char (padding for parentheses) -> drop both
  s = s.replace(/_./g, '');
  // Asterisk + char (fill char) -> drop both
  s = s.replace(/\*./g, '');
  return s;
}

function isDateFormat(rawFmt: string): boolean {
  const fmt = sanitizeFormat(rawFmt);
  // Date / time tokens
  return /[ymdhsYMDHS]/.test(fmt) && !/^[#0.,%\s]+$/.test(fmt);
}

function pad(n: number, w: number): string {
  return String(n).padStart(w, '0');
}

const MONTHS_LONG = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];
const MONTHS_SHORT = MONTHS_LONG.map(m => m.slice(0, 3));
const DAYS_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAYS_SHORT = DAYS_LONG.map(d => d.slice(0, 3));

/**
 * Convert Excel serial number to JS Date.
 * Excel epoch is 1899-12-30 (1900 leap-year bug compensated).
 */
function excelSerialToDate(serial: number): Date {
  // Excel 1900-based: integer = days since 1899-12-30 in UTC.
  const ms = Math.round(serial * 86400000);
  return new Date(Date.UTC(1899, 11, 30) + ms);
}

/**
 * Format an Excel serial number as a date string per the supplied format.
 * Supports m, mm, mmm, mmmm, d, dd, ddd, dddd, yy, yyyy, h, hh, m, mm
 * (minutes when next to h or s), s, ss, AM/PM. Falls back to ISO if all
 * else fails.
 */
function formatDate(value: number, rawFmt: string): string {
  const fmt = sanitizeFormat(rawFmt).trim();
  if (!fmt) return String(value);

  const d = excelSerialToDate(value);
  const yyyy = d.getUTCFullYear();
  const mo = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  const dow = d.getUTCDay();
  const hh = d.getUTCHours();
  const mm = d.getUTCMinutes();
  const ss = d.getUTCSeconds();

  const has12hr = /\bAM\/PM\b|\ba\/p\b/i.test(fmt);
  const h12 = (() => { const h = hh % 12; return h === 0 ? 12 : h; })();

  let out = '';
  let i = 0;
  while (i < fmt.length) {
    const ch = fmt[i];
    // Match longest run of same letter (case insensitive)
    if (/[ymdhsap]/i.test(ch)) {
      let j = i;
      while (j < fmt.length && fmt[j].toLowerCase() === ch.toLowerCase()) j++;
      const run = fmt.slice(i, j).toLowerCase();
      // Disambiguate 'm/mm' as month vs minute: if previous non-space token
      // was h/hh or next is s/ss, treat as minute.
      const prev = out.slice(-3).toLowerCase();
      const next = fmt.slice(j).toLowerCase().replace(/^[\s:.,/-]+/, '');
      const followedByS = next.startsWith('s');
      const precededByH = /h\d*$/.test(prev) || /h\s*$/.test(prev) || /\d{1,2}$/.test(prev) && /h/.test(out.slice(-5).toLowerCase());

      if (run === 'm' || run === 'mm') {
        // Check if surrounded by hour/second context
        const looksLikeMinute = precededByH || followedByS ||
          /h+$/.test(fmt.slice(0, i).replace(/[\s:]+$/, '')) ||
          /^[\s:]*s/.test(fmt.slice(j));
        if (looksLikeMinute) {
          out += run === 'mm' ? pad(mm, 2) : String(mm);
          i = j; continue;
        }
      }

      switch (run) {
        case 'yy': out += pad(yyyy % 100, 2); break;
        case 'yyyy': case 'yyy': out += pad(yyyy, 4); break;
        case 'm': out += String(mo); break;
        case 'mm': out += pad(mo, 2); break;
        case 'mmm': out += MONTHS_SHORT[mo - 1]; break;
        case 'mmmm': out += MONTHS_LONG[mo - 1]; break;
        case 'mmmmm': out += MONTHS_LONG[mo - 1][0]; break;
        case 'd': out += String(day); break;
        case 'dd': out += pad(day, 2); break;
        case 'ddd': out += DAYS_SHORT[dow]; break;
        case 'dddd': out += DAYS_LONG[dow]; break;
        case 'h': out += String(has12hr ? h12 : hh); break;
        case 'hh': out += pad(has12hr ? h12 : hh, 2); break;
        case 's': out += String(ss); break;
        case 'ss': out += pad(ss, 2); break;
        case 'am/pm': case 'a/p':
          out += hh < 12 ? (run === 'a/p' ? 'A' : 'AM') : (run === 'a/p' ? 'P' : 'PM');
          break;
        default:
          // Fallback: emit raw
          out += run;
      }
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

function formatNumber(value: number, rawFmt: string): string {
  // Percent
  if (rawFmt.includes('%')) {
    const pct = value * 100;
    const decMatch = rawFmt.match(/\.(0+)/);
    const decimals = decMatch ? decMatch[1].length : 0;
    return pct.toFixed(decimals) + '%';
  }

  const hasThousandSep = /[#0],[#0]/.test(rawFmt);
  const decimalMatch = rawFmt.match(/\.(0+)/);
  const decimals = decimalMatch ? decimalMatch[1].length : 0;
  const absValue = Math.abs(value);
  let formatted: string;
  if (hasThousandSep) {
    formatted = absValue.toLocaleString('en-US', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping: true,
    });
  } else {
    formatted = absValue.toFixed(decimals);
  }
  if (value < 0) formatted = '-' + formatted;
  return formatted;
}

function formatCellValue(value: any, numberFormat: string): string {
  if (value == null) return '';
  // Date objects → ISO short
  if (value instanceof Date) {
    const y = value.getUTCFullYear();
    const m = pad(value.getUTCMonth() + 1, 2);
    const d = pad(value.getUTCDate(), 2);
    return `${m}/${d}/${y}`;
  }

  if (typeof value !== 'number') return String(value);

  if (!numberFormat || numberFormat === 'General' || numberFormat === '@') {
    // Use a clean numeric string (avoid "3.5" → "3,5" locale surprises)
    if (Number.isInteger(value)) return String(value);
    return String(value);
  }

  try {
    if (isDateFormat(numberFormat)) {
      return formatDate(value, numberFormat);
    }
    return formatNumber(value, numberFormat);
  } catch {
    return String(value);
  }
}

// ─── Drawing overlay helpers ────────────────────────────────────────────

function colWidthPx(ws: Worksheet, col: number): number {
  const w = ws._columnWidths?.[col];
  return Math.round((w ?? DEFAULT_COL_WIDTH) * PX_PER_UNIT);
}

function rowHeightPx(ws: Worksheet, row: number): number {
  const h = ws._rowHeights?.[row];
  return Math.round((h ?? DEFAULT_ROW_HEIGHT_PT) * PT_TO_PX);
}

/**
 * Convert a DrawingML anchor (cell + EMU offset, optionally to+ext) into
 * pixel coordinates relative to the table's top-left corner (col 1, row 1
 * of the rendered range).
 */
function anchorToPx(
  ws: Worksheet,
  anchor: DrawingAnchor,
  minRow: number,
  minCol: number,
  maxCol: number,
): { left: number; top: number; width: number; height: number } {
  // OOXML uses 0-based rows/cols
  const fromColXl = anchor.fromCol + 1;
  const fromRowXl = anchor.fromRow + 1;

  // Sum widths for cols [minCol, fromColXl-1]
  let left = 0;
  for (let c = minCol; c < fromColXl; c++) left += colWidthPx(ws, c);
  left += anchor.fromColOff / EMU_PER_PX;

  let top = 0;
  for (let r = minRow; r < fromRowXl; r++) top += rowHeightPx(ws, r);
  top += anchor.fromRowOff / EMU_PER_PX;

  let width: number;
  let height: number;

  if (anchor.toCol != null && anchor.toRow != null) {
    const toColXl = anchor.toCol + 1;
    const toRowXl = anchor.toRow + 1;
    let right = 0;
    for (let c = minCol; c < toColXl; c++) right += colWidthPx(ws, c);
    right += (anchor.toColOff || 0) / EMU_PER_PX;
    let bottom = 0;
    for (let r = minRow; r < toRowXl; r++) bottom += rowHeightPx(ws, r);
    bottom += (anchor.toRowOff || 0) / EMU_PER_PX;
    width = Math.max(1, right - left);
    height = Math.max(1, bottom - top);
  } else {
    width = Math.max(1, (anchor.cx || 0) / EMU_PER_PX);
    height = Math.max(1, (anchor.cy || 0) / EMU_PER_PX);
  }

  return { left: Math.round(left), top: Math.round(top), width: Math.round(width), height: Math.round(height) };
}

function geomToBorderRadius(geom: string, w: number, h: number): string {
  switch (geom) {
    case 'roundRect':
      return `${Math.min(w, h) * 0.1}px`;
    case 'ellipse':
      return '50%';
    default:
      return '0';
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function renderShapeText(paragraphs: ShapeParagraph[], defaultFont: string): string {
  if (!paragraphs || paragraphs.length === 0) return '';
  let html = '';
  for (const p of paragraphs) {
    const align = p.align || 'left';
    let inner = '';
    for (const r of p.runs) {
      const styles: string[] = [];
      if (r.bold) styles.push('font-weight:bold');
      if (r.italic) styles.push('font-style:italic');
      if (r.underline) styles.push('text-decoration:underline');
      if (r.strike) styles.push('text-decoration:line-through');
      if (r.size) styles.push(`font-size:${r.size}pt`);
      if (r.color) styles.push(`color:${r.color}`);
      if (r.font) styles.push(`font-family:${r.font},${defaultFont}`);
      const styleAttr = styles.length ? ` style="${styles.join(';')}"` : '';
      inner += `<span${styleAttr}>${escapeHtml(r.text)}</span>`;
    }
    if (!inner) inner = '&nbsp;';
    html += `<div style="text-align:${align};margin:0;line-height:1.2">${inner}</div>`;
  }
  return html;
}

function renderPicture(p: DrawingPicture, box: { left: number; top: number; width: number; height: number }): string {
  const transforms: string[] = [];
  if (p.rotation) transforms.push(`rotate(${p.rotation}deg)`);
  if (p.flipH) transforms.push('scaleX(-1)');
  if (p.flipV) transforms.push('scaleY(-1)');
  const trs = transforms.length ? `;transform:${transforms.join(' ')}` : '';
  return `<img src="${p.imageDataUri}" style="position:absolute;left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px;object-fit:fill${trs}"/>`;
}

function renderShape(
  s: DrawingShape,
  box: { left: number; top: number; width: number; height: number },
  defaultFont: string,
): string {
  const styles: string[] = [
    'position:absolute',
    `left:${box.left}px`,
    `top:${box.top}px`,
    `width:${box.width}px`,
    `height:${box.height}px`,
    'box-sizing:border-box',
    'overflow:hidden',
  ];
  if (s.fillColor) styles.push(`background-color:${s.fillColor}`);
  if (s.strokeColor) {
    styles.push(`border:${s.strokeWidth || 1}px solid ${s.strokeColor}`);
  }
  const radius = geomToBorderRadius(s.geom, box.width, box.height);
  if (radius !== '0') styles.push(`border-radius:${radius}`);
  if (s.rotation) styles.push(`transform:rotate(${s.rotation}deg)`);

  // Vertical alignment via flexbox
  const vAlign = s.vertAlign || 'middle';
  const justify = vAlign === 'top' ? 'flex-start' : vAlign === 'bottom' ? 'flex-end' : 'center';
  styles.push('display:flex', 'flex-direction:column', `justify-content:${justify}`);
  // Inner padding so text doesn't touch the border
  styles.push('padding:4px');

  const inner = renderShapeText(s.paragraphs, defaultFont);
  return `<div style="${styles.join(';')}">${inner}</div>`;
}

function renderDrawings(
  ws: Worksheet,
  drawings: DrawingItem[],
  minRow: number,
  minCol: number,
  maxRow: number,
  maxCol: number,
  defaultFont: string,
): string {
  if (!drawings || drawings.length === 0) return '';
  const parts: string[] = [];
  for (const d of drawings) {
    // Skip anchors clearly outside the rendered range
    const fromColXl = d.anchor.fromCol + 1;
    const fromRowXl = d.anchor.fromRow + 1;
    if (fromColXl > maxCol + 1 || fromRowXl > maxRow + 1) continue;

    const box = anchorToPx(ws, d.anchor, minRow, minCol, maxCol);
    if (d.type === 'pic') {
      parts.push(renderPicture(d, box));
    } else if (d.type === 'sp') {
      parts.push(renderShape(d, box, defaultFont));
    }
  }
  return parts.join('');
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
    // Also include drawings in bounds detection
    if (!options.skipDrawings) {
      const drawings: DrawingItem[] | undefined = (ws as any)._drawings;
      if (drawings) {
        for (const d of drawings) {
          const fc = d.anchor.fromCol + 1;
          const fr = d.anchor.fromRow + 1;
          if (fc > maxCol) maxCol = fc;
          if (fr > maxRow) maxRow = fr;
          if (d.anchor.toCol != null) {
            const tc = d.anchor.toCol + 1;
            if (tc > maxCol) maxCol = tc;
          }
          if (d.anchor.toRow != null) {
            const tr = d.anchor.toRow + 1;
            if (tr > maxRow) maxRow = tr;
          }
        }
      }
    }
  }
  if (maxRow < 1) maxRow = 1;
  if (maxCol < 1) maxCol = 1;

  // Build colgroup for column widths
  let colgroup = '';
  let totalTableWidth = 0;
  for (let c = minCol; c <= maxCol; c++) {
    totalTableWidth += colWidthPx(ws, c);
  }
  colgroup = '<colgroup>';
  for (let c = minCol; c <= maxCol; c++) {
    colgroup += `<col style="width:${colWidthPx(ws, c)}px">`;
  }
  colgroup += '</colgroup>';

  // Build rows
  let totalTableHeight = 0;
  let tbody = '';
  for (let r = minRow; r <= maxRow; r++) {
    const rh = rowHeightPx(ws, r);
    totalTableHeight += rh;
    tbody += `<tr style="height:${rh}px">`;

    for (let c = minCol; c <= maxCol; c++) {
      const key = `${r},${c}`;
      if (hidden.has(key)) continue;

      const ref = `${colToLetter(c)}${r}`;
      const cell = cellMap.get(ref);
      const merge = origins.get(key);

      const styles: string[] = [];
      let spanAttr = '';

      if (merge) {
        if (merge.colSpan > 1) spanAttr += ` colspan="${merge.colSpan}"`;
        if (merge.rowSpan > 1) spanAttr += ` rowspan="${merge.rowSpan}"`;
      }

      if (cell) {
        const style = cell.style;
        const font = style.font;
        const fill = style.fill;
        const borders = style.borders;

        if (font.bold) styles.push('font-weight:bold');
        if (font.italic) styles.push('font-style:italic');
        if (font.strikethrough) styles.push('text-decoration:line-through');
        if (font.size && font.size !== 11) styles.push(`font-size:${font.size}pt`);
        if (!isDefaultFontColor(font.color)) {
          const hex = argbToHex(font.color);
          if (hex) styles.push(`color:${hex}`);
        }

        if (shouldRenderFill(fill.patternType, fill.foregroundColor)) {
          const hex = '#' + fill.foregroundColor.slice(-6);
          styles.push(`background-color:${hex}`);
        }

        if (borders) {
          // For merged cells, each cell inside the merge can hold its own
          // borders. The visible perimeter borders of the merged range come
          // from the cells that actually sit on that perimeter, not from
          // the anchor cell. So:
          //   top    -> anchor cell (this one)
          //   left   -> anchor cell (this one)
          //   right  -> rightmost cell within the merge on the same row
          //   bottom -> bottom-most cell within the merge on the same column
          const colSpan = merge?.colSpan ?? 1;
          const rowSpan = merge?.rowSpan ?? 1;
          const rightInside = cellMap.get(`${colToLetter(c + colSpan - 1)}${r}`);
          const bottomInside = cellMap.get(`${colToLetter(c)}${r + rowSpan - 1}`);
          const rightBorders = (rightInside ?? cell).style.borders;
          const bottomBorders = (bottomInside ?? cell).style.borders;

          let bt = borderToCss(borders.top?.lineStyle, borders.top?.color);
          let bb = borderToCss(bottomBorders?.bottom?.lineStyle, bottomBorders?.bottom?.color);
          let bl = borderToCss(borders.left?.lineStyle, borders.left?.color);
          let br2 = borderToCss(rightBorders?.right?.lineStyle, rightBorders?.right?.color);

          // border-collapse:collapse means a shared edge between two cells
          // is rendered only once. Some Excel files only set the "left"
          // border of cell N+1 (or only the "top" of row N+1) and leave
          // the matching "right" / "bottom" of the previous cell empty.
          // Mirror neighbour borders so every cell edge has a defined CSS border.
          const rightCol = c + colSpan;
          const bottomRow = r + rowSpan;

          if (!br2) {
            const rightNeighbour = cellMap.get(`${colToLetter(rightCol)}${r}`);
            const rb = rightNeighbour?.style.borders?.left;
            if (rb?.lineStyle) {
              const mirrored = borderToCss(rb.lineStyle, rb.color);
              if (mirrored) br2 = mirrored;
            }
          }
          if (!bb) {
            const belowNeighbour = cellMap.get(`${colToLetter(c)}${bottomRow}`);
            const bbn = belowNeighbour?.style.borders?.top;
            if (bbn?.lineStyle) {
              const mirrored = borderToCss(bbn.lineStyle, bbn.color);
              if (mirrored) bb = mirrored;
            }
          }

          if (bt) styles.push(`border-top:${bt}`);
          if (bb) styles.push(`border-bottom:${bb}`);
          if (bl) styles.push(`border-left:${bl}`);
          if (br2) styles.push(`border-right:${br2}`);
        }

        const align = style.alignment;
        if (align) {
          if (align.horizontal === 'general' || !align.horizontal) {
            if (typeof cell.value === 'number') styles.push('text-align:right');
          } else {
            styles.push(`text-align:${align.horizontal}`);
          }
          if (align.vertical && align.vertical !== 'bottom') {
            const vMap: Record<string, string> = { top: 'top', center: 'middle', justify: 'middle', distributed: 'middle' };
            styles.push(`vertical-align:${vMap[align.vertical] || 'middle'}`);
          }
          if (align.wrapText) {
            styles.push('white-space:normal;word-wrap:break-word');
          }
        }
      }

      const styleAttr = styles.length ? ` style="${styles.join(';')}"` : '';
      const value = cell?.value != null ? escapeHtml(formatCellValue(cell.value, cell.style.numberFormat)) : '';
      tbody += `<td${spanAttr}${styleAttr}>${value}</td>`;
    }
    tbody += '</tr>';
  }

  // With border-collapse:collapse, outer borders extend ~1px outside the
  // sum of column widths. Add a small buffer so the right/bottom borders
  // are not clipped by the wrapper or by the body bounding-box screenshot.
  const BORDER_BUFFER = 2;
  const wrapperWidth = totalTableWidth + BORDER_BUFFER;
  const wrapperHeight = totalTableHeight + BORDER_BUFFER;

  const tableLayoutStyle = totalTableWidth > 0
    ? `table-layout:fixed;width:${totalTableWidth}px;`
    : '';
  const defaultFont = options.defaultFont || 'Calibri,Arial,sans-serif';
  const tableHtml = `<table style="border-collapse:collapse;${tableLayoutStyle}font-family:${defaultFont}">${colgroup}${tbody}</table>`;

  // Drawings overlay
  let overlayHtml = '';
  if (!options.skipDrawings) {
    const drawings: DrawingItem[] | undefined = (ws as any)._drawings;
    if (drawings && drawings.length) {
      const inner = renderDrawings(ws, drawings, minRow, minCol, maxRow, maxCol, defaultFont);
      if (inner) {
        overlayHtml = `<div style="position:absolute;left:0;top:0;width:${totalTableWidth}px;height:${totalTableHeight}px;pointer-events:none">${inner}</div>`;
      }
    }
  }

  // Always wrap in an inline-block container with a small padding-right /
  // padding-bottom so collapsed borders are not cut off by the page bounds.
  const wrapped = overlayHtml
    ? `<div style="position:relative;display:inline-block;width:${wrapperWidth}px;height:${wrapperHeight}px">${tableHtml}${overlayHtml}</div>`
    : `<div style="display:inline-block;padding-right:${BORDER_BUFFER}px;padding-bottom:${BORDER_BUFFER}px">${tableHtml}</div>`;

  if (options.fullPage) {
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{margin:0;padding:8px;background:#fff}td{padding:${CELL_PADDING_TB_PX}px ${CELL_PADDING_LR_PX}px;vertical-align:middle;background-color:#fff;white-space:nowrap;overflow:hidden}img{max-width:none}</style></head><body>${wrapped}</body></html>`;
  }

  return wrapped;
}
