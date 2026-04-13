/**
 * Comment XML Handler
 *
 * Reads and writes Excel comments (annotations) in XLSX format.
 * Comments are stored in xl/comments{n}.xml with VML drawings for positioning.
 */

import type JSZip from 'jszip';
import type { Worksheet } from '../core/Worksheet';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Sort key for cell references: row-major order */
function cellReferenceSortKey(ref: string): [number, number] {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return [0, 0];
  const col = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  let colNum = 0;
  for (const ch of col) {
    colNum = colNum * 26 + (ch.charCodeAt(0) - 64);
  }
  return [row, colNum];
}

function colFromRef(ref: string): number {
  const match = ref.match(/^([A-Z]+)/i);
  if (!match) return 0;
  let col = 0;
  for (const ch of match[1].toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return col - 1; // 0-based
}

function rowFromRef(ref: string): number {
  const match = ref.match(/(\d+)$/);
  return match ? parseInt(match[1], 10) - 1 : 0; // 0-based
}

// ==================== Writer ====================

export class CommentXmlWriter {
  /**
   * Check if worksheet has any comments
   */
  static worksheetHasComments(worksheet: Worksheet): boolean {
    const cellsObj = (worksheet as any)._cells;
    const cells: Map<string, any> = cellsObj._cells ?? cellsObj;
    for (const cell of cells.values()) {
      if (cell.comment) return true;
    }
    return false;
  }

  /**
   * Write comments XML to the ZIP archive
   */
  writeCommentsXml(zip: JSZip, worksheet: Worksheet, sheetNum: number): void {
    const cellsObj = (worksheet as any)._cells;
    const cells: Map<string, any> = cellsObj._cells ?? cellsObj;
    const commentCells: Array<{ ref: string; text: string; author: string }> = [];

    for (const [ref, cell] of cells.entries()) {
      if (cell.comment) {
        commentCells.push({
          ref,
          text: cell.comment.text ?? '',
          author: cell.comment.author ?? 'None',
        });
      }
    }

    if (commentCells.length === 0) return;

    // Sort by cell reference
    commentCells.sort((a, b) => {
      const ka = cellReferenceSortKey(a.ref);
      const kb = cellReferenceSortKey(b.ref);
      return ka[0] - kb[0] || ka[1] - kb[1];
    });

    // Collect unique authors
    const authors = [...new Set(commentCells.map((c) => c.author))];

    const lines: string[] = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<authors>',
    ];

    for (const author of authors) {
      lines.push(`<author>${escapeXml(author)}</author>`);
    }

    lines.push('</authors>');
    lines.push('<commentList>');

    for (const cc of commentCells) {
      const authorIdx = authors.indexOf(cc.author);
      lines.push(`<comment ref="${cc.ref}" authorId="${authorIdx}">`);
      lines.push('<text>');
      lines.push(`<r><t>${escapeXml(cc.text)}</t></r>`);
      lines.push('</text>');
      lines.push('</comment>');
    }

    lines.push('</commentList>');
    lines.push('</comments>');

    zip.file(`xl/comments${sheetNum}.xml`, lines.join(''));
  }

  /**
   * Write VML drawing XML for comment shapes
   */
  writeVmlDrawingXml(zip: JSZip, worksheet: Worksheet, sheetNum: number): void {
    const cellsObj = (worksheet as any)._cells;
    const cells: Map<string, any> = cellsObj._cells ?? cellsObj;
    const commentCells: Array<{ ref: string; width: number; height: number }> = [];

    for (const [ref, cell] of cells.entries()) {
      if (cell.comment) {
        const width = cell.comment.width ?? 108;
        const height = cell.comment.height ?? 59;
        commentCells.push({ ref, width, height });
      }
    }

    if (commentCells.length === 0) return;

    commentCells.sort((a, b) => {
      const ka = cellReferenceSortKey(a.ref);
      const kb = cellReferenceSortKey(b.ref);
      return ka[0] - kb[0] || ka[1] - kb[1];
    });

    const lines: string[] = [
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
      '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
      '</v:shapetype>',
    ];

    for (let i = 0; i < commentCells.length; i++) {
      const cc = commentCells[i];
      const row = rowFromRef(cc.ref);
      const col = colFromRef(cc.ref);
      const anchor = this._calculateAnchor(row, col, cc.width, cc.height);

      lines.push(`<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" `);
      lines.push(`style="position:absolute;margin-left:0;margin-top:0;width:${cc.width}pt;height:${cc.height}pt;z-index:${i + 1};visibility:hidden" `);
      lines.push('fillcolor="#ffffe1" o:insetmode="auto">');
      lines.push('<v:fill color="#ffffe1"/>');
      lines.push('<v:shadow on="t" color="black" obscured="t"/>');
      lines.push('<v:path o:connecttype="none"/>');
      lines.push('<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>');
      lines.push('<x:ClientData ObjectType="Note">');
      lines.push('<x:MoveWithCells/>');
      lines.push('<x:SizeWithCells/>');
      lines.push(`<x:Anchor>${anchor}</x:Anchor>`);
      lines.push('<x:AutoFill>False</x:AutoFill>');
      lines.push(`<x:Row>${row}</x:Row>`);
      lines.push(`<x:Column>${col}</x:Column>`);
      lines.push('</x:ClientData>');
      lines.push('</v:shape>');
    }

    lines.push('</xml>');
    zip.file(`xl/drawings/vmlDrawing${sheetNum}.vml`, lines.join(''));
  }

  private _calculateAnchor(row: number, col: number, width: number, height: number): string {
    // Calculate anchor coordinates for VML
    // Format: leftCol, leftOffset, topRow, topOffset, rightCol, rightOffset, bottomRow, bottomOffset
    const colWidth = 64; // default column width in pixels approx
    const rowHeight = 20; // default row height in pixels approx

    const leftCol = col + 1;
    const leftOffset = 15;
    const topRow = row;
    const topOffset = 10;

    const widthPx = Math.round(width * 1.33);
    const heightPx = Math.round(height * 1.33);

    const rightCol = leftCol + Math.max(1, Math.floor(widthPx / colWidth));
    const rightOffset = widthPx % colWidth;
    const bottomRow = topRow + Math.max(1, Math.floor(heightPx / rowHeight));
    const bottomOffset = heightPx % rowHeight;

    return `${leftCol}, ${leftOffset}, ${topRow}, ${topOffset}, ${rightCol}, ${rightOffset}, ${bottomRow}, ${bottomOffset}`;
  }
}

// ==================== Reader ====================

export class CommentXmlReader {
  loadComments(zip: JSZip, worksheet: Worksheet, sheetNum: number): void {
    const commentsFile = zip.file(`xl/comments${sheetNum}.xml`);
    if (!commentsFile) return;

    // We'll load this asynchronously in the main loader
    // This is a sync placeholder - actual loading happens through the XmlLoader integration
  }

  /**
   * Parse comments XML and apply to worksheet
   */
  parseAndApplyComments(commentsRoot: any, worksheet: Worksheet): void {
    if (!commentsRoot?.comments) return;

    const comments = commentsRoot.comments;

    // Load authors
    const authors: string[] = [];
    const authorElems = this._ensureArray(comments?.authors?.author);
    for (const a of authorElems) {
      authors.push(typeof a === 'string' ? a : String(a));
    }

    // Load comment list
    const commentElems = this._ensureArray(comments?.commentList?.comment);
    for (const ce of commentElems) {
      const ref = ce?.['@_ref'];
      if (!ref) continue;

      const authorId = parseInt(ce?.['@_authorId'] ?? '0', 10);
      const author = authors[authorId] ?? 'None';

      // Extract text
      let text = '';
      const textElem = ce?.text;
      if (textElem) {
        const runs = this._ensureArray(textElem.r);
        if (runs.length > 0) {
          for (const r of runs) {
            const tVal = r?.t;
            if (tVal != null) {
              // t may be an array (from isArray parser config) or a single value
              const tArr = Array.isArray(tVal) ? tVal : [tVal];
              for (const t of tArr) {
                if (t != null) {
                  text += typeof t === 'object' ? (t['#text'] ?? '') : String(t);
                }
              }
            }
          }
        } else if (textElem.t != null) {
          // t may be an array here too
          const tArr = Array.isArray(textElem.t) ? textElem.t : [textElem.t];
          for (const t of tArr) {
            if (t != null) {
              text += typeof t === 'object' ? (t['#text'] ?? '') : String(t);
            }
          }
        }
      }

      if (text) {
        // Access the cells map to set comment
        const cellsObj2 = (worksheet as any)._cells;
        const cells2: Map<string, any> = cellsObj2._cells ?? cellsObj2;
        const cell = cells2.get(ref);
        if (cell) {
          cell.setComment(text, author);
        } else {
          // Create cell with comment
          const { Cell } = require('../core/Cell');
          const newCell = new Cell();
          newCell.setComment(text, author);
          cells2.set(ref, newCell);
        }
      }
    }
  }

  /**
   * Parse VML drawing to get comment dimensions
   */
  parseAndApplyVmlDrawing(vmlContent: string, worksheet: Worksheet): void {
    // Simple regex-based parsing of VML to extract comment sizes
    const shapeRegex = /<v:shape[^>]*>[\s\S]*?<\/v:shape>/gi;
    let match;

    while ((match = shapeRegex.exec(vmlContent)) !== null) {
      const shape = match[0];

      // Extract row and column
      const rowMatch = shape.match(/<x:Row>(\d+)<\/x:Row>/i);
      const colMatch = shape.match(/<x:Column>(\d+)<\/x:Column>/i);
      if (!rowMatch || !colMatch) continue;

      const row = parseInt(rowMatch[1], 10);
      const col = parseInt(colMatch[1], 10);

      // Extract width/height from style
      const widthMatch = shape.match(/width:\s*([\d.]+)pt/i);
      const heightMatch = shape.match(/height:\s*([\d.]+)pt/i);

      if (widthMatch && heightMatch) {
        const width = parseFloat(widthMatch[1]);
        const height = parseFloat(heightMatch[1]);

        // Convert col/row to cell reference
        const colLetter = this._colToLetter(col);
        const ref = `${colLetter}${row + 1}`;

        const cellsObj3 = (worksheet as any)._cells;
        const cells3: Map<string, any> = cellsObj3._cells ?? cellsObj3;
        const cell = cells3.get(ref);
        if (cell?.comment) {
          cell.comment.width = width;
          cell.comment.height = height;
        }
      }
    }
  }

  private _ensureArray<T>(val: T | T[] | undefined | null): T[] {
    if (val == null) return [];
    return Array.isArray(val) ? val : [val];
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
