/**
 * XLSX XML Saver – writes a Workbook to an OOXML ZIP (.xlsx) archive.
 *
 * Ported from Python `xml_saver.py`. Phase 2 scope:
 *   ✔ Content types, relationships, workbook.xml
 *   ✔ Worksheet XML (sheetData, dimension, cols, merged cells, freeze panes,
 *     sheet protection, page setup, page margins)
 *   ✔ Styles XML (fonts, fills, borders, alignments, protection, number formats, cellXfs)
 *   ✔ Shared strings XML
 *   ✔ Core & extended properties (docProps)
 *   ✗ Charts, pictures, shapes, tables, sparklines, conditional formatting,
 *     hyperlinks, data validations, comments, autofilter, calc chain (future phases)
 *
 * @module io/XmlSaver
 */

import JSZip from 'jszip';
import { Workbook } from '../core/Workbook';
import { Worksheet } from '../core/Worksheet';
import { Cell } from '../core/Cell';
import { Cells } from '../core/Cells';
import { SharedStringTable } from './SharedStrings';
import {
  CellValueHandler,
  CELL_TYPE_STRING,
} from './CellValueHandler';
import {
  Font,
  Fill,
  Borders,
  Alignment,
  Protection,
  NumberFormat,
} from '../styling/Style';

// Phase 4 feature savers
import { DataValidationXmlSaver } from './XmlDataValidationSaver';
import { ConditionalFormatXmlSaver } from './XmlConditionalFormatSaver';
import { HyperlinkXmlSaver } from './XmlHyperlinkHandler';
import { AutoFilterXmlSaver } from './XmlAutoFilterSaver';
import { CommentXmlWriter } from './CommentXml';

// ── Style data interfaces (internal, dict-like) ────────────────────────────

interface FontData {
  name: string;
  size: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
}

interface FillData {
  patternType: string;
  fgColor: string;
  bgColor: string;
}

interface BorderSideData {
  style: string;
  color: string;
}

interface BorderData {
  top: BorderSideData;
  bottom: BorderSideData;
  left: BorderSideData;
  right: BorderSideData;
}

interface AlignmentData {
  horizontal: string;
  vertical: string;
  wrapText: boolean;
  indent: number;
  textRotation: number;
  shrinkToFit: boolean;
  readingOrder: number;
  relativeIndent: number;
}

interface ProtectionData {
  locked: boolean;
  hidden: boolean;
}

/** A cell style composite key: [fontIdx, fillIdx, borderIdx, numFmtIdx, alignIdx, protIdx] */
type CellStyleKey = [number, number, number, number, number, number];

// ── XMLSaver class ─────────────────────────────────────────────────────────

export class XmlSaver {
  private _workbook: Workbook;
  private _sharedStringTable: SharedStringTable;

  // Style registries: index → data
  private _fontStyles: Map<number, FontData> = new Map();
  private _fillStyles: Map<number, FillData> = new Map();
  private _borderStyles: Map<number, BorderData> = new Map();
  private _alignmentStyles: Map<number, AlignmentData> = new Map();
  private _protectionStyles: Map<number, ProtectionData> = new Map();
  private _numFormats: Map<number, string> = new Map();

  // Composite cell style key → xf index
  private _cellStyles: Map<string, number> = new Map();

  // Phase 4: DXF entries collected from conditional formatting
  private _dxfEntries: string[] = [];

  constructor(workbook: Workbook) {
    this._workbook = workbook;
    this._sharedStringTable = new SharedStringTable();
  }

  // ── Public API ──────────────────────────────────────────────────────────

  /**
   * Saves the workbook as an XLSX buffer (in-memory ZIP).
   *
   * @returns A `Buffer` containing the XLSX data.
   */
  async saveToBuffer(): Promise<Buffer> {
    const zip = new JSZip();

    // 1. Content Types
    this._writeContentTypes(zip);

    // 2. Root relationships
    this._writeRootRelationships(zip);

    // 3. Workbook relationships
    this._writeWorkbookRelationships(zip);

    // 4. Workbook XML
    this._writeWorkbookXml(zip);

    // 5. Worksheets (this populates shared strings + style registries + dxf entries)
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      this._writeWorksheetXml(zip, this._workbook.worksheets[i], i + 1);
    }

    // 5b. Comments (after worksheets, writes xl/commentsN.xml + VML drawings)
    const commentWriter = new CommentXmlWriter();
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      const ws = this._workbook.worksheets[i];
      if (CommentXmlWriter.worksheetHasComments(ws)) {
        commentWriter.writeCommentsXml(zip, ws, i + 1);
        commentWriter.writeVmlDrawingXml(zip, ws, i + 1);
      }
    }

    // 6. Styles XML (must come after worksheets so all styles + dxf entries are registered)
    this._writeStylesXml(zip);

    // 7. Shared strings XML
    this._writeSharedStringsXml(zip);

    // 8. Theme XML (minimal – always write to satisfy Excel)
    this._writeThemeXml(zip);

    // 9. Document properties
    this._writeCorePropertiesXml(zip);
    this._writeAppPropertiesXml(zip);

    // Generate ZIP buffer
    const buf = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
    return buf;
  }

  /**
   * Saves the workbook to a file path.
   *
   * @param filePath - Destination file path.
   */
  async saveToFile(filePath: string): Promise<void> {
    const fs = await import('fs');
    const path = await import('path');

    // Ensure output directory exists
    const dir = path.dirname(filePath);
    if (dir && dir !== '.') {
      fs.mkdirSync(dir, { recursive: true });
    }

    const buffer = await this.saveToBuffer();
    fs.writeFileSync(filePath, buffer);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // [Content_Types].xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeContentTypes(zip: JSZip): void {
    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    );

    // Default extensions
    lines.push(
      '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    );
    lines.push('  <Default Extension="xml" ContentType="application/xml"/>');

    // Overrides
    lines.push(
      '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    );
    lines.push(
      '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    );
    lines.push(
      '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    );

    // Worksheet overrides
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      lines.push(
        `  <Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
      );
    }

    // Comment and VML drawing overrides
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      const ws = this._workbook.worksheets[i];
      if (CommentXmlWriter.worksheetHasComments(ws)) {
        lines.push(
          `  <Override PartName="/xl/comments${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
        );
        lines.push(
          `  <Override PartName="/xl/drawings/vmlDrawing${i + 1}.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`,
        );
      }
    }

    // Theme override
    lines.push(
      '  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    );

    // Document properties
    lines.push(
      '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    );
    lines.push(
      '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    );

    lines.push('</Types>');
    zip.file('[Content_Types].xml', lines.join('\n'));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // _rels/.rels
  // ══════════════════════════════════════════════════════════════════════════

  private _writeRootRelationships(zip: JSZip): void {
    const content = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
      '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
      '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
      '</Relationships>',
    ].join('\n');
    zip.file('_rels/.rels', content);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/_rels/workbook.xml.rels
  // ══════════════════════════════════════════════════════════════════════════

  private _writeWorkbookRelationships(zip: JSZip): void {
    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    );

    // Worksheet relationships
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      lines.push(
        `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      );
    }

    // Styles, shared strings, theme
    lines.push(
      '  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    );
    lines.push(
      '  <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    );
    lines.push(
      '  <Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    );

    lines.push('</Relationships>');
    zip.file('xl/_rels/workbook.xml.rels', lines.join('\n'));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/workbook.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeWorkbookXml(zip: JSZip): void {
    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );

    // File version
    lines.push(
      '  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>',
    );

    // Workbook properties
    lines.push('  <workbookPr defaultThemeVersion="124226"/>');

    // Book views
    const activeTab = this._workbook.properties.view.activeTab;
    lines.push('  <bookViews>');
    lines.push(
      `    <workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192" activeTab="${activeTab}"/>`,
    );
    lines.push('  </bookViews>');

    // Sheets
    lines.push('  <sheets>');
    for (let i = 0; i < this._workbook.worksheets.length; i++) {
      const ws = this._workbook.worksheets[i];
      let stateAttr = '';
      if (ws.visible === false) {
        stateAttr = ' state="hidden"';
      } else if (ws.visible === 'veryHidden') {
        stateAttr = ' state="veryHidden"';
      }
      lines.push(
        `    <sheet name="${escapeXml(ws.name)}" sheetId="${i + 1}"${stateAttr} r:id="rId${i + 1}"/>`,
      );
    }
    lines.push('  </sheets>');

    // Defined names (ECMA-376 Section 18.2.6)
    const definedNames = this._workbook.properties.definedNames;
    if (definedNames.count > 0) {
      lines.push('  <definedNames>');
      for (const dn of definedNames) {
        const dnAttrs: string[] = [`name="${escapeXml(dn.name)}"`];
        if (dn.localSheetId !== null) {
          dnAttrs.push(`localSheetId="${dn.localSheetId}"`);
        }
        if (dn.hidden) {
          dnAttrs.push('hidden="1"');
        }
        if (dn.comment) {
          dnAttrs.push(`comment="${escapeXml(dn.comment)}"`);
        }
        lines.push(`    <definedName ${dnAttrs.join(' ')}>${escapeXml(dn.refersTo)}</definedName>`);
      }
      lines.push('  </definedNames>');
    }

    // Calculation properties
    const calcMode = this._workbook.properties.calcMode || 'auto';
    lines.push(`  <calcPr calcId="145621" calcMode="${calcMode}"/>`);

    lines.push('</workbook>');
    zip.file('xl/workbook.xml', lines.join('\n'));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/worksheets/sheetN.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeWorksheetXml(
    zip: JSZip,
    worksheet: Worksheet,
    sheetNum: number,
  ): void {
    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );

    // Dimension
    const dimRef = this._computeDimensionRef(worksheet);
    if (dimRef) {
      lines.push(`  <dimension ref="${escapeXml(dimRef)}"/>`);
    }

    // Sheet views
    const activeTab = this._workbook.properties.view.activeTab;
    const isActive = sheetNum - 1 === activeTab;
    lines.push('  <sheetViews>');

    let sheetViewAttrs = `workbookViewId="0"`;
    if (isActive) sheetViewAttrs += ' tabSelected="1"';

    // Freeze pane
    const fp = worksheet.freezePane;
    if (fp) {
      lines.push(`    <sheetView ${sheetViewAttrs}>`);
      const topLeftCell = Cells.coordinateToString(
        fp.freezedRows + 1,
        fp.freezedColumns + 1,
      );
      lines.push(
        `      <pane xSplit="${fp.column}" ySplit="${fp.row}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`,
      );
      lines.push('    </sheetView>');
    } else {
      lines.push(`    <sheetView ${sheetViewAttrs}/>`);
    }
    lines.push('  </sheetViews>');

    // Sheet format properties
    lines.push('  <sheetFormatPr defaultRowHeight="15"/>');

    // Column widths / hidden columns
    const colsXml = this._formatColsXml(worksheet);
    if (colsXml) {
      lines.push(colsXml);
    }

    // Sheet data
    lines.push('  <sheetData>');

    // Collect all cells
    const cellsMap = worksheet.cells._cells;
    const sortedRefs = Array.from(cellsMap.keys()).sort((a, b) => {
      const [ra, ca] = Cells.coordinateFromString(a);
      const [rb, cb] = Cells.coordinateFromString(b);
      return ra !== rb ? ra - rb : ca - cb;
    });

    // Group by row
    const rows = new Map<number, Array<[string, Cell]>>();
    for (const ref of sortedRefs) {
      const [row] = Cells.coordinateFromString(ref);
      if (!rows.has(row)) rows.set(row, []);
      rows.get(row)!.push([ref, cellsMap.get(ref)!]);
    }

    // Include rows with custom heights or hidden even if they have no cells
    const rowHeights: Record<number, number> = (worksheet as any)._rowHeights ?? {};
    for (const rowNum of Object.keys(rowHeights).map(Number)) {
      if (!rows.has(rowNum)) rows.set(rowNum, []);
    }
    const hiddenRows: Set<number> = (worksheet as any)._hiddenRows ?? new Set();
    for (const rowNum of hiddenRows) {
      if (!rows.has(rowNum)) rows.set(rowNum, []);
    }

    // Write rows in order
    const sortedRowNums = Array.from(rows.keys()).sort((a, b) => a - b);
    for (const rowNum of sortedRowNums) {
      const rowAttrs: string[] = [`r="${rowNum}"`];

      // Custom row height
      const rh = rowHeights[rowNum];
      if (rh !== undefined) {
        rowAttrs.push(`ht="${rh}"`);
        rowAttrs.push('customHeight="1"');
      }

      // Hidden row
      if (hiddenRows.has(rowNum)) {
        rowAttrs.push('hidden="1"');
      }

      const cellEntries = rows.get(rowNum)!;
      if (cellEntries.length === 0) {
        lines.push(`    <row ${rowAttrs.join(' ')}/>`);
      } else {
        lines.push(`    <row ${rowAttrs.join(' ')}>`);
        for (const [ref, cell] of cellEntries) {
          lines.push(this._formatCellXml(ref, cell));
        }
        lines.push('    </row>');
      }
    }

    lines.push('  </sheetData>');

    // Sheet protection
    const prot = (worksheet as any)._protection;
    if (prot && prot.sheet) {
      const protAttrs: string[] = ['sheet="1"'];
      if (prot.password) protAttrs.push(`password="${escapeXml(prot.password)}"`);
      if (!prot.formatCells) protAttrs.push('formatCells="0"');
      if (!prot.formatColumns) protAttrs.push('formatColumns="0"');
      if (!prot.formatRows) protAttrs.push('formatRows="0"');
      if (!prot.insertColumns) protAttrs.push('insertColumns="0"');
      if (!prot.insertRows) protAttrs.push('insertRows="0"');
      if (!prot.deleteColumns) protAttrs.push('deleteColumns="0"');
      if (!prot.deleteRows) protAttrs.push('deleteRows="0"');
      if (!prot.sort) protAttrs.push('sort="0"');
      if (!prot.autoFilter) protAttrs.push('autoFilter="0"');
      if (!prot.insertHyperlinks) protAttrs.push('insertHyperlinks="0"');
      if (!prot.pivotTables) protAttrs.push('pivotTables="0"');
      if (prot.selectLockedCells) protAttrs.push('selectLockedCells="1"');
      if (prot.selectUnlockedCells) protAttrs.push('selectUnlockedCells="1"');
      if (prot.objects) protAttrs.push('objects="1"');
      if (prot.scenarios) protAttrs.push('scenarios="1"');
      lines.push(`  <sheetProtection ${protAttrs.join(' ')}/>`);
    }

    // AutoFilter (ECMA-376 Section 18.3.1.2 - after sheetProtection, before mergeCells)
    if (worksheet.autoFilter.range !== null) {
      const autoFilterXml = new AutoFilterXmlSaver().formatAutoFilterXml(worksheet.autoFilter);
      if (autoFilterXml) lines.push(autoFilterXml);
    }

    // Merged cells
    const mergedCells: string[] = (worksheet as any)._mergedCells ?? [];
    if (mergedCells.length > 0) {
      lines.push(`  <mergeCells count="${mergedCells.length}">`);
      for (const mc of mergedCells) {
        lines.push(`    <mergeCell ref="${escapeXml(mc)}"/>`);
      }
      lines.push('  </mergeCells>');
    }

    // Conditional formatting (ECMA-376 Section 18.3.1.18 - after mergeCells)
    if (worksheet.conditionalFormatting.count > 0) {
      const cfSaver = new ConditionalFormatXmlSaver(escapeXml);
      const startDxfId = this._dxfEntries.length;
      const cfResult = cfSaver.formatConditionalFormattingXml(
        worksheet.conditionalFormatting,
        startDxfId,
      );
      if (cfResult.xml) lines.push(cfResult.xml);
      this._dxfEntries.push(...cfResult.dxfEntries);
    }

    // Data validations (ECMA-376 Section 18.3.1.30 - after conditionalFormatting)
    if (worksheet.dataValidations.count > 0) {
      const dvXml = new DataValidationXmlSaver(escapeXml).formatDataValidationsXml(worksheet.dataValidations);
      if (dvXml) lines.push(dvXml);
    }

    // Hyperlinks (ECMA-376 Section 18.3.1.48 - after dataValidations)
    const hlSaver = new HyperlinkXmlSaver();
    const nextRelId = CommentXmlWriter.worksheetHasComments(worksheet) ? 3 : 1;
    if (worksheet.hyperlinks.count > 0) {
      hlSaver.resetRelationshipCounter(nextRelId);
      const hlXml = hlSaver.formatHyperlinksXml(worksheet.hyperlinks);
      if (hlXml) lines.push(hlXml);
    }

    // Page margins
    const pm = worksheet.pageMargins;
    lines.push(
      `  <pageMargins left="${pm.left}" right="${pm.right}" top="${pm.top}" bottom="${pm.bottom}" header="${pm.header}" footer="${pm.footer}"/>`,
    );

    // Page setup
    const ps = worksheet.pageSetup;
    const psAttrs: string[] = [];
    if (ps.orientation) psAttrs.push(`orientation="${ps.orientation}"`);
    if (ps.paperSize != null) psAttrs.push(`paperSize="${ps.paperSize}"`);
    if (ps.fitToPage) {
      psAttrs.push('fitToPage="1"');
    }
    if (ps.scale != null) psAttrs.push(`scale="${ps.scale}"`);
    if (ps.fitToWidth != null) psAttrs.push(`fitToWidth="${ps.fitToWidth}"`);
    if (ps.fitToHeight != null) psAttrs.push(`fitToHeight="${ps.fitToHeight}"`);
    if (psAttrs.length > 0) {
      lines.push(`  <pageSetup ${psAttrs.join(' ')}/>`);
    }

    // Row breaks / horizontal page breaks (ECMA-376 Section 18.3.1.73)
    const hBreaks = Array.from(worksheet.horizontalPageBreaks).sort((a, b) => a - b);
    if (hBreaks.length > 0) {
      lines.push(`  <rowBreaks count="${hBreaks.length}" manualBreakCount="${hBreaks.length}">`);
      for (const row of hBreaks) {
        lines.push(`    <brk id="${row}" max="16383" man="1"/>`);
      }
      lines.push('  </rowBreaks>');
    }

    // Column breaks / vertical page breaks (ECMA-376 Section 18.3.1.17)
    const vBreaks = Array.from(worksheet.verticalPageBreaks).sort((a, b) => a - b);
    if (vBreaks.length > 0) {
      lines.push(`  <colBreaks count="${vBreaks.length}" manualBreakCount="${vBreaks.length}">`);
      for (const col of vBreaks) {
        lines.push(`    <brk id="${col}" max="1048575" man="1"/>`);
      }
      lines.push('  </colBreaks>');
    }

    // Legacy drawing reference for comments
    if (CommentXmlWriter.worksheetHasComments(worksheet)) {
      lines.push('  <legacyDrawing r:id="rId1"/>');
    }

    lines.push('</worksheet>');
    zip.file(`xl/worksheets/sheet${sheetNum}.xml`, lines.join('\n'));

    // Write sheet-level .rels file if needed (hyperlinks + comments)
    const sheetRelsLines: string[] = [];

    // Comment relationships (rId1 = VML, rId2 = comments)
    if (CommentXmlWriter.worksheetHasComments(worksheet)) {
      sheetRelsLines.push(
        `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${sheetNum}.vml"/>`,
      );
      sheetRelsLines.push(
        `  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${sheetNum}.xml"/>`,
      );
    }

    // Hyperlink relationships
    if (worksheet.hyperlinks.count > 0) {
      const hlRels = hlSaver.getHyperlinkRelationships(worksheet.hyperlinks);
      for (const rel of hlRels) {
        sheetRelsLines.push(
          `  <Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(rel.target)}" TargetMode="External"/>`,
        );
      }
    }

    if (sheetRelsLines.length > 0) {
      const relsContent = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ...sheetRelsLines,
        '</Relationships>',
      ].join('\n');
      zip.file(`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`, relsContent);
    }
  }

  // ── Worksheet helpers ───────────────────────────────────────────────────

  /**
   * Computes the dimension reference string (e.g. "A1:C10") for a worksheet.
   */
  private _computeDimensionRef(worksheet: Worksheet): string | null {
    const cellsMap = worksheet.cells._cells;
    if (cellsMap.size === 0) return 'A1';

    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;

    for (const ref of cellsMap.keys()) {
      const [r, c] = Cells.coordinateFromString(ref);
      if (r < minRow) minRow = r;
      if (r > maxRow) maxRow = r;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }

    // Also consider merged ranges
    const merged: string[] = (worksheet as any)._mergedCells ?? [];
    for (const mc of merged) {
      if (mc.includes(':')) {
        const [startRef, endRef] = mc.split(':');
        const [sr, sc] = Cells.coordinateFromString(startRef);
        const [er, ec] = Cells.coordinateFromString(endRef);
        if (sr < minRow) minRow = sr;
        if (er > maxRow) maxRow = er;
        if (sc < minCol) minCol = sc;
        if (ec > maxCol) maxCol = ec;
      }
    }

    if (!isFinite(minRow)) return 'A1';

    const startCell = Cells.coordinateToString(minRow, minCol);
    const endCell = Cells.coordinateToString(maxRow, maxCol);
    return startCell === endCell ? startCell : `${startCell}:${endCell}`;
  }

  /**
   * Formats <cols> XML for column widths and hidden columns.
   */
  private _formatColsXml(worksheet: Worksheet): string {
    const colWidths: Record<number, number> =
      (worksheet as any)._columnWidths ?? {};
    const hiddenCols: Set<number> =
      (worksheet as any)._hiddenColumns ?? new Set();

    const allColIndices = new Set<number>();
    for (const k of Object.keys(colWidths).map(Number)) allColIndices.add(k);
    for (const k of hiddenCols) allColIndices.add(k);

    if (allColIndices.size === 0) return '';

    const sorted = Array.from(allColIndices).sort((a, b) => a - b);
    const lines: string[] = ['  <cols>'];

    for (const colIdx of sorted) {
      const attrs: string[] = [`min="${colIdx}"`, `max="${colIdx}"`];
      const w = colWidths[colIdx];
      if (w !== undefined) {
        attrs.push(`width="${w}"`);
        attrs.push('customWidth="1"');
      }
      if (hiddenCols.has(colIdx)) {
        attrs.push('hidden="1"');
      }
      lines.push(`    <col ${attrs.join(' ')}/>`);
    }

    lines.push('  </cols>');
    return lines.join('\n');
  }

  /**
   * Formats a single cell as XML.
   */
  private _formatCellXml(ref: string, cell: Cell): string {
    const styleIdx = this._getOrCreateCellStyle(cell);
    const [valueStr, cellType] = CellValueHandler.formatValueForXml(cell.value);

    // Handle shared strings: strings are stored in the SST
    let actualValue = valueStr;
    let actualType = cellType;
    if (cellType === CELL_TYPE_STRING && valueStr !== null) {
      const ssIdx = this._sharedStringTable.addString(valueStr);
      actualValue = String(ssIdx);
      actualType = CELL_TYPE_STRING;
    }

    // Build cell XML
    const attrs: string[] = [`r="${ref}"`];
    if (styleIdx > 0) attrs.push(`s="${styleIdx}"`);
    if (actualType) attrs.push(`t="${actualType}"`);

    let xml = `      <c ${attrs.join(' ')}>`;

    // Formula (must come before value per ECMA-376)
    const formula = cell.formula;
    if (formula) {
      xml += `<f>${escapeXml(formula)}</f>`;
    }

    // Value
    if (actualValue !== null) {
      xml += `<v>${escapeXml(actualValue)}</v>`;
    }

    xml += '</c>';
    return xml;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/styles.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeStylesXml(zip: JSZip): void {
    // Register default styles first
    this._registerDefaultStyles();

    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    );

    // Number formats (custom only, id >= 164)
    const customFmts = new Map<number, string>();
    for (const [id, code] of this._numFormats) {
      if (id >= 164) customFmts.set(id, code);
    }
    lines.push(`  <numFmts count="${customFmts.size}">`);
    for (const [id, code] of Array.from(customFmts.entries()).sort(
      (a, b) => a[0] - b[0],
    )) {
      lines.push(
        `    <numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`,
      );
    }
    lines.push('  </numFmts>');

    // Fonts
    const fontCount = this._fontStyles.size;
    lines.push(`  <fonts count="${fontCount}">`);
    for (const idx of sortedKeys(this._fontStyles)) {
      lines.push(this._formatFontXml(this._fontStyles.get(idx)!));
    }
    lines.push('  </fonts>');

    // Fills
    const fillCount = this._fillStyles.size;
    lines.push(`  <fills count="${fillCount}">`);
    for (const idx of sortedKeys(this._fillStyles)) {
      lines.push(this._formatFillXml(this._fillStyles.get(idx)!));
    }
    lines.push('  </fills>');

    // Borders
    const borderCount = this._borderStyles.size;
    lines.push(`  <borders count="${borderCount}">`);
    for (const idx of sortedKeys(this._borderStyles)) {
      lines.push(this._formatBorderXml(this._borderStyles.get(idx)!));
    }
    lines.push('  </borders>');

    // cellStyleXfs (base style)
    lines.push('  <cellStyleXfs count="1">');
    lines.push(
      '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
    );
    lines.push('  </cellStyleXfs>');

    // cellXfs
    const xfCount = this._cellStyles.size + 1; // +1 for default xf at index 0
    lines.push(`  <cellXfs count="${xfCount}">`);
    // Default xf at index 0
    lines.push(
      '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    );
    // Custom xfs sorted by index
    const xfEntries = Array.from(this._cellStyles.entries()).sort(
      (a, b) => a[1] - b[1],
    );
    for (const [keyStr, _xfIdx] of xfEntries) {
      const key = JSON.parse(keyStr) as CellStyleKey;
      const [fontIdx, fillIdx, borderIdx, numFmtIdx, alignIdx, protIdx] = key;
      const applyNumFmt =
        numFmtIdx !== 0 ? ' applyNumberFormat="1"' : '';
      const applyProt = protIdx !== 0 ? ' applyProtection="1"' : '';
      lines.push(
        `    <xf numFmtId="${numFmtIdx}" fontId="${fontIdx}" fillId="${fillIdx}" borderId="${borderIdx}" xfId="0"${applyNumFmt}${applyProt}>`,
      );
      if (alignIdx > 0) {
        const ad = this._alignmentStyles.get(alignIdx);
        if (ad) lines.push(this._formatAlignmentXml(ad));
      }
      if (protIdx > 0) {
        const pd = this._protectionStyles.get(protIdx);
        if (pd) lines.push(this._formatProtectionXml(pd));
      }
      lines.push('    </xf>');
    }
    lines.push('  </cellXfs>');

    // Cell styles
    lines.push('  <cellStyles count="1">');
    lines.push('    <cellStyle name="Normal" xfId="0" builtinId="0"/>');
    lines.push('  </cellStyles>');

    // Differential formatting (Phase 4: real DXF entries from conditional formatting)
    if (this._dxfEntries.length > 0) {
      lines.push(`  <dxfs count="${this._dxfEntries.length}">`);
      for (const dxfXml of this._dxfEntries) {
        lines.push(dxfXml);
      }
      lines.push('  </dxfs>');
    } else {
      lines.push('  <dxfs count="0"/>');
    }

    lines.push('</styleSheet>');
    zip.file('xl/styles.xml', lines.join('\n'));
  }

  // ── Style XML formatters ────────────────────────────────────────────────

  private _formatFontXml(fd: FontData): string {
    const parts: string[] = ['    <font>'];
    if (fd.bold) parts.push('      <b/>');
    if (fd.italic) parts.push('      <i/>');
    if (fd.underline) parts.push('      <u/>');
    if (fd.strikethrough) parts.push('      <strike/>');
    parts.push(`      <sz val="${fd.size}"/>`);
    parts.push(`      <color rgb="${fd.color}"/>`);
    parts.push(`      <name val="${escapeXml(fd.name)}"/>`);
    parts.push('    </font>');
    return parts.join('\n');
  }

  private _formatFillXml(fd: FillData): string {
    const pt = fd.patternType;
    if (pt === 'none' || pt === 'gray125') {
      return [
        '    <fill>',
        `      <patternFill patternType="${pt}"/>`,
        '    </fill>',
      ].join('\n');
    }
    return [
      '    <fill>',
      `      <patternFill patternType="${pt}">`,
      `        <fgColor rgb="${fd.fgColor}"/>`,
      `        <bgColor rgb="${fd.bgColor}"/>`,
      '      </patternFill>',
      '    </fill>',
    ].join('\n');
  }

  private _formatBorderXml(bd: BorderData): string {
    const parts: string[] = ['    <border>'];
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      const sd = bd[side];
      if (sd.style !== 'none') {
        parts.push(`      <${side} style="${sd.style}">`);
        parts.push(`        <color rgb="${sd.color}"/>`);
        parts.push(`      </${side}>`);
      } else {
        parts.push(`      <${side}/>`);
      }
    }
    parts.push('      <diagonal/>');
    parts.push('    </border>');
    return parts.join('\n');
  }

  private _formatAlignmentXml(ad: AlignmentData): string {
    const attrs: string[] = [];
    if (ad.horizontal !== 'general') attrs.push(`horizontal="${ad.horizontal}"`);
    if (ad.vertical !== 'bottom') attrs.push(`vertical="${ad.vertical}"`);
    if (ad.wrapText) attrs.push('wrapText="1"');
    if (ad.indent !== 0) attrs.push(`indent="${ad.indent}"`);
    if (ad.textRotation !== 0) attrs.push(`textRotation="${ad.textRotation}"`);
    if (ad.shrinkToFit) attrs.push('shrinkToFit="1"');
    if (ad.readingOrder !== 0) attrs.push(`readingOrder="${ad.readingOrder}"`);
    if (ad.relativeIndent !== 0)
      attrs.push(`relativeIndent="${ad.relativeIndent}"`);
    return attrs.length > 0
      ? `      <alignment ${attrs.join(' ')}/>`
      : '      <alignment/>';
  }

  private _formatProtectionXml(pd: ProtectionData): string {
    const attrs: string[] = [];
    if (!pd.locked) attrs.push('locked="0"');
    if (pd.hidden) attrs.push('hidden="1"');
    if (attrs.length === 0) return '';
    return `      <protection ${attrs.join(' ')}/>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // Style registration
  // ══════════════════════════════════════════════════════════════════════════

  private _registerDefaultStyles(): void {
    // Default font: Calibri 11pt black
    if (!this._fontStyles.has(0)) {
      this._fontStyles.set(0, {
        name: 'Calibri',
        size: 11,
        color: 'FF000000',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
      });
    }

    // Default fills: none + gray125
    if (!this._fillStyles.has(0)) {
      this._fillStyles.set(0, {
        patternType: 'none',
        fgColor: 'FFFFFFFF',
        bgColor: 'FFFFFFFF',
      });
    }
    if (!this._fillStyles.has(1)) {
      this._fillStyles.set(1, {
        patternType: 'gray125',
        fgColor: 'FFFFFFFF',
        bgColor: 'FFFFFFFF',
      });
    }

    // Default borders: none
    if (!this._borderStyles.has(0)) {
      this._borderStyles.set(0, {
        top: { style: 'none', color: 'FF000000' },
        bottom: { style: 'none', color: 'FF000000' },
        left: { style: 'none', color: 'FF000000' },
        right: { style: 'none', color: 'FF000000' },
      });
    }

    // Default alignment: general / bottom
    if (!this._alignmentStyles.has(0)) {
      this._alignmentStyles.set(0, {
        horizontal: 'general',
        vertical: 'bottom',
        wrapText: false,
        indent: 0,
        textRotation: 0,
        shrinkToFit: false,
        readingOrder: 0,
        relativeIndent: 0,
      });
    }

    // Default protection: locked=true, hidden=false
    if (!this._protectionStyles.has(0)) {
      this._protectionStyles.set(0, { locked: true, hidden: false });
    }
  }

  private _getOrCreateFontStyle(font: Font): number {
    // Check existing
    for (const [idx, fd] of this._fontStyles) {
      if (
        fd.name === font.name &&
        fd.size === font.size &&
        fd.color === font.color &&
        fd.bold === font.bold &&
        fd.italic === font.italic &&
        fd.underline === font.underline &&
        fd.strikethrough === font.strikethrough
      ) {
        return idx;
      }
    }
    const newIdx = this._fontStyles.size;
    this._fontStyles.set(newIdx, {
      name: font.name,
      size: font.size,
      color: font.color,
      bold: font.bold,
      italic: font.italic,
      underline: font.underline,
      strikethrough: font.strikethrough,
    });
    return newIdx;
  }

  private _getOrCreateFillStyle(fill: Fill): number {
    for (const [idx, fd] of this._fillStyles) {
      if (
        fd.patternType === fill.patternType &&
        fd.fgColor === fill.foregroundColor &&
        fd.bgColor === fill.backgroundColor
      ) {
        return idx;
      }
    }
    const newIdx = this._fillStyles.size;
    this._fillStyles.set(newIdx, {
      patternType: fill.patternType,
      fgColor: fill.foregroundColor,
      bgColor: fill.backgroundColor,
    });
    return newIdx;
  }

  private _getOrCreateBorderStyle(borders: Borders): number {
    for (const [idx, bd] of this._borderStyles) {
      if (
        bd.top.style === borders.top.lineStyle &&
        bd.top.color === borders.top.color &&
        bd.bottom.style === borders.bottom.lineStyle &&
        bd.bottom.color === borders.bottom.color &&
        bd.left.style === borders.left.lineStyle &&
        bd.left.color === borders.left.color &&
        bd.right.style === borders.right.lineStyle &&
        bd.right.color === borders.right.color
      ) {
        return idx;
      }
    }
    const newIdx = this._borderStyles.size;
    this._borderStyles.set(newIdx, {
      top: { style: borders.top.lineStyle, color: borders.top.color },
      bottom: { style: borders.bottom.lineStyle, color: borders.bottom.color },
      left: { style: borders.left.lineStyle, color: borders.left.color },
      right: { style: borders.right.lineStyle, color: borders.right.color },
    });
    return newIdx;
  }

  private _getOrCreateAlignmentStyle(alignment: Alignment): number {
    for (const [idx, ad] of this._alignmentStyles) {
      if (
        ad.horizontal === alignment.horizontal &&
        ad.vertical === alignment.vertical &&
        ad.wrapText === alignment.wrapText &&
        ad.indent === alignment.indent &&
        ad.textRotation === alignment.textRotation &&
        ad.shrinkToFit === alignment.shrinkToFit &&
        ad.readingOrder === alignment.readingOrder &&
        ad.relativeIndent === alignment.relativeIndent
      ) {
        return idx;
      }
    }
    const newIdx = this._alignmentStyles.size;
    this._alignmentStyles.set(newIdx, {
      horizontal: alignment.horizontal,
      vertical: alignment.vertical,
      wrapText: alignment.wrapText,
      indent: alignment.indent,
      textRotation: alignment.textRotation,
      shrinkToFit: alignment.shrinkToFit,
      readingOrder: alignment.readingOrder,
      relativeIndent: alignment.relativeIndent,
    });
    return newIdx;
  }

  private _getOrCreateProtectionStyle(protection: Protection): number {
    for (const [idx, pd] of this._protectionStyles) {
      if (
        pd.locked === protection.locked &&
        pd.hidden === protection.hidden
      ) {
        return idx;
      }
    }
    const newIdx = this._protectionStyles.size;
    this._protectionStyles.set(newIdx, {
      locked: protection.locked,
      hidden: protection.hidden,
    });
    return newIdx;
  }

  private _getOrCreateNumberFormatStyle(numberFormat: string): number {
    // Check built-in formats
    const builtinId = NumberFormat.lookupBuiltinFormat(numberFormat);
    if (builtinId !== null) return builtinId;

    // Also handle 'General' explicitly
    if (numberFormat === 'General') return 0;

    // Check existing custom formats
    for (const [id, code] of this._numFormats) {
      if (code === numberFormat) return id;
    }

    // Create new custom format (IDs start at 164)
    let newId = 164;
    for (const id of this._numFormats.keys()) {
      if (id >= newId) newId = id + 1;
    }
    this._numFormats.set(newId, numberFormat);
    return newId;
  }

  /**
   * Gets or creates a composite cellXfs style index for a cell.
   */
  private _getOrCreateCellStyle(cell: Cell): number {
    this._registerDefaultStyles();

    const style = cell.style;
    const fontIdx = this._getOrCreateFontStyle(style.font);
    const fillIdx = this._getOrCreateFillStyle(style.fill);
    const borderIdx = this._getOrCreateBorderStyle(style.borders);
    const numFmtIdx = this._getOrCreateNumberFormatStyle(
      style.numberFormat,
    );
    const alignIdx = this._getOrCreateAlignmentStyle(style.alignment);
    const protIdx = this._getOrCreateProtectionStyle(style.protection);

    // Default style → index 0
    if (
      fontIdx === 0 &&
      fillIdx === 0 &&
      borderIdx === 0 &&
      numFmtIdx === 0 &&
      alignIdx === 0 &&
      protIdx === 0
    ) {
      return 0;
    }

    const key: CellStyleKey = [
      fontIdx,
      fillIdx,
      borderIdx,
      numFmtIdx,
      alignIdx,
      protIdx,
    ];
    const keyStr = JSON.stringify(key);

    if (!this._cellStyles.has(keyStr)) {
      // Index starts at 1 (0 is the default)
      const nextIdx =
        this._cellStyles.size > 0
          ? Math.max(...Array.from(this._cellStyles.values())) + 1
          : 1;
      this._cellStyles.set(keyStr, nextIdx);
    }
    return this._cellStyles.get(keyStr)!;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/sharedStrings.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeSharedStringsXml(zip: JSZip): void {
    zip.file('xl/sharedStrings.xml', this._sharedStringTable.toXml());
  }

  // ══════════════════════════════════════════════════════════════════════════
  // xl/theme/theme1.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeThemeXml(zip: JSZip): void {
    zip.file('xl/theme/theme1.xml', DEFAULT_THEME_XML);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // docProps/core.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeCorePropertiesXml(zip: JSZip): void {
    const dp = this._workbook.documentProperties;
    const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');

    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
        'xmlns:dcterms="http://purl.org/dc/terms/" ' +
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' +
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    );

    if (dp.title) lines.push(`  <dc:title>${escapeXml(dp.title)}</dc:title>`);
    if (dp.subject)
      lines.push(`  <dc:subject>${escapeXml(dp.subject)}</dc:subject>`);
    if (dp.creator)
      lines.push(`  <dc:creator>${escapeXml(dp.creator)}</dc:creator>`);
    if (dp.keywords)
      lines.push(`  <cp:keywords>${escapeXml(dp.keywords)}</cp:keywords>`);
    if (dp.description)
      lines.push(
        `  <dc:description>${escapeXml(dp.description)}</dc:description>`,
      );
    if (dp.lastModifiedBy)
      lines.push(
        `  <cp:lastModifiedBy>${escapeXml(dp.lastModifiedBy)}</cp:lastModifiedBy>`,
      );
    if (dp.category)
      lines.push(`  <cp:category>${escapeXml(dp.category)}</cp:category>`);

    lines.push(
      `  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`,
    );
    lines.push(
      `  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`,
    );

    lines.push('</cp:coreProperties>');
    zip.file('docProps/core.xml', lines.join('\n'));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // docProps/app.xml
  // ══════════════════════════════════════════════════════════════════════════

  private _writeAppPropertiesXml(zip: JSZip): void {
    const wsCount = this._workbook.worksheets.length;

    const lines: string[] = [];
    lines.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    lines.push(
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    );
    lines.push('  <Application>Microsoft Excel</Application>');
    lines.push('  <DocSecurity>0</DocSecurity>');
    lines.push('  <ScaleCrop>false</ScaleCrop>');
    lines.push('  <LinksUpToDate>false</LinksUpToDate>');
    lines.push('  <SharedDoc>false</SharedDoc>');

    // Heading pairs & titles
    lines.push('  <HeadingPairs>');
    lines.push('    <vt:vector size="2" baseType="variant">');
    lines.push('      <vt:variant>');
    lines.push('        <vt:lpstr>Worksheets</vt:lpstr>');
    lines.push('      </vt:variant>');
    lines.push('      <vt:variant>');
    lines.push(`        <vt:i4>${wsCount}</vt:i4>`);
    lines.push('      </vt:variant>');
    lines.push('    </vt:vector>');
    lines.push('  </HeadingPairs>');

    lines.push('  <TitlesOfParts>');
    lines.push(`    <vt:vector size="${wsCount}" baseType="lpstr">`);
    for (const ws of this._workbook.worksheets) {
      lines.push(`      <vt:lpstr>${escapeXml(ws.name)}</vt:lpstr>`);
    }
    lines.push('    </vt:vector>');
    lines.push('  </TitlesOfParts>');

    lines.push('</Properties>');
    zip.file('docProps/app.xml', lines.join('\n'));
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// Utilities
// ══════════════════════════════════════════════════════════════════════════════

/** Escape XML special characters. */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Sorted numeric keys of a Map<number, T>. */
function sortedKeys<T>(map: Map<number, T>): number[] {
  return Array.from(map.keys()).sort((a, b) => a - b);
}

// ══════════════════════════════════════════════════════════════════════════════
// Default theme XML (minimal Office theme for XLSX validity)
// ══════════════════════════════════════════════════════════════════════════════

const DEFAULT_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
