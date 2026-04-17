/**
 * tinyweb-office-cells – XML Loader Module (Phase 3)
 *
 * Reads OOXML (.xlsx) ZIP archives and populates Workbook objects.
 * Port of the Python `xml_loader.py`.
 *
 * ECMA-376 Compliant cell value import.
 */

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { Workbook } from '../core/Workbook';
import { Worksheet } from '../core/Worksheet';
import { Cell } from '../core/Cell';
import { Cells } from '../core/Cells';
import { CellValueHandler } from './CellValueHandler';
import { AutoFilterXmlLoader } from './XmlAutoFilterLoader';
import { ConditionalFormatXmlLoader } from './XmlConditionalFormatLoader';
import { DataValidationXmlLoader } from './XmlDataValidationLoader';
import { HyperlinkXmlLoader } from './XmlHyperlinkHandler';
import { CommentXmlReader } from './CommentXml';
import { DefinedName } from '../features/DefinedName';
import type { CellValue } from '../core/Cell';

// ── Style data interfaces (loader-specific, richer than saver's) ───────────

export interface LoaderFontData {
  name: string;
  size: number;
  color: string;
  color_type: string | null;
  color_value: string | null;
  color_tint: string | null;
  family: string | null;
  charset: string | null;
  scheme: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
}

export interface LoaderFillData {
  pattern_type: string;
  fg_color: string;
  bg_color: string;
  fg_color_type: string | null;
  fg_color_value: string | null;
  fg_color_tint: number | null;
  bg_color_type: string | null;
  bg_color_value: string | null;
  bg_color_tint: number | null;
}

export interface LoaderBorderSideData {
  style: string;
  color: string;
}

export interface LoaderBorderData {
  top: LoaderBorderSideData;
  bottom: LoaderBorderSideData;
  left: LoaderBorderSideData;
  right: LoaderBorderSideData;
}

export interface LoaderAlignmentData {
  horizontal: string;
  vertical: string;
  wrap_text: boolean;
  indent: number;
  text_rotation: number;
  shrink_to_fit: boolean;
  reading_order: number;
  relative_indent: number;
}

export interface LoaderProtectionData {
  locked: boolean;
  hidden: boolean;
}

/** Cell style composite key: [fontIdx, fillIdx, borderIdx, numFmtIdx, alignIdx, protIdx] */
type CellStyleKey = [number, number, number, number, number, number];

// ── Augment Workbook with loader-specific internal state ───────────────────

/* eslint-disable @typescript-eslint/no-explicit-any */
interface WorkbookInternal {
  _worksheets: Worksheet[];
  _sharedStrings: string[];
  _styles: any[];
  _properties: any;
  _documentProperties: any;
  _sourceXml: Record<string, string>;

  // Loader style registries (numeric index → data)
  _loaderFontStyles: Map<number, LoaderFontData>;
  _loaderFillStyles: Map<number, LoaderFillData>;
  _loaderBorderStyles: Map<number, LoaderBorderData>;
  _loaderAlignmentStyles: Map<number, LoaderAlignmentData>;
  _loaderProtectionStyles: Map<number, LoaderProtectionData>;
  _loaderNumFormats: Map<number, string>;
  _loaderCellStyles: Map<string, number>;
  _loaderCellXfByIndex: Map<number, CellStyleKey>;

  // Source XML preservation for round-trip
  _sourceStylesXml: Buffer | null;
  _sourceCellXfsCount: number;
  _themeXml: Buffer | null;
  _contentTypeDefaults: Record<string, string>;
  _contentTypeOverrides: Record<string, string>;
  _sourceExtraWorkbookRels: any[];
  _sourceCalcChainBytes: Buffer | null;
  _sourceCalcChainRel: Record<string, any> | null;
  _dxfStyles: any[];
}
/* eslint-enable @typescript-eslint/no-explicit-any */

// ── Built-in number formats (ECMA-376 §18.8.30) ───────────────────────────

const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

// ── XML Parser configuration ───────────────────────────────────────────────

function createParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    isArray: (name: string) => {
      // Force these elements to always be arrays even when there's only one
      const arrayElements = new Set([
        'sheet', 'row', 'c', 'si', 't', 'r', 'font', 'fill', 'border',
        'xf', 'numFmt', 'mergeCell', 'col', 'Relationship',
        'Override', 'Default', 'dxf', 'definedName',
      ]);
      return arrayElements.has(name);
    },
    removeNSPrefix: true,
    parseTagValue: false,
    trimValues: false,
    processEntities: false,
  });
}

// ── Helper: extract text from attribute or return default ──────────────────

function attr(elem: any, name: string, defaultVal = ''): string {
  if (!elem) return defaultVal;
  const key = `@_${name}`;
  return elem[key] !== undefined ? String(elem[key]) : defaultVal;
}

function attrOpt(elem: any, name: string): string | null {
  if (!elem) return null;
  const key = `@_${name}`;
  return elem[key] !== undefined ? String(elem[key]) : null;
}

// ── Helper: ensure value is always an array ────────────────────────────────

function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val === undefined || val === null) return [];
  return Array.isArray(val) ? val : [val];
}

// ── Helper: regex-based raw XML extraction (for round-trip) ────────────────

function extractRootExtraAttrs(xmlText: string, tagName: string): string {
  const regex = new RegExp(`<${tagName}\\b([^>]*)>`, 's');
  const match = regex.exec(xmlText);
  if (!match) return '';
  let attrs = match[1];
  attrs = attrs.replace(/\s+xmlns="[^"]*"/g, '');
  attrs = attrs.replace(/\s+xmlns:r="[^"]*"/g, '');
  return attrs.trim();
}

function extractRawElement(xmlText: string, tagName: string): string | null {
  // Escaped tag name for regex
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const patterns = [
    new RegExp(`(<${escaped}\\b[^>]*/>)`, 's'),
    new RegExp(`(<${escaped}\\b[^>]*>.*?</${escaped}>)`, 's'),
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(xmlText);
    if (match) return match[1];
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════
// XmlLoader class
// ═══════════════════════════════════════════════════════════════════════════

export class XmlLoader {
  private _workbook: Workbook;
  private _wb: WorkbookInternal;
  private _parser: XMLParser;

  // Content-type tracking
  private _contentTypeOverrides: Record<string, string> = {};
  private _contentTypeDefaults: Record<string, string> = {};

  constructor(workbook: Workbook) {
    this._workbook = workbook;
    this._wb = workbook as unknown as WorkbookInternal;
    this._parser = createParser();

    // Initialize loader-specific style registries on workbook
    this._wb._loaderFontStyles = new Map();
    this._wb._loaderFillStyles = new Map();
    this._wb._loaderBorderStyles = new Map();
    this._wb._loaderAlignmentStyles = new Map();
    this._wb._loaderProtectionStyles = new Map();
    this._wb._loaderNumFormats = new Map();
    this._wb._loaderCellStyles = new Map();
    this._wb._loaderCellXfByIndex = new Map();
    this._wb._sourceStylesXml = null;
    this._wb._sourceCellXfsCount = 0;
    this._wb._themeXml = null;
    this._wb._contentTypeDefaults = {};
    this._wb._contentTypeOverrides = {};
    this._wb._sourceExtraWorkbookRels = [];
    this._wb._sourceCalcChainBytes = null;
    this._wb._sourceCalcChainRel = null;
    this._wb._dxfStyles = [];
  }

  // ── Public API ──────────────────────────────────────────────────────────

  /**
   * Loads workbook data from a JSZip instance.
   */
  async loadWorkbook(zip: JSZip): Promise<void> {
    // Load workbook.xml
    const wbXmlBuf = await zip.file('xl/workbook.xml')?.async('nodebuffer');
    if (!wbXmlBuf) throw new Error('xl/workbook.xml not found in XLSX archive');
    const wbXmlText = wbXmlBuf.toString('utf-8');
    const wbRoot = this._parser.parse(wbXmlText);
    const workbookNode = wbRoot.workbook || wbRoot;

    // Preserve raw XML for round-trip
    this._wb._sourceXml['workbook_root_extra_attrs'] =
      extractRootExtraAttrs(wbXmlText, 'workbook');
    const altContent = extractRawElement(wbXmlText, 'mc:AlternateContent');
    if (altContent) this._wb._sourceXml['workbook_alt_content'] = altContent;
    const revPtr = extractRawElement(wbXmlText, 'xr:revisionPtr');
    if (revPtr) this._wb._sourceXml['workbook_revision_ptr'] = revPtr;
    const extLst = extractRawElement(wbXmlText, 'extLst');
    if (extLst) this._wb._sourceXml['workbook_extlst'] = extLst;

    // Load workbook properties
    this._loadWorkbookProperties(workbookNode);

    // Load document properties
    await this._loadDocumentProperties(zip);

    // Load content type overrides
    await this._loadContentTypeOverrides(zip);

    // Load theme
    await this._loadTheme(zip);

    // Load worksheet info (creates worksheet objects)
    this._loadWorksheetInfo(workbookNode);

    // Apply print areas from defined names
    this._applyPrintAreasFromDefinedNames(workbookNode);

    // Load extra workbook rels
    await this._loadExtraWorkbookRels(zip);

    // Load shared strings
    await this._loadSharedStrings(zip);

    // Load styles
    await this._loadStyles(zip);

    // Load worksheet data
    await this._loadWorksheetsData(zip);
  }

  /**
   * Loads a workbook from a Buffer.
   */
  static async loadFromBuffer(buffer: Buffer): Promise<Workbook> {
    const zip = await JSZip.loadAsync(buffer);
    const workbook = Object.create(Workbook.prototype) as Workbook;

    // Manually initialize internal state (bypass constructor's default sheet creation)
    const wb = workbook as unknown as WorkbookInternal;
    wb._worksheets = [];
    wb._styles = [];
    wb._sharedStrings = [];
    (workbook as any)._filePath = null;
    (workbook as any)._properties = new (
      await import('../core/Workbook')
    ).WorkbookProperties();
    (workbook as any)._documentProperties = null;
    wb._sourceXml = {};

    // Initialize saver-compatible style maps (Map<string, number>)
    (workbook as any)._fontStyles = new Map();
    (workbook as any)._fillStyles = new Map();
    (workbook as any)._borderStyles = new Map();
    (workbook as any)._alignmentStyles = new Map();
    (workbook as any)._protectionStyles = new Map();
    (workbook as any)._cellStyles = new Map();
    (workbook as any)._numFormats = new Map();

    // Push default style
    const { Style } = await import('../styling/Style');
    wb._styles.push(new Style());

    const loader = new XmlLoader(workbook);
    await loader.loadWorkbook(zip);
    return workbook;
  }

  /**
   * Loads a workbook from a file path.
   */
  static async loadFromFile(filePath: string): Promise<Workbook> {
    const fs = await import('fs');
    const buffer = fs.readFileSync(filePath);
    const workbook = await XmlLoader.loadFromBuffer(buffer);
    (workbook as any)._filePath = filePath;
    return workbook;
  }

  // ── Content Types ───────────────────────────────────────────────────────

  private async _loadContentTypeOverrides(zip: JSZip): Promise<void> {
    this._contentTypeDefaults = {};
    this._contentTypeOverrides = {};
    this._wb._contentTypeDefaults = {};
    this._wb._contentTypeOverrides = {};

    const file = zip.file('[Content_Types].xml');
    if (!file) return;

    try {
      const buf = await file.async('nodebuffer');
      const root = this._parser.parse(buf.toString('utf-8'));
      const types = root.Types || root;

      for (const def of ensureArray(types.Default)) {
        const ext = attrOpt(def, 'Extension');
        const ct = attrOpt(def, 'ContentType');
        if (ext && ct) {
          this._contentTypeDefaults[ext.toLowerCase()] = ct;
          this._wb._contentTypeDefaults[ext.toLowerCase()] = ct;
        }
      }

      for (const ov of ensureArray(types.Override)) {
        const partName = attrOpt(ov, 'PartName');
        const ct = attrOpt(ov, 'ContentType');
        if (partName && ct) {
          this._contentTypeOverrides[partName] = ct;
          this._wb._contentTypeOverrides[partName] = ct;
        }
      }
    } catch {
      // Ignore parse errors
    }
  }

  // ── Theme ───────────────────────────────────────────────────────────────

  private async _loadTheme(zip: JSZip): Promise<void> {
    const file = zip.file('xl/theme/theme1.xml');
    if (file) {
      this._wb._themeXml = await file.async('nodebuffer');
    } else {
      this._wb._themeXml = null;
    }
  }

  // ── Workbook Properties ─────────────────────────────────────────────────

  private _loadWorkbookProperties(workbookNode: any): void {
    const props = this._wb._properties;

    // File version
    const fv = workbookNode.fileVersion;
    if (fv && props.file_version) {
      props.file_version.appName = attr(fv, 'appName', 'xl');
      props.file_version.lastEdited = attr(fv, 'lastEdited', '5');
      props.file_version.lowestEdited = attr(fv, 'lowestEdited', '5');
      props.file_version.rupBuild = attr(fv, 'rupBuild', '9302');
    }

    // Workbook Pr
    const wbPr = workbookNode.workbookPr;
    if (wbPr && props.workbook_pr) {
      const dtvStr = attrOpt(wbPr, 'defaultThemeVersion');
      if (dtvStr) props.workbook_pr.defaultThemeVersion = dtvStr;
    }

    // Book views
    const bookViews = workbookNode.bookViews;
    if (bookViews) {
      const wbView = bookViews.workbookView;
      const view = Array.isArray(wbView) ? wbView[0] : wbView;
      if (view) {
        const tab = attrOpt(view, 'activeTab');
        if (tab !== null) props.view.activeTab = parseInt(tab, 10) || 0;
      }
    }

    // Calc properties
    const calcPr = workbookNode.calcPr;
    if (calcPr && props.calculation !== undefined) {
      const calcMode = attrOpt(calcPr, 'calcMode');
      if (calcMode) props.calcMode = calcMode;
    }

    // Load defined names (ECMA-376 Section 18.2.6)
    this._loadDefinedNames(workbookNode);
  }

  private _loadDefinedNames(workbookNode: any): void {
    const dnContainer = workbookNode.definedNames;
    if (!dnContainer) return;

    const dnElems = ensureArray(dnContainer.definedName);
    const collection = this._wb._properties.definedNames;

    for (const dnElem of dnElems) {
      // definedName element: name in attribute, value in text
      const name = dnElem?.['@_name'] ?? '';
      if (!name) continue;

      const refersTo = typeof dnElem === 'string'
        ? dnElem
        : (dnElem?.['#text'] ?? '');

      let localSheetId: number | null = null;
      const lsId = dnElem?.['@_localSheetId'];
      if (lsId != null) {
        localSheetId = parseInt(String(lsId), 10);
      }

      const dn = new DefinedName(name, String(refersTo), localSheetId);
      dn.hidden = dnElem?.['@_hidden'] === '1';
      const comment = dnElem?.['@_comment'];
      if (comment) dn.comment = String(comment);

      (collection as any)._names.push(dn);
    }
  }

  // ── Worksheet Info ──────────────────────────────────────────────────────

  private _loadWorksheetInfo(workbookNode: any): void {
    const sheetsNode = workbookNode.sheets;
    if (!sheetsNode) return;

    const sheets = ensureArray(sheetsNode.sheet);
    for (const sheet of sheets) {
      const sheetName = attr(sheet, 'name', 'Sheet');
      const ws = new Worksheet(sheetName);
      (ws as any)._workbook = this._workbook;

      // Load visibility state
      const state = attrOpt(sheet, 'state');
      if (state === 'hidden') {
        (ws as any)._visible = false;
      } else if (state === 'veryHidden') {
        (ws as any)._visible = 'veryHidden';
      }

      this._wb._worksheets.push(ws);
    }
  }

  // ── Print Areas from Defined Names ──────────────────────────────────────

  private _applyPrintAreasFromDefinedNames(workbookNode: any): void {
    const dnNode = workbookNode.definedNames;
    if (!dnNode) return;

    const names = ensureArray(dnNode.definedName);
    for (const dn of names) {
      const name = attr(dn, 'name');
      if (name !== '_xlnm.Print_Area') continue;

      const sheetIdStr = attrOpt(dn, 'localSheetId');
      if (sheetIdStr === null) continue;
      const sheetId = parseInt(sheetIdStr, 10);
      if (isNaN(sheetId) || sheetId < 0 || sheetId >= this._wb._worksheets.length) continue;

      const refersTo = typeof dn === 'object' && dn['#text']
        ? String(dn['#text'])
        : '';
      if (!refersTo) continue;

      const ws = this._wb._worksheets[sheetId];
      const printArea = this._extractPrintArea(refersTo);
      if (printArea) {
        (ws as any)._printArea = printArea;
      }
    }
  }

  private _extractPrintArea(refersTo: string): string | null {
    if (!refersTo) return null;
    const parts: string[] = [];
    for (let token of refersTo.split(',')) {
      token = token.trim();
      if (!token) continue;
      let addr = token;
      if (token.includes('!')) {
        addr = token.split('!')[1];
      }
      addr = addr.replace(/\$/g, '').trim().toUpperCase();
      parts.push(addr);
    }
    return parts.length > 0 ? parts.join(',') : null;
  }

  // ── Shared Strings ──────────────────────────────────────────────────────

  private async _loadSharedStrings(zip: JSZip): Promise<void> {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) {
      this._wb._sharedStrings = [];
      return;
    }

    try {
      const buf = await file.async('nodebuffer');
      const root = this._parser.parse(buf.toString('utf-8'));
      const sst = root.sst || root;
      this._wb._sharedStrings = [];

      const siList = ensureArray(sst.si);
      for (const si of siList) {
        // Simple text: <si><t>text</t></si>
        if (si.t !== undefined) {
          const tArr = ensureArray(si.t);
          const texts: string[] = [];
          for (const t of tArr) {
            if (typeof t === 'string') {
              texts.push(t);
            } else if (t && t['#text'] !== undefined) {
              texts.push(String(t['#text']));
            } else if (typeof t === 'number') {
              texts.push(String(t));
            } else {
              texts.push('');
            }
          }
          this._wb._sharedStrings.push(texts.join(''));
          continue;
        }

        // Rich text: <si><r><t>text</t></r>...</si>
        if (si.r !== undefined) {
          const runs = ensureArray(si.r);
          const texts: string[] = [];
          for (const run of runs) {
            if (run.t !== undefined) {
              const tArr = ensureArray(run.t);
              for (const t of tArr) {
                if (typeof t === 'string') {
                  texts.push(t);
                } else if (t && t['#text'] !== undefined) {
                  texts.push(String(t['#text']));
                } else if (typeof t === 'number') {
                  texts.push(String(t));
                } else {
                  texts.push('');
                }
              }
            }
          }
          this._wb._sharedStrings.push(texts.join(''));
          continue;
        }

        // Fallback: empty string
        this._wb._sharedStrings.push('');
      }
    } catch {
      this._wb._sharedStrings = [];
    }
  }

  // ── Styles ──────────────────────────────────────────────────────────────

  private async _loadStyles(zip: JSZip): Promise<void> {
    const file = zip.file('xl/styles.xml');
    if (!file) return;

    try {
      const buf = await file.async('nodebuffer');
      this._wb._sourceStylesXml = buf;
      const root = this._parser.parse(buf.toString('utf-8'));
      const stylesRoot = root.styleSheet || root;

      // Count cellXfs
      const cellXfsElem = stylesRoot.cellXfs;
      if (cellXfsElem) {
        this._wb._sourceCellXfsCount = parseInt(attr(cellXfsElem, 'count', '0'), 10);
      }

      this._loadStylesXml(stylesRoot);
      this._loadDxfStyles(stylesRoot);
    } catch {
      // Use defaults on parse error
    }
  }

  private _loadStylesXml(stylesRoot: any): void {
    // Register built-in number formats
    for (const [id, code] of Object.entries(BUILTIN_FORMATS)) {
      this._wb._loaderNumFormats.set(Number(id), code);
    }

    // Load custom number formats
    const numFmtsNode = stylesRoot.numFmts;
    if (numFmtsNode) {
      const fmts = ensureArray(numFmtsNode.numFmt);
      for (const fmt of fmts) {
        const id = parseInt(attr(fmt, 'numFmtId', '0'), 10);
        const code = attr(fmt, 'formatCode', 'General');
        this._wb._loaderNumFormats.set(id, code);
      }
    }

    // Load fonts
    this._loadFonts(stylesRoot);

    // Load fills
    this._loadFills(stylesRoot);

    // Load borders
    this._loadBorders(stylesRoot);

    // Load cellXfs
    this._loadCellXfs(stylesRoot);
  }

  // ── Fonts ───────────────────────────────────────────────────────────────

  private _loadFonts(stylesRoot: any): void {
    const fontsNode = stylesRoot.fonts;
    if (!fontsNode) return;

    const fonts = ensureArray(fontsNode.font);
    for (let i = 0; i < fonts.length; i++) {
      const fontElem = fonts[i];

      const szElem = fontElem.sz;
      const colorElem = fontElem.color;
      const nameElem = fontElem.name;
      const familyElem = fontElem.family;
      const charsetElem = fontElem.charset;
      const schemeElem = fontElem.scheme;
      const bElem = fontElem.b;
      const iElem = fontElem.i;
      const uElem = fontElem.u;
      const strikeElem = fontElem.strike;

      let colorType: string | null = null;
      let colorValue: string | null = null;
      let colorTint: string | null = null;

      if (colorElem) {
        if (attrOpt(colorElem, 'rgb') !== null) {
          colorType = 'rgb';
          colorValue = attr(colorElem, 'rgb');
        } else if (attrOpt(colorElem, 'theme') !== null) {
          colorType = 'theme';
          colorValue = attr(colorElem, 'theme');
        } else if (attrOpt(colorElem, 'indexed') !== null) {
          colorType = 'indexed';
          colorValue = attr(colorElem, 'indexed');
        } else if (attrOpt(colorElem, 'auto') !== null) {
          colorType = 'auto';
          colorValue = attr(colorElem, 'auto');
        }
        if (attrOpt(colorElem, 'tint') !== null) {
          colorTint = attr(colorElem, 'tint');
        }
      }

      const fontData: LoaderFontData = {
        name: nameElem ? attr(nameElem, 'val', 'Calibri') : 'Calibri',
        size: szElem ? parseFloat(attr(szElem, 'val', '11')) : 11,
        color: colorType === 'rgb' && colorValue ? colorValue : 'FF000000',
        color_type: colorType,
        color_value: colorValue,
        color_tint: colorTint,
        family: familyElem ? attr(familyElem, 'val') : null,
        charset: charsetElem ? attr(charsetElem, 'val') : null,
        scheme: schemeElem ? attr(schemeElem, 'val') : null,
        bold: bElem !== undefined,
        italic: iElem !== undefined,
        underline: uElem !== undefined,
        strikethrough: strikeElem !== undefined,
      };

      this._wb._loaderFontStyles.set(i, fontData);

      // Update default style font (fontId=0)
      if (i === 0 && this._wb._styles.length > 0) {
        const defaultFont = this._wb._styles[0].font;
        defaultFont.name = fontData.name;
        defaultFont.size = fontData.size;
        defaultFont.bold = fontData.bold;
        defaultFont.italic = fontData.italic;
        defaultFont.underline = fontData.underline;
        defaultFont.strikethrough = fontData.strikethrough;
        if (fontData.color_type === 'rgb' && fontData.color_value) {
          defaultFont.color = fontData.color_value;
        }
      }
    }
  }

  // ── Fills ───────────────────────────────────────────────────────────────

  private _loadFills(stylesRoot: any): void {
    const fillsNode = stylesRoot.fills;
    if (!fillsNode) return;

    const fills = ensureArray(fillsNode.fill);
    for (let i = 0; i < fills.length; i++) {
      const fillElem = fills[i];
      const patternElem = fillElem.patternFill;
      const fgColorElem = patternElem ? patternElem.fgColor : null;
      const bgColorElem = patternElem ? patternElem.bgColor : null;

      const [fgType, fgValue, fgTint] = this._extractColorAttrs(fgColorElem);
      const [bgType, bgValue, bgTint] = this._extractColorAttrs(bgColorElem);

      const fgRgb = fgType === 'rgb' && fgValue ? fgValue : 'FFFFFFFF';
      const bgRgb = bgType === 'rgb' && bgValue ? bgValue : 'FFFFFFFF';

      const fillData: LoaderFillData = {
        pattern_type: patternElem ? attr(patternElem, 'patternType', 'none') : 'none',
        fg_color: fgRgb,
        bg_color: bgRgb,
        fg_color_type: fgType,
        fg_color_value: fgValue,
        fg_color_tint: fgTint,
        bg_color_type: bgType,
        bg_color_value: bgValue,
        bg_color_tint: bgTint,
      };

      this._wb._loaderFillStyles.set(i, fillData);
    }
  }

  private _extractColorAttrs(
    colorElem: any,
  ): [string, string | null, number | null] {
    if (!colorElem) return ['rgb', 'FFFFFFFF', null];

    const tintStr = attrOpt(colorElem, 'tint');
    const tint = tintStr !== null ? parseFloat(tintStr) : null;

    if (attrOpt(colorElem, 'rgb') !== null) {
      return ['rgb', attr(colorElem, 'rgb'), tint];
    }
    if (attrOpt(colorElem, 'theme') !== null) {
      return ['theme', attr(colorElem, 'theme'), tint];
    }
    if (attrOpt(colorElem, 'indexed') !== null) {
      return ['indexed', attr(colorElem, 'indexed'), tint];
    }
    if (attrOpt(colorElem, 'auto') !== null) {
      return ['auto', attr(colorElem, 'auto'), tint];
    }
    return ['rgb', 'FFFFFFFF', tint];
  }

  // ── Borders ─────────────────────────────────────────────────────────────

  private _loadBorders(stylesRoot: any): void {
    const bordersNode = stylesRoot.borders;
    if (!bordersNode) return;

    const borders = ensureArray(bordersNode.border);
    for (let i = 0; i < borders.length; i++) {
      if (i === 0) continue; // Skip default border

      const borderElem = borders[i];
      const loadSide = (side: string): LoaderBorderSideData => {
        const sideElem = borderElem[side];
        if (!sideElem) return { style: 'none', color: 'FF000000' };
        const style = attr(sideElem, 'style', 'none');
        const colorElem = sideElem.color;
        const color = colorElem ? attr(colorElem, 'rgb', 'FF000000') : 'FF000000';
        return { style, color };
      };

      const borderData: LoaderBorderData = {
        top: loadSide('top'),
        bottom: loadSide('bottom'),
        left: loadSide('left'),
        right: loadSide('right'),
      };

      this._wb._loaderBorderStyles.set(i, borderData);
    }
  }

  // ── Cell XFs ────────────────────────────────────────────────────────────

  private _loadCellXfs(stylesRoot: any): void {
    // Ensure default alignment/protection exist at index 0
    if (!this._wb._loaderAlignmentStyles.has(0)) {
      this._wb._loaderAlignmentStyles.set(0, {
        horizontal: 'general',
        vertical: 'bottom',
        wrap_text: false,
        indent: 0,
        text_rotation: 0,
        shrink_to_fit: false,
        reading_order: 0,
        relative_indent: 0,
      });
    }
    if (!this._wb._loaderProtectionStyles.has(0)) {
      this._wb._loaderProtectionStyles.set(0, {
        locked: true,
        hidden: false,
      });
    }

    const cellXfsNode = stylesRoot.cellXfs;
    if (!cellXfsNode) return;

    const xfs = ensureArray(cellXfsNode.xf);
    for (let i = 0; i < xfs.length; i++) {
      const xfElem = xfs[i];

      const fontIdx = parseInt(attr(xfElem, 'fontId', '0'), 10);
      const fillIdx = parseInt(attr(xfElem, 'fillId', '0'), 10);
      const borderIdx = parseInt(attr(xfElem, 'borderId', '0'), 10);
      const numFmtIdx = parseInt(attr(xfElem, 'numFmtId', '0'), 10);

      // Load alignment if present
      let alignmentIdx = 0;
      const alignElem = xfElem.alignment;
      if (alignElem) {
        const horizontal = attr(alignElem, 'horizontal', 'general');
        const vertical = attr(alignElem, 'vertical', 'bottom');
        const textRotation = parseInt(attr(alignElem, 'textRotation', '0'), 10);
        const wrapText = attr(alignElem, 'wrapText') === '1';
        const shrinkToFit = attr(alignElem, 'shrinkToFit') === '1';
        const indent = parseInt(attr(alignElem, 'indent', '0'), 10);
        const readingOrder = parseInt(attr(alignElem, 'readingOrder', '0'), 10);
        const relativeIndent = parseInt(attr(alignElem, 'relativeIndent', '0'), 10);

        // Check if this alignment already exists
        let found = false;
        for (const [idx, ad] of this._wb._loaderAlignmentStyles) {
          if (
            ad.horizontal === horizontal &&
            ad.vertical === vertical &&
            ad.wrap_text === wrapText &&
            ad.indent === indent &&
            ad.text_rotation === textRotation &&
            ad.shrink_to_fit === shrinkToFit &&
            ad.reading_order === readingOrder &&
            ad.relative_indent === relativeIndent
          ) {
            alignmentIdx = idx;
            found = true;
            break;
          }
        }

        if (!found) {
          alignmentIdx = this._wb._loaderAlignmentStyles.size;
          this._wb._loaderAlignmentStyles.set(alignmentIdx, {
            horizontal,
            vertical,
            wrap_text: wrapText,
            indent,
            text_rotation: textRotation,
            shrink_to_fit: shrinkToFit,
            reading_order: readingOrder,
            relative_indent: relativeIndent,
          });
        }
      }

      // Load protection if present
      let protectionIdx = 0;
      const protElem = xfElem.protection;
      if (protElem) {
        const locked = attr(protElem, 'locked', '1') === '1';
        const hidden = attr(protElem, 'hidden', '0') === '1';

        // Check if this protection already exists
        let found = false;
        for (const [idx, pd] of this._wb._loaderProtectionStyles) {
          if (pd.locked === locked && pd.hidden === hidden) {
            protectionIdx = idx;
            found = true;
            break;
          }
        }

        if (!found && !(locked && !hidden)) {
          protectionIdx = this._wb._loaderProtectionStyles.size;
          this._wb._loaderProtectionStyles.set(protectionIdx, {
            locked,
            hidden,
          });
        }
      }

      const cellStyleKey: CellStyleKey = [
        fontIdx, fillIdx, borderIdx, numFmtIdx, alignmentIdx, protectionIdx,
      ];
      const keyStr = JSON.stringify(cellStyleKey);
      this._wb._loaderCellStyles.set(keyStr, i);
      this._wb._loaderCellXfByIndex.set(i, cellStyleKey);
    }
  }

  // ── DXF Styles (conditional formatting) ─────────────────────────────────

  private _loadDxfStyles(stylesRoot: any): void {
    this._wb._dxfStyles = [];

    const dxfsNode = stylesRoot.dxfs;
    if (!dxfsNode) return;

    const dxfs = ensureArray(dxfsNode.dxf);
    for (const dxfElem of dxfs) {
      const dxfData: Record<string, any> = {};

      // Font
      const fontElem = dxfElem.font;
      if (fontElem) {
        const fontData: Record<string, any> = {};
        if (fontElem.b !== undefined) {
          fontData.bold = attr(fontElem.b, 'val', '1') !== '0';
        }
        if (fontElem.i !== undefined) {
          fontData.italic = attr(fontElem.i, 'val', '1') !== '0';
        }
        if (fontElem.u !== undefined) {
          fontData.underline = true;
        }
        if (fontElem.strike !== undefined) {
          fontData.strikethrough = true;
        }
        if (fontElem.color) {
          fontData.color = attr(fontElem.color, 'rgb', 'FF000000');
        }
        if (Object.keys(fontData).length > 0) {
          dxfData.font = fontData;
        }
      }

      // Fill
      const fillElem = dxfElem.fill;
      if (fillElem) {
        const patternElem = fillElem.patternFill;
        if (patternElem) {
          const fillData: Record<string, string> = {
            pattern_type: attr(patternElem, 'patternType', 'solid'),
          };
          if (patternElem.fgColor) {
            fillData.fg_color = attr(patternElem.fgColor, 'rgb', 'FFFFFFFF');
          }
          if (patternElem.bgColor) {
            fillData.bg_color = attr(patternElem.bgColor, 'rgb', 'FFFFFFFF');
          }
          dxfData.fill = fillData;
        }
      }

      // Border
      const borderElem = dxfElem.border;
      if (borderElem) {
        for (const side of ['left', 'right', 'top', 'bottom']) {
          const sideElem = borderElem[side];
          if (sideElem) {
            const style = attr(sideElem, 'style', 'thin');
            let color = 'FF000000';
            if (sideElem.color) {
              color = attr(sideElem.color, 'rgb', 'FF000000');
            }
            dxfData.border = { style, color };
            break;
          }
        }
      }

      this._wb._dxfStyles.push(dxfData);
    }
  }

  // ── Extra Workbook Rels ─────────────────────────────────────────────────

  private async _loadExtraWorkbookRels(zip: JSZip): Promise<void> {
    const file = zip.file('xl/_rels/workbook.xml.rels');
    if (!file) return;

    const buf = await file.async('nodebuffer');
    const root = this._parser.parse(buf.toString('utf-8'));
    const relsNode = root.Relationships || root;

    const NATIVE_TYPES = new Set([
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    ]);

    const extraRels: any[] = [];
    const rels = ensureArray(relsNode.Relationship);

    for (const rel of rels) {
      const relType = attr(rel, 'Type');
      const target = attr(rel, 'Target');
      const relId = attr(rel, 'Id');
      const targetMode = attr(rel, 'TargetMode');

      // Handle calcChain
      if (relType === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain') {
        const partPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
        try {
          const partFile = zip.file(partPath);
          if (partFile) {
            this._wb._sourceCalcChainBytes = await partFile.async('nodebuffer');
            this._wb._sourceCalcChainRel = {
              rel_id: relId,
              target,
              part_path: partPath,
              content_type:
                this._contentTypeOverrides[`/${partPath}`] ||
                'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
            };
          }
        } catch {
          // ignore
        }
        continue;
      }

      if (NATIVE_TYPES.has(relType)) continue;
      if (targetMode === 'External') continue;

      const partPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`;

      let partBytes: Buffer | null = null;
      try {
        const partFile = zip.file(partPath);
        if (partFile) {
          partBytes = await partFile.async('nodebuffer');
        }
      } catch {
        // ignore
      }

      // Read part's own rels file
      const lastSlash = partPath.lastIndexOf('/');
      const dirPart = lastSlash >= 0 ? partPath.substring(0, lastSlash) : '';
      const filePart = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
      const partRelsPath = `${dirPart}/_rels/${filePart}.rels`;

      let partRelsBytes: Buffer | null = null;
      try {
        const rlsFile = zip.file(partRelsPath);
        if (rlsFile) {
          partRelsBytes = await rlsFile.async('nodebuffer');
        }
      } catch {
        // ignore
      }

      const contentType = this._contentTypeOverrides[`/${partPath}`] || null;

      extraRels.push({
        rel_id: relId,
        rel_type: relType,
        target,
        part_path: partPath,
        part_bytes: partBytes,
        part_rels_path: partRelsPath,
        part_rels_bytes: partRelsBytes,
        content_type: contentType,
      });
    }

    if (extraRels.length > 0) {
      this._wb._sourceExtraWorkbookRels = extraRels;
    }
  }

  // ── Worksheets Data ─────────────────────────────────────────────────────

  private async _loadWorksheetsData(zip: JSZip): Promise<void> {
    interface WsAndRoot {
      idx: number;
      ws: Worksheet;
      root: any | null;
    }

    const worksheetsAndRoots: WsAndRoot[] = [];

    for (let i = 0; i < this._wb._worksheets.length; i++) {
      const ws = this._wb._worksheets[i];
      const file = zip.file(`xl/worksheets/sheet${i + 1}.xml`);
      if (!file) {
        worksheetsAndRoots.push({ idx: i, ws, root: null });
        continue;
      }

      try {
        const buf = await file.async('nodebuffer');
        const xmlText = buf.toString('utf-8');
        const parsed = this._parser.parse(xmlText);
        const wsRoot = parsed.worksheet || parsed;

        // Preserve raw XML for round-trip
        (ws as any)._sourceXml = xmlText;

        const extraAttrs = extractRootExtraAttrs(xmlText, 'worksheet');
        if (extraAttrs) {
          (ws as any)._source_root_extra_attrs = extraAttrs;
        }

        const sheetPr = extractRawElement(xmlText, 'sheetPr');
        if (sheetPr) (ws as any)._source_sheet_pr_xml = sheetPr;

        const phoneticPr = extractRawElement(xmlText, 'phoneticPr');
        if (phoneticPr) (ws as any)._source_phonetic_pr_xml = phoneticPr;

        const dimensionXml = extractRawElement(xmlText, 'dimension');
        if (dimensionXml) {
          const m = /ref="([^"]+)"/.exec(dimensionXml);
          if (m) (ws as any)._source_dimension_ref = m[1];
        }

        // Load extra sheet rels
        await this._loadExtraSheetRels(zip, ws, i + 1);

        worksheetsAndRoots.push({ idx: i, ws, root: wsRoot });
      } catch {
        worksheetsAndRoots.push({ idx: i, ws, root: null });
      }
    }

    // Build map: comments part_path -> the worksheet that owns it via its rels.
    const claimedCommentsPaths = new Map<string, Worksheet>();
    for (const { ws } of worksheetsAndRoots) {
      for (const r of ((ws as any)._source_extra_sheet_rels ?? [])) {
        if ((r.rel_type ?? '').toLowerCase().includes('comments')) {
          claimedCommentsPaths.set(r.part_path, ws);
        }
      }
    }

    // Pre-load sheet rels for hyperlinks
    const sheetRelsMap = new Map<number, any>();
    for (const { idx } of worksheetsAndRoots) {
      const relsPath = `xl/worksheets/_rels/sheet${idx + 1}.xml.rels`;
      const relsFile = zip.file(relsPath);
      if (relsFile) {
        try {
          const relsBuf = await relsFile.async('nodebuffer');
          const relsRoot = this._parser.parse(relsBuf.toString('utf-8'));
          sheetRelsMap.set(idx, relsRoot);
        } catch {
          // ignore
        }
      }
    }

    // Main pass: load cell data, comments, hyperlinks
    const commentReader = new CommentXmlReader();
    const hlLoader = new HyperlinkXmlLoader();

    for (const { idx, ws, root } of worksheetsAndRoots) {
      if (root === null) continue;
      try {
        this._loadWorksheetData(ws, root);

        // Load comments for this worksheet
        const defaultCommentsPath = `xl/comments${idx + 1}.xml`;
        const commentsOwner = claimedCommentsPaths.get(defaultCommentsPath);
        if (commentsOwner == null || commentsOwner === ws) {
          await this._loadComments(zip, ws, idx + 1, commentReader);
        }

        // Load hyperlinks for this worksheet
        const relsRoot = sheetRelsMap.get(idx);
        const relationships = relsRoot
          ? hlLoader.loadRelationships(relsRoot)
          : new Map<string, string>();
        hlLoader.loadHyperlinks(root, ws.hyperlinks, relationships);
      } catch {
        // Skip worksheet on error
      }
    }
  }

  private async _loadComments(
    zip: JSZip,
    worksheet: Worksheet,
    sheetNum: number,
    reader: CommentXmlReader,
  ): Promise<void> {
    const commentsFile = zip.file(`xl/comments${sheetNum}.xml`);
    if (!commentsFile) return;

    try {
      const buf = await commentsFile.async('nodebuffer');
      const xmlText = buf.toString('utf-8');
      const parsed = this._parser.parse(xmlText);
      reader.parseAndApplyComments(parsed, worksheet);
    } catch {
      // Skip comments on error
    }

    // Load VML drawing for comment sizes
    const vmlFile = zip.file(`xl/drawings/vmlDrawing${sheetNum}.vml`);
    if (vmlFile) {
      try {
        const vmlBuf = await vmlFile.async('nodebuffer');
        const vmlContent = vmlBuf.toString('utf-8');
        reader.parseAndApplyVmlDrawing(vmlContent, worksheet);
      } catch {
        // Skip VML on error
      }
    }
  }

  // ── Extra Sheet Rels ────────────────────────────────────────────────────

  private async _loadExtraSheetRels(
    zip: JSZip,
    worksheet: Worksheet,
    sheetNum: number,
  ): Promise<void> {
    const relsPath = `xl/worksheets/_rels/sheet${sheetNum}.xml.rels`;
    const file = zip.file(relsPath);
    if (!file) return;

    const buf = await file.async('nodebuffer');
    const root = this._parser.parse(buf.toString('utf-8'));
    const relsNode = root.Relationships || root;

    const PRESERVED_TYPES = new Set([
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    ]);

    const extraRels: any[] = [];
    const rels = ensureArray(relsNode.Relationship);

    for (const rel of rels) {
      const relType = attr(rel, 'Type');
      const target = attr(rel, 'Target');
      const targetMode = attr(rel, 'TargetMode');
      const relId = attr(rel, 'Id');

      if (!PRESERVED_TYPES.has(relType)) continue;
      if (targetMode === 'External') continue;

      let partPath: string;
      if (target.startsWith('../')) {
        partPath = 'xl/' + target.slice(3);
      } else if (target.startsWith('/')) {
        partPath = target.slice(1);
      } else {
        partPath = 'xl/worksheets/' + target;
      }

      let partBytes: Buffer | null = null;
      try {
        const partFile = zip.file(partPath);
        if (partFile) {
          partBytes = await partFile.async('nodebuffer');
        }
      } catch {
        // ignore
      }

      extraRels.push({
        rel_id: relId,
        rel_type: relType,
        target,
        target_mode: targetMode,
        part_path: partPath,
        part_bytes: partBytes,
      });
    }

    if (extraRels.length > 0) {
      (worksheet as any)._source_extra_sheet_rels = extraRels;
    }
  }

  // ── Worksheet Data (cells, merges, columns, rows) ───────────────────────

  private _loadWorksheetData(worksheet: Worksheet, wsRoot: any): void {
    // Load column dimensions
    this._loadColumnDimensions(worksheet, wsRoot);

    // Load row heights
    this._loadRowHeights(worksheet, wsRoot);

    // Load merged cells
    this._loadMergedCells(worksheet, wsRoot);

    // Load freeze pane from sheetViews
    this._loadFreezePane(worksheet, wsRoot);

    // Load auto filter settings (ECMA-376 Section 18.3.1.2)
    new AutoFilterXmlLoader().loadAutoFilter(wsRoot, worksheet.autoFilter);

    // Load conditional formatting (ECMA-376 Section 18.3.1.18)
    new ConditionalFormatXmlLoader(this._wb._dxfStyles).loadConditionalFormatting(
      wsRoot, worksheet.conditionalFormatting,
    );

    // Load data validations (ECMA-376 Section 18.3.1.30, 18.3.1.31)
    new DataValidationXmlLoader().loadDataValidations(wsRoot, worksheet.dataValidations);

    // Load page breaks (ECMA-376 Section 18.3.1.14, 18.3.1.73)
    this._loadPageBreaks(worksheet, wsRoot);

    // Load cell data
    const sharedStrings = this._wb._sharedStrings;
    const sheetData = wsRoot.sheetData;
    if (!sheetData) return;

    const rows = ensureArray(sheetData.row);
    for (const rowElem of rows) {
      const cells = ensureArray(rowElem.c);
      for (const cellElem of cells) {
        const cellRef = attr(cellElem, 'r');
        if (!cellRef) continue;

        const cellType = attr(cellElem, 't', 'n'); // Default to numeric per ECMA-376

        // Check for formula
        let formula: string | null = null;
        if (cellElem.f !== undefined) {
          const fValue = cellElem.f;
          if (typeof fValue === 'string') {
            formula = fValue;
          } else if (fValue && fValue['#text'] !== undefined) {
            formula = String(fValue['#text']);
          } else if (typeof fValue === 'object') {
            // Self-closing <f/> with no text
            formula = null;
          }
          if (formula !== null && !formula.startsWith('=')) {
            formula = '=' + formula;
          }
        }

        // Get style index
        const sAttr = attrOpt(cellElem, 's');
        const styleIdx = sAttr !== null ? parseInt(sAttr, 10) : 0;

        // Get cell value
        let value: CellValue = null;
        if (cellElem.v !== undefined) {
          let valueStr: string | null = null;
          if (typeof cellElem.v === 'string') {
            valueStr = cellElem.v;
          } else if (typeof cellElem.v === 'number') {
            valueStr = String(cellElem.v);
          } else if (cellElem.v && cellElem.v['#text'] !== undefined) {
            valueStr = String(cellElem.v['#text']);
          }

          if (valueStr !== null) {
            value = CellValueHandler.parseValueFromXml(
              valueStr,
              cellType,
              sharedStrings,
            );
          }
        }

        // Create cell
        const cell = new Cell(value, formula);

        // Apply style
        this._applyCellStyle(cell, styleIdx);

        // Set cell on worksheet
        worksheet.cells.set(cellRef, cell);
      }
    }
  }

  // ── Merged Cells ────────────────────────────────────────────────────────

  private _loadMergedCells(worksheet: Worksheet, wsRoot: any): void {
    const mergeCellsNode = wsRoot.mergeCells;
    if (!mergeCellsNode) {
      (worksheet as any)._mergedCells = [];
      return;
    }

    const merged: string[] = [];
    const cells = ensureArray(mergeCellsNode.mergeCell);
    for (const mc of cells) {
      const ref = attrOpt(mc, 'ref');
      if (ref) {
        merged.push(ref.toUpperCase());
      }
    }
    (worksheet as any)._mergedCells = merged;
  }

  // ── Column Dimensions ───────────────────────────────────────────────────

  private _loadColumnDimensions(worksheet: Worksheet, wsRoot: any): void {
    const colsNode = wsRoot.cols;
    if (!colsNode) return;

    const wsAny = worksheet as any;
    if (!wsAny._columnWidths) wsAny._columnWidths = {};
    if (!wsAny._hiddenColumns) wsAny._hiddenColumns = new Set();

    const cols = ensureArray(colsNode.col);
    for (const colElem of cols) {
      const minVal = attrOpt(colElem, 'min');
      const maxVal = attrOpt(colElem, 'max');
      const widthVal = attrOpt(colElem, 'width');
      const hiddenVal = attrOpt(colElem, 'hidden');

      if (minVal === null || maxVal === null) continue;

      const minCol = Math.floor(parseFloat(minVal));
      const maxCol = Math.floor(parseFloat(maxVal));
      if (minCol < 1 || maxCol < minCol) continue;

      let width: number | null = null;
      if (widthVal !== null) {
        width = parseFloat(widthVal);
        if (isNaN(width) || width <= 0) width = null;
      }

      for (let colIdx = minCol; colIdx <= maxCol; colIdx++) {
        if (width !== null) {
          wsAny._columnWidths[colIdx] = width;
        }
        if (hiddenVal === '1' || hiddenVal === 'true' || hiddenVal === 'True') {
          wsAny._hiddenColumns.add(colIdx);
        }
      }
    }
  }

  // ── Row Heights ─────────────────────────────────────────────────────────

  private _loadRowHeights(worksheet: Worksheet, wsRoot: any): void {
    const wsAny = worksheet as any;
    if (!wsAny._rowHeights) wsAny._rowHeights = {};
    if (!wsAny._hiddenRows) wsAny._hiddenRows = new Set();

    const sheetData = wsRoot.sheetData;
    if (!sheetData) return;

    const rows = ensureArray(sheetData.row);
    for (const rowElem of rows) {
      const ht = attrOpt(rowElem, 'ht');
      const hiddenVal = attrOpt(rowElem, 'hidden');
      const rowNum = attrOpt(rowElem, 'r');

      if (ht === null && !(hiddenVal === '1' || hiddenVal === 'true' || hiddenVal === 'True')) {
        continue;
      }
      if (rowNum === null) continue;

      const rowIdx = parseInt(rowNum, 10);
      if (isNaN(rowIdx) || rowIdx < 1) continue;

      if (ht !== null) {
        const height = parseFloat(ht);
        if (!isNaN(height) && height > 0) {
          wsAny._rowHeights[rowIdx] = height;
        }
      }
      if (hiddenVal === '1' || hiddenVal === 'true' || hiddenVal === 'True') {
        wsAny._hiddenRows.add(rowIdx);
      }
    }
  }

  // ── Freeze Pane ─────────────────────────────────────────────────────────

  private _loadFreezePane(worksheet: Worksheet, wsRoot: any): void {
    const sheetViews = wsRoot.sheetViews;
    if (!sheetViews) return;

    const views = ensureArray(sheetViews.sheetView || sheetViews);
    if (views.length === 0) return;

    // Use the first sheetView (or sheetViews itself if it's the view)
    let view = views[0];
    if (view.sheetView) {
      view = Array.isArray(view.sheetView) ? view.sheetView[0] : view.sheetView;
    }

    const pane = view.pane;
    if (!pane) return;

    const state = attrOpt(pane, 'state');
    if (state !== 'frozen') return;

    const xSplit = parseInt(attr(pane, 'xSplit', '0'), 10);
    const ySplit = parseInt(attr(pane, 'ySplit', '0'), 10);

    if (xSplit > 0 || ySplit > 0) {
      worksheet.setFreezePane(ySplit, xSplit, ySplit, xSplit);
    }
  }

  /**
   * Loads manual page breaks from rowBreaks / colBreaks elements.
   * ECMA-376 Section 18.3.1.14 (colBreaks), 18.3.1.73 (rowBreaks).
   */
  private _loadPageBreaks(worksheet: Worksheet, wsRoot: any): void {
    // Horizontal page breaks (rowBreaks)
    const rowBreaks = wsRoot.rowBreaks;
    if (rowBreaks) {
      const brks = ensureArray(rowBreaks.brk);
      for (const brk of brks) {
        const id = brk?.['@_id'];
        if (id != null) {
          worksheet.horizontalPageBreaks.add(parseInt(String(id), 10));
        }
      }
    }

    // Vertical page breaks (colBreaks)
    const colBreaks = wsRoot.colBreaks;
    if (colBreaks) {
      const brks = ensureArray(colBreaks.brk);
      for (const brk of brks) {
        const id = brk?.['@_id'];
        if (id != null) {
          worksheet.verticalPageBreaks.add(parseInt(String(id), 10));
        }
      }
    }
  }

  // ── Apply Cell Style ────────────────────────────────────────────────────

  private _applyCellStyle(cell: Cell, styleIdx: number): void {
    // Look up the cellXf key by index
    let cellStyleKey: CellStyleKey | undefined =
      this._wb._loaderCellXfByIndex.get(styleIdx);

    if (!cellStyleKey) {
      // Fallback: search by value
      for (const [keyStr, idx] of this._wb._loaderCellStyles) {
        if (idx === styleIdx) {
          cellStyleKey = JSON.parse(keyStr) as CellStyleKey;
          break;
        }
      }
    }

    if (!cellStyleKey) return;

    // Preserve original style index
    (cell as any)._sourceStyleIdx = styleIdx;

    const [fontKey, fillKey, borderKey, numFmtKey, alignmentKey, protectionKey] =
      cellStyleKey;

    // Apply font
    const fontData = this._wb._loaderFontStyles.get(fontKey);
    if (fontData) {
      cell.style.font.name = fontData.name;
      cell.style.font.size = fontData.size;
      cell.style.font.color = fontData.color;
      cell.style.font.bold = fontData.bold;
      cell.style.font.italic = fontData.italic;
      cell.style.font.underline = fontData.underline;
      cell.style.font.strikethrough = fontData.strikethrough;
    }

    // Apply fill
    const fillData = this._wb._loaderFillStyles.get(fillKey);
    if (fillData) {
      cell.style.fill.patternType = fillData.pattern_type;
      cell.style.fill.foregroundColor = fillData.fg_color;
      cell.style.fill.backgroundColor = fillData.bg_color;
    }

    // Apply border
    const borderData = this._wb._loaderBorderStyles.get(borderKey);
    if (borderData) {
      cell.style.borders.top.lineStyle = borderData.top.style;
      cell.style.borders.top.color = borderData.top.color;
      cell.style.borders.bottom.lineStyle = borderData.bottom.style;
      cell.style.borders.bottom.color = borderData.bottom.color;
      cell.style.borders.left.lineStyle = borderData.left.style;
      cell.style.borders.left.color = borderData.left.color;
      cell.style.borders.right.lineStyle = borderData.right.style;
      cell.style.borders.right.color = borderData.right.color;
    }

    // Apply number format
    const numFmtCode = this._wb._loaderNumFormats.get(numFmtKey);
    if (numFmtCode !== undefined) {
      cell.style.numberFormat = numFmtCode;
    }

    // Apply alignment
    const alignData = this._wb._loaderAlignmentStyles.get(alignmentKey);
    if (alignData) {
      cell.style.alignment.horizontal = alignData.horizontal;
      cell.style.alignment.vertical = alignData.vertical;
      cell.style.alignment.wrapText = alignData.wrap_text;
      cell.style.alignment.indent = alignData.indent;
      cell.style.alignment.textRotation = alignData.text_rotation;
      cell.style.alignment.shrinkToFit = alignData.shrink_to_fit;
      cell.style.alignment.readingOrder = alignData.reading_order;
      cell.style.alignment.relativeIndent = alignData.relative_indent;
    }

    // Apply protection
    const protData = this._wb._loaderProtectionStyles.get(protectionKey);
    if (protData) {
      cell.style.protection.locked = protData.locked;
      cell.style.protection.hidden = protData.hidden;
    }
  }

  // ── Document Properties ─────────────────────────────────────────────────

  private async _loadDocumentProperties(zip: JSZip): Promise<void> {
    await this._loadCoreProperties(zip);
    await this._loadAppProperties(zip);
  }

  private async _loadCoreProperties(zip: JSZip): Promise<void> {
    const file = zip.file('docProps/core.xml');
    if (!file) return;

    try {
      const buf = await file.async('nodebuffer');
      // Use a separate parser for Dublin Core namespaces
      const coreParser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        textNodeName: '#text',
        removeNSPrefix: true,
        parseTagValue: false,
        trimValues: false,
    processEntities: false,
      });
      const root = coreParser.parse(buf.toString('utf-8'));
      const coreProps = root.coreProperties || root['cp:coreProperties'] || root;

      const docProps = this._workbook.documentProperties;

      // Title, Subject, Creator, Description (Dublin Core)
      if (coreProps.title) {
        docProps.title = typeof coreProps.title === 'string'
          ? coreProps.title
          : coreProps.title['#text'] || '';
      }
      if (coreProps.subject) {
        docProps.subject = typeof coreProps.subject === 'string'
          ? coreProps.subject
          : coreProps.subject['#text'] || '';
      }
      if (coreProps.creator) {
        docProps.creator = typeof coreProps.creator === 'string'
          ? coreProps.creator
          : coreProps.creator['#text'] || '';
      }
      if (coreProps.description) {
        docProps.description = typeof coreProps.description === 'string'
          ? coreProps.description
          : coreProps.description['#text'] || '';
      }

      // OPC Core Properties
      if (coreProps.keywords) {
        docProps.keywords = typeof coreProps.keywords === 'string'
          ? coreProps.keywords
          : coreProps.keywords['#text'] || '';
      }
      if (coreProps.lastModifiedBy) {
        docProps.lastModifiedBy = typeof coreProps.lastModifiedBy === 'string'
          ? coreProps.lastModifiedBy
          : coreProps.lastModifiedBy['#text'] || '';
      }
      if (coreProps.revision) {
        docProps.revision = typeof coreProps.revision === 'string'
          ? coreProps.revision
          : coreProps.revision['#text'] || '';
      }
      if (coreProps.category) {
        docProps.category = typeof coreProps.category === 'string'
          ? coreProps.category
          : coreProps.category['#text'] || '';
      }

      // Dates
      if (coreProps.created) {
        const dateStr = typeof coreProps.created === 'string'
          ? coreProps.created
          : coreProps.created['#text'] || '';
        if (dateStr) docProps.created = this._parseDatetime(dateStr);
      }
      if (coreProps.modified) {
        const dateStr = typeof coreProps.modified === 'string'
          ? coreProps.modified
          : coreProps.modified['#text'] || '';
        if (dateStr) docProps.modified = this._parseDatetime(dateStr);
      }
    } catch {
      // Ignore parse errors
    }
  }

  private async _loadAppProperties(zip: JSZip): Promise<void> {
    const file = zip.file('docProps/app.xml');
    if (!file) return;

    try {
      const buf = await file.async('nodebuffer');
      const appParser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        textNodeName: '#text',
        removeNSPrefix: true,
        parseTagValue: false,
        trimValues: false,
    processEntities: false,
      });
      const root = appParser.parse(buf.toString('utf-8'));
      const appProps = root.Properties || root;

      const docProps = this._workbook.documentProperties;

      if (appProps.Application) {
        docProps.application = typeof appProps.Application === 'string'
          ? appProps.Application
          : appProps.Application['#text'] || '';
      }
      if (appProps.AppVersion) {
        docProps.appVersion = typeof appProps.AppVersion === 'string'
          ? appProps.AppVersion
          : appProps.AppVersion['#text'] || '';
      }
      if (appProps.Company) {
        docProps.company = typeof appProps.Company === 'string'
          ? appProps.Company
          : appProps.Company['#text'] || '';
      }
      if (appProps.Manager) {
        docProps.manager = typeof appProps.Manager === 'string'
          ? appProps.Manager
          : appProps.Manager['#text'] || '';
      }
    } catch {
      // Ignore parse errors
    }
  }

  private _parseDatetime(dateStr: string): Date | string {
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      return d;
    } catch {
      return dateStr;
    }
  }
}
