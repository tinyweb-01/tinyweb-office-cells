/**
 * tinyweb-office-cells – Public API surface
 *
 * Re-exports every class, enum, interface, and type that consumers need.
 */

// ---------------------------------------------------------------------------
// Styling
// ---------------------------------------------------------------------------
export {
  Font,
  Fill,
  Border,
  BorderType,
  Borders,
  Alignment,
  NumberFormat,
  Protection,
  Style,
} from './styling/Style';

// ---------------------------------------------------------------------------
// Core
// ---------------------------------------------------------------------------
export {
  Cell,
  CellComment,
  CellDataType,
  CellValue,
} from './core/Cell';

export {
  Cells,
  WorksheetLike,
} from './core/Cells';

export {
  Worksheet,
  SheetProtection,
  SheetProtectionDictWrapper,
  PageSetup,
  PageMargins,
  FreezePane,
} from './core/Worksheet';

export {
  Workbook,
  SaveFormat,
  saveFormatFromExtension,
  save_format_from_extension,
  WorkbookProtection,
  WorkbookProperties,
  WorkbookView,
  DocumentProperties,
} from './core/Workbook';

// ---------------------------------------------------------------------------
// Features (Phase 4)
// ---------------------------------------------------------------------------
export {
  DataValidation,
  DataValidationCollection,
  DataValidationType,
  DataValidationOperator,
  DataValidationAlertStyle,
  DataValidationImeMode,
} from './features/DataValidation';

export {
  ConditionalFormat,
  ConditionalFormatCollection,
} from './features/ConditionalFormat';

export {
  Hyperlink,
  HyperlinkCollection,
} from './features/Hyperlink';
export type { HyperlinkType } from './features/Hyperlink';

export {
  AutoFilter,
  FilterColumn,
} from './features/AutoFilter';
export type {
  CustomFilter,
  ColorFilter,
  DynamicFilter,
  Top10Filter,
  SortState,
} from './features/AutoFilter';

export {
  HorizontalPageBreakCollection,
  VerticalPageBreakCollection,
} from './features/PageBreak';

export {
  DefinedName,
  DefinedNameCollection,
} from './features/DefinedName';

export {
  FormulaEvaluator,
} from './features/FormulaEvaluator';

// ---------------------------------------------------------------------------
// I/O – Comment XML
// ---------------------------------------------------------------------------
export { CommentXmlWriter, CommentXmlReader } from './io/CommentXml';

// ---------------------------------------------------------------------------
// I/O – Feature Loaders / Savers
// ---------------------------------------------------------------------------
export { AutoFilterXmlLoader } from './io/XmlAutoFilterLoader';
export { AutoFilterXmlSaver } from './io/XmlAutoFilterSaver';
export { ConditionalFormatXmlLoader } from './io/XmlConditionalFormatLoader';
export { ConditionalFormatXmlSaver } from './io/XmlConditionalFormatSaver';
export { DataValidationXmlLoader } from './io/XmlDataValidationLoader';
export { DataValidationXmlSaver } from './io/XmlDataValidationSaver';
export {
  HyperlinkXmlLoader,
  HyperlinkXmlSaver,
  HyperlinkRelationshipWriter,
} from './io/XmlHyperlinkHandler';
export type { HyperlinkRelationship } from './io/XmlHyperlinkHandler';

// ---------------------------------------------------------------------------
// I/O (Phase 2 + Phase 3)
// ---------------------------------------------------------------------------
export { SharedStringTable } from './io/SharedStrings';
export {
  CellValueHandler,
  CELL_TYPE_STRING,
  CELL_TYPE_NUMBER,
  CELL_TYPE_BOOLEAN,
  CELL_TYPE_ERROR,
} from './io/CellValueHandler';
export { XmlSaver } from './io/XmlSaver';
export { XmlLoader } from './io/XmlLoader';

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------
export {
  worksheetToHtml,
  worksheetToPng,
} from './rendering';
export type {
  HtmlRenderOptions,
  PngRenderOptions,
} from './rendering';
