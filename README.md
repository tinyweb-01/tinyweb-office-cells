# tinyweb-office-cells

> Open-source Node.js/TypeScript library for reading, writing, and manipulating Excel XLSX files with full formatting, formula evaluation, and Aspose Cells-compatible API.

[![npm version](https://img.shields.io/npm/v/tinyweb-office-cells.svg)](https://www.npmjs.com/package/tinyweb-office-cells)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- ✅ **Read/Write XLSX files** — full round-trip fidelity (preserves unmodified XML)
- ✅ **Cell formatting** — fonts, colors, borders, fills, number formats, alignment
- ✅ **Merge cells** — column widths, row heights
- ✅ **Formula evaluation** — 43 built-in Excel functions
- ✅ **Data validation** — dropdown lists, numeric ranges, date ranges, custom formulas
- ✅ **Conditional formatting** — cell value rules, color scales, data bars, formulas
- ✅ **Named ranges** — workbook and sheet-scoped defined names
- ✅ **Hyperlinks** — external URLs, internal sheet references, email links
- ✅ **Comments** — cell comments with author and size support
- ✅ **AutoFilter** — column filters with custom, color, dynamic, and top-10 filters
- ✅ **Page breaks** — horizontal and vertical page break management
- ✅ **Sheet/workbook protection** — password-based protection with granular permissions
- ✅ **Document properties** — title, author, subject, keywords, and custom properties
- ✅ **Page setup** — orientation, paper size, margins, print area, headers/footers
- ✅ **Freeze panes** — freeze rows and columns
- ✅ **Dual API naming** — both `camelCase` and `snake_case` method names
- ✅ **TypeScript-first** — full type declarations included
- ✅ **Zero native dependencies** — pure JavaScript, works everywhere Node.js runs

## Installation

```bash
npm install tinyweb-office-cells
```

## Quick Start

### Create a Workbook and Save

```typescript
import { Workbook, SaveFormat } from 'tinyweb-office-cells';

// Create a new workbook
const workbook = new Workbook();
const sheet = workbook.worksheets.get(0);

// Set cell values
sheet.cells.get('A1').putValue('Name');
sheet.cells.get('B1').putValue('Score');
sheet.cells.get('A2').putValue('Alice');
sheet.cells.get('B2').putValue(95);
sheet.cells.get('A3').putValue('Bob');
sheet.cells.get('B3').putValue(87);

// Add a formula
sheet.cells.get('B4').setFormula('=SUM(B2:B3)');

// Save to file
await workbook.save('output.xlsx');
```

### Read an Existing File

```typescript
import { Workbook } from 'tinyweb-office-cells';

const workbook = new Workbook();
await workbook.loadFile('input.xlsx');

const sheet = workbook.worksheets.get(0);
console.log(sheet.cells.get('A1').value); // Read cell value
console.log(sheet.name);                  // Sheet name
```

### Apply Formatting

```typescript
import { Workbook } from 'tinyweb-office-cells';

const workbook = new Workbook();
const sheet = workbook.worksheets.get(0);
const cell = sheet.cells.get('A1');
cell.putValue('Hello, Excel!');

// Get cell style and modify
const style = cell.getStyle();
style.font.name = 'Arial';
style.font.size = 14;
style.font.bold = true;
style.font.color = '#FF0000';
style.fill.patternType = 'solid';
style.fill.fgColor = '#FFFF00';
style.borders.top.lineStyle = 'thin';
style.borders.top.color = '#000000';
style.alignment.horizontal = 'center';
cell.setStyle(style);

await workbook.save('formatted.xlsx');
```

## API Reference

### Workbook

The main entry point for creating and manipulating Excel files.

```typescript
const workbook = new Workbook();

// Load from file
await workbook.loadFile('path/to/file.xlsx');

// Load from buffer
await workbook.loadBuffer(buffer);

// Save to file
await workbook.save('output.xlsx');
await workbook.save('output.xlsx', SaveFormat.XLSX);

// Save to buffer
const buffer = await workbook.saveToBuffer(SaveFormat.XLSX);

// Access worksheets
workbook.worksheets;           // WorksheetCollection
workbook.worksheets.get(0);    // By index
workbook.worksheets.get('Sheet1'); // By name
workbook.worksheets.add('NewSheet');

// Workbook protection
workbook.properties.protection.lockStructure = true;
workbook.properties.protection.password = 'secret';

// Document properties
workbook.documentProperties.title = 'My Report';
workbook.documentProperties.author = 'John Doe';
workbook.documentProperties.subject = 'Q4 Sales';
```

### Worksheet

Represents a single sheet in the workbook.

```typescript
const sheet = workbook.worksheets.get(0);

// Properties
sheet.name = 'DataSheet';
sheet.isVisible = true;
sheet.tabColor = '#FF0000';

// Cells
sheet.cells.get('A1').putValue('Hello');

// Merge cells
sheet.cells.merge(0, 0, 2, 3); // startRow, startCol, rowCount, colCount

// Column width and row height
sheet.cells.setColumnWidth(0, 20);  // Column A = 20
sheet.cells.setRowHeight(0, 30);    // Row 1 = 30

// Freeze panes
sheet.freezePanes(1, 0);  // Freeze first row

// Page setup
sheet.pageSetup.orientation = 'landscape';
sheet.pageSetup.paperSize = 'A4';

// Sheet protection
sheet.protect('password');
sheet.protection.sheet = true;
sheet.protection.formatCells = false;
```

### Cell / Cells

Access and manipulate individual cells.

```typescript
const cells = sheet.cells;

// Get a cell (creates if not exists)
const cell = cells.get('A1');
// Also supports: cells.get(0, 0)

// Set values
cell.putValue('text');      // String
cell.putValue(42);          // Number
cell.putValue(true);        // Boolean
cell.putValue(new Date());  // Date

// Read values
cell.value;          // Raw value
cell.stringValue;    // String representation
cell.type;           // CellDataType

// Formulas
cell.setFormula('=SUM(A1:A10)');
cell.formula;        // Get formula string
cell.hasFormula();   // Check if cell has formula

// Styling
const style = cell.getStyle();
style.font.bold = true;
cell.setStyle(style);

// Comments
cell.comment = { text: 'Note', author: 'Admin' };
```

### Style (Font, Fill, Borders, Alignment, NumberFormat)

Full cell formatting support.

```typescript
const style = cell.getStyle();

// Font
style.font.name = 'Calibri';
style.font.size = 12;
style.font.bold = true;
style.font.italic = true;
style.font.underline = 'single';
style.font.strikethrough = true;
style.font.color = '#0000FF';

// Fill
style.fill.patternType = 'solid';
style.fill.fgColor = '#FFFF00';
style.fill.bgColor = '#FFFFFF';

// Borders
style.borders.top.lineStyle = 'thin';      // thin, medium, thick, dashed, dotted, double
style.borders.top.color = '#000000';
style.borders.bottom.lineStyle = 'double';
style.borders.left.lineStyle = 'medium';
style.borders.right.lineStyle = 'thin';
style.borders.diagonal.lineStyle = 'thin';

// Alignment
style.alignment.horizontal = 'center';   // left, center, right, fill, justify
style.alignment.vertical = 'center';     // top, center, bottom
style.alignment.wrapText = true;
style.alignment.textRotation = 45;
style.alignment.indent = 2;
style.alignment.shrinkToFit = true;

// Number format
style.numberFormat.formatCode = '#,##0.00';
style.numberFormat.formatCode = '0.00%';
style.numberFormat.formatCode = 'yyyy-mm-dd';
style.numberFormat.formatCode = '$#,##0.00';

// Protection
style.protection.locked = true;
style.protection.hidden = true;

cell.setStyle(style);
```

### FormulaEvaluator

Evaluate formulas in cells.

```typescript
import { FormulaEvaluator } from 'tinyweb-office-cells';

// Set up formulas
sheet.cells.get('A1').putValue(10);
sheet.cells.get('A2').putValue(20);
sheet.cells.get('A3').setFormula('=SUM(A1:A2)');

// Evaluate all formulas
const evaluator = new FormulaEvaluator(workbook);
evaluator.evaluateAll();

console.log(sheet.cells.get('A3').value); // 30
```

### Data Validation

Add data validation rules to cells.

```typescript
import { DataValidation, DataValidationType, DataValidationOperator } from 'tinyweb-office-cells';

const validations = sheet.dataValidations;

// Dropdown list
const dv = validations.add();
dv.sqref = 'A1:A10';
dv.type = DataValidationType.LIST;
dv.formula1 = '"Apple,Banana,Cherry"';

// Numeric range
const numDv = validations.add();
numDv.sqref = 'B1:B10';
numDv.type = DataValidationType.WHOLE;
numDv.operator = DataValidationOperator.BETWEEN;
numDv.formula1 = '1';
numDv.formula2 = '100';
numDv.showErrorMessage = true;
numDv.errorTitle = 'Invalid Input';
numDv.error = 'Please enter a number between 1 and 100';
```

### Conditional Formatting

Apply conditional formatting rules.

```typescript
const cfCollection = sheet.conditionalFormattings;

// Cell value rule
const cf = cfCollection.add();
cf.sqref = 'A1:A20';
cf.type = 'cellIs';
cf.operator = 'greaterThan';
cf.formula1 = '90';
cf.priority = 1;
cf.dxf = {
  font: { color: '#006100' },
  fill: { bgColor: '#C6EFCE' },
};
```

### Hyperlinks

Add hyperlinks to cells.

```typescript
const hyperlinks = sheet.hyperlinks;

// External URL
hyperlinks.add('A1', 1, 1, 'https://example.com');
const link = hyperlinks.get(0);
link.textToDisplay = 'Visit Example';
link.screenTip = 'Click to open';

// Internal reference
hyperlinks.add('A2', 1, 1, '#Sheet2!A1');
```

### Named Ranges

Define and use named ranges.

```typescript
// Add a defined name
workbook.definedNames.add('SalesData', 'Sheet1!$A$1:$D$100');
workbook.definedNames.add('TaxRate', '0.08');

// Use in formulas
sheet.cells.get('E1').setFormula('=SUM(SalesData)');
```

### AutoFilter

Apply auto-filtering to data ranges.

```typescript
const autoFilter = sheet.autoFilter;
autoFilter.range = 'A1:D100';

// Add a filter column
const fc = autoFilter.addFilterColumn(0);
fc.setValues(['Active', 'Pending']);
```

### Comments

Add comments to cells.

```typescript
const cell = sheet.cells.get('A1');
cell.comment = {
  text: 'This is a comment',
  author: 'Admin',
  width: 200,
  height: 100,
};
```

### Page Setup

Configure print settings.

```typescript
sheet.pageSetup.orientation = 'landscape';
sheet.pageSetup.paperSize = 'A4';
sheet.pageSetup.printArea = 'A1:H50';
sheet.pageSetup.fitToPagesTall = 1;
sheet.pageSetup.fitToPagesWide = 1;
```

## Formula Support

The built-in formula evaluator supports **43 Excel functions** across 7 categories:

### Math & Statistics (16 functions)
`SUM` · `AVERAGE` · `MIN` · `MAX` · `COUNT` · `COUNTA` · `ABS` · `ROUND` · `ROUNDUP` · `ROUNDDOWN` · `INT` · `MOD` · `POWER` · `SQRT` · `CEILING` · `FLOOR`

### String Functions (11 functions)
`CONCATENATE` · `CONCAT` · `TEXT` · `LEN` · `TRIM` · `UPPER` · `LOWER` · `LEFT` · `RIGHT` · `MID` · `SUBSTITUTE` · `REPT`

### Logic Functions (4 functions)
`IF` · `AND` · `OR` · `NOT`

### Lookup & Reference (4 functions)
`VLOOKUP` · `HLOOKUP` · `INDEX` · `MATCH`

### Date & Time (6 functions)
`TODAY` · `NOW` · `DATE` · `YEAR` · `MONTH` · `DAY`

### Information Functions (5 functions)
`ISNUMBER` · `ISTEXT` · `ISBLANK` · `ISERROR` · `ISNA`

### Other Functions (2 functions)
`VALUE` · `CHOOSE`

### Also Supports
- Arithmetic operators: `+` `-` `*` `/` `^`
- Comparison operators: `=` `<>` `<` `>` `<=` `>=`
- String concatenation: `&`
- Cell references: `A1`, `$A$1`, `Sheet1!A1`
- Range references: `A1:B5`
- Defined names / named ranges
- Nested formulas with recursive evaluation
- Error values: `#VALUE!`, `#REF!`, `#NAME?`, `#DIV/0!`, `#N/A`

## Examples

### Create and Save a Workbook

```typescript
import { Workbook } from 'tinyweb-office-cells';

const wb = new Workbook();
const ws = wb.worksheets.get(0);
ws.name = 'Sales Report';

// Headers
['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'].forEach((h, i) => {
  const cell = ws.cells.get(0, i);
  cell.putValue(h);
  const style = cell.getStyle();
  style.font.bold = true;
  style.fill.patternType = 'solid';
  style.fill.fgColor = '#4472C4';
  style.font.color = '#FFFFFF';
  cell.setStyle(style);
});

// Data
const data = [
  ['Widget A', 1200, 1350, 1100, 1500],
  ['Widget B', 800, 920, 870, 1050],
];

data.forEach((row, r) => {
  row.forEach((val, c) => {
    ws.cells.get(r + 1, c).putValue(val);
  });
  // Total formula
  ws.cells.get(r + 1, 5).setFormula(`=SUM(B${r + 2}:E${r + 2})`);
});

await wb.save('sales-report.xlsx');
```

### Read and Modify an Existing File

```typescript
import { Workbook } from 'tinyweb-office-cells';

const wb = new Workbook();
await wb.loadFile('existing.xlsx');

const ws = wb.worksheets.get(0);

// Read all data
for (let r = 0; r < 100; r++) {
  for (let c = 0; c < 10; c++) {
    const cell = ws.cells.get(r, c);
    if (cell.value !== null) {
      console.log(`${cell.name}: ${cell.value}`);
    }
  }
}

// Modify a cell
ws.cells.get('A1').putValue('Updated!');
await wb.save('modified.xlsx');
```

### Data Validation with Dropdown

```typescript
import { Workbook, DataValidationType } from 'tinyweb-office-cells';

const wb = new Workbook();
const ws = wb.worksheets.get(0);

ws.cells.get('A1').putValue('Select a fruit:');

const dv = ws.dataValidations.add();
dv.sqref = 'B1';
dv.type = DataValidationType.LIST;
dv.formula1 = '"Apple,Banana,Cherry,Date,Elderberry"';
dv.showDropDown = true;
dv.showInputMessage = true;
dv.inputTitle = 'Fruit Selection';
dv.inputMessage = 'Choose a fruit from the list';

await wb.save('dropdown.xlsx');
```

### Conditional Formatting

```typescript
import { Workbook } from 'tinyweb-office-cells';

const wb = new Workbook();
const ws = wb.worksheets.get(0);

// Add some scores
for (let i = 0; i < 20; i++) {
  ws.cells.get(i, 0).putValue(Math.floor(Math.random() * 100));
}

// Highlight scores above 80 in green
const cf = ws.conditionalFormattings.add();
cf.sqref = 'A1:A20';
cf.type = 'cellIs';
cf.operator = 'greaterThanOrEqual';
cf.formula1 = '80';
cf.dxf = {
  fill: { bgColor: '#C6EFCE' },
  font: { color: '#006100' },
};

await wb.save('conditional.xlsx');
```

## API Compatibility

This library provides **dual naming conventions** for maximum compatibility:

| camelCase (TypeScript) | snake_case (Python-compat) |
|---|---|
| `cell.putValue(v)` | `cell.put_value(v)` |
| `cell.getStyle()` | `cell.get_style()` |
| `cell.setStyle(s)` | `cell.set_style(s)` |
| `cell.hasFormula()` | `cell.has_formula()` |
| `cell.setFormula(f)` | `cell.set_formula(f)` |
| `cell.stringValue` | `cell.string_value` |
| `cells.setColumnWidth()` | `cells.set_column_width()` |
| `cells.setRowHeight()` | `cells.set_row_height()` |
| `workbook.loadFile()` | `workbook.load_file()` |
| `workbook.saveToBuffer()` | `workbook.save_to_buffer()` |

Both naming styles work identically — use whichever fits your codebase.

## TypeScript

Full TypeScript declarations are included. All classes, interfaces, enums, and types are exported:

```typescript
import {
  Workbook,
  Worksheet,
  Cell,
  Cells,
  Style,
  Font,
  Fill,
  Border,
  Borders,
  Alignment,
  NumberFormat,
  Protection,
  SaveFormat,
  FormulaEvaluator,
  DataValidation,
  DataValidationCollection,
  DataValidationType,
  DataValidationOperator,
  ConditionalFormat,
  ConditionalFormatCollection,
  Hyperlink,
  HyperlinkCollection,
  AutoFilter,
  FilterColumn,
  DefinedName,
  DefinedNameCollection,
  HorizontalPageBreakCollection,
  VerticalPageBreakCollection,
  SheetProtection,
  PageSetup,
  FreezePane,
  WorkbookProtection,
  DocumentProperties,
} from 'tinyweb-office-cells';
```

## Dependencies

| Package | Purpose |
|---|---|
| [jszip](https://www.npmjs.com/package/jszip) | XLSX ZIP container read/write |
| [fast-xml-parser](https://www.npmjs.com/package/fast-xml-parser) | XML parsing and generation |

No native/binary dependencies — works on any platform where Node.js runs.

## License

[MIT](./LICENSE)
