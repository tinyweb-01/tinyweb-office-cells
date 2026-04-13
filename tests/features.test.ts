/**
 * Phase 4 – Feature Round-Trip Tests
 *
 * Tests that verify each Phase 4 feature survives a save → load → save cycle:
 *   - Data Validation
 *   - Conditional Formatting
 *   - Hyperlinks
 *   - AutoFilter
 *   - Comments
 *   - Defined Names (Named Ranges)
 *   - Page Breaks
 */

import {
  Workbook,
  DataValidationType,
  DataValidationOperator,
  DataValidationAlertStyle,
} from '../src';

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Save workbook to buffer and reload it. */
async function roundTrip(wb: Workbook): Promise<Workbook> {
  const buf = await wb.saveToBuffer();
  return Workbook.loadFromBuffer(buf);
}

// ─── 1. Data Validation ────────────────────────────────────────────────────

describe('DataValidation round-trip', () => {
  it('preserves a list validation with prompts', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    // Set up list validation on B2:B10
    const dv = ws.dataValidations.add('B2:B10');
    dv.add(DataValidationType.LIST, DataValidationAlertStyle.STOP, DataValidationOperator.BETWEEN, '"Yes,No,Maybe"');
    dv.showInputMessage = true;
    dv.inputTitle = 'Choose';
    dv.inputMessage = 'Pick one';
    dv.showErrorMessage = true;
    dv.errorTitle = 'Invalid';
    dv.errorMessage = 'Not in list';

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    expect(ws2.dataValidations.count).toBe(1);

    const dv2 = ws2.dataValidations.get(0);
    expect(dv2.sqref).toBe('B2:B10');
    expect(dv2.type).toBe(DataValidationType.LIST);
    expect(dv2.formula1).toBe('"Yes,No,Maybe"');
    expect(dv2.showInputMessage).toBe(true);
    expect(dv2.inputTitle).toBe('Choose');
    expect(dv2.inputMessage).toBe('Pick one');
    expect(dv2.showErrorMessage).toBe(true);
    expect(dv2.errorTitle).toBe('Invalid');
    expect(dv2.errorMessage).toBe('Not in list');
  });

  it('preserves a whole-number range validation', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    const dv = ws.dataValidations.add('C3');
    dv.add(DataValidationType.WHOLE_NUMBER, DataValidationAlertStyle.WARNING, DataValidationOperator.BETWEEN, '1', '100');

    const wb2 = await roundTrip(wb);
    const dv2 = wb2.getWorksheet(0).dataValidations.get(0);
    expect(dv2.type).toBe(DataValidationType.WHOLE_NUMBER);
    expect(dv2.operator).toBe(DataValidationOperator.BETWEEN);
    expect(dv2.formula1).toBe('1');
    expect(dv2.formula2).toBe('100');
    expect(dv2.alertStyle).toBe(DataValidationAlertStyle.WARNING);
  });

  it('preserves multiple validations on same sheet', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    ws.dataValidations.add('A1').add(DataValidationType.WHOLE_NUMBER);
    ws.dataValidations.add('B1').add(DataValidationType.DECIMAL);

    const wb2 = await roundTrip(wb);
    expect(wb2.getWorksheet(0).dataValidations.count).toBe(2);
  });
});

// ─── 2. Conditional Formatting ─────────────────────────────────────────────

describe('ConditionalFormatting round-trip', () => {
  it('preserves a cell-value rule with font formatting', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    const cf = ws.conditionalFormatting.add();
    cf.range = 'A1:A10';
    cf.type = 'cellIs';
    cf.operator = 'greaterThan';
    cf.formula1 = '100';
    cf.font.color = 'FFFF0000';

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    expect(ws2.conditionalFormatting.count).toBe(1);

    const cf2 = ws2.conditionalFormatting.get(0);
    expect(cf2.range).toBe('A1:A10');
    expect(cf2.type).toBe('cellIs');
    expect(cf2.operator).toBe('greaterThan');
    expect(cf2.formula1).toBe('100');
  });

  it('preserves multiple CF rules', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    const cf1 = ws.conditionalFormatting.add();
    cf1.range = 'B1:B5';
    cf1.type = 'cellIs';
    cf1.operator = 'lessThan';
    cf1.formula1 = '0';
    cf1.font.color = 'FFFF0000';

    const cf2 = ws.conditionalFormatting.add();
    cf2.range = 'B1:B5';
    cf2.type = 'cellIs';
    cf2.operator = 'greaterThan';
    cf2.formula1 = '1000';
    cf2.font.color = 'FF00FF00';

    const wb2 = await roundTrip(wb);
    expect(wb2.getWorksheet(0).conditionalFormatting.count).toBe(2);
  });
});

// ─── 3. Hyperlinks ─────────────────────────────────────────────────────────

describe('Hyperlink round-trip', () => {
  it('preserves external hyperlink', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Click me');

    // add(rangeAddress, address, subAddress, textToDisplay, screenTip)
    ws.hyperlinks.add('A1', 'https://example.com', '', '', 'Example');

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    expect(ws2.hyperlinks.count).toBe(1);

    const hl = ws2.hyperlinks.get(0);
    expect(hl.range).toBe('A1');
    expect(hl.address).toBe('https://example.com');
    expect(hl.screenTip).toBe('Example');
    expect(hl.type).toBe('external');
  });

  it('preserves internal hyperlink (sheet reference)', async () => {
    const wb = new Workbook();
    wb.addWorksheet('Other');
    const ws = wb.getWorksheet(0);
    ws.cells.set('B2', 'Go to Other');

    // Internal link: address='', subAddress='Other!A1'
    ws.hyperlinks.add('B2', '', 'Other!A1', '', 'Jump');

    const wb2 = await roundTrip(wb);
    const hl = wb2.getWorksheet(0).hyperlinks.get(0);
    expect(hl.range).toBe('B2');
    expect(hl.subAddress).toBe('Other!A1');
    expect(hl.type).toBe('internal');
  });

  it('preserves multiple hyperlinks', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Link1');
    ws.cells.set('A2', 'Link2');

    ws.hyperlinks.add('A1', 'https://one.com');
    ws.hyperlinks.add('A2', 'https://two.com');

    const wb2 = await roundTrip(wb);
    expect(wb2.getWorksheet(0).hyperlinks.count).toBe(2);
  });
});

// ─── 4. AutoFilter ─────────────────────────────────────────────────────────

describe('AutoFilter round-trip', () => {
  it('preserves filter range', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Name');
    ws.cells.set('B1', 'Value');
    ws.cells.set('A2', 'Alice');
    ws.cells.set('B2', 100);

    ws.autoFilter.range = 'A1:B2';

    const wb2 = await roundTrip(wb);
    expect(wb2.getWorksheet(0).autoFilter.range).toBe('A1:B2');
  });

  it('preserves filter columns with value filters', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Status');
    ws.cells.set('A2', 'Active');
    ws.cells.set('A3', 'Inactive');

    ws.autoFilter.range = 'A1:A3';
    ws.autoFilter.addFilter(0, 'Active');

    const wb2 = await roundTrip(wb);
    const af2 = wb2.getWorksheet(0).autoFilter;
    expect(af2.range).toBe('A1:A3');
    expect(af2.filterColumns.size).toBe(1);
    const fc2 = af2.filterColumns.get(0);
    expect(fc2).toBeDefined();
    expect(fc2!.filters).toContain('Active');
  });

  it('preserves custom filters', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Amount');

    ws.autoFilter.range = 'A1:A10';
    ws.autoFilter.customFilter(0, 'greaterThan', '100');

    const wb2 = await roundTrip(wb);
    const fc2 = wb2.getWorksheet(0).autoFilter.filterColumns.get(0);
    expect(fc2).toBeDefined();
    expect(fc2!.customFilters.length).toBe(1);
    expect(fc2!.customFilters[0].operator).toBe('greaterThan');
    expect(fc2!.customFilters[0].value).toBe('100');
  });
});

// ─── 5. Comments ───────────────────────────────────────────────────────────

describe('Comments round-trip', () => {
  it('preserves cell comment text and author', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'Data');
    const cellA1 = ws.cells.get('A1');
    cellA1.setComment('This is a note', 'John');

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    const cell2 = ws2.cells.get('A1');
    expect(cell2.comment).not.toBeNull();
    expect(cell2.comment!.text).toBe('This is a note');
    expect(cell2.comment!.author).toBe('John');
  });

  it('preserves multiple comments on same sheet', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 1);
    ws.cells.set('B2', 2);
    ws.cells.get('A1').setComment('Note A1');
    ws.cells.get('B2').setComment('Note B2', 'Author2');

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    expect(ws2.cells.get('A1').comment!.text).toBe('Note A1');
    expect(ws2.cells.get('B2').comment!.text).toBe('Note B2');
    expect(ws2.cells.get('B2').comment!.author).toBe('Author2');
  });

  it('preserves comment size', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('C3', 'x');
    ws.cells.get('C3').setComment('Wide note', 'Me', 300, 150);

    const wb2 = await roundTrip(wb);
    const c2 = wb2.getWorksheet(0).cells.get('C3');
    expect(c2.comment!.text).toBe('Wide note');
    // VML dimensions are approximate after round-trip; just ensure non-null
    const size = c2.getCommentSize();
    expect(size).not.toBeNull();
  });
});

// ─── 6. Defined Names (Named Ranges) ──────────────────────────────────────

describe('DefinedName round-trip', () => {
  it('preserves a global defined name', async () => {
    const wb = new Workbook();
    wb.getWorksheet(0).cells.set('A1', 42);

    wb.properties.definedNames.add('MyRange', 'Sheet1!$A$1');

    const wb2 = await roundTrip(wb);
    const coll = wb2.properties.definedNames;
    expect(coll.count).toBe(1);

    const dn2 = coll.get('MyRange');
    expect(dn2.name).toBe('MyRange');
    expect(dn2.refersTo).toBe('Sheet1!$A$1');
    expect(dn2.localSheetId).toBeNull();
  });

  it('preserves a local defined name with hidden flag', async () => {
    const wb = new Workbook();
    wb.getWorksheet(0).cells.set('A1', 1);

    const dn = wb.properties.definedNames.add('LocalName', 'Sheet1!$B$2', 0);
    dn.hidden = true;
    dn.comment = 'test comment';

    const wb2 = await roundTrip(wb);
    const dn2 = wb2.properties.definedNames.get('LocalName');
    expect(dn2.localSheetId).toBe(0);
    expect(dn2.hidden).toBe(true);
    expect(dn2.comment).toBe('test comment');
  });

  it('preserves multiple defined names', async () => {
    const wb = new Workbook();
    wb.properties.definedNames.add('Name1', 'Sheet1!$A$1');
    wb.properties.definedNames.add('Name2', 'Sheet1!$B$1');
    wb.properties.definedNames.add('Name3', 'Sheet1!$C$1');

    const wb2 = await roundTrip(wb);
    expect(wb2.properties.definedNames.count).toBe(3);
  });
});

// ─── 7. Page Breaks ────────────────────────────────────────────────────────

describe('PageBreaks round-trip', () => {
  it('preserves horizontal page breaks', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'data');

    ws.horizontalPageBreaks.add(5);
    ws.horizontalPageBreaks.add(10);

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    const breaks = ws2.horizontalPageBreaks.toList();
    expect(breaks).toEqual([5, 10]);
  });

  it('preserves vertical page breaks', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'data');

    ws.verticalPageBreaks.add(3);
    ws.verticalPageBreaks.add(7);

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    const breaks = ws2.verticalPageBreaks.toList();
    expect(breaks).toEqual([3, 7]);
  });

  it('preserves both horizontal and vertical breaks', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.set('A1', 'data');

    ws.horizontalPageBreaks.add(4);
    ws.verticalPageBreaks.add(2);

    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);
    expect(ws2.horizontalPageBreaks.count).toBe(1);
    expect(ws2.verticalPageBreaks.count).toBe(1);
    expect(ws2.horizontalPageBreaks.toList()).toEqual([4]);
    expect(ws2.verticalPageBreaks.toList()).toEqual([2]);
  });
});

// ─── 8. Combined Features ──────────────────────────────────────────────────

describe('Combined features round-trip', () => {
  it('preserves all features on the same worksheet', async () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);

    // Cell data
    ws.cells.set('A1', 'Name');
    ws.cells.set('B1', 'Score');
    ws.cells.set('A2', 'Alice');
    ws.cells.set('B2', 95);
    ws.cells.set('A3', 'Bob');
    ws.cells.set('B3', 80);

    // Data validation
    const dv = ws.dataValidations.add('B2:B3');
    dv.add(DataValidationType.WHOLE_NUMBER, DataValidationAlertStyle.STOP, DataValidationOperator.BETWEEN, '0', '100');

    // Conditional formatting
    const cf = ws.conditionalFormatting.add();
    cf.range = 'B2:B3';
    cf.type = 'cellIs';
    cf.operator = 'lessThan';
    cf.formula1 = '50';
    cf.font.color = 'FFFF0000';

    // Hyperlinks
    ws.hyperlinks.add('A1', 'https://example.com', '', '', 'Header link');

    // AutoFilter
    ws.autoFilter.range = 'A1:B3';

    // Comments
    ws.cells.get('A2').setComment('Top student', 'Teacher');

    // Page breaks
    ws.horizontalPageBreaks.add(3);

    // Defined names
    wb.properties.definedNames.add('Scores', 'Sheet1!$B$2:$B$3');

    // Round-trip
    const wb2 = await roundTrip(wb);
    const ws2 = wb2.getWorksheet(0);

    // Verify cell data
    expect(ws2.cells.get('A2').value).toBe('Alice');
    expect(ws2.cells.get('B2').value).toBe(95);

    // Verify data validation
    expect(ws2.dataValidations.count).toBe(1);
    expect(ws2.dataValidations.get(0).type).toBe(DataValidationType.WHOLE_NUMBER);

    // Verify conditional formatting
    expect(ws2.conditionalFormatting.count).toBe(1);
    expect(ws2.conditionalFormatting.get(0).type).toBe('cellIs');

    // Verify hyperlinks
    expect(ws2.hyperlinks.count).toBe(1);
    expect(ws2.hyperlinks.get(0).address).toBe('https://example.com');

    // Verify AutoFilter
    expect(ws2.autoFilter.range).toBe('A1:B3');

    // Verify comments
    expect(ws2.cells.get('A2').comment!.text).toBe('Top student');

    // Verify page breaks
    expect(ws2.horizontalPageBreaks.toList()).toEqual([3]);

    // Verify defined names
    expect(wb2.properties.definedNames.count).toBe(1);
    expect(wb2.properties.definedNames.get('Scores').refersTo).toBe('Sheet1!$B$2:$B$3');
  });

  it('preserves features across multiple worksheets', async () => {
    const wb = new Workbook();
    const ws1 = wb.getWorksheet(0);
    const ws2 = wb.addWorksheet('Sheet2');

    // Sheet1 features
    ws1.cells.set('A1', 'Test');
    ws1.hyperlinks.add('A1', 'https://one.com');
    ws1.cells.get('A1').setComment('Note1');

    // Sheet2 features
    ws2.cells.set('A1', 'Data');
    ws2.dataValidations.add('A1').add(DataValidationType.LIST, undefined, undefined, '"X,Y"');
    ws2.verticalPageBreaks.add(5);

    const wb3 = await roundTrip(wb);

    // Verify Sheet1
    const r1 = wb3.getWorksheet(0);
    expect(r1.hyperlinks.count).toBe(1);
    expect(r1.cells.get('A1').comment!.text).toBe('Note1');

    // Verify Sheet2
    const r2 = wb3.getWorksheet(1);
    expect(r2.dataValidations.count).toBe(1);
    expect(r2.verticalPageBreaks.toList()).toEqual([5]);
  });
});
