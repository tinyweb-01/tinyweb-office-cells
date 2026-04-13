/**
 * Phase 5 – FormulaEvaluator Tests
 *
 * Comprehensive tests for the formula evaluation engine covering:
 *   - Arithmetic & operator precedence
 *   - Comparison & concatenation operators
 *   - String functions (CONCATENATE, CONCAT, TEXT, LEN, TRIM, UPPER, LOWER, LEFT, RIGHT, MID, SUBSTITUTE, REPT)
 *   - Logic functions (IF, AND, OR, NOT)
 *   - Math/stat functions (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, CEILING, FLOOR)
 *   - Lookup functions (VLOOKUP, HLOOKUP, INDEX, MATCH)
 *   - Date functions (TODAY, NOW, DATE, YEAR, MONTH, DAY)
 *   - Info functions (ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISNA)
 *   - Other functions (VALUE, CHOOSE)
 *   - Cell references (A1, $A$1, ranges, cross-sheet)
 *   - Defined names
 *   - Error handling
 *   - Workbook.calculateFormula() and Worksheet.calculateFormula() integration
 */

import { Workbook } from '../src';
import { FormulaEvaluator } from '../src/features/FormulaEvaluator';

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Create a workbook with a single sheet and return { wb, ws, evaluator }. */
function setup(sheetName = 'Sheet1') {
  const wb = new Workbook();
  const ws = wb.getWorksheet(0);
  // Rename if needed
  if (sheetName !== 'Sheet1') {
    (ws as any)._name = sheetName;
  }
  const evaluator = new FormulaEvaluator(wb);
  return { wb, ws, evaluator };
}

// ─── 1. Arithmetic & Operator Precedence ───────────────────────────────────

describe('FormulaEvaluator – Arithmetic', () => {
  test('simple addition', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1+2', ws)).toBe(3);
  });

  test('simple subtraction', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('10-3', ws)).toBe(7);
  });

  test('multiplication', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('4*5', ws)).toBe(20);
  });

  test('division', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('20/4', ws)).toBe(5);
  });

  test('exponentiation', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('2^3', ws)).toBe(8);
  });

  test('operator precedence (* before +)', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('2+3*4', ws)).toBe(14);
  });

  test('operator precedence (^ before *)', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('2*3^2', ws)).toBe(18);
  });

  test('parentheses override precedence', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('(2+3)*4', ws)).toBe(20);
  });

  test('nested parentheses', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('((1+2)*(3+4))', ws)).toBe(21);
  });

  test('unary minus', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('-5+3', ws)).toBe(-2);
  });

  test('unary plus', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('+5', ws)).toBe(5);
  });

  test('division by zero yields #DIV/0!', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1/0', ws)).toBe('#DIV/0!');
  });

  test('decimal numbers', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1.5+2.5', ws)).toBe(4);
  });

  test('percentage (unary %)', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('50%', ws)).toBeCloseTo(0.5);
  });
});

// ─── 2. Comparison Operators ───────────────────────────────────────────────

describe('FormulaEvaluator – Comparisons', () => {
  test('equal', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1=1', ws)).toBe(true);
    expect(evaluator.evaluate('1=2', ws)).toBe(false);
  });

  test('not equal <>', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1<>2', ws)).toBe(true);
    expect(evaluator.evaluate('1<>1', ws)).toBe(false);
  });

  test('less than', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('1<2', ws)).toBe(true);
    expect(evaluator.evaluate('2<1', ws)).toBe(false);
  });

  test('greater than', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('3>2', ws)).toBe(true);
  });

  test('less than or equal', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('2<=2', ws)).toBe(true);
    expect(evaluator.evaluate('3<=2', ws)).toBe(false);
  });

  test('greater than or equal', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('3>=2', ws)).toBe(true);
    expect(evaluator.evaluate('1>=2', ws)).toBe(false);
  });
});

// ─── 3. String Concatenation Operator ──────────────────────────────────────

describe('FormulaEvaluator – Concatenation', () => {
  test('ampersand concatenation', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('"Hello"&" "&"World"', ws)).toBe('Hello World');
  });

  test('number to string concatenation', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('"Value: "&123', ws)).toBe('Value: 123');
  });
});

// ─── 4. Literals ───────────────────────────────────────────────────────────

describe('FormulaEvaluator – Literals', () => {
  test('string literal', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('"hello"', ws)).toBe('hello');
  });

  test('numeric literal', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('42', ws)).toBe(42);
  });

  test('boolean TRUE', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('TRUE', ws)).toBe(true);
  });

  test('boolean FALSE', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('FALSE', ws)).toBe(false);
  });

  test('error literal #N/A', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('#N/A', ws)).toBe('#N/A');
  });
});

// ─── 5. Cell References ────────────────────────────────────────────────────

describe('FormulaEvaluator – Cell References', () => {
  test('simple cell reference', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    expect(evaluator.evaluate('A1', ws)).toBe(10);
  });

  test('absolute cell reference $A$1', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(42);
    expect(evaluator.evaluate('$A$1', ws)).toBe(42);
  });

  test('arithmetic with cell references', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('B1').putValue(20);
    expect(evaluator.evaluate('A1+B1', ws)).toBe(30);
  });

  test('empty cell reference returns 0 for arithmetic', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(5);
    // B1 is empty, should be treated as 0 in arithmetic
    expect(evaluator.evaluate('A1+B1', ws)).toBe(5);
  });

  test('cell with formula chain', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(5);
    ws.cells.get('B1').value = null;
    ws.cells.get('B1').formula = 'A1*2';
    // First evaluate B1's formula
    const b1Result = evaluator.evaluate('A1*2', ws);
    ws.cells.get('B1').value = b1Result;
    expect(evaluator.evaluate('B1+1', ws)).toBe(11);
  });
});

// ─── 6. String Functions ───────────────────────────────────────────────────

describe('FormulaEvaluator – String Functions', () => {
  test('CONCATENATE', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('CONCATENATE("Hello"," ","World")', ws)).toBe('Hello World');
  });

  test('CONCAT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('CONCAT("A","B","C")', ws)).toBe('ABC');
  });

  test('LEN', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('LEN("hello")', ws)).toBe(5);
  });

  test('LEN with cell reference', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue('test');
    expect(evaluator.evaluate('LEN(A1)', ws)).toBe(4);
  });

  test('TRIM', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('TRIM("  hello  ")', ws)).toBe('hello');
  });

  test('UPPER', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('UPPER("hello")', ws)).toBe('HELLO');
  });

  test('LOWER', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('LOWER("HELLO")', ws)).toBe('hello');
  });

  test('LEFT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('LEFT("Hello",3)', ws)).toBe('Hel');
  });

  test('LEFT default 1 char', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('LEFT("Hello")', ws)).toBe('H');
  });

  test('RIGHT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('RIGHT("Hello",3)', ws)).toBe('llo');
  });

  test('MID', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('MID("Hello World",7,5)', ws)).toBe('World');
  });

  test('SUBSTITUTE', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('SUBSTITUTE("Hello World","World","Earth")', ws)).toBe('Hello Earth');
  });

  test('REPT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('REPT("ab",3)', ws)).toBe('ababab');
  });

  test('TEXT with number format', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('TEXT(1234.5,"0.00")', ws);
    expect(result).toBe('1234.50');
  });

  test('TEXT with integer format', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('TEXT(42,"0")', ws)).toBe('42');
  });
});

// ─── 7. Logic Functions ────────────────────────────────────────────────────

describe('FormulaEvaluator – Logic Functions', () => {
  test('IF true branch', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('IF(TRUE,"yes","no")', ws)).toBe('yes');
  });

  test('IF false branch', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('IF(FALSE,"yes","no")', ws)).toBe('no');
  });

  test('IF with comparison', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    expect(evaluator.evaluate('IF(A1>5,"big","small")', ws)).toBe('big');
  });

  test('IF with only 2 args (false returns FALSE)', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('IF(FALSE,"yes")', ws)).toBe(false);
  });

  test('AND all true', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('AND(TRUE,TRUE,TRUE)', ws)).toBe(true);
  });

  test('AND one false', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('AND(TRUE,FALSE,TRUE)', ws)).toBe(false);
  });

  test('OR all false', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('OR(FALSE,FALSE)', ws)).toBe(false);
  });

  test('OR one true', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('OR(FALSE,TRUE)', ws)).toBe(true);
  });

  test('NOT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('NOT(TRUE)', ws)).toBe(false);
    expect(evaluator.evaluate('NOT(FALSE)', ws)).toBe(true);
  });

  test('nested IF with AND', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('B1').putValue(20);
    expect(evaluator.evaluate('IF(AND(A1>5,B1>15),"both","not both")', ws)).toBe('both');
  });
});

// ─── 8. Math/Stat Functions ────────────────────────────────────────────────

describe('FormulaEvaluator – Math Functions', () => {
  test('SUM of values', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('SUM(1,2,3)', ws)).toBe(6);
  });

  test('SUM with cell range', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue(20);
    ws.cells.get('A3').putValue(30);
    expect(evaluator.evaluate('SUM(A1:A3)', ws)).toBe(60);
  });

  test('AVERAGE', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue(20);
    ws.cells.get('A3').putValue(30);
    expect(evaluator.evaluate('AVERAGE(A1:A3)', ws)).toBe(20);
  });

  test('MIN', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('MIN(5,3,8,1,9)', ws)).toBe(1);
  });

  test('MAX', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('MAX(5,3,8,1,9)', ws)).toBe(9);
  });

  test('COUNT (counts numbers only)', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue('text');
    ws.cells.get('A3').putValue(30);
    expect(evaluator.evaluate('COUNT(A1:A3)', ws)).toBe(2);
  });

  test('COUNTA (counts non-empty)', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue('text');
    ws.cells.get('A3').putValue(30);
    expect(evaluator.evaluate('COUNTA(A1:A3)', ws)).toBe(3);
  });

  test('ABS', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ABS(-5)', ws)).toBe(5);
    expect(evaluator.evaluate('ABS(5)', ws)).toBe(5);
  });

  test('ROUND', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ROUND(3.14159,2)', ws)).toBeCloseTo(3.14);
  });

  test('ROUNDUP', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ROUNDUP(3.141,2)', ws)).toBeCloseTo(3.15);
  });

  test('ROUNDDOWN', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ROUNDDOWN(3.149,2)', ws)).toBeCloseTo(3.14);
  });

  test('INT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('INT(3.7)', ws)).toBe(3);
    expect(evaluator.evaluate('INT(-3.7)', ws)).toBe(-4);
  });

  test('MOD', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('MOD(10,3)', ws)).toBe(1);
  });

  test('POWER', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('POWER(2,10)', ws)).toBe(1024);
  });

  test('SQRT', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('SQRT(16)', ws)).toBe(4);
  });

  test('SQRT of negative yields #NUM!', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('SQRT(-1)', ws)).toBe('#NUM!');
  });

  test('CEILING', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('CEILING(2.1,1)', ws)).toBe(3);
    expect(evaluator.evaluate('CEILING(2.5,0.5)', ws)).toBe(2.5);
  });

  test('FLOOR', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('FLOOR(2.9,1)', ws)).toBe(2);
  });
});

// ─── 9. Lookup Functions ───────────────────────────────────────────────────

describe('FormulaEvaluator – Lookup Functions', () => {
  function setupLookupData() {
    const { wb, ws, evaluator } = setup();
    // A1:C3 lookup table
    // A1: 1    B1: "Apple"   C1: 100
    // A2: 2    B2: "Banana"  C2: 200
    // A3: 3    B3: "Cherry"  C3: 300
    ws.cells.get('A1').putValue(1);
    ws.cells.get('B1').putValue('Apple');
    ws.cells.get('C1').putValue(100);
    ws.cells.get('A2').putValue(2);
    ws.cells.get('B2').putValue('Banana');
    ws.cells.get('C2').putValue(200);
    ws.cells.get('A3').putValue(3);
    ws.cells.get('B3').putValue('Cherry');
    ws.cells.get('C3').putValue(300);
    return { wb, ws, evaluator };
  }

  test('VLOOKUP exact match', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('VLOOKUP(2,A1:C3,2,FALSE)', ws)).toBe('Banana');
  });

  test('VLOOKUP returns value from column index', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('VLOOKUP(3,A1:C3,3,FALSE)', ws)).toBe(300);
  });

  test('VLOOKUP not found yields #N/A', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('VLOOKUP(99,A1:C3,2,FALSE)', ws)).toBe('#N/A');
  });

  test('HLOOKUP exact match', () => {
    const { evaluator, ws } = setupLookupData();
    // HLOOKUP(1, A1:C3, 2, FALSE) searches first row for 1, finds it at col A, returns row 2 col A = 2
    expect(evaluator.evaluate('HLOOKUP(1,A1:C3,2,FALSE)', ws)).toBe(2);
    // Search for "Apple" in first row, return row 3 (Cherry row) => 100
    // Actually: search for "Banana" in first row won't find it; instead test a column header:
    // Let's use a proper HLOOKUP: search row 1 for "Apple", return row 2
    expect(evaluator.evaluate('HLOOKUP("Apple",B1:C3,2,FALSE)', ws)).toBe('Banana');
  });

  test('INDEX', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('INDEX(A1:C3,2,2)', ws)).toBe('Banana');
  });

  test('MATCH exact', () => {
    const { evaluator, ws } = setupLookupData();
    // Look for 2 in A1:A3 -> position 2
    expect(evaluator.evaluate('MATCH(2,A1:A3,0)', ws)).toBe(2);
  });

  test('MATCH not found yields #N/A', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('MATCH(99,A1:A3,0)', ws)).toBe('#N/A');
  });

  test('INDEX + MATCH combined', () => {
    const { evaluator, ws } = setupLookupData();
    expect(evaluator.evaluate('INDEX(B1:B3,MATCH(3,A1:A3,0),1)', ws)).toBe('Cherry');
  });
});

// ─── 10. Date Functions ────────────────────────────────────────────────────

describe('FormulaEvaluator – Date Functions', () => {
  test('DATE returns a number', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('DATE(2024,1,15)', ws);
    expect(typeof result).toBe('number');
    expect(result).toBeGreaterThan(0);
  });

  test('YEAR extracts year', () => {
    const { evaluator, ws } = setup();
    const dateSerial = evaluator.evaluate('DATE(2024,6,15)', ws) as number;
    ws.cells.get('A1').putValue(dateSerial);
    expect(evaluator.evaluate('YEAR(A1)', ws)).toBe(2024);
  });

  test('MONTH extracts month', () => {
    const { evaluator, ws } = setup();
    const dateSerial = evaluator.evaluate('DATE(2024,6,15)', ws) as number;
    ws.cells.get('A1').putValue(dateSerial);
    expect(evaluator.evaluate('MONTH(A1)', ws)).toBe(6);
  });

  test('DAY extracts day', () => {
    const { evaluator, ws } = setup();
    const dateSerial = evaluator.evaluate('DATE(2024,6,15)', ws) as number;
    ws.cells.get('A1').putValue(dateSerial);
    expect(evaluator.evaluate('DAY(A1)', ws)).toBe(15);
  });

  test('TODAY returns a number', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('TODAY()', ws);
    expect(typeof result).toBe('number');
    expect(result).toBeGreaterThan(40000); // After year 2009
  });

  test('NOW returns a number', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('NOW()', ws);
    expect(typeof result).toBe('number');
    expect(result).toBeGreaterThan(40000);
  });
});

// ─── 11. Info Functions ────────────────────────────────────────────────────

describe('FormulaEvaluator – Info Functions', () => {
  test('ISNUMBER', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(42);
    ws.cells.get('A2').putValue('text');
    expect(evaluator.evaluate('ISNUMBER(A1)', ws)).toBe(true);
    expect(evaluator.evaluate('ISNUMBER(A2)', ws)).toBe(false);
  });

  test('ISTEXT', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue('hello');
    ws.cells.get('A2').putValue(42);
    expect(evaluator.evaluate('ISTEXT(A1)', ws)).toBe(true);
    expect(evaluator.evaluate('ISTEXT(A2)', ws)).toBe(false);
  });

  test('ISBLANK', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(42);
    expect(evaluator.evaluate('ISBLANK(A1)', ws)).toBe(false);
    // A2 has never been written to; its value is null (blank)
    ws.cells.get('A2'); // ensure cell exists but value is null
    expect(evaluator.evaluate('ISBLANK(A2)', ws)).toBe(true);
  });

  test('ISERROR', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ISERROR(1/0)', ws)).toBe(true);
    expect(evaluator.evaluate('ISERROR(1+1)', ws)).toBe(false);
  });

  test('ISNA', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('ISNA(#N/A)', ws)).toBe(true);
    expect(evaluator.evaluate('ISNA(1)', ws)).toBe(false);
  });
});

// ─── 12. Other Functions ───────────────────────────────────────────────────

describe('FormulaEvaluator – Other Functions', () => {
  test('VALUE converts string to number', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('VALUE("123")', ws)).toBe(123);
    expect(evaluator.evaluate('VALUE("3.14")', ws)).toBe(3.14);
  });

  test('VALUE with non-numeric yields #VALUE!', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('VALUE("abc")', ws)).toBe('#VALUE!');
  });

  test('CHOOSE', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('CHOOSE(2,"a","b","c")', ws)).toBe('b');
  });

  test('CHOOSE out of bounds yields #VALUE!', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('CHOOSE(5,"a","b","c")', ws)).toBe('#VALUE!');
  });
});

// ─── 13. Defined Names ────────────────────────────────────────────────────

describe('FormulaEvaluator – Defined Names', () => {
  test('resolves a defined name to a cell value', () => {
    const { wb, ws, evaluator } = setup();
    ws.cells.get('A1').putValue(99);
    wb.properties.definedNames.add('MyValue', 'Sheet1!A1');
    expect(evaluator.evaluate('MyValue', ws)).toBe(99);
  });

  test('defined name in a formula', () => {
    const { wb, ws, evaluator } = setup();
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue(20);
    wb.properties.definedNames.add('Price', 'Sheet1!A1');
    wb.properties.definedNames.add('Qty', 'Sheet1!A2');
    expect(evaluator.evaluate('Price*Qty', ws)).toBe(200);
  });
});

// ─── 14. Cross-Sheet References ────────────────────────────────────────────

describe('FormulaEvaluator – Cross-Sheet References', () => {
  test('Sheet2!A1 reference', () => {
    const wb = new Workbook();
    const ws1 = wb.getWorksheet(0);
    const ws2 = wb.addWorksheet('Sheet2');
    ws2.cells.get('A1').putValue(42);
    const evaluator = new FormulaEvaluator(wb);
    expect(evaluator.evaluate('Sheet2!A1', ws1)).toBe(42);
  });

  test('SUM across sheets', () => {
    const wb = new Workbook();
    const ws1 = wb.getWorksheet(0);
    const ws2 = wb.addWorksheet('Data');
    ws2.cells.get('A1').putValue(10);
    ws2.cells.get('A2').putValue(20);
    ws2.cells.get('A3').putValue(30);
    const evaluator = new FormulaEvaluator(wb);
    expect(evaluator.evaluate('SUM(Data!A1:A3)', ws1)).toBe(60);
  });
});

// ─── 15. Error Handling ────────────────────────────────────────────────────

describe('FormulaEvaluator – Error Handling', () => {
  test('unknown function returns null', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('UNKNOWNFUNC(1)', ws);
    // Unknown functions should return null or an error
    expect(result === null || (typeof result === 'string' && result.startsWith('#'))).toBe(true);
  });

  test('malformed formula returns null', () => {
    const { evaluator, ws } = setup();
    const result = evaluator.evaluate('((((', ws);
    expect(result).toBeNull();
  });

  test('empty formula returns null', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('', ws)).toBeNull();
  });
});

// ─── 16. evaluateAll – Workbook & Worksheet integration ────────────────────

describe('FormulaEvaluator – evaluateAll integration', () => {
  test('evaluateAll populates cell values from formulas', () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.get('A1').putValue(10);
    ws.cells.get('A2').putValue(20);

    const cellB1 = ws.cells.get('B1');
    cellB1.formula = 'A1+A2';

    const evaluator = new FormulaEvaluator(wb);
    evaluator.evaluateAll();

    expect(cellB1.value).toBe(30);
  });

  test('evaluateAll with single sheet parameter', () => {
    const wb = new Workbook();
    const ws1 = wb.getWorksheet(0);
    const ws2 = wb.addWorksheet('Sheet2');

    ws1.cells.get('A1').putValue(5);
    const cellB1 = ws1.cells.get('B1');
    cellB1.formula = 'A1*3';

    ws2.cells.get('A1').putValue(7);
    const cellB1s2 = ws2.cells.get('B1');
    cellB1s2.formula = 'A1*4';

    const evaluator = new FormulaEvaluator(wb);
    // Only evaluate Sheet1
    evaluator.evaluateAll(ws1);

    expect(cellB1.value).toBe(15);
    // Sheet2 should NOT have been evaluated
    expect(cellB1s2.value).toBeNull();
  });

  test('Workbook.calculateFormula() works end-to-end', () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.get('A1').putValue(100);
    ws.cells.get('A2').putValue(200);

    const cellA3 = ws.cells.get('A3');
    cellA3.formula = 'SUM(A1:A2)';

    wb.calculateFormula();

    expect(cellA3.value).toBe(300);
  });

  test('Worksheet.calculateFormula() works end-to-end', () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.get('A1').putValue(5);
    ws.cells.get('A2').putValue(10);

    const cellA3 = ws.cells.get('A3');
    cellA3.formula = 'A1+A2';

    ws.calculateFormula();

    expect(cellA3.value).toBe(15);
  });

  test('snake_case aliases work', () => {
    const wb = new Workbook();
    const ws = wb.getWorksheet(0);
    ws.cells.get('A1').putValue(3);
    ws.cells.get('A2').putValue(4);
    const cellA3 = ws.cells.get('A3');
    cellA3.formula = 'A1+A2';

    wb.calculate_formula();
    expect(cellA3.value).toBe(7);
  });
});

// ─── 17. Complex Formulas ──────────────────────────────────────────────────

describe('FormulaEvaluator – Complex Formulas', () => {
  test('nested function calls', () => {
    const { evaluator, ws } = setup();
    expect(evaluator.evaluate('UPPER(TRIM("  hello  "))', ws)).toBe('HELLO');
  });

  test('IF with arithmetic comparison', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(85);
    expect(evaluator.evaluate('IF(A1>=90,"A",IF(A1>=80,"B","C"))', ws)).toBe('B');
  });

  test('SUM mixed with individual values', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue(1);
    ws.cells.get('A2').putValue(2);
    ws.cells.get('A3').putValue(3);
    expect(evaluator.evaluate('SUM(A1:A3,10)', ws)).toBe(16);
  });

  test('CONCATENATE with cell values and text', () => {
    const { evaluator, ws } = setup();
    ws.cells.get('A1').putValue('John');
    ws.cells.get('B1').putValue('Doe');
    expect(evaluator.evaluate('CONCATENATE(A1," ",B1)', ws)).toBe('John Doe');
  });

  test('complex arithmetic with multiple operations', () => {
    const { evaluator, ws } = setup();
    // (10 + 5) * 2 - 3 / (1 + 2) = 15 * 2 - 1 = 30 - 1 = 29
    expect(evaluator.evaluate('(10+5)*2-3/(1+2)', ws)).toBe(29);
  });
});
