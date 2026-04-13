/**
 * Aspose.Cells for Node.js – Formula Evaluator Module
 *
 * Provides formula evaluation functionality for cells that have formulas.
 * Ported from Python `formula_evaluator.py` with significant enhancements.
 *
 * Supports:
 * - String, numeric, and boolean literals
 * - Cell references (A1 style, absolute $A$1, sheet-qualified Sheet1!A1)
 * - Range references (A1:B5)
 * - Arithmetic operators: + - * / ^
 * - Comparison operators: = <> < > <= >=
 * - String concatenation: &
 * - Built-in functions (SUM, IF, VLOOKUP, etc.)
 * - Defined names
 * - Nested formulas with recursive evaluation
 * - Error handling (#VALUE!, #REF!, #NAME?, #DIV/0!, #N/A)
 *
 * @module features/FormulaEvaluator
 */

import { Cells } from '../core/Cells';
import { CellValueHandler } from '../io/CellValueHandler';

// Use interfaces to avoid circular imports
interface WorksheetLike {
  name: string;
  cells: CellsLike;
}

interface CellLike {
  value: any;
  formula: string | null;
  hasFormula(): boolean;
}

interface CellsLike {
  get(key: string): CellLike;
  _cells: Map<string, CellLike>;
}

interface DefinedNameLike {
  name: string;
  refersTo: string;
}

interface DefinedNameCollectionLike {
  toArray(): DefinedNameLike[];
}

interface WorkbookPropertiesLike {
  definedNames: DefinedNameCollectionLike;
}

interface WorkbookLike {
  worksheets: WorksheetLike[];
  properties: WorkbookPropertiesLike;
}

// ─── Token Types ──────────────────────────────────────────────────────────

enum TokenType {
  NUMBER = 'NUMBER',
  STRING = 'STRING',
  BOOLEAN = 'BOOLEAN',
  CELL_REF = 'CELL_REF',
  RANGE_REF = 'RANGE_REF',
  SHEET_REF = 'SHEET_REF',
  FUNCTION = 'FUNCTION',
  OPERATOR = 'OPERATOR',
  COMPARISON = 'COMPARISON',
  LPAREN = 'LPAREN',
  RPAREN = 'RPAREN',
  COMMA = 'COMMA',
  CONCAT = 'CONCAT',
  UNARY_MINUS = 'UNARY_MINUS',
  UNARY_PLUS = 'UNARY_PLUS',
  ERROR = 'ERROR',
  NAME = 'NAME',
}

interface Token {
  type: TokenType;
  value: string;
}

// ─── Excel Errors ─────────────────────────────────────────────────────────

const ERRORS = {
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NUM: '#NUM!',
  NULL: '#NULL!',
} as const;

function isError(val: any): boolean {
  if (typeof val !== 'string') return false;
  return val === ERRORS.VALUE || val === ERRORS.REF || val === ERRORS.NAME ||
    val === ERRORS.DIV0 || val === ERRORS.NA || val === ERRORS.NUM || val === ERRORS.NULL;
}

// ─── Tokenizer ────────────────────────────────────────────────────────────

/**
 * Tokenizes a formula string into an array of tokens.
 */
function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const len = formula.length;

  while (i < len) {
    const ch = formula[i];

    // Skip whitespace
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }

    // String literal
    if (ch === '"') {
      let str = '';
      i++; // skip opening quote
      while (i < len) {
        if (formula[i] === '"') {
          if (i + 1 < len && formula[i + 1] === '"') {
            // Escaped quote
            str += '"';
            i += 2;
          } else {
            break;
          }
        } else {
          str += formula[i];
          i++;
        }
      }
      i++; // skip closing quote
      tokens.push({ type: TokenType.STRING, value: str });
      continue;
    }

    // Error values like #DIV/0!, #VALUE!, #REF!, #NAME?, #N/A, #NUM!, #NULL!
    if (ch === '#') {
      // Try to match known error patterns first
      const remaining = formula.substring(i);
      const errMatch = remaining.match(/^(#DIV\/0!|#VALUE!|#REF!|#NAME\?|#N\/A|#NUM!|#NULL!)/i);
      if (errMatch) {
        tokens.push({ type: TokenType.ERROR, value: errMatch[1].toUpperCase() });
        i += errMatch[1].length;
        continue;
      }
      // Fallback: consume until delimiter
      let errStr = '#';
      i++;
      while (i < len && formula[i] !== ' ' && formula[i] !== ',' && formula[i] !== ')' && formula[i] !== '+' && formula[i] !== '-' && formula[i] !== '*' && formula[i] !== '&' && formula[i] !== '^' && formula[i] !== '=' && formula[i] !== '<' && formula[i] !== '>') {
        errStr += formula[i];
        i++;
      }
      tokens.push({ type: TokenType.ERROR, value: errStr });
      continue;
    }

    // Number
    if ((ch >= '0' && ch <= '9') || (ch === '.' && i + 1 < len && formula[i + 1] >= '0' && formula[i + 1] <= '9')) {
      let num = '';
      while (i < len && ((formula[i] >= '0' && formula[i] <= '9') || formula[i] === '.')) {
        num += formula[i];
        i++;
      }
      // Handle scientific notation
      if (i < len && (formula[i] === 'e' || formula[i] === 'E')) {
        num += formula[i];
        i++;
        if (i < len && (formula[i] === '+' || formula[i] === '-')) {
          num += formula[i];
          i++;
        }
        while (i < len && formula[i] >= '0' && formula[i] <= '9') {
          num += formula[i];
          i++;
        }
      }
      // Check for percent
      if (i < len && formula[i] === '%') {
        const val = parseFloat(num) / 100;
        tokens.push({ type: TokenType.NUMBER, value: String(val) });
        i++;
      } else {
        tokens.push({ type: TokenType.NUMBER, value: num });
      }
      continue;
    }

    // Operators and special chars
    if (ch === '(') {
      tokens.push({ type: TokenType.LPAREN, value: '(' });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: TokenType.RPAREN, value: ')' });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ type: TokenType.COMMA, value: ',' });
      i++;
      continue;
    }
    if (ch === '&') {
      tokens.push({ type: TokenType.CONCAT, value: '&' });
      i++;
      continue;
    }
    if (ch === '+') {
      // Check if unary
      if (tokens.length === 0 || tokens[tokens.length - 1].type === TokenType.LPAREN ||
          tokens[tokens.length - 1].type === TokenType.COMMA ||
          tokens[tokens.length - 1].type === TokenType.OPERATOR ||
          tokens[tokens.length - 1].type === TokenType.COMPARISON ||
          tokens[tokens.length - 1].type === TokenType.CONCAT) {
        tokens.push({ type: TokenType.UNARY_PLUS, value: '+' });
      } else {
        tokens.push({ type: TokenType.OPERATOR, value: '+' });
      }
      i++;
      continue;
    }
    if (ch === '-') {
      // Check if unary
      if (tokens.length === 0 || tokens[tokens.length - 1].type === TokenType.LPAREN ||
          tokens[tokens.length - 1].type === TokenType.COMMA ||
          tokens[tokens.length - 1].type === TokenType.OPERATOR ||
          tokens[tokens.length - 1].type === TokenType.COMPARISON ||
          tokens[tokens.length - 1].type === TokenType.CONCAT) {
        tokens.push({ type: TokenType.UNARY_MINUS, value: '-' });
      } else {
        tokens.push({ type: TokenType.OPERATOR, value: '-' });
      }
      i++;
      continue;
    }
    if (ch === '*' || ch === '/' || ch === '^') {
      tokens.push({ type: TokenType.OPERATOR, value: ch });
      i++;
      continue;
    }

    // Comparison operators
    if (ch === '<') {
      if (i + 1 < len && formula[i + 1] === '>') {
        tokens.push({ type: TokenType.COMPARISON, value: '<>' });
        i += 2;
      } else if (i + 1 < len && formula[i + 1] === '=') {
        tokens.push({ type: TokenType.COMPARISON, value: '<=' });
        i += 2;
      } else {
        tokens.push({ type: TokenType.COMPARISON, value: '<' });
        i++;
      }
      continue;
    }
    if (ch === '>') {
      if (i + 1 < len && formula[i + 1] === '=') {
        tokens.push({ type: TokenType.COMPARISON, value: '>=' });
        i += 2;
      } else {
        tokens.push({ type: TokenType.COMPARISON, value: '>' });
        i++;
      }
      continue;
    }
    if (ch === '=') {
      tokens.push({ type: TokenType.COMPARISON, value: '=' });
      i++;
      continue;
    }

    // Identifiers: cell refs, function names, booleans, named ranges
    if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '_' || ch === '$' || ch === '\'') {
      // Handle quoted sheet names: 'Sheet Name'!A1
      if (ch === '\'') {
        let sheetName = '';
        i++; // skip opening quote
        while (i < len && formula[i] !== '\'') {
          sheetName += formula[i];
          i++;
        }
        i++; // skip closing quote
        if (i < len && formula[i] === '!') {
          i++; // skip !
          let cellPart = '';
          while (i < len && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z') || (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '$' || formula[i] === ':')) {
            cellPart += formula[i];
            i++;
          }
          if (cellPart.includes(':')) {
            tokens.push({ type: TokenType.RANGE_REF, value: `'${sheetName}'!${cellPart}` });
          } else {
            tokens.push({ type: TokenType.SHEET_REF, value: `'${sheetName}'!${cellPart}` });
          }
        }
        continue;
      }

      let ident = '';
      while (i < len && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z') ||
             (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '_' || formula[i] === '$' || formula[i] === '.')) {
        ident += formula[i];
        i++;
      }

      // Check for sheet reference: Sheet1!A1 or Sheet1!A1:B5
      if (i < len && formula[i] === '!') {
        const sheetName = ident;
        i++; // skip !
        let cellPart = '';
        while (i < len && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z') || (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '$' || formula[i] === ':')) {
          cellPart += formula[i];
          i++;
        }
        if (cellPart.includes(':')) {
          tokens.push({ type: TokenType.RANGE_REF, value: `${sheetName}!${cellPart}` });
        } else {
          tokens.push({ type: TokenType.SHEET_REF, value: `${sheetName}!${cellPart}` });
        }
        continue;
      }

      // Check for range reference: A1:B5
      if (i < len && formula[i] === ':') {
        const startRef = ident;
        i++; // skip :
        let endRef = '';
        while (i < len && ((formula[i] >= 'A' && formula[i] <= 'Z') || (formula[i] >= 'a' && formula[i] <= 'z') || (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '$')) {
          endRef += formula[i];
          i++;
        }
        tokens.push({ type: TokenType.RANGE_REF, value: `${startRef}:${endRef}` });
        continue;
      }

      // Check for function call (followed by parenthesis)
      if (i < len && formula[i] === '(') {
        tokens.push({ type: TokenType.FUNCTION, value: ident.toUpperCase() });
        continue;
      }

      // Check for boolean
      const upper = ident.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ type: TokenType.BOOLEAN, value: upper });
        continue;
      }

      // Check for cell reference pattern (letters followed by digits, possibly with $)
      const cleanIdent = ident.replace(/\$/g, '');
      if (/^[A-Za-z]+[0-9]+$/.test(cleanIdent)) {
        tokens.push({ type: TokenType.CELL_REF, value: ident });
        continue;
      }

      // Must be a name (defined name or unknown identifier)
      tokens.push({ type: TokenType.NAME, value: ident });
      continue;
    }

    // Unknown character - skip
    i++;
  }

  return tokens;
}

// ─── Operator Precedence ──────────────────────────────────────────────────

function precedence(op: string): number {
  switch (op) {
    case '^': return 5;
    case '*': case '/': return 4;
    case '+': case '-': return 3;
    case '&': return 2;
    case '=': case '<>': case '<': case '>': case '<=': case '>=': return 1;
    default: return 0;
  }
}

function isLeftAssociative(op: string): boolean {
  return op !== '^'; // ^ is right-associative
}

// ─── FormulaEvaluator ─────────────────────────────────────────────────────

/**
 * Evaluates Excel formulas within a workbook context.
 *
 * Supports arithmetic, comparisons, string concatenation, cell references,
 * range references, defined names, and a comprehensive set of built-in functions.
 *
 * @example
 * ```ts
 * const evaluator = new FormulaEvaluator(workbook);
 * const result = evaluator.evaluate('=SUM(A1:A5)', worksheet);
 * ```
 */
export class FormulaEvaluator {
  private _workbook: WorkbookLike;
  private _definedNamesCache: Map<string, string> | null = null;
  private _evaluatingCells: Set<string> = new Set();
  private _maxDepth = 100;

  constructor(workbook: WorkbookLike) {
    this._workbook = workbook;
  }

  // ── Main API ────────────────────────────────────────────────────────────

  /**
   * Evaluate a formula and return the result.
   *
   * @param formula - The formula string (with or without leading '=').
   * @param worksheet - The worksheet context for cell references.
   * @returns The evaluated value, or an error string on failure.
   */
  evaluate(formula: string, worksheet?: WorksheetLike | null): any {
    if (formula == null) return null;

    // Remove leading '='
    let expr = formula;
    if (expr.startsWith('=')) {
      expr = expr.substring(1);
    }

    try {
      this._evaluatingCells.clear();
      const tokens = tokenize(expr);
      const result = this._evaluateTokens(tokens, worksheet ?? null);
      return result;
    } catch (e: any) {
      if (isError(e?.message)) return e.message;
      return ERRORS.VALUE;
    }
  }

  /**
   * Evaluate all formulas in all worksheets of the workbook.
   * Sets each cell's value to the evaluated result.
   */
  evaluateAll(singleSheet?: WorksheetLike): void {
    const sheets = singleSheet ? [singleSheet] : this._workbook.worksheets;
    for (const ws of sheets) {
      for (const [ref, cell] of ws.cells._cells.entries()) {
        if (cell.formula) {
          const result = this.evaluate(cell.formula, ws);
          if (result !== null && result !== undefined) {
            cell.value = result;
          }
        }
      }
    }
  }

  /** snake_case alias */
  evaluate_all(singleSheet?: WorksheetLike): void { this.evaluateAll(singleSheet); }

  // ── Token evaluation (recursive descent parser) ─────────────────────────

  private _evaluateTokens(tokens: Token[], worksheet: WorksheetLike | null): any {
    if (tokens.length === 0) return null;

    const { result } = this._parseExpression(tokens, 0, worksheet, 0);
    return result;
  }

  /**
   * Parse an expression with operator precedence (Pratt parser style).
   * Returns { result, pos } where pos is the next token index.
   */
  private _parseExpression(
    tokens: Token[],
    pos: number,
    worksheet: WorksheetLike | null,
    minPrec: number,
  ): { result: any; pos: number } {
    let { result: left, pos: nextPos } = this._parseAtom(tokens, pos, worksheet);

    while (nextPos < tokens.length) {
      const tok = tokens[nextPos];

      if (tok.type === TokenType.OPERATOR) {
        const prec = precedence(tok.value);
        if (prec < minPrec) break;
        const nextMinPrec = isLeftAssociative(tok.value) ? prec + 1 : prec;
        nextPos++;
        const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, nextMinPrec);
        nextPos = rPos;
        left = this._applyArithmetic(tok.value, left, right);
        continue;
      }

      if (tok.type === TokenType.COMPARISON) {
        const prec = precedence(tok.value);
        if (prec < minPrec) break;
        nextPos++;
        const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, prec + 1);
        nextPos = rPos;
        left = this._applyComparison(tok.value, left, right);
        continue;
      }

      if (tok.type === TokenType.CONCAT) {
        const prec = precedence('&');
        if (prec < minPrec) break;
        nextPos++;
        const { result: right, pos: rPos } = this._parseExpression(tokens, nextPos, worksheet, prec + 1);
        nextPos = rPos;
        left = String(left ?? '') + String(right ?? '');
        continue;
      }

      break;
    }

    return { result: left, pos: nextPos };
  }

  /**
   * Parse an atomic value (number, string, cell ref, function call, parenthesised expression, etc.)
   */
  private _parseAtom(
    tokens: Token[],
    pos: number,
    worksheet: WorksheetLike | null,
  ): { result: any; pos: number } {
    if (pos >= tokens.length) return { result: null, pos };

    const tok = tokens[pos];

    switch (tok.type) {
      case TokenType.NUMBER:
        return { result: parseFloat(tok.value), pos: pos + 1 };

      case TokenType.STRING:
        return { result: tok.value, pos: pos + 1 };

      case TokenType.BOOLEAN:
        return { result: tok.value === 'TRUE', pos: pos + 1 };

      case TokenType.ERROR:
        return { result: tok.value, pos: pos + 1 };

      case TokenType.UNARY_MINUS: {
        const { result, pos: nPos } = this._parseAtom(tokens, pos + 1, worksheet);
        const num = this._toNumber(result);
        if (isError(num)) return { result: num, pos: nPos };
        return { result: -(num as number), pos: nPos };
      }

      case TokenType.UNARY_PLUS: {
        const { result, pos: nPos } = this._parseAtom(tokens, pos + 1, worksheet);
        const num = this._toNumber(result);
        if (isError(num)) return { result: num, pos: nPos };
        return { result: +(num as number), pos: nPos };
      }

      case TokenType.LPAREN: {
        const { result, pos: nPos } = this._parseExpression(tokens, pos + 1, worksheet, 0);
        // Consume RPAREN
        const rPos = nPos < tokens.length && tokens[nPos].type === TokenType.RPAREN ? nPos + 1 : nPos;
        return { result, pos: rPos };
      }

      case TokenType.CELL_REF: {
        const val = this._resolveCellRef(tok.value, worksheet);
        return { result: val, pos: pos + 1 };
      }

      case TokenType.SHEET_REF: {
        const val = this._resolveSheetRef(tok.value, worksheet);
        return { result: val, pos: pos + 1 };
      }

      case TokenType.RANGE_REF: {
        // Range refs are typically only useful as function arguments
        // When used alone, return the array of values
        const vals = this._resolveRange(tok.value, worksheet);
        if (Array.isArray(vals) && vals.length === 1) return { result: vals[0], pos: pos + 1 };
        return { result: vals, pos: pos + 1 };
      }

      case TokenType.FUNCTION: {
        // Parse function arguments
        const funcName = tok.value;
        pos++; // skip function name
        // Consume LPAREN
        if (pos < tokens.length && tokens[pos].type === TokenType.LPAREN) {
          pos++;
        }

        const args: any[] = [];
        const argTokenGroups: Token[][] = [];
        let currentArgTokens: Token[] = [];
        let parenDepth = 0;

        // Collect raw tokens for each argument
        while (pos < tokens.length) {
          const t = tokens[pos];
          if (t.type === TokenType.RPAREN && parenDepth === 0) {
            if (currentArgTokens.length > 0) {
              argTokenGroups.push(currentArgTokens);
            }
            pos++;
            break;
          }
          if (t.type === TokenType.LPAREN) parenDepth++;
          if (t.type === TokenType.RPAREN) parenDepth--;
          if (t.type === TokenType.COMMA && parenDepth === 0) {
            argTokenGroups.push(currentArgTokens);
            currentArgTokens = [];
            pos++;
            continue;
          }
          currentArgTokens.push(t);
          pos++;
        }

        // Handle no-argument functions
        if (argTokenGroups.length === 0 && currentArgTokens.length === 0) {
          return { result: this._callFunction(funcName, [], worksheet), pos };
        }

        const result = this._callFunctionWithTokens(funcName, argTokenGroups, worksheet);
        return { result, pos };
      }

      case TokenType.NAME: {
        // Try to resolve as a defined name
        const names = this._getDefinedNames();
        const upperName = tok.value.toUpperCase();
        for (const [name, refersTo] of names.entries()) {
          if (name.toUpperCase() === upperName) {
            const resolved = this._resolveDefinedName(refersTo, worksheet);
            return { result: resolved, pos: pos + 1 };
          }
        }
        // Unknown name
        return { result: ERRORS.NAME, pos: pos + 1 };
      }

      default:
        return { result: null, pos: pos + 1 };
    }
  }

  // ── Defined Names ───────────────────────────────────────────────────────

  private _getDefinedNames(): Map<string, string> {
    if (this._definedNamesCache === null) {
      this._definedNamesCache = new Map();
      try {
        const names = this._workbook.properties.definedNames.toArray();
        for (const dn of names) {
          this._definedNamesCache.set(dn.name, dn.refersTo);
        }
      } catch {
        // No defined names available
      }
    }
    return this._definedNamesCache;
  }

  private _resolveDefinedName(refersTo: string, worksheet: WorksheetLike | null): any {
    // refersTo could be like "Sheet1!$A$1:$B$5" or "Sheet1!$A$1"
    const cleaned = refersTo.replace(/\$/g, '');

    if (cleaned.includes(':')) {
      // It's a range
      return this._resolveRange(cleaned, worksheet);
    }

    if (cleaned.includes('!')) {
      return this._resolveSheetRef(cleaned, worksheet);
    }

    return this._resolveCellRef(cleaned, worksheet);
  }

  // ── Cell Reference Resolution ───────────────────────────────────────────

  private _resolveCellRef(ref: string, worksheet: WorksheetLike | null): any {
    if (!worksheet) return ERRORS.REF;
    const cleanRef = ref.replace(/\$/g, '').toUpperCase();

    // Check for circular reference
    const cellKey = `${worksheet.name}!${cleanRef}`;
    if (this._evaluatingCells.has(cellKey)) {
      return 0; // Break circular ref
    }

    const cell = worksheet.cells._cells.get(cleanRef);
    if (!cell) return 0; // Empty cell returns 0 for numeric context, but we return 0

    if (cell.value !== null && cell.value !== undefined) {
      return cell.value;
    }

    // If cell has formula but no value, evaluate recursively
    if (cell.formula) {
      this._evaluatingCells.add(cellKey);
      try {
        const result = this.evaluate(cell.formula, worksheet);
        return result;
      } finally {
        this._evaluatingCells.delete(cellKey);
      }
    }

    return null;
  }

  private _resolveSheetRef(ref: string, _worksheet: WorksheetLike | null): any {
    let sheetPart: string;
    let cellPart: string;

    if (ref.includes('!')) {
      [sheetPart, cellPart] = ref.split('!', 2);
    } else {
      return ERRORS.REF;
    }

    // Remove quotes from sheet name
    sheetPart = sheetPart.replace(/^'|'$/g, '');

    const targetSheet = this._workbook.worksheets.find(ws => ws.name === sheetPart);
    if (!targetSheet) return ERRORS.REF;

    const cleanCell = cellPart.replace(/\$/g, '');

    if (cleanCell.includes(':')) {
      return this._resolveRange(`${sheetPart}!${cleanCell}`, _worksheet);
    }

    return this._resolveCellRef(cleanCell, targetSheet);
  }

  /**
   * Resolves a range reference into an array of cell values.
   */
  private _resolveRange(rangeRef: string, worksheet: WorksheetLike | null): any[] {
    let sheetName: string | null = null;
    let rangePart: string;

    if (rangeRef.includes('!')) {
      const parts = rangeRef.split('!', 2);
      sheetName = parts[0].replace(/^'|'$/g, '');
      rangePart = parts[1];
    } else {
      rangePart = rangeRef;
    }

    const targetSheet = sheetName
      ? this._workbook.worksheets.find(ws => ws.name === sheetName)
      : worksheet;

    if (!targetSheet) return [];

    // Clean and parse range
    const cleanRange = rangePart.replace(/\$/g, '');
    const [startRef, endRef] = cleanRange.split(':');

    try {
      const [startRow, startCol] = Cells.coordinateFromString(startRef.toUpperCase());
      const [endRow, endCol] = Cells.coordinateFromString(endRef.toUpperCase());

      const values: any[] = [];
      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);
      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);

      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const cellRef = Cells.coordinateToString(r, c);
          const val = this._resolveCellRef(cellRef, targetSheet);
          values.push(val);
        }
      }

      return values;
    } catch {
      return [];
    }
  }

  /**
   * Resolves a range reference returning an array of arrays (rows x cols) for lookup functions.
   */
  private _resolveRange2D(rangeRef: string, worksheet: WorksheetLike | null): any[][] {
    let sheetName: string | null = null;
    let rangePart: string;

    if (rangeRef.includes('!')) {
      const parts = rangeRef.split('!', 2);
      sheetName = parts[0].replace(/^'|'$/g, '');
      rangePart = parts[1];
    } else {
      rangePart = rangeRef;
    }

    const targetSheet = sheetName
      ? this._workbook.worksheets.find(ws => ws.name === sheetName)
      : worksheet;

    if (!targetSheet) return [];

    const cleanRange = rangePart.replace(/\$/g, '');
    const [startRef, endRef] = cleanRange.split(':');

    try {
      const [startRow, startCol] = Cells.coordinateFromString(startRef.toUpperCase());
      const [endRow, endCol] = Cells.coordinateFromString(endRef.toUpperCase());

      const rows: any[][] = [];
      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);
      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);

      for (let r = minRow; r <= maxRow; r++) {
        const row: any[] = [];
        for (let c = minCol; c <= maxCol; c++) {
          const cellRef = Cells.coordinateToString(r, c);
          const val = this._resolveCellRef(cellRef, targetSheet);
          row.push(val);
        }
        rows.push(row);
      }

      return rows;
    } catch {
      return [];
    }
  }

  // ── Arithmetic ──────────────────────────────────────────────────────────

  private _applyArithmetic(op: string, left: any, right: any): any {
    if (isError(left)) return left;
    if (isError(right)) return right;

    const l = this._toNumber(left);
    const r = this._toNumber(right);
    if (isError(l)) return l;
    if (isError(r)) return r;

    switch (op) {
      case '+': return (l as number) + (r as number);
      case '-': return (l as number) - (r as number);
      case '*': return (l as number) * (r as number);
      case '/':
        if ((r as number) === 0) return ERRORS.DIV0;
        return (l as number) / (r as number);
      case '^': return Math.pow(l as number, r as number);
      default: return ERRORS.VALUE;
    }
  }

  // ── Comparison ──────────────────────────────────────────────────────────

  private _applyComparison(op: string, left: any, right: any): any {
    if (isError(left)) return left;
    if (isError(right)) return right;

    // Try numeric comparison first
    const lNum = typeof left === 'number' ? left : (typeof left === 'boolean' ? (left ? 1 : 0) : NaN);
    const rNum = typeof right === 'number' ? right : (typeof right === 'boolean' ? (right ? 1 : 0) : NaN);

    if (!isNaN(lNum) && !isNaN(rNum)) {
      switch (op) {
        case '=': return lNum === rNum;
        case '<>': return lNum !== rNum;
        case '<': return lNum < rNum;
        case '>': return lNum > rNum;
        case '<=': return lNum <= rNum;
        case '>=': return lNum >= rNum;
      }
    }

    // String comparison
    const lStr = String(left ?? '').toLowerCase();
    const rStr = String(right ?? '').toLowerCase();

    switch (op) {
      case '=': return lStr === rStr;
      case '<>': return lStr !== rStr;
      case '<': return lStr < rStr;
      case '>': return lStr > rStr;
      case '<=': return lStr <= rStr;
      case '>=': return lStr >= rStr;
    }

    return false;
  }

  // ── Type Conversion ─────────────────────────────────────────────────────

  private _toNumber(val: any): number | string {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (typeof val === 'string') {
      if (isError(val)) return val;
      if (val === '') return 0;
      const num = Number(val);
      if (!isNaN(num)) return num;
      return ERRORS.VALUE;
    }
    return ERRORS.VALUE;
  }

  private _toBool(val: any): boolean {
    if (typeof val === 'boolean') return val;
    if (typeof val === 'number') return val !== 0;
    if (typeof val === 'string') {
      if (val.toUpperCase() === 'TRUE') return true;
      if (val.toUpperCase() === 'FALSE') return false;
      return val !== '';
    }
    return val != null;
  }

  // ── Flatten helper (for range arguments) ────────────────────────────────

  private _flattenArgs(args: any[]): any[] {
    const result: any[] = [];
    for (const arg of args) {
      if (Array.isArray(arg)) {
        result.push(...this._flattenArgs(arg));
      } else {
        result.push(arg);
      }
    }
    return result;
  }

  private _flattenNumeric(args: any[]): number[] {
    const flat = this._flattenArgs(args);
    const nums: number[] = [];
    for (const v of flat) {
      if (typeof v === 'number') nums.push(v);
      else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
      else if (typeof v === 'string' && v !== '' && !isError(v)) {
        const n = Number(v);
        if (!isNaN(n)) nums.push(n);
      }
    }
    return nums;
  }

  // ── Function Dispatch ───────────────────────────────────────────────────

  /** Functions that need 2D arrays for their table/range argument */
  private static readonly _2D_FUNCS = new Set(['VLOOKUP', 'HLOOKUP', 'INDEX']);

  private _callFunctionWithTokens(
    funcName: string,
    argTokenGroups: Token[][],
    worksheet: WorksheetLike | null,
  ): any {
    // For functions that need raw range tokens (SUM, AVERAGE, etc.)
    // evaluate each argument group, keeping ranges as arrays
    const needs2D = FormulaEvaluator._2D_FUNCS.has(funcName);
    const evaluatedArgs: any[] = [];
    for (let i = 0; i < argTokenGroups.length; i++) {
      const argTokens = argTokenGroups[i];
      if (argTokens.length === 1 && argTokens[0].type === TokenType.RANGE_REF) {
        // For VLOOKUP/HLOOKUP/INDEX, the table array arg (index 1 for VLOOKUP/HLOOKUP, index 0 for INDEX) needs 2D
        const isTableArg =
          (needs2D && funcName !== 'INDEX' && i === 1) ||
          (funcName === 'INDEX' && i === 0);
        if (isTableArg) {
          evaluatedArgs.push(this._resolveRange2D(argTokens[0].value, worksheet));
        } else {
          evaluatedArgs.push(this._resolveRange(argTokens[0].value, worksheet));
        }
      } else {
        // Evaluate as expression
        const { result } = this._parseExpression(argTokens, 0, worksheet, 0);
        evaluatedArgs.push(result);
      }
    }

    return this._callFunction(funcName, evaluatedArgs, worksheet);
  }

  private _callFunction(funcName: string, args: any[], worksheet: WorksheetLike | null): any {
    switch (funcName) {
      // ── String Functions ──────────────────────────────────────────────
      case 'CONCATENATE':
      case 'CONCAT':
        return this._funcConcatenate(args);
      case 'TEXT':
        return this._funcText(args);
      case 'LEN':
        return this._funcLen(args);
      case 'TRIM':
        return this._funcTrim(args);
      case 'UPPER':
        return this._funcUpper(args);
      case 'LOWER':
        return this._funcLower(args);
      case 'LEFT':
        return this._funcLeft(args);
      case 'RIGHT':
        return this._funcRight(args);
      case 'MID':
        return this._funcMid(args);
      case 'SUBSTITUTE':
        return this._funcSubstitute(args);
      case 'REPT':
        return this._funcRept(args);

      // ── Logic Functions ───────────────────────────────────────────────
      case 'IF':
        return this._funcIf(args);
      case 'AND':
        return this._funcAnd(args);
      case 'OR':
        return this._funcOr(args);
      case 'NOT':
        return this._funcNot(args);

      // ── Math Functions ────────────────────────────────────────────────
      case 'SUM':
        return this._funcSum(args);
      case 'AVERAGE':
        return this._funcAverage(args);
      case 'MIN':
        return this._funcMin(args);
      case 'MAX':
        return this._funcMax(args);
      case 'COUNT':
        return this._funcCount(args);
      case 'COUNTA':
        return this._funcCountA(args);
      case 'ABS':
        return this._funcAbs(args);
      case 'ROUND':
        return this._funcRound(args);
      case 'ROUNDUP':
        return this._funcRoundUp(args);
      case 'ROUNDDOWN':
        return this._funcRoundDown(args);
      case 'INT':
        return this._funcInt(args);
      case 'MOD':
        return this._funcMod(args);
      case 'POWER':
        return this._funcPower(args);
      case 'SQRT':
        return this._funcSqrt(args);
      case 'CEILING':
        return this._funcCeiling(args);
      case 'FLOOR':
        return this._funcFloor(args);

      // ── Lookup Functions ──────────────────────────────────────────────
      case 'VLOOKUP':
        return this._funcVlookup(args, worksheet);
      case 'HLOOKUP':
        return this._funcHlookup(args, worksheet);
      case 'INDEX':
        return this._funcIndex(args);
      case 'MATCH':
        return this._funcMatch(args);

      // ── Date Functions ────────────────────────────────────────────────
      case 'TODAY':
        return this._funcToday();
      case 'NOW':
        return this._funcNow();
      case 'DATE':
        return this._funcDate(args);
      case 'YEAR':
        return this._funcYear(args);
      case 'MONTH':
        return this._funcMonth(args);
      case 'DAY':
        return this._funcDay(args);

      // ── Info Functions ────────────────────────────────────────────────
      case 'ISNUMBER':
        return this._funcIsNumber(args);
      case 'ISTEXT':
        return this._funcIsText(args);
      case 'ISBLANK':
        return this._funcIsBlank(args);
      case 'ISERROR':
        return this._funcIsError(args);
      case 'ISNA':
        return this._funcIsNa(args);

      // ── Other Functions ───────────────────────────────────────────────
      case 'VALUE':
        return this._funcValue(args);
      case 'CHOOSE':
        return this._funcChoose(args);

      default:
        return ERRORS.NAME;
    }
  }

  // ── String Function Implementations ─────────────────────────────────────

  private _funcConcatenate(args: any[]): string {
    const flat = this._flattenArgs(args);
    return flat.map(v => v == null ? '' : String(v)).join('');
  }

  private _funcText(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const value = args[0];
    const format = String(args[1] ?? '');
    if (value == null) return '';
    // Basic TEXT format support
    if (typeof value === 'number') {
      if (format === '0') return Math.round(value).toString();
      if (format === '0.00') return value.toFixed(2);
      if (/^0\.0+$/.test(format)) {
        const decimals = format.length - 2;
        return value.toFixed(decimals);
      }
    }
    return String(value);
  }

  private _funcLen(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const val = args[0];
    if (val == null) return 0;
    return String(val).length;
  }

  private _funcTrim(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const val = args[0];
    if (val == null) return '';
    return String(val).split(/\s+/).filter(s => s).join(' ');
  }

  private _funcUpper(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const val = args[0];
    if (val == null) return '';
    return String(val).toUpperCase();
  }

  private _funcLower(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const val = args[0];
    if (val == null) return '';
    return String(val).toLowerCase();
  }

  private _funcLeft(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const text = String(args[0] ?? '');
    const numChars = args.length >= 2 ? this._toNumber(args[1]) : 1;
    if (isError(numChars)) return numChars;
    return text.substring(0, numChars as number);
  }

  private _funcRight(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const text = String(args[0] ?? '');
    const numChars = args.length >= 2 ? this._toNumber(args[1]) : 1;
    if (isError(numChars)) return numChars;
    return text.substring(text.length - (numChars as number));
  }

  private _funcMid(args: any[]): any {
    if (args.length < 3) return ERRORS.VALUE;
    const text = String(args[0] ?? '');
    const startNum = this._toNumber(args[1]);
    const numChars = this._toNumber(args[2]);
    if (isError(startNum)) return startNum;
    if (isError(numChars)) return numChars;
    // Excel MID is 1-based
    return text.substring((startNum as number) - 1, (startNum as number) - 1 + (numChars as number));
  }

  private _funcSubstitute(args: any[]): any {
    if (args.length < 3) return ERRORS.VALUE;
    const text = String(args[0] ?? '');
    const oldText = String(args[1] ?? '');
    const newText = String(args[2] ?? '');
    if (args.length >= 4) {
      const instanceNum = this._toNumber(args[3]);
      if (isError(instanceNum)) return instanceNum;
      let count = 0;
      let result = '';
      let searchFrom = 0;
      while (true) {
        const idx = text.indexOf(oldText, searchFrom);
        if (idx === -1) {
          result += text.substring(searchFrom);
          break;
        }
        count++;
        if (count === instanceNum) {
          result += text.substring(searchFrom, idx) + newText;
          result += text.substring(idx + oldText.length);
          break;
        }
        result += text.substring(searchFrom, idx + oldText.length);
        searchFrom = idx + oldText.length;
      }
      return result;
    }
    return text.split(oldText).join(newText);
  }

  private _funcRept(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const text = String(args[0] ?? '');
    const times = this._toNumber(args[1]);
    if (isError(times)) return times;
    if ((times as number) < 0) return ERRORS.VALUE;
    return text.repeat(times as number);
  }

  // ── Logic Function Implementations ──────────────────────────────────────

  private _funcIf(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const condition = this._toBool(args[0]);
    if (condition) {
      return args[1];
    }
    return args.length > 2 ? args[2] : false;
  }

  private _funcAnd(args: any[]): any {
    const flat = this._flattenArgs(args);
    if (flat.length === 0) return ERRORS.VALUE;
    for (const v of flat) {
      if (isError(v)) return v;
      if (!this._toBool(v)) return false;
    }
    return true;
  }

  private _funcOr(args: any[]): any {
    const flat = this._flattenArgs(args);
    if (flat.length === 0) return ERRORS.VALUE;
    for (const v of flat) {
      if (isError(v)) return v;
      if (this._toBool(v)) return true;
    }
    return false;
  }

  private _funcNot(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    if (isError(args[0])) return args[0];
    return !this._toBool(args[0]);
  }

  // ── Math Function Implementations ───────────────────────────────────────

  private _funcSum(args: any[]): any {
    const nums = this._flattenNumeric(args);
    let sum = 0;
    for (const n of nums) sum += n;
    return sum;
  }

  private _funcAverage(args: any[]): any {
    const nums = this._flattenNumeric(args);
    if (nums.length === 0) return ERRORS.DIV0;
    let sum = 0;
    for (const n of nums) sum += n;
    return sum / nums.length;
  }

  private _funcMin(args: any[]): any {
    const nums = this._flattenNumeric(args);
    if (nums.length === 0) return 0;
    return Math.min(...nums);
  }

  private _funcMax(args: any[]): any {
    const nums = this._flattenNumeric(args);
    if (nums.length === 0) return 0;
    return Math.max(...nums);
  }

  private _funcCount(args: any[]): number {
    const flat = this._flattenArgs(args);
    let count = 0;
    for (const v of flat) {
      if (typeof v === 'number') count++;
      else if (typeof v === 'string' && v !== '' && !isError(v)) {
        const n = Number(v);
        if (!isNaN(n)) count++;
      }
    }
    return count;
  }

  private _funcCountA(args: any[]): number {
    const flat = this._flattenArgs(args);
    let count = 0;
    for (const v of flat) {
      if (v !== null && v !== undefined && v !== '') count++;
    }
    return count;
  }

  private _funcAbs(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    if (isError(num)) return num;
    return Math.abs(num as number);
  }

  private _funcRound(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const digits = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(digits)) return digits;
    const factor = Math.pow(10, digits as number);
    return Math.round((num as number) * factor) / factor;
  }

  private _funcRoundUp(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const digits = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(digits)) return digits;
    const factor = Math.pow(10, digits as number);
    const n = num as number;
    if (n >= 0) {
      return Math.ceil(n * factor) / factor;
    } else {
      return -Math.ceil(-n * factor) / factor;
    }
  }

  private _funcRoundDown(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const digits = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(digits)) return digits;
    const factor = Math.pow(10, digits as number);
    return Math.trunc((num as number) * factor) / factor;
  }

  private _funcInt(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    if (isError(num)) return num;
    return Math.floor(num as number);
  }

  private _funcMod(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const divisor = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(divisor)) return divisor;
    if ((divisor as number) === 0) return ERRORS.DIV0;
    // Excel MOD: result has same sign as divisor
    const n = num as number;
    const d = divisor as number;
    return n - d * Math.floor(n / d);
  }

  private _funcPower(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const base = this._toNumber(args[0]);
    const exp = this._toNumber(args[1]);
    if (isError(base)) return base;
    if (isError(exp)) return exp;
    return Math.pow(base as number, exp as number);
  }

  private _funcSqrt(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    if (isError(num)) return num;
    if ((num as number) < 0) return ERRORS.NUM;
    return Math.sqrt(num as number);
  }

  private _funcCeiling(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const significance = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(significance)) return significance;
    if ((significance as number) === 0) return 0;
    return Math.ceil((num as number) / (significance as number)) * (significance as number);
  }

  private _funcFloor(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const num = this._toNumber(args[0]);
    const significance = this._toNumber(args[1]);
    if (isError(num)) return num;
    if (isError(significance)) return significance;
    if ((significance as number) === 0) return ERRORS.DIV0;
    return Math.floor((num as number) / (significance as number)) * (significance as number);
  }

  // ── Lookup Function Implementations ─────────────────────────────────────

  private _funcVlookup(args: any[], worksheet: WorksheetLike | null): any {
    if (args.length < 3) return ERRORS.VALUE;
    const lookupValue = args[0];
    const tableArray = args[1]; // Should be an array (from range)
    const colIndexNum = this._toNumber(args[2]);
    if (isError(colIndexNum)) return colIndexNum;
    const rangeLookup = args.length >= 4 ? this._toBool(args[3]) : true;

    // tableArray should be a 2D-like flat array. We need the original range ref.
    // Since we already flattened it, we need to reconstruct.
    // For VLOOKUP, we need the 2D structure. We'll handle this via the token-level dispatch.
    if (!Array.isArray(tableArray)) return ERRORS.VALUE;

    // If tableArray is already flat, we can't determine dimensions reliably.
    // For the test scenarios, we'll attempt to work with 2D arrays
    // Since _callFunctionWithTokens resolves ranges to flat arrays,
    // we need a different approach for VLOOKUP.

    // Actually let's detect: if the first element is an array, it's 2D
    const is2D = tableArray.length > 0 && Array.isArray(tableArray[0]);
    let table2D: any[][];

    if (is2D) {
      table2D = tableArray as any[][];
    } else {
      // Flat array - can't determine columns. Assume it's a single column.
      table2D = tableArray.map((v: any) => [v]);
    }

    const colIdx = (colIndexNum as number) - 1;
    if (colIdx < 0) return ERRORS.VALUE;

    if (rangeLookup) {
      // Approximate match (sorted data, find largest value <= lookupValue)
      let lastMatch = -1;
      for (let r = 0; r < table2D.length; r++) {
        const cellVal = table2D[r][0];
        if (cellVal == null) continue;
        if (typeof lookupValue === 'number' && typeof cellVal === 'number') {
          if (cellVal <= lookupValue) lastMatch = r;
          else break;
        } else {
          const cmp = String(cellVal).toLowerCase().localeCompare(String(lookupValue).toLowerCase());
          if (cmp <= 0) lastMatch = r;
          else break;
        }
      }
      if (lastMatch === -1) return ERRORS.NA;
      if (colIdx >= table2D[lastMatch].length) return ERRORS.REF;
      return table2D[lastMatch][colIdx];
    } else {
      // Exact match
      for (let r = 0; r < table2D.length; r++) {
        const cellVal = table2D[r][0];
        if (this._valuesEqual(cellVal, lookupValue)) {
          if (colIdx >= table2D[r].length) return ERRORS.REF;
          return table2D[r][colIdx];
        }
      }
      return ERRORS.NA;
    }
  }

  private _funcHlookup(args: any[], worksheet: WorksheetLike | null): any {
    if (args.length < 3) return ERRORS.VALUE;
    const lookupValue = args[0];
    const tableArray = args[1];
    const rowIndexNum = this._toNumber(args[2]);
    if (isError(rowIndexNum)) return rowIndexNum;
    const rangeLookup = args.length >= 4 ? this._toBool(args[3]) : true;

    if (!Array.isArray(tableArray)) return ERRORS.VALUE;

    const is2D = tableArray.length > 0 && Array.isArray(tableArray[0]);
    let table2D: any[][];

    if (is2D) {
      table2D = tableArray as any[][];
    } else {
      // Single row
      table2D = [tableArray];
    }

    const rowIdx = (rowIndexNum as number) - 1;
    if (rowIdx < 0 || rowIdx >= table2D.length) return ERRORS.REF;

    if (table2D.length === 0) return ERRORS.NA;

    const firstRow = table2D[0];

    if (rangeLookup) {
      let lastMatch = -1;
      for (let c = 0; c < firstRow.length; c++) {
        const cellVal = firstRow[c];
        if (cellVal == null) continue;
        if (typeof lookupValue === 'number' && typeof cellVal === 'number') {
          if (cellVal <= lookupValue) lastMatch = c;
          else break;
        } else {
          const cmp = String(cellVal).toLowerCase().localeCompare(String(lookupValue).toLowerCase());
          if (cmp <= 0) lastMatch = c;
          else break;
        }
      }
      if (lastMatch === -1) return ERRORS.NA;
      return table2D[rowIdx][lastMatch];
    } else {
      for (let c = 0; c < firstRow.length; c++) {
        if (this._valuesEqual(firstRow[c], lookupValue)) {
          return table2D[rowIdx][c];
        }
      }
      return ERRORS.NA;
    }
  }

  private _funcIndex(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const arr = args[0];
    const rowNum = this._toNumber(args[1]);
    if (isError(rowNum)) return rowNum;

    if (!Array.isArray(arr)) return ERRORS.REF;

    const is2D = arr.length > 0 && Array.isArray(arr[0]);
    if (is2D) {
      const table2D = arr as any[][];
      const rIdx = (rowNum as number) - 1;
      if (rIdx < 0 || rIdx >= table2D.length) return ERRORS.REF;
      if (args.length >= 3) {
        const colNum = this._toNumber(args[2]);
        if (isError(colNum)) return colNum;
        const cIdx = (colNum as number) - 1;
        if (cIdx < 0 || cIdx >= table2D[rIdx].length) return ERRORS.REF;
        return table2D[rIdx][cIdx];
      }
      if (table2D[rIdx].length === 1) return table2D[rIdx][0];
      return table2D[rIdx];
    }

    // 1D array
    const idx = (rowNum as number) - 1;
    if (idx < 0 || idx >= arr.length) return ERRORS.REF;
    return arr[idx];
  }

  private _funcMatch(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const lookupValue = args[0];
    const lookupArray = args[1];
    const matchType = args.length >= 3 ? this._toNumber(args[2]) : 1;
    if (isError(matchType)) return matchType;

    if (!Array.isArray(lookupArray)) return ERRORS.NA;

    const flat = this._flattenArgs([lookupArray]);

    if ((matchType as number) === 0) {
      // Exact match
      for (let i = 0; i < flat.length; i++) {
        if (this._valuesEqual(flat[i], lookupValue)) return i + 1;
      }
      return ERRORS.NA;
    }

    if ((matchType as number) === 1) {
      // Largest value <= lookupValue (sorted ascending)
      let lastMatch = -1;
      for (let i = 0; i < flat.length; i++) {
        if (typeof lookupValue === 'number' && typeof flat[i] === 'number') {
          if (flat[i] <= lookupValue) lastMatch = i;
          else break;
        }
      }
      if (lastMatch === -1) return ERRORS.NA;
      return lastMatch + 1;
    }

    if ((matchType as number) === -1) {
      // Smallest value >= lookupValue (sorted descending)
      let lastMatch = -1;
      for (let i = 0; i < flat.length; i++) {
        if (typeof lookupValue === 'number' && typeof flat[i] === 'number') {
          if (flat[i] >= lookupValue) lastMatch = i;
          else break;
        }
      }
      if (lastMatch === -1) return ERRORS.NA;
      return lastMatch + 1;
    }

    return ERRORS.NA;
  }

  private _valuesEqual(a: any, b: any): boolean {
    if (a === b) return true;
    if (a == null || b == null) return false;
    if (typeof a === 'number' && typeof b === 'number') return a === b;
    return String(a).toLowerCase() === String(b).toLowerCase();
  }

  // ── Date Function Implementations ───────────────────────────────────────

  private _funcToday(): number {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    return CellValueHandler.dateToExcelSerial(today);
  }

  private _funcNow(): number {
    return CellValueHandler.dateToExcelSerial(new Date());
  }

  private _funcDate(args: any[]): any {
    if (args.length < 3) return ERRORS.VALUE;
    const year = this._toNumber(args[0]);
    const month = this._toNumber(args[1]);
    const day = this._toNumber(args[2]);
    if (isError(year)) return year;
    if (isError(month)) return month;
    if (isError(day)) return day;
    const dt = new Date(year as number, (month as number) - 1, day as number);
    return CellValueHandler.dateToExcelSerial(dt);
  }

  private _funcYear(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const serial = this._toNumber(args[0]);
    if (isError(serial)) return serial;
    const dt = CellValueHandler.excelSerialToDate(serial as number);
    return dt.getFullYear();
  }

  private _funcMonth(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const serial = this._toNumber(args[0]);
    if (isError(serial)) return serial;
    const dt = CellValueHandler.excelSerialToDate(serial as number);
    return dt.getMonth() + 1;
  }

  private _funcDay(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const serial = this._toNumber(args[0]);
    if (isError(serial)) return serial;
    const dt = CellValueHandler.excelSerialToDate(serial as number);
    return dt.getDate();
  }

  // ── Info Function Implementations ───────────────────────────────────────

  private _funcIsNumber(args: any[]): boolean {
    if (args.length < 1) return false;
    return typeof args[0] === 'number';
  }

  private _funcIsText(args: any[]): boolean {
    if (args.length < 1) return false;
    return typeof args[0] === 'string' && !isError(args[0]);
  }

  private _funcIsBlank(args: any[]): boolean {
    if (args.length < 1) return true;
    return args[0] === null || args[0] === undefined || args[0] === '';
  }

  private _funcIsError(args: any[]): boolean {
    if (args.length < 1) return false;
    return isError(args[0]);
  }

  private _funcIsNa(args: any[]): boolean {
    if (args.length < 1) return false;
    return args[0] === ERRORS.NA;
  }

  // ── Other Function Implementations ──────────────────────────────────────

  private _funcValue(args: any[]): any {
    if (args.length < 1) return ERRORS.VALUE;
    const val = args[0];
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      const num = Number(val);
      if (isNaN(num)) return ERRORS.VALUE;
      return num;
    }
    return ERRORS.VALUE;
  }

  private _funcChoose(args: any[]): any {
    if (args.length < 2) return ERRORS.VALUE;
    const indexNum = this._toNumber(args[0]);
    if (isError(indexNum)) return indexNum;
    const idx = indexNum as number;
    if (idx < 1 || idx >= args.length) return ERRORS.VALUE;
    return args[idx];
  }
}
