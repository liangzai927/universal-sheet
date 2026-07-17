import { describe, expect, it } from 'vitest';

import {
  createFormulaRecord,
  evaluateFormula,
  FormulaError,
  fromExcelFormulaText,
  normalizeFormulaText,
  parseFormula,
  toExcelFormulaText,
  tokenizeFormula,
} from '../index';

describe('tokenizeFormula', () => {
  it('should tokenize a function call with a range argument', () => {
    const tokens = tokenizeFormula('=SUM(A1:B2, 3)');

    expect(tokens.map((token) => token.type)).toEqual([
      'identifier',
      'leftParen',
      'identifier',
      'colon',
      'identifier',
      'comma',
      'number',
      'rightParen',
      'eof',
    ]);
  });

  it('should tokenize Excel error literals', () => {
    const tokens = tokenizeFormula('=#REF!');

    expect(tokens.map((token) => token.type)).toEqual(['error', 'eof']);
  });
});

describe('parseFormula', () => {
  it('should parse a binary expression with cell references', () => {
    const ast = parseFormula('=A1+B2');

    expect(ast).toEqual({
      type: 'binary',
      operator: '+',
      left: { type: 'cell', ref: 'A1' },
      right: { type: 'cell', ref: 'B2' },
    });
  });

  it('should parse a function call with a range reference', () => {
    const ast = parseFormula('=SUM(A1:A3)');

    expect(ast).toEqual({
      type: 'function',
      name: 'SUM',
      args: [{ type: 'range', start: 'A1', end: 'A3' }],
    });
  });

  it('should parse Excel error literals', () => {
    expect(parseFormula('=#REF!')).toEqual({ type: 'error', code: '#REF!' });
  });
});

describe('evaluateFormula', () => {
  it('should respect arithmetic precedence', () => {
    expect(evaluateFormula('=1+2*3')).toBe(7);
  });

  it('should resolve cell references from context', () => {
    const values = new Map([
      ['A1', 2],
      ['B1', 5],
    ]);

    expect(evaluateFormula('=A1+B1', { getCellValue: (ref) => values.get(ref) ?? null })).toBe(7);
  });

  it('should calculate functions with range values', () => {
    expect(
      evaluateFormula('=SUM(A1:A3)', {
        getRangeValues: () => [1, 2, 3],
      }),
    ).toBe(6);
  });

  it('should resolve range values from cell context', () => {
    const values = new Map([
      ['A1', 1],
      ['B1', 2],
      ['A2', 3],
      ['B2', 4],
    ]);

    expect(evaluateFormula('=SUM(A1:B2)', { getCellValue: (ref) => values.get(ref) ?? null })).toBe(
      10,
    );
  });

  it('should choose the true branch for IF formulas', () => {
    expect(evaluateFormula('=IF(A1>0,"yes","no")', { getCellValue: () => 1 })).toBe('yes');
  });

  it('should throw FormulaError when dividing by zero', () => {
    expect(() => evaluateFormula('=1/0')).toThrow(FormulaError);
  });

  it('should throw name error for unknown functions', () => {
    try {
      evaluateFormula('=UNKNOWN(1)');
    } catch (error) {
      expect(error).toMatchObject({ code: '#NAME?' });
    }
  });

  it('should throw FormulaError for error literals', () => {
    expect(() => evaluateFormula('=#REF!')).toThrow(FormulaError);
  });
});

describe('normalizeFormulaText', () => {
  it('should normalize formula text with one leading equals sign', () => {
    expect(normalizeFormulaText('SUM(A1:A2)')).toBe('=SUM(A1:A2)');
  });
});

describe('toExcelFormulaText', () => {
  it('should remove the leading equals sign for Excel storage', () => {
    expect(toExcelFormulaText('=SUM(A1:A2)')).toBe('SUM(A1:A2)');
  });
});

describe('fromExcelFormulaText', () => {
  it('should add the leading equals sign from Excel storage text', () => {
    expect(fromExcelFormulaText('SUM(A1:A2)')).toBe('=SUM(A1:A2)');
  });
});

describe('createFormulaRecord', () => {
  it('should create a parsed formula record for supported formulas', () => {
    const record = createFormulaRecord('SUM(A1:A2)', { cachedValue: 3 });

    expect(record).toMatchObject({
      formula: '=SUM(A1:A2)',
      excelFormula: 'SUM(A1:A2)',
      cachedValue: 3,
      ast: {
        type: 'function',
        name: 'SUM',
      },
    });
  });

  it('should preserve unsupported Excel formulas without parsing', () => {
    const record = createFormulaRecord("'Sheet 1'!A1+1", { parse: false });

    expect(record).toEqual({
      formula: "='Sheet 1'!A1+1",
      excelFormula: "'Sheet 1'!A1+1",
    });
  });
});
