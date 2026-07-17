import { describe, expect, it } from 'vitest';

import {
  applyFormulaReferenceChange,
  offsetFormulaReferences,
  parseFormula,
  serializeFormula,
  serializeFormulaNode,
} from '../index';

describe('serializeFormulaNode', () => {
  it('should serialize formula AST nodes into expression text', () => {
    expect(serializeFormulaNode(parseFormula('=SUM(A1:B2,3)'))).toBe('SUM(A1:B2,3)');
  });

  it('should escape string literals for formula text', () => {
    expect(serializeFormulaNode(parseFormula('="a"'))).toBe('"a"');
  });
});

describe('serializeFormula', () => {
  it('should serialize formula AST nodes with leading equals sign', () => {
    expect(serializeFormula(parseFormula('=A1+B1'))).toBe('=A1+B1');
  });

  it('should preserve precedence with parentheses for nested binary nodes', () => {
    expect(serializeFormula(parseFormula('=(A1+B1)*C1'))).toBe('=(A1+B1)*C1');
  });

  it('should serialize Excel error literals', () => {
    expect(serializeFormula(parseFormula('=#REF!'))).toBe('=#REF!');
  });
});

describe('offsetFormulaReferences', () => {
  it('should offset relative cell references for copy fill', () => {
    expect(offsetFormulaReferences('=A1+B2', { rowOffset: 1, colOffset: 2 })).toBe('=C2+D3');
  });

  it('should keep absolute cell parts unchanged by default', () => {
    expect(offsetFormulaReferences('=$A1+B$2+$C$3', { rowOffset: 2, colOffset: 3 })).toBe(
      '=$A3+E$2+$C$3',
    );
  });

  it('should offset range references', () => {
    expect(offsetFormulaReferences('=SUM(A1:B2)', { rowOffset: 1, colOffset: 1 })).toBe(
      '=SUM(B2:C3)',
    );
  });

  it('should offset absolute references when requested', () => {
    expect(
      offsetFormulaReferences('=$A$1', { rowOffset: 1, colOffset: 1, respectAbsolute: false }),
    ).toBe('=$B$2');
  });
});

describe('applyFormulaReferenceChange', () => {
  it('should move row references after inserted rows', () => {
    expect(
      applyFormulaReferenceChange('=A1+A3', { type: 'insertRows', startIndex: 1, count: 2 }),
    ).toBe('=A1+A5');
  });

  it('should move column references after inserted columns', () => {
    expect(
      applyFormulaReferenceChange('=A1+C1', { type: 'insertColumns', startIndex: 1, count: 2 }),
    ).toBe('=A1+E1');
  });

  it('should move references after deleted rows', () => {
    expect(
      applyFormulaReferenceChange('=A1+A5', { type: 'deleteRows', startIndex: 1, count: 2 }),
    ).toBe('=A1+A3');
  });

  it('should replace deleted cell references with REF error literals', () => {
    expect(
      applyFormulaReferenceChange('=A2+B4', { type: 'deleteRows', startIndex: 1, count: 1 }),
    ).toBe('=#REF!+B3');
  });

  it('should replace ranges with REF when an endpoint is deleted', () => {
    expect(
      applyFormulaReferenceChange('=SUM(A1:A3)', { type: 'deleteRows', startIndex: 2, count: 1 }),
    ).toBe('=SUM(#REF!)');
  });
});
