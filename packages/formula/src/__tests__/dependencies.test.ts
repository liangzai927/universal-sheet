import { describe, expect, it } from 'vitest';

import { collectFormulaDependencies, collectFormulaNodeDependencies, parseFormula } from '../index';

describe('collectFormulaDependencies', () => {
  it('should collect direct cell references', () => {
    expect(collectFormulaDependencies('=A1+B2')).toEqual({
      cells: ['A1', 'B2'],
      ranges: [],
      expandedCells: ['A1', 'B2'],
    });
  });

  it('should collect range references and expanded cells', () => {
    expect(collectFormulaDependencies('=SUM(B2:A1)')).toEqual({
      cells: [],
      ranges: ['A1:B2'],
      expandedCells: ['A1', 'B1', 'A2', 'B2'],
    });
  });

  it('should deduplicate repeated dependencies while preserving first-seen order', () => {
    expect(collectFormulaDependencies('=A1+SUM(A1:A2)+A1')).toEqual({
      cells: ['A1'],
      ranges: ['A1:A2'],
      expandedCells: ['A1', 'A2'],
    });
  });

  it('should ignore literals and Excel error literals', () => {
    expect(collectFormulaDependencies('=IF(TRUE,#REF!,1)')).toEqual({
      cells: [],
      ranges: [],
      expandedCells: [],
    });
  });
});

describe('collectFormulaNodeDependencies', () => {
  it('should collect dependencies from parsed AST nodes', () => {
    const ast = parseFormula('=MAX(C1:C2,D1)');

    expect(collectFormulaNodeDependencies(ast)).toEqual({
      cells: ['D1'],
      ranges: ['C1:C2'],
      expandedCells: ['C1', 'C2', 'D1'],
    });
  });
});
