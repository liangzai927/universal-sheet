import { describe, expect, it } from 'vitest';

import { calculateFormulaCells } from '../index';

describe('calculateFormulaCells', () => {
  it('should calculate formulas after their dependencies', () => {
    const result = calculateFormulaCells([
      { ref: 'A1', formula: '=B1+1' },
      { ref: 'B1', value: 2 },
    ]);

    expect(result.values.get('A1')).toBe(3);
    expect(result.order).toEqual(['A1']);
  });

  it('should calculate chained formula dependencies in order', () => {
    const result = calculateFormulaCells([
      { ref: 'A1', formula: '=B1+1' },
      { ref: 'B1', formula: '=C1+1' },
      { ref: 'C1', value: 1 },
    ]);

    expect(result.values.get('A1')).toBe(3);
    expect(result.values.get('B1')).toBe(2);
    expect(result.order).toEqual(['B1', 'A1']);
  });

  it('should calculate range formulas from formula cell dependencies', () => {
    const result = calculateFormulaCells([
      { ref: 'A1', formula: '=SUM(B1:B2)' },
      { ref: 'B1', formula: '=1+1' },
      { ref: 'B2', value: 3 },
    ]);

    expect(result.values.get('A1')).toBe(5);
    expect(result.order).toEqual(['B1', 'A1']);
  });

  it('should return null for missing references', () => {
    const result = calculateFormulaCells([{ ref: 'A1', formula: '=B1+1' }]);

    expect(result.values.get('A1')).toBe(1);
  });

  it('should report circular reference errors', () => {
    const result = calculateFormulaCells([
      { ref: 'A1', formula: '=B1+1' },
      { ref: 'B1', formula: '=A1+1' },
    ]);

    expect(result.errors.get('A1')?.message).toContain('Circular formula reference');
    expect(result.errors.get('B1')?.message).toContain('Circular formula reference');
  });

  it('should propagate dependency calculation errors', () => {
    const result = calculateFormulaCells([
      { ref: 'A1', formula: '=B1+1' },
      { ref: 'B1', formula: '=1/0' },
    ]);

    expect(result.errors.get('A1')?.code).toBe('#DIV/0!');
    expect(result.errors.get('B1')?.code).toBe('#DIV/0!');
  });
});
