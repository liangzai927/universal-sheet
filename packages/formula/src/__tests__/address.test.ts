import { describe, expect, it } from 'vitest';

import {
  columnIndexToLabel,
  columnLabelToIndex,
  enumerateRangeAddress,
  formatCellAddress,
  formatRangeAddress,
  isCellAddress,
  parseCellAddress,
  parseRangeAddress,
} from '../index';

describe('columnLabelToIndex', () => {
  it('should convert Excel column labels to zero-based indexes', () => {
    expect(columnLabelToIndex('A')).toBe(0);
    expect(columnLabelToIndex('Z')).toBe(25);
    expect(columnLabelToIndex('AA')).toBe(26);
  });
});

describe('columnIndexToLabel', () => {
  it('should convert zero-based indexes to Excel column labels', () => {
    expect(columnIndexToLabel(0)).toBe('A');
    expect(columnIndexToLabel(25)).toBe('Z');
    expect(columnIndexToLabel(26)).toBe('AA');
  });
});

describe('isCellAddress', () => {
  it('should detect valid A1 cell addresses', () => {
    expect(isCellAddress('A1')).toBe(true);
    expect(isCellAddress('$B$2')).toBe(true);
    expect(isCellAddress('SUM')).toBe(false);
  });
});

describe('parseCellAddress', () => {
  it('should parse A1 addresses with absolute flags', () => {
    expect(parseCellAddress('$B$2')).toEqual({
      row: 1,
      col: 1,
      absoluteRow: true,
      absoluteCol: true,
    });
  });
});

describe('formatCellAddress', () => {
  it('should format zero-based coordinates as A1 addresses', () => {
    expect(formatCellAddress({ row: 4, col: 27, absoluteRow: false, absoluteCol: true })).toBe(
      '$AB5',
    );
  });
});

describe('parseRangeAddress', () => {
  it('should normalize reversed ranges', () => {
    expect(parseRangeAddress('B2:A1')).toEqual({
      start: { row: 0, col: 0, absoluteRow: false, absoluteCol: false },
      end: { row: 1, col: 1, absoluteRow: false, absoluteCol: false },
    });
  });
});

describe('formatRangeAddress', () => {
  it('should format ranges as A1 range text', () => {
    expect(
      formatRangeAddress({
        start: { row: 0, col: 0, absoluteRow: false, absoluteCol: false },
        end: { row: 1, col: 1, absoluteRow: true, absoluteCol: true },
      }),
    ).toBe('A1:$B$2');
  });
});

describe('enumerateRangeAddress', () => {
  it('should enumerate cells in row-major order', () => {
    const cells = enumerateRangeAddress(parseRangeAddress('A1:B2'));

    expect(cells.map((cell) => formatCellAddress(cell))).toEqual(['A1', 'B1', 'A2', 'B2']);
  });
});
