import { describe, expect, it } from 'vitest';

import {
  extractCellRangeText,
  pasteCellRangeStructured,
  pasteCellRangeText,
  serializeCellRange,
  tryParseClipboardPayload,
} from '../clipboard';
import { createSheetData, getCellData, setCellValue } from '../sheet-model';
import { cellKey } from '../types';

function buildSheet(
  values: Array<Array<string | number | boolean | null>>,
): ReturnType<typeof createSheetData> {
  let sheet = createSheetData({ colCount: values[0]?.length ?? 1, rowCount: values.length });
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < (values[r]?.length ?? 0); c++) {
      sheet = setCellValue(sheet, r, c, values[r]![c]!);
    }
  }
  return sheet;
}

describe('extractCellRangeText', () => {
  it('should extract a single cell value as plain text', () => {
    const sheet = buildSheet([['hello']]);
    const result = extractCellRangeText(sheet, { startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(result).toBe('hello');
  });

  it('should extract a multi-cell range with tabs and newlines', () => {
    const sheet = buildSheet([
      ['A', 'B'],
      ['C', 'D'],
    ]);
    const result = extractCellRangeText(sheet, { startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    expect(result).toBe('A\tB\nC\tD');
  });

  it('should produce empty string for empty cells', () => {
    const sheet = buildSheet([
      ['A', null],
      [null, 'D'],
    ]);
    const result = extractCellRangeText(sheet, { startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    expect(result).toBe('A\t\n\tD');
  });

  it('should handle numeric and boolean values', () => {
    const sheet = buildSheet([['hello', 42, true, false]]);
    const result = extractCellRangeText(sheet, { startRow: 0, startCol: 0, endRow: 0, endCol: 3 });
    expect(result).toBe('hello\t42\ttrue\tfalse');
  });
});

describe('pasteCellRangeText', () => {
  it('should paste a single value into one cell', () => {
    const sheet = createSheetData({ colCount: 3, rowCount: 3 });
    const updated = pasteCellRangeText(sheet, 0, 0, 'hello');
    expect(getCellData(updated, 0, 0)?.value).toBe('hello');
  });

  it('should paste tab-separated values across a row', () => {
    const sheet = createSheetData({ colCount: 5, rowCount: 3 });
    const updated = pasteCellRangeText(sheet, 0, 0, 'A\tB\tC');
    expect(getCellData(updated, 0, 0)?.value).toBe('A');
    expect(getCellData(updated, 0, 1)?.value).toBe('B');
    expect(getCellData(updated, 0, 2)?.value).toBe('C');
  });

  it('should paste multi-line text across rows', () => {
    const sheet = createSheetData({ colCount: 3, rowCount: 5 });
    const updated = pasteCellRangeText(sheet, 1, 1, 'X\tY\nZ\tW');
    expect(getCellData(updated, 1, 1)?.value).toBe('X');
    expect(getCellData(updated, 1, 2)?.value).toBe('Y');
    expect(getCellData(updated, 2, 1)?.value).toBe('Z');
    expect(getCellData(updated, 2, 2)?.value).toBe('W');
  });

  it('should clamp paste to sheet bounds', () => {
    const sheet = createSheetData({ colCount: 2, rowCount: 2 });
    const updated = pasteCellRangeText(sheet, 1, 1, 'A\tB\tC\nD\tE\tF');
    expect(getCellData(updated, 1, 1)?.value).toBe('A');
    expect(getCellData(updated, 1, 2)).toBeNull();
    expect(getCellData(updated, 2, 0)).toBeNull();
  });

  it('should not mutate the original sheet', () => {
    const sheet = createSheetData({ colCount: 3, rowCount: 3 });
    pasteCellRangeText(sheet, 0, 0, 'data');
    expect(getCellData(sheet, 0, 0)).toBeNull();
  });

  it('should clear cells when pasting empty strings', () => {
    const sheet = buildSheet([['before']]);
    const updated = pasteCellRangeText(sheet, 0, 0, '');
    expect(getCellData(updated, 0, 0)?.value).toBeNull();
  });
});

describe('serializeCellRange', () => {
  it('should include source range origin for formula paste offset', () => {
    const sheet = buildSheet([['A']]);
    const payload = tryParseClipboardPayload(
      serializeCellRange(sheet, { startRow: 2, startCol: 3, endRow: 2, endCol: 3 }),
    );

    expect(payload).toMatchObject({
      sourceStartRow: 2,
      sourceStartCol: 3,
    });
  });
});

describe('pasteCellRangeStructured', () => {
  it('should preserve formula metadata when pasting structured cells', () => {
    const cells = new Map([
      [
        cellKey(0, 0),
        {
          value: 3,
          displayValue: '3',
          formula: '=A1+B1',
        },
      ],
    ]);
    const sheet = {
      ...createSheetData({ colCount: 4, rowCount: 4 }),
      cells,
    };
    const payload = tryParseClipboardPayload(
      serializeCellRange(sheet, { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }),
    );

    if (!payload) throw new Error('Expected clipboard payload');

    const updated = pasteCellRangeStructured(sheet, 1, 1, payload);

    expect(getCellData(updated, 1, 1)).toMatchObject({
      value: 3,
      displayValue: '3',
      formula: '=A1+B1',
    });
  });
});
