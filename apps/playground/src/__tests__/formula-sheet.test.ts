import { createSheetData, getCellData } from '@universal-sheet/core';
import { describe, expect, it } from 'vitest';

import {
  applyCellInput,
  deleteRowsWithFormulas,
  insertColumnsWithFormulas,
} from '../formula-sheet';

describe('insertColumnsWithFormulas', () => {
  it('should shift formula references when inserting columns before referenced cells', () => {
    const sheet = applyCellInput(
      applyCellInput(createSheetData({ rowCount: 3, colCount: 3 }), 0, 0, '2'),
      0,
      1,
      '=A1',
    );

    const updated = insertColumnsWithFormulas(sheet, 0, 1);
    const cell = getCellData(updated, 0, 2);

    expect(cell).toMatchObject({
      formula: '=B1',
      value: '2',
      displayValue: '2',
    });
  });
});

describe('deleteRowsWithFormulas', () => {
  it('should convert deleted cell references to REF errors', () => {
    const sheet = applyCellInput(
      applyCellInput(createSheetData({ rowCount: 3, colCount: 3 }), 0, 0, '5'),
      1,
      0,
      '=A1',
    );

    const updated = deleteRowsWithFormulas(sheet, 0, 1);
    const cell = getCellData(updated, 0, 0);

    expect(cell).toMatchObject({
      formula: '=#REF!',
      value: '#REF!',
      displayValue: '#REF!',
    });
  });
});
