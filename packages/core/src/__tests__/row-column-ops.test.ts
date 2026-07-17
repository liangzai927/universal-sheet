import { describe, expect, it } from 'vitest';

import {
  deleteColumns,
  deleteRows,
  hideColumns,
  hideRows,
  insertColumns,
  insertRows,
  unhideColumns,
  unhideRows,
} from '../row-column-ops';
import {
  getCellData,
  getColumnWidth,
  getMergeAt,
  getRowHeight,
  mergeCells,
  setCellValue,
  setColumnWidth,
  setRowHeight,
} from '../sheet-model';
import { createSheetData } from '../sheet-model';
import type { SheetData } from '../types';

describe('insertRows', () => {
  it('should shift cells and row settings when inserting rows', () => {
    const sheet = setRowHeight(setCellValue(createSheetData({ rowCount: 4 }), 2, 0, 'A'), 2, 40);

    const updated = insertRows(sheet, 1, 2);

    expect(updated.config.rowCount).toBe(6);
    expect(getCellData(updated, 4, 0)?.value).toBe('A');
    expect(getRowHeight(updated, 4)).toBe(40);
    expect(getCellData(sheet, 2, 0)?.value).toBe('A');
  });

  it('should expand a merge when rows are inserted inside it', () => {
    const sheet = buildMergedSheet({ startRow: 1, startCol: 1, endRow: 3, endCol: 2 });

    const updated = insertRows(sheet, 2, 1);

    expect(getMergeAt(updated, 3, 1)).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 4,
      endCol: 2,
    });
  });
});

describe('deleteRows', () => {
  it('should delete row cells and shift following cells upward', () => {
    const sheet = setCellValue(
      setCellValue(createSheetData({ rowCount: 5 }), 1, 0, 'delete'),
      3,
      0,
      'keep',
    );

    const updated = deleteRows(sheet, 1, 2);

    expect(updated.config.rowCount).toBe(3);
    expect(getCellData(updated, 1, 0)?.value).toBe('keep');
    expect(getCellData(updated, 3, 0)).toBeNull();
  });

  it('should keep at least one row when deleting too many rows', () => {
    const sheet = setCellValue(createSheetData({ rowCount: 2 }), 1, 0, 'last');

    const updated = deleteRows(sheet, 0, 10);

    expect(updated.config.rowCount).toBe(1);
    expect(getCellData(updated, 0, 0)?.value).toBe('last');
  });

  it('should shrink or remove merges touched by deleted rows', () => {
    const sheet = buildMergedSheet({ startRow: 1, startCol: 1, endRow: 3, endCol: 2 });

    const updated = deleteRows(sheet, 2, 1);

    expect(getMergeAt(updated, 2, 1)).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2,
    });
  });
});

describe('insertColumns', () => {
  it('should shift cells and column settings when inserting columns', () => {
    const sheet = setColumnWidth(setCellValue(createSheetData({ colCount: 4 }), 0, 2, 'A'), 2, 120);

    const updated = insertColumns(sheet, 1, 2);

    expect(updated.config.colCount).toBe(6);
    expect(getCellData(updated, 0, 4)?.value).toBe('A');
    expect(getColumnWidth(updated, 4)).toBe(120);
  });

  it('should expand a merge when columns are inserted inside it', () => {
    const sheet = buildMergedSheet({ startRow: 1, startCol: 1, endRow: 2, endCol: 3 });

    const updated = insertColumns(sheet, 2, 1);

    expect(getMergeAt(updated, 1, 3)).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 4,
    });
  });
});

describe('deleteColumns', () => {
  it('should delete column cells and shift following cells left', () => {
    const sheet = setCellValue(
      setCellValue(createSheetData({ colCount: 5 }), 0, 1, 'delete'),
      0,
      3,
      'keep',
    );

    const updated = deleteColumns(sheet, 1, 2);

    expect(updated.config.colCount).toBe(3);
    expect(getCellData(updated, 0, 1)?.value).toBe('keep');
    expect(getCellData(updated, 0, 3)).toBeNull();
  });

  it('should keep at least one column when deleting too many columns', () => {
    const sheet = setCellValue(createSheetData({ colCount: 2 }), 0, 1, 'last');

    const updated = deleteColumns(sheet, 0, 10);

    expect(updated.config.colCount).toBe(1);
    expect(getCellData(updated, 0, 0)?.value).toBe('last');
  });

  it('should remove a merge when deleted columns collapse it to one cell', () => {
    const sheet = buildMergedSheet({ startRow: 1, startCol: 1, endRow: 1, endCol: 2 });

    const updated = deleteColumns(sheet, 2, 1);

    expect(getMergeAt(updated, 1, 1)).toBeNull();
  });
});

describe('hideRows', () => {
  it('should set hidden state while preserving row height', () => {
    const sheet = setRowHeight(createSheetData({ rowCount: 4 }), 2, 44);

    const hidden = hideRows(sheet, 2, 1);
    const visible = unhideRows(hidden, 2, 1);

    expect(hidden.rows.get(2)?.hidden).toBe(true);
    expect(getRowHeight(visible, 2)).toBe(44);
    expect(visible.rows.get(2)?.hidden).toBe(false);
  });
});

describe('hideColumns', () => {
  it('should set hidden state while preserving column width', () => {
    const sheet = setColumnWidth(createSheetData({ colCount: 4 }), 2, 144);

    const hidden = hideColumns(sheet, 2, 1);
    const visible = unhideColumns(hidden, 2, 1);

    expect(hidden.columns.get(2)?.hidden).toBe(true);
    expect(getColumnWidth(visible, 2)).toBe(144);
    expect(visible.columns.get(2)?.hidden).toBe(false);
  });
});

/**
 * 创建包含指定合并区域的测试工作表。
 *
 * @param range - 需要创建的合并区域
 * @returns 包含合并区域的工作表数据
 * @author liangzai927
 */
function buildMergedSheet(range: Parameters<typeof mergeCells>[1]): SheetData {
  const merged = mergeCells(createSheetData({ rowCount: 6, colCount: 6 }), range);
  if (merged) return merged;
  return createSheetData({ rowCount: 6, colCount: 6 });
}
