import { describe, expect, it } from 'vitest';

import {
  createSheetData,
  getCellData,
  getColumnWidth,
  getRowHeight,
  setCellValue,
  setColumnWidth,
  setRowHeight,
} from '../sheet-model';
import { cellKey, createPosition } from '../types';

describe('cellKey', () => {
  it('should produce a key string from row and col', () => {
    expect(cellKey(0, 0)).toBe('R0C0');
    expect(cellKey(5, 10)).toBe('R5C10');
  });
});

describe('createPosition', () => {
  it('should create a CellPosition from row and col', () => {
    const pos = createPosition(3, 7);
    expect(pos.row).toBe(3);
    expect(pos.col).toBe(7);
  });
});

describe('createSheetData', () => {
  it('should create an empty sheet with default config', () => {
    const sheet = createSheetData();
    expect(sheet.config.colCount).toBe(26);
    expect(sheet.config.rowCount).toBe(100);
    expect(sheet.config.defaultColWidth).toBe(100);
    expect(sheet.config.defaultRowHeight).toBe(28);
  });

  it('should merge overrides into config', () => {
    const sheet = createSheetData({ colCount: 5, defaultColWidth: 80 });
    expect(sheet.config.colCount).toBe(5);
    expect(sheet.config.defaultColWidth).toBe(80);
    expect(sheet.config.rowCount).toBe(100); // default
  });
});

describe('setCellValue', () => {
  it('should store a value in the sheet', () => {
    const sheet = createSheetData();
    const updated = setCellValue(sheet, 0, 0, 'hello');
    const cell = getCellData(updated, 0, 0);
    expect(cell?.value).toBe('hello');
  });

  it('should not mutate the original sheet', () => {
    const sheet = createSheetData();
    setCellValue(sheet, 0, 0, 'hello');
    expect(getCellData(sheet, 0, 0)).toBeNull();
  });

  it('should return null for an empty cell', () => {
    const sheet = createSheetData();
    expect(getCellData(sheet, 0, 0)).toBeNull();
  });
});

describe('setColumnWidth', () => {
  it('should set a per-column width', () => {
    const sheet = createSheetData();
    const updated = setColumnWidth(sheet, 2, 150);
    expect(getColumnWidth(updated, 2)).toBe(150);
  });

  it('should not affect other columns', () => {
    const sheet = createSheetData();
    const updated = setColumnWidth(sheet, 2, 150);
    expect(getColumnWidth(updated, 1)).toBe(100); // default
  });
});

describe('setRowHeight', () => {
  it('should set a per-row height', () => {
    const sheet = createSheetData();
    const updated = setRowHeight(sheet, 5, 40);
    expect(getRowHeight(updated, 5)).toBe(40);
  });

  it('should not affect other rows', () => {
    const sheet = createSheetData();
    const updated = setRowHeight(sheet, 5, 40);
    expect(getRowHeight(updated, 4)).toBe(28); // default
  });
});
