import { describe, expect, it } from 'vitest';

import {
  deserializeSheetData,
  deserializeWorkbookData,
  parseWorkbookData,
  serializeSheetData,
  serializeWorkbookData,
  stringifyWorkbookData,
} from '../serialization';
import {
  getCellData,
  getColumnWidth,
  getMergeAt,
  getRowHeight,
  mergeCells,
  setCellStyle,
  setCellValue,
  setColumnWidth,
  setRowHeight,
} from '../sheet-model';
import { createWorkbookData } from '../workbook-model';

describe('serializeSheetData', () => {
  it('should serialize sheet maps into JSON-safe arrays', () => {
    const sheet = buildRichSheet();
    const serialized = serializeSheetData(sheet);

    expect(serialized.cells).toHaveLength(1);
    expect(serialized.columns).toHaveLength(1);
    expect(serialized.rows).toHaveLength(1);
    expect(serialized.merges).toHaveLength(1);
    expect(serialized.config.colCount).toBe(8);
  });
});

describe('deserializeSheetData', () => {
  it('should restore serialized sheet maps and values', () => {
    const sheet = buildRichSheet();
    const restored = deserializeSheetData(serializeSheetData(sheet));

    expect(getCellData(restored, 0, 0)?.value).toBe('hello');
    expect(getCellData(restored, 0, 0)?.style?.bold).toBe(true);
    expect(getColumnWidth(restored, 2)).toBe(120);
    expect(getRowHeight(restored, 3)).toBe(42);
    expect(getMergeAt(restored, 1, 1)).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2,
    });
  });
});

describe('serializeWorkbookData', () => {
  it('should serialize all workbook sheets', () => {
    const workbook = createWorkbookData({
      activeSheetId: 'report',
      sheets: [
        { id: 'raw', name: 'Raw', sheet: buildRichSheet() },
        { id: 'report', name: 'Report' },
      ],
    });

    const serialized = serializeWorkbookData(workbook);

    expect(serialized.version).toBe(1);
    expect(serialized.activeSheetId).toBe('report');
    expect(serialized.sheets.map((sheet) => sheet.id)).toEqual(['raw', 'report']);
  });
});

describe('deserializeWorkbookData', () => {
  it('should restore workbook sheets and active sheet id', () => {
    const workbook = createWorkbookData({
      activeSheetId: 'raw',
      sheets: [{ id: 'raw', name: 'Raw', sheet: buildRichSheet() }],
    });

    const restored = deserializeWorkbookData(serializeWorkbookData(workbook));

    expect(restored.activeSheetId).toBe('raw');
    expect(restored.sheets[0]?.name).toBe('Raw');
    expect(getCellData(restored.sheets[0]!.sheet, 0, 0)?.value).toBe('hello');
  });
});

describe('stringifyWorkbookData', () => {
  it('should create a JSON string from workbook data', () => {
    const workbook = createWorkbookData();
    const json = stringifyWorkbookData(workbook);

    expect(JSON.parse(json)).toMatchObject({
      version: 1,
      activeSheetId: 'sheet-1',
    });
  });
});

describe('parseWorkbookData', () => {
  it('should parse a workbook JSON string', () => {
    const workbook = createWorkbookData({
      sheets: [{ id: 'raw', name: 'Raw', sheet: buildRichSheet() }],
    });

    const parsed = parseWorkbookData(stringifyWorkbookData(workbook));

    expect(parsed.sheets[0]?.id).toBe('raw');
    expect(getCellData(parsed.sheets[0]!.sheet, 0, 0)?.value).toBe('hello');
  });
});

/**
 * 创建包含单元格、样式、行列尺寸和合并单元格的测试工作表。
 *
 * @returns 测试用工作表数据
 * @author liangzai927
 */
function buildRichSheet(): ReturnType<typeof setRowHeight> {
  let sheet = createWorkbookData({ sheetConfig: { colCount: 8, rowCount: 8 } }).sheets[0]!.sheet;
  sheet = setCellValue(sheet, 0, 0, 'hello');
  sheet = setCellStyle(sheet, 0, 0, { bold: true, backgroundColor: '#ffffaa' });
  sheet = setColumnWidth(sheet, 2, 120);
  sheet = setRowHeight(sheet, 3, 42);

  const merged = mergeCells(sheet, { startRow: 1, startCol: 1, endRow: 2, endCol: 2 });
  if (merged) return merged;
  return sheet;
}
