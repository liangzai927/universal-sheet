import {
  shiftMergesForColumnDelete,
  shiftMergesForColumnInsert,
  shiftMergesForRowDelete,
  shiftMergesForRowInsert,
} from './row-column-merge';
import {
  shiftCellsForColumnDelete,
  shiftCellsForColumnInsert,
  shiftCellsForRowDelete,
  shiftCellsForRowInsert,
  shiftColumnsForDelete,
  shiftColumnsForInsert,
  shiftRowsForDelete,
  shiftRowsForInsert,
} from './row-column-shift';
import { clampInsertIndex, normalizeBoundedRange, normalizeDeleteRange } from './row-column-utils';
import type { SheetData } from './types';

/**
 * 在指定位置插入行，并同步移动单元格、行配置和合并区域。
 *
 * @param sheet - 需要更新的工作表数据
 * @param rowIndex - 插入行的起始位置
 * @param count - 插入行数量
 * @returns 插入行后的工作表数据
 * @author liangzai927
 */
export function insertRows(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  if (count <= 0) return sheet;

  const insertIndex = clampInsertIndex(rowIndex, sheet.config.rowCount);
  return {
    ...sheet,
    config: { ...sheet.config, rowCount: sheet.config.rowCount + count },
    cells: shiftCellsForRowInsert(sheet.cells, insertIndex, count),
    rows: shiftRowsForInsert(sheet.rows, insertIndex, count),
    merges: shiftMergesForRowInsert(sheet.merges, insertIndex, count),
  };
}

/**
 * 删除指定范围内的行，并同步移动单元格、行配置和合并区域。
 *
 * @param sheet - 需要更新的工作表数据
 * @param rowIndex - 删除行的起始位置
 * @param count - 删除行数量
 * @returns 删除行后的工作表数据
 * @author liangzai927
 */
export function deleteRows(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  const deletion = normalizeDeleteRange(rowIndex, count, sheet.config.rowCount);
  if (!deletion) return sheet;

  return {
    ...sheet,
    config: { ...sheet.config, rowCount: sheet.config.rowCount - deletion.count },
    cells: shiftCellsForRowDelete(sheet.cells, deletion.start, deletion.end, deletion.count),
    rows: shiftRowsForDelete(sheet.rows, deletion.start, deletion.end, deletion.count),
    merges: shiftMergesForRowDelete(sheet.merges, deletion.start, deletion.end, deletion.count),
  };
}

/**
 * 在指定位置插入列，并同步移动单元格、列配置和合并区域。
 *
 * @param sheet - 需要更新的工作表数据
 * @param colIndex - 插入列的起始位置
 * @param count - 插入列数量
 * @returns 插入列后的工作表数据
 * @author liangzai927
 */
export function insertColumns(sheet: SheetData, colIndex: number, count = 1): SheetData {
  if (count <= 0) return sheet;

  const insertIndex = clampInsertIndex(colIndex, sheet.config.colCount);
  return {
    ...sheet,
    config: { ...sheet.config, colCount: sheet.config.colCount + count },
    cells: shiftCellsForColumnInsert(sheet.cells, insertIndex, count),
    columns: shiftColumnsForInsert(sheet.columns, insertIndex, count),
    merges: shiftMergesForColumnInsert(sheet.merges, insertIndex, count),
  };
}

/**
 * 删除指定范围内的列，并同步移动单元格、列配置和合并区域。
 *
 * @param sheet - 需要更新的工作表数据
 * @param colIndex - 删除列的起始位置
 * @param count - 删除列数量
 * @returns 删除列后的工作表数据
 * @author liangzai927
 */
export function deleteColumns(sheet: SheetData, colIndex: number, count = 1): SheetData {
  const deletion = normalizeDeleteRange(colIndex, count, sheet.config.colCount);
  if (!deletion) return sheet;

  return {
    ...sheet,
    config: { ...sheet.config, colCount: sheet.config.colCount - deletion.count },
    cells: shiftCellsForColumnDelete(sheet.cells, deletion.start, deletion.end, deletion.count),
    columns: shiftColumnsForDelete(sheet.columns, deletion.start, deletion.end, deletion.count),
    merges: shiftMergesForColumnDelete(sheet.merges, deletion.start, deletion.end, deletion.count),
  };
}

/**
 * 隐藏指定范围内的行。
 *
 * @param sheet - 需要更新的工作表数据
 * @param rowIndex - 隐藏行的起始位置
 * @param count - 隐藏行数量
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function hideRows(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  return setRowsHidden(sheet, rowIndex, count, true);
}

/**
 * 取消隐藏指定范围内的行。
 *
 * @param sheet - 需要更新的工作表数据
 * @param rowIndex - 取消隐藏行的起始位置
 * @param count - 取消隐藏行数量
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function unhideRows(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  return setRowsHidden(sheet, rowIndex, count, false);
}

/**
 * 隐藏指定范围内的列。
 *
 * @param sheet - 需要更新的工作表数据
 * @param colIndex - 隐藏列的起始位置
 * @param count - 隐藏列数量
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function hideColumns(sheet: SheetData, colIndex: number, count = 1): SheetData {
  return setColumnsHidden(sheet, colIndex, count, true);
}

/**
 * 取消隐藏指定范围内的列。
 *
 * @param sheet - 需要更新的工作表数据
 * @param colIndex - 取消隐藏列的起始位置
 * @param count - 取消隐藏列数量
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function unhideColumns(sheet: SheetData, colIndex: number, count = 1): SheetData {
  return setColumnsHidden(sheet, colIndex, count, false);
}

/**
 * 设置指定范围内行的隐藏状态。
 *
 * @param sheet - 需要更新的工作表数据
 * @param rowIndex - 更新行的起始位置
 * @param count - 更新行数量
 * @param hidden - 是否隐藏
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function setRowsHidden(
  sheet: SheetData,
  rowIndex: number,
  count: number,
  hidden: boolean,
): SheetData {
  const range = normalizeBoundedRange(rowIndex, count, sheet.config.rowCount);
  if (!range) return sheet;

  const rows = new Map(sheet.rows);
  for (let row = range.start; row <= range.end; row++) {
    const existing = rows.get(row);
    rows.set(row, {
      height: existing?.height ?? sheet.config.defaultRowHeight,
      hidden,
    });
  }

  return { ...sheet, rows };
}

/**
 * 设置指定范围内列的隐藏状态。
 *
 * @param sheet - 需要更新的工作表数据
 * @param colIndex - 更新列的起始位置
 * @param count - 更新列数量
 * @param hidden - 是否隐藏
 * @returns 更新隐藏状态后的工作表数据
 * @author liangzai927
 */
export function setColumnsHidden(
  sheet: SheetData,
  colIndex: number,
  count: number,
  hidden: boolean,
): SheetData {
  const range = normalizeBoundedRange(colIndex, count, sheet.config.colCount);
  if (!range) return sheet;

  const columns = new Map(sheet.columns);
  for (let col = range.start; col <= range.end; col++) {
    const existing = columns.get(col);
    columns.set(col, {
      width: existing?.width ?? sheet.config.defaultColWidth,
      hidden,
    });
  }

  return { ...sheet, columns };
}
