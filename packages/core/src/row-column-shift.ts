import { parseCellKey } from './row-column-utils';
import type { CellData, ColumnData, RowData } from './types';
import { cellKey } from './types';

/**
 * 插入行时移动单元格数据。
 *
 * @param cells - 原始单元格 Map
 * @param insertIndex - 插入行位置
 * @param count - 插入行数量
 * @returns 移动后的单元格 Map
 * @author liangzai927
 */
export function shiftCellsForRowInsert(
  cells: ReadonlyMap<string, CellData>,
  insertIndex: number,
  count: number,
): ReadonlyMap<string, CellData> {
  const next = new Map<string, CellData>();
  for (const [key, cell] of cells) {
    const pos = parseCellKey(key);
    if (!pos) {
      next.set(key, cell);
      continue;
    }
    const row = pos.row >= insertIndex ? pos.row + count : pos.row;
    next.set(cellKey(row, pos.col), cell);
  }
  return next;
}

/**
 * 删除行时移动单元格数据。
 *
 * @param cells - 原始单元格 Map
 * @param start - 删除起始行
 * @param end - 删除结束行
 * @param count - 删除行数量
 * @returns 移动后的单元格 Map
 * @author liangzai927
 */
export function shiftCellsForRowDelete(
  cells: ReadonlyMap<string, CellData>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<string, CellData> {
  const next = new Map<string, CellData>();
  for (const [key, cell] of cells) {
    const pos = parseCellKey(key);
    if (!pos) {
      next.set(key, cell);
      continue;
    }
    if (pos.row >= start && pos.row <= end) continue;

    const row = pos.row > end ? pos.row - count : pos.row;
    next.set(cellKey(row, pos.col), cell);
  }
  return next;
}

/**
 * 插入列时移动单元格数据。
 *
 * @param cells - 原始单元格 Map
 * @param insertIndex - 插入列位置
 * @param count - 插入列数量
 * @returns 移动后的单元格 Map
 * @author liangzai927
 */
export function shiftCellsForColumnInsert(
  cells: ReadonlyMap<string, CellData>,
  insertIndex: number,
  count: number,
): ReadonlyMap<string, CellData> {
  const next = new Map<string, CellData>();
  for (const [key, cell] of cells) {
    const pos = parseCellKey(key);
    if (!pos) {
      next.set(key, cell);
      continue;
    }
    const col = pos.col >= insertIndex ? pos.col + count : pos.col;
    next.set(cellKey(pos.row, col), cell);
  }
  return next;
}

/**
 * 删除列时移动单元格数据。
 *
 * @param cells - 原始单元格 Map
 * @param start - 删除起始列
 * @param end - 删除结束列
 * @param count - 删除列数量
 * @returns 移动后的单元格 Map
 * @author liangzai927
 */
export function shiftCellsForColumnDelete(
  cells: ReadonlyMap<string, CellData>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<string, CellData> {
  const next = new Map<string, CellData>();
  for (const [key, cell] of cells) {
    const pos = parseCellKey(key);
    if (!pos) {
      next.set(key, cell);
      continue;
    }
    if (pos.col >= start && pos.col <= end) continue;

    const col = pos.col > end ? pos.col - count : pos.col;
    next.set(cellKey(pos.row, col), cell);
  }
  return next;
}

/**
 * 插入行时移动行配置。
 *
 * @param rows - 原始行配置 Map
 * @param insertIndex - 插入行位置
 * @param count - 插入行数量
 * @returns 移动后的行配置 Map
 * @author liangzai927
 */
export function shiftRowsForInsert(
  rows: ReadonlyMap<number, RowData>,
  insertIndex: number,
  count: number,
): ReadonlyMap<number, RowData> {
  const next = new Map<number, RowData>();
  for (const [row, data] of rows) {
    next.set(row >= insertIndex ? row + count : row, data);
  }
  return next;
}

/**
 * 删除行时移动行配置。
 *
 * @param rows - 原始行配置 Map
 * @param start - 删除起始行
 * @param end - 删除结束行
 * @param count - 删除行数量
 * @returns 移动后的行配置 Map
 * @author liangzai927
 */
export function shiftRowsForDelete(
  rows: ReadonlyMap<number, RowData>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<number, RowData> {
  const next = new Map<number, RowData>();
  for (const [row, data] of rows) {
    if (row >= start && row <= end) continue;
    next.set(row > end ? row - count : row, data);
  }
  return next;
}

/**
 * 插入列时移动列配置。
 *
 * @param columns - 原始列配置 Map
 * @param insertIndex - 插入列位置
 * @param count - 插入列数量
 * @returns 移动后的列配置 Map
 * @author liangzai927
 */
export function shiftColumnsForInsert(
  columns: ReadonlyMap<number, ColumnData>,
  insertIndex: number,
  count: number,
): ReadonlyMap<number, ColumnData> {
  const next = new Map<number, ColumnData>();
  for (const [col, data] of columns) {
    next.set(col >= insertIndex ? col + count : col, data);
  }
  return next;
}

/**
 * 删除列时移动列配置。
 *
 * @param columns - 原始列配置 Map
 * @param start - 删除起始列
 * @param end - 删除结束列
 * @param count - 删除列数量
 * @returns 移动后的列配置 Map
 * @author liangzai927
 */
export function shiftColumnsForDelete(
  columns: ReadonlyMap<number, ColumnData>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<number, ColumnData> {
  const next = new Map<number, ColumnData>();
  for (const [col, data] of columns) {
    if (col >= start && col <= end) continue;
    next.set(col > end ? col - count : col, data);
  }
  return next;
}
