/** 单元格键的行列位置。 */
export interface CellKeyPosition {
  readonly row: number;
  readonly col: number;
}

const MIN_GRID_SIZE = 1;
const CELL_KEY_PATTERN = /^R(\d+)C(\d+)$/;

/**
 * 将插入位置裁剪到有效范围。
 *
 * @param index - 原始插入位置
 * @param size - 当前行数或列数
 * @returns 裁剪后的插入位置
 * @author liangzai927
 */
export function clampInsertIndex(index: number, size: number): number {
  return Math.max(0, Math.min(index, size));
}

/**
 * 规范化删除范围，并保证至少保留一行或一列。
 *
 * @param index - 删除起始位置
 * @param count - 删除数量
 * @param size - 当前行数或列数
 * @returns 规范化后的删除范围；无效时返回 null
 * @author liangzai927
 */
export function normalizeDeleteRange(
  index: number,
  count: number,
  size: number,
): { readonly start: number; readonly end: number; readonly count: number } | null {
  if (count <= 0 || size <= MIN_GRID_SIZE || index >= size) return null;

  const start = Math.max(0, index);
  const maxDeleteCount = size - MIN_GRID_SIZE;
  const actualCount = Math.min(count, maxDeleteCount, size - start);
  if (actualCount <= 0) return null;

  return {
    start,
    end: start + actualCount - 1,
    count: actualCount,
  };
}

/**
 * 规范化受边界限制的范围。
 *
 * @param index - 范围起始位置
 * @param count - 范围长度
 * @param size - 当前行数或列数
 * @returns 规范化后的范围；无效时返回 null
 * @author liangzai927
 */
export function normalizeBoundedRange(
  index: number,
  count: number,
  size: number,
): { readonly start: number; readonly end: number } | null {
  if (count <= 0 || size <= 0 || index >= size) return null;

  const start = Math.max(0, index);
  const end = Math.min(size - 1, start + count - 1);
  return start <= end ? { start, end } : null;
}

/**
 * 解析单元格 Map 的键。
 *
 * @param key - 单元格键
 * @returns 解析后的行列位置；格式非法时返回 null
 * @author liangzai927
 */
export function parseCellKey(key: string): CellKeyPosition | null {
  const match = CELL_KEY_PATTERN.exec(key);
  if (!match) return null;

  const rowText = match[1];
  const colText = match[2];
  if (rowText === undefined || colText === undefined) return null;

  return {
    row: Number.parseInt(rowText, 10),
    col: Number.parseInt(colText, 10),
  };
}
