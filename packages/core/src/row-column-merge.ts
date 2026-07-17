import type { CellRange } from './types';
import { cellKey } from './types';

/**
 * 插入行时调整合并区域。
 *
 * @param merges - 原始合并区域 Map
 * @param insertIndex - 插入行位置
 * @param count - 插入行数量
 * @returns 调整后的合并区域 Map
 * @author liangzai927
 */
export function shiftMergesForRowInsert(
  merges: ReadonlyMap<string, CellRange>,
  insertIndex: number,
  count: number,
): ReadonlyMap<string, CellRange> {
  return mapMerges(merges, (range) => {
    if (range.startRow >= insertIndex) {
      return {
        ...range,
        startRow: range.startRow + count,
        endRow: range.endRow + count,
      };
    }
    if (range.endRow >= insertIndex) {
      return { ...range, endRow: range.endRow + count };
    }
    return range;
  });
}

/**
 * 删除行时调整合并区域。
 *
 * @param merges - 原始合并区域 Map
 * @param start - 删除起始行
 * @param end - 删除结束行
 * @param count - 删除行数量
 * @returns 调整后的合并区域 Map
 * @author liangzai927
 */
export function shiftMergesForRowDelete(
  merges: ReadonlyMap<string, CellRange>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<string, CellRange> {
  return mapMerges(merges, (range) => shrinkRangeByDeletedRows(range, start, end, count));
}

/**
 * 插入列时调整合并区域。
 *
 * @param merges - 原始合并区域 Map
 * @param insertIndex - 插入列位置
 * @param count - 插入列数量
 * @returns 调整后的合并区域 Map
 * @author liangzai927
 */
export function shiftMergesForColumnInsert(
  merges: ReadonlyMap<string, CellRange>,
  insertIndex: number,
  count: number,
): ReadonlyMap<string, CellRange> {
  return mapMerges(merges, (range) => {
    if (range.startCol >= insertIndex) {
      return {
        ...range,
        startCol: range.startCol + count,
        endCol: range.endCol + count,
      };
    }
    if (range.endCol >= insertIndex) {
      return { ...range, endCol: range.endCol + count };
    }
    return range;
  });
}

/**
 * 删除列时调整合并区域。
 *
 * @param merges - 原始合并区域 Map
 * @param start - 删除起始列
 * @param end - 删除结束列
 * @param count - 删除列数量
 * @returns 调整后的合并区域 Map
 * @author liangzai927
 */
export function shiftMergesForColumnDelete(
  merges: ReadonlyMap<string, CellRange>,
  start: number,
  end: number,
  count: number,
): ReadonlyMap<string, CellRange> {
  return mapMerges(merges, (range) => shrinkRangeByDeletedColumns(range, start, end, count));
}

/**
 * 映射合并区域并按新锚点重建 Map。
 *
 * @param merges - 原始合并区域 Map
 * @param mapper - 合并区域映射函数
 * @returns 重建后的合并区域 Map
 * @author liangzai927
 */
function mapMerges(
  merges: ReadonlyMap<string, CellRange>,
  mapper: (range: CellRange) => CellRange | null,
): ReadonlyMap<string, CellRange> {
  const next = new Map<string, CellRange>();
  for (const range of merges.values()) {
    const mapped = mapper(range);
    if (!mapped || !isMergeRange(mapped)) continue;
    next.set(cellKey(mapped.startRow, mapped.startCol), mapped);
  }
  return next;
}

/**
 * 根据删除行范围收缩或移动合并区域。
 *
 * @param range - 原始合并区域
 * @param start - 删除起始行
 * @param end - 删除结束行
 * @param count - 删除行数量
 * @returns 调整后的合并区域；完全删除时返回 null
 * @author liangzai927
 */
function shrinkRangeByDeletedRows(
  range: CellRange,
  start: number,
  end: number,
  count: number,
): CellRange | null {
  if (range.endRow < start) return range;
  if (range.startRow > end) {
    return {
      ...range,
      startRow: range.startRow - count,
      endRow: range.endRow - count,
    };
  }
  if (range.startRow >= start && range.endRow <= end) return null;
  if (range.startRow < start && range.endRow > end) {
    return { ...range, endRow: range.endRow - count };
  }
  if (range.startRow < start) {
    return { ...range, endRow: start - 1 };
  }
  return {
    ...range,
    startRow: start,
    endRow: range.endRow - count,
  };
}

/**
 * 根据删除列范围收缩或移动合并区域。
 *
 * @param range - 原始合并区域
 * @param start - 删除起始列
 * @param end - 删除结束列
 * @param count - 删除列数量
 * @returns 调整后的合并区域；完全删除时返回 null
 * @author liangzai927
 */
function shrinkRangeByDeletedColumns(
  range: CellRange,
  start: number,
  end: number,
  count: number,
): CellRange | null {
  if (range.endCol < start) return range;
  if (range.startCol > end) {
    return {
      ...range,
      startCol: range.startCol - count,
      endCol: range.endCol - count,
    };
  }
  if (range.startCol >= start && range.endCol <= end) return null;
  if (range.startCol < start && range.endCol > end) {
    return { ...range, endCol: range.endCol - count };
  }
  if (range.startCol < start) {
    return { ...range, endCol: start - 1 };
  }
  return {
    ...range,
    startCol: start,
    endCol: range.endCol - count,
  };
}

/**
 * 判断范围是否仍然是有效合并区域。
 *
 * @param range - 需要检查的单元格范围
 * @returns 当范围跨越至少两格时返回 true
 * @author liangzai927
 */
function isMergeRange(range: CellRange): boolean {
  return (
    range.startRow <= range.endRow &&
    range.startCol <= range.endCol &&
    (range.startRow !== range.endRow || range.startCol !== range.endCol)
  );
}
