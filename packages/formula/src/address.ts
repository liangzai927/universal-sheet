import { createFormulaParseError } from './errors';

const CELL_ADDRESS_PATTERN = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/;
const MIN_EXCEL_ROW_NUMBER = 1;
const MAX_EXCEL_COLUMN_INDEX = 16383;

/** A1 单元格地址。 */
export interface CellAddress {
  readonly row: number;
  readonly col: number;
  readonly absoluteRow: boolean;
  readonly absoluteCol: boolean;
}

/** A1 区域地址。 */
export interface RangeAddress {
  readonly start: CellAddress;
  readonly end: CellAddress;
}

/**
 * 判断文本是否是 A1 单元格地址。
 *
 * @param value - 需要检查的文本
 * @returns 是 A1 单元格地址时返回 true
 * @author liangzai927
 */
export function isCellAddress(value: string): boolean {
  return CELL_ADDRESS_PATTERN.test(value);
}

/**
 * 将 A1 单元格地址解析为零基行列坐标。
 *
 * @param value - A1 单元格地址，例如 A1 或 $B$2
 * @returns 零基行列坐标和绝对引用标记
 * @throws {FormulaError} 当地址格式非法时抛出
 * @author liangzai927
 */
export function parseCellAddress(value: string): CellAddress {
  const match = CELL_ADDRESS_PATTERN.exec(value);
  if (!match) throw createFormulaParseError(`Invalid cell address "${value}"`);

  const absoluteCol = match[1] === '$';
  const colLabel = match[2];
  const absoluteRow = match[3] === '$';
  const rowText = match[4];
  if (colLabel === undefined || rowText === undefined) {
    throw createFormulaParseError(`Invalid cell address "${value}"`);
  }

  const rowNumber = Number(rowText);
  if (!Number.isInteger(rowNumber) || rowNumber < MIN_EXCEL_ROW_NUMBER) {
    throw createFormulaParseError(`Invalid row number "${rowText}"`);
  }

  return {
    row: rowNumber - 1,
    col: columnLabelToIndex(colLabel),
    absoluteRow,
    absoluteCol,
  };
}

/**
 * 将零基行列坐标格式化为 A1 单元格地址。
 *
 * @param address - 零基行列坐标和绝对引用标记
 * @returns A1 单元格地址
 * @throws {FormulaError} 当行列坐标越界时抛出
 * @author liangzai927
 */
export function formatCellAddress(address: CellAddress): string {
  if (!Number.isInteger(address.row) || address.row < 0) {
    throw createFormulaParseError(`Invalid row index "${address.row}"`);
  }
  return `${address.absoluteCol ? '$' : ''}${columnIndexToLabel(address.col)}${
    address.absoluteRow ? '$' : ''
  }${address.row + 1}`;
}

/**
 * 将 A1 区域地址解析为零基范围坐标。
 *
 * @param value - A1 区域地址，例如 A1:B2
 * @returns 零基范围坐标
 * @throws {FormulaError} 当区域格式非法时抛出
 * @author liangzai927
 */
export function parseRangeAddress(value: string): RangeAddress {
  const parts = value.split(':');
  if (parts.length !== 2) throw createFormulaParseError(`Invalid range address "${value}"`);

  const start = parts[0];
  const end = parts[1];
  if (start === undefined || end === undefined) {
    throw createFormulaParseError(`Invalid range address "${value}"`);
  }

  return normalizeRangeAddress({
    start: parseCellAddress(start),
    end: parseCellAddress(end),
  });
}

/**
 * 将零基范围坐标格式化为 A1 区域地址。
 *
 * @param range - 零基范围坐标
 * @returns A1 区域地址
 * @throws {FormulaError} 当行列坐标越界时抛出
 * @author liangzai927
 */
export function formatRangeAddress(range: RangeAddress): string {
  const normalized = normalizeRangeAddress(range);
  return `${formatCellAddress(normalized.start)}:${formatCellAddress(normalized.end)}`;
}

/**
 * 枚举区域内的所有单元格地址。
 *
 * @param range - 零基范围坐标
 * @returns 按行优先顺序排列的单元格地址
 * @author liangzai927
 */
export function enumerateRangeAddress(range: RangeAddress): Array<CellAddress> {
  const normalized = normalizeRangeAddress(range);
  const cells: Array<CellAddress> = [];
  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
      cells.push({
        row,
        col,
        absoluteRow: normalized.start.absoluteRow,
        absoluteCol: normalized.start.absoluteCol,
      });
    }
  }
  return cells;
}

/**
 * 将 Excel 列标转换为零基列索引。
 *
 * @param label - Excel 列标，例如 A 或 AA
 * @returns 零基列索引
 * @throws {FormulaError} 当列标非法时抛出
 * @author liangzai927
 */
export function columnLabelToIndex(label: string): number {
  const normalized = label.toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    throw createFormulaParseError(`Invalid column label "${label}"`);
  }

  let index = 0;
  for (const char of normalized) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }

  const zeroBasedIndex = index - 1;
  if (zeroBasedIndex > MAX_EXCEL_COLUMN_INDEX) {
    throw createFormulaParseError(`Column label out of Excel range "${label}"`);
  }
  return zeroBasedIndex;
}

/**
 * 将零基列索引转换为 Excel 列标。
 *
 * @param index - 零基列索引
 * @returns Excel 列标
 * @throws {FormulaError} 当列索引越界时抛出
 * @author liangzai927
 */
export function columnIndexToLabel(index: number): string {
  if (!Number.isInteger(index) || index < 0 || index > MAX_EXCEL_COLUMN_INDEX) {
    throw createFormulaParseError(`Invalid column index "${index}"`);
  }

  let remaining = index + 1;
  let label = '';
  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}

/**
 * 归一化区域起止坐标，保证 start 在左上角。
 *
 * @param range - 需要归一化的范围坐标
 * @returns 归一化后的范围坐标
 * @author liangzai927
 */
function normalizeRangeAddress(range: RangeAddress): RangeAddress {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endCol = Math.max(range.start.col, range.end.col);

  return {
    start: {
      row: startRow,
      col: startCol,
      absoluteRow: range.start.absoluteRow,
      absoluteCol: range.start.absoluteCol,
    },
    end: {
      row: endRow,
      col: endCol,
      absoluteRow: range.end.absoluteRow,
      absoluteCol: range.end.absoluteCol,
    },
  };
}
