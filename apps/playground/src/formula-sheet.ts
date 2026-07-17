import type { CellData, ClipboardPayload, SerializedCell, SheetData } from '@universal-sheet/core';
import {
  cellKey,
  deleteColumns,
  deleteRows,
  insertColumns,
  insertRows,
  pasteCellRangeStructured,
} from '@universal-sheet/core';
import type {
  FormulaCellInput,
  FormulaPrimitive,
  FormulaReferenceChange,
} from '@universal-sheet/formula';
import {
  applyFormulaReferenceChange,
  calculateFormulaCells,
  formatCellAddress,
  normalizeFormulaText,
  offsetFormulaReferences,
} from '@universal-sheet/formula';

const CELL_KEY_PATTERN = /^R(\d+)C(\d+)$/;

interface NormalizedDeleteRange {
  readonly startIndex: number;
  readonly count: number;
}

/**
 * 获取单元格编辑时应该展示的原始文本。
 *
 * @param cell - 单元格数据
 * @returns 公式原文或普通单元格值文本
 * @author liangzai927
 */
export function getCellInputText(cell: CellData | null): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  return cell.value != null ? String(cell.value) : '';
}

/**
 * 将用户输入写入工作表并重算公式。
 *
 * @param sheet - 原始工作表数据
 * @param row - 单元格行索引
 * @param col - 单元格列索引
 * @param input - 用户输入文本
 * @returns 更新并完成公式重算后的工作表数据
 * @author liangzai927
 */
export function applyCellInput(
  sheet: SheetData,
  row: number,
  col: number,
  input: string,
): SheetData {
  const updated = writeCellInput(sheet, row, col, input);
  return recalculateSheetFormulas(updated);
}

/**
 * 粘贴结构化单元格并按目标位置偏移公式相对引用。
 *
 * @param sheet - 原始工作表数据
 * @param startRow - 粘贴起始行
 * @param startCol - 粘贴起始列
 * @param payload - 结构化剪贴板数据
 * @returns 粘贴并重算公式后的工作表数据
 * @author liangzai927
 */
export function pasteFormulaAwareStructured(
  sheet: SheetData,
  startRow: number,
  startCol: number,
  payload: ClipboardPayload,
): SheetData {
  const adjustedPayload: ClipboardPayload = {
    ...payload,
    cells: payload.cells.map((cell) => offsetClipboardFormula(cell, startRow, startCol, payload)),
  };
  return recalculateSheetFormulas(
    pasteCellRangeStructured(sheet, startRow, startCol, adjustedPayload),
  );
}

/**
 * 插入行并同步更新公式引用。
 *
 * @param sheet - 原始工作表数据
 * @param rowIndex - 插入行起始索引
 * @param count - 插入行数量
 * @returns 插入行并重算公式后的工作表数据
 * @author liangzai927
 */
export function insertRowsWithFormulas(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  if (count <= 0) return sheet;

  const startIndex = clampInsertIndex(rowIndex, sheet.config.rowCount);
  return applyFormulaAwareReferenceChange(insertRows(sheet, startIndex, count), {
    type: 'insertRows',
    startIndex,
    count,
  });
}

/**
 * 删除行并同步更新公式引用。
 *
 * @param sheet - 原始工作表数据
 * @param rowIndex - 删除行起始索引
 * @param count - 删除行数量
 * @returns 删除行并重算公式后的工作表数据
 * @author liangzai927
 */
export function deleteRowsWithFormulas(sheet: SheetData, rowIndex: number, count = 1): SheetData {
  const deletion = normalizeDeleteRange(rowIndex, count, sheet.config.rowCount);
  if (!deletion) return sheet;

  return applyFormulaAwareReferenceChange(deleteRows(sheet, deletion.startIndex, deletion.count), {
    type: 'deleteRows',
    startIndex: deletion.startIndex,
    count: deletion.count,
  });
}

/**
 * 插入列并同步更新公式引用。
 *
 * @param sheet - 原始工作表数据
 * @param colIndex - 插入列起始索引
 * @param count - 插入列数量
 * @returns 插入列并重算公式后的工作表数据
 * @author liangzai927
 */
export function insertColumnsWithFormulas(
  sheet: SheetData,
  colIndex: number,
  count = 1,
): SheetData {
  if (count <= 0) return sheet;

  const startIndex = clampInsertIndex(colIndex, sheet.config.colCount);
  return applyFormulaAwareReferenceChange(insertColumns(sheet, startIndex, count), {
    type: 'insertColumns',
    startIndex,
    count,
  });
}

/**
 * 删除列并同步更新公式引用。
 *
 * @param sheet - 原始工作表数据
 * @param colIndex - 删除列起始索引
 * @param count - 删除列数量
 * @returns 删除列并重算公式后的工作表数据
 * @author liangzai927
 */
export function deleteColumnsWithFormulas(
  sheet: SheetData,
  colIndex: number,
  count = 1,
): SheetData {
  const deletion = normalizeDeleteRange(colIndex, count, sheet.config.colCount);
  if (!deletion) return sheet;

  return applyFormulaAwareReferenceChange(
    deleteColumns(sheet, deletion.startIndex, deletion.count),
    {
      type: 'deleteColumns',
      startIndex: deletion.startIndex,
      count: deletion.count,
    },
  );
}

/**
 * 按结构变更更新工作表内所有公式引用。
 *
 * @param sheet - 已完成结构移动的工作表数据
 * @param change - 公式引用结构变更描述
 * @returns 更新公式引用并重算后的工作表数据
 * @author liangzai927
 */
export function applyFormulaAwareReferenceChange(
  sheet: SheetData,
  change: FormulaReferenceChange,
): SheetData {
  const cells = new Map(sheet.cells);

  for (const [key, cell] of sheet.cells) {
    if (!cell.formula) continue;

    const formula = safeApplyFormulaReferenceChange(cell.formula, change);
    cells.set(key, {
      ...cell,
      formula,
      value: null,
      displayValue: formula,
    });
  }

  return recalculateSheetFormulas({ ...sheet, cells });
}

/**
 * 将用户输入写入单元格。
 *
 * @param sheet - 原始工作表数据
 * @param row - 单元格行索引
 * @param col - 单元格列索引
 * @param input - 用户输入文本
 * @returns 写入后的工作表数据
 * @author liangzai927
 */
function writeCellInput(sheet: SheetData, row: number, col: number, input: string): SheetData {
  const key = cellKey(row, col);
  const cells = new Map(sheet.cells);
  const existing = cells.get(key);

  if (input.length === 0) {
    if (existing?.style) {
      cells.set(key, { value: null, style: existing.style });
    } else {
      cells.delete(key);
    }
    return { ...sheet, cells };
  }

  cells.set(
    key,
    input.startsWith('=') ? createFormulaCell(existing, input) : createValueCell(existing, input),
  );
  return { ...sheet, cells };
}

/**
 * 创建公式单元格数据。
 *
 * @param existing - 原有单元格数据
 * @param input - 用户输入的公式文本
 * @returns 公式单元格数据
 * @author liangzai927
 */
function createFormulaCell(existing: CellData | undefined, input: string): CellData {
  const formula = safeNormalizeFormulaText(input);
  return {
    value: null,
    displayValue: formula,
    formula,
    style: existing?.style,
  };
}

/**
 * 安全归一化公式文本，语法不完整时保留用户原始输入。
 *
 * @param input - 用户输入的公式文本
 * @returns 可保存到单元格的公式文本
 * @author liangzai927
 */
function safeNormalizeFormulaText(input: string): string {
  try {
    return normalizeFormulaText(input);
  } catch {
    return input.startsWith('=') ? input : `=${input}`;
  }
}

/**
 * 按粘贴目标位置偏移剪贴板单元格公式。
 *
 * @param cell - 剪贴板单元格
 * @param startRow - 粘贴起始行
 * @param startCol - 粘贴起始列
 * @param payload - 结构化剪贴板数据
 * @returns 偏移公式后的剪贴板单元格
 * @author liangzai927
 */
function offsetClipboardFormula(
  cell: SerializedCell,
  startRow: number,
  startCol: number,
  payload: ClipboardPayload,
): SerializedCell {
  if (!cell.formula) return cell;

  const sourceRow =
    payload.sourceStartRow === undefined ? startRow + cell.r : payload.sourceStartRow + cell.r;
  const sourceCol =
    payload.sourceStartCol === undefined ? startCol + cell.c : payload.sourceStartCol + cell.c;
  const targetRow = startRow + cell.r;
  const targetCol = startCol + cell.c;

  try {
    const formula = offsetFormulaReferences(cell.formula, {
      rowOffset: targetRow - sourceRow,
      colOffset: targetCol - sourceCol,
    });
    return {
      ...cell,
      formula,
      displayValue: formula,
    };
  } catch {
    return cell;
  }
}

/**
 * 安全应用公式引用变更，失败时保留原公式。
 *
 * @param formula - 原始公式文本
 * @param change - 公式引用结构变更描述
 * @returns 更新后的公式文本
 * @author liangzai927
 */
function safeApplyFormulaReferenceChange(formula: string, change: FormulaReferenceChange): string {
  try {
    return applyFormulaReferenceChange(formula, change);
  } catch {
    return formula;
  }
}

/**
 * 将插入索引限制在可用范围内。
 *
 * @param index - 原始插入索引
 * @param total - 当前总行数或总列数
 * @returns 可用于结构变更的插入索引
 * @author liangzai927
 */
function clampInsertIndex(index: number, total: number): number {
  return Math.max(0, Math.min(index, total));
}

/**
 * 归一化删除范围。
 *
 * @param startIndex - 原始删除起始索引
 * @param count - 删除数量
 * @param total - 当前总行数或总列数
 * @returns 可用于结构变更的删除范围；无效时返回 null
 * @author liangzai927
 */
function normalizeDeleteRange(
  startIndex: number,
  count: number,
  total: number,
): NormalizedDeleteRange | null {
  if (count <= 0 || total <= 0) return null;

  const normalizedStart = Math.max(0, Math.min(startIndex, total - 1));
  const endIndex = Math.min(total - 1, normalizedStart + count - 1);
  const normalizedCount = endIndex - normalizedStart + 1;
  return normalizedCount > 0 ? { startIndex: normalizedStart, count: normalizedCount } : null;
}

/**
 * 创建普通值单元格数据。
 *
 * @param existing - 原有单元格数据
 * @param input - 用户输入文本
 * @returns 普通值单元格数据
 * @author liangzai927
 */
function createValueCell(existing: CellData | undefined, input: string): CellData {
  return {
    value: input,
    style: existing?.style,
  };
}

/**
 * 重算工作表中的所有公式单元格。
 *
 * @param sheet - 原始工作表数据
 * @returns 完成公式重算后的工作表数据
 * @author liangzai927
 */
export function recalculateSheetFormulas(sheet: SheetData): SheetData {
  const formulaInputs = sheetToFormulaInputs(sheet);
  if (!formulaInputs.some((cell) => Boolean(cell.formula))) return sheet;

  const result = calculateFormulaCells(formulaInputs);
  const cells = new Map(sheet.cells);

  for (const [key, cell] of sheet.cells) {
    if (!cell.formula) continue;
    const ref = cellKeyToFormulaRef(key);
    if (!ref) continue;

    const error = result.errors.get(ref);
    const value = error ? error.code : (result.values.get(ref) ?? null);
    cells.set(key, {
      ...cell,
      value,
      displayValue: formatFormulaDisplayValue(value),
    });
  }

  return { ...sheet, cells };
}

/**
 * 将工作表单元格转换为公式计算输入。
 *
 * @param sheet - 原始工作表数据
 * @returns 公式计算输入列表
 * @author liangzai927
 */
function sheetToFormulaInputs(sheet: SheetData): Array<FormulaCellInput> {
  const inputs: Array<FormulaCellInput> = [];
  for (const [key, cell] of sheet.cells) {
    const ref = cellKeyToFormulaRef(key);
    if (!ref) continue;
    inputs.push({
      ref,
      value: coerceFormulaValue(cell.value),
      formula: cell.formula,
    });
  }
  return inputs;
}

/**
 * 将 core 单元格 key 转换为公式 A1 引用。
 *
 * @param key - core 单元格 key
 * @returns A1 单元格引用；无法解析时返回 null
 * @author liangzai927
 */
function cellKeyToFormulaRef(key: string): string | null {
  const match = CELL_KEY_PATTERN.exec(key);
  const rowText = match?.[1];
  const colText = match?.[2];
  if (!rowText || !colText) return null;
  return formatCellAddress({
    row: Number(rowText),
    col: Number(colText),
    absoluteRow: false,
    absoluteCol: false,
  });
}

/**
 * 将单元格值转换为公式计算值。
 *
 * @param value - 单元格原始值
 * @returns 公式计算值
 * @author liangzai927
 */
function coerceFormulaValue(value: CellData['value']): FormulaPrimitive {
  return value;
}

/**
 * 格式化公式显示值。
 *
 * @param value - 公式计算结果
 * @returns 公式显示文本
 * @author liangzai927
 */
function formatFormulaDisplayValue(value: FormulaPrimitive): string | undefined {
  return value != null ? String(value) : undefined;
}
