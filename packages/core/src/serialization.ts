import type {
  CellData,
  CellRange,
  CellStyle,
  ColumnData,
  RowData,
  SheetConfig,
  SheetData,
} from './types';
import type { WorkbookData, WorkbookSheet } from './workbook-model';

/** 用于序列化 Map 数据的 JSON 安全键值条目。 */
export interface SerializedMapEntry<TKey extends string | number, TValue> {
  readonly key: TKey;
  readonly value: TValue;
}

/** SheetData 的 JSON 安全结构。 */
export interface SerializedSheetData {
  readonly config: SheetConfig;
  readonly cells: Array<SerializedMapEntry<string, CellData>>;
  readonly columns: Array<SerializedMapEntry<number, ColumnData>>;
  readonly rows: Array<SerializedMapEntry<number, RowData>>;
  readonly merges: Array<SerializedMapEntry<string, CellRange>>;
}

/** 工作簿工作表条目的 JSON 安全结构。 */
export interface SerializedWorkbookSheet {
  readonly id: string;
  readonly name: string;
  readonly sheet: SerializedSheetData;
}

/** WorkbookData 的 JSON 安全结构。 */
export interface SerializedWorkbookData {
  readonly version: 1;
  readonly activeSheetId: string;
  readonly sheets: Array<SerializedWorkbookSheet>;
}

/**
 * 将 SheetData 序列化为 JSON 安全对象。
 *
 * @param sheet - 需要序列化的工作表数据
 * @returns JSON 安全的工作表对象
 * @author liangzai927
 */
export function serializeSheetData(sheet: SheetData): SerializedSheetData {
  return {
    config: copySheetConfig(sheet.config),
    cells: Array.from(sheet.cells, ([key, value]) => ({ key, value: copyCellData(value) })),
    columns: Array.from(sheet.columns, ([key, value]) => ({ key, value: copyColumnData(value) })),
    rows: Array.from(sheet.rows, ([key, value]) => ({ key, value: copyRowData(value) })),
    merges: Array.from(sheet.merges, ([key, value]) => ({ key, value: copyCellRange(value) })),
  };
}

/**
 * 将 JSON 安全工作表对象反序列化为 SheetData。
 *
 * @param data - 已序列化的工作表数据
 * @returns 恢复 Map 实例后的 SheetData 对象
 * @author liangzai927
 */
export function deserializeSheetData(data: SerializedSheetData): SheetData {
  return {
    config: copySheetConfig(data.config),
    cells: new Map(data.cells.map((entry) => [entry.key, copyCellData(entry.value)])),
    columns: new Map(data.columns.map((entry) => [entry.key, copyColumnData(entry.value)])),
    rows: new Map(data.rows.map((entry) => [entry.key, copyRowData(entry.value)])),
    merges: new Map(data.merges.map((entry) => [entry.key, copyCellRange(entry.value)])),
  };
}

/**
 * 将 WorkbookData 序列化为 JSON 安全对象。
 *
 * @param workbook - 需要序列化的工作簿数据
 * @returns JSON 安全的工作簿对象
 * @author liangzai927
 */
export function serializeWorkbookData(workbook: WorkbookData): SerializedWorkbookData {
  return {
    version: 1,
    activeSheetId: workbook.activeSheetId,
    sheets: workbook.sheets.map((sheet) => serializeWorkbookSheet(sheet)),
  };
}

/**
 * 将 JSON 安全工作簿对象反序列化为 WorkbookData。
 *
 * @param data - 已序列化的工作簿数据
 * @returns 恢复工作表 Map 数据后的 WorkbookData 对象
 * @author liangzai927
 */
export function deserializeWorkbookData(data: SerializedWorkbookData): WorkbookData {
  return {
    activeSheetId: data.activeSheetId,
    sheets: data.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      sheet: deserializeSheetData(sheet.sheet),
    })),
  };
}

/**
 * 将 WorkbookData 转换为 JSON 字符串。
 *
 * @param workbook - 需要转换的工作簿数据
 * @returns 包含序列化工作簿的 JSON 字符串
 * @author liangzai927
 */
export function stringifyWorkbookData(workbook: WorkbookData): string {
  return JSON.stringify(serializeWorkbookData(workbook));
}

/**
 * 将 JSON 字符串解析为 WorkbookData。
 *
 * @param json - 由 stringifyWorkbookData 创建的 JSON 字符串
 * @returns 解析后的工作簿数据
 * @throws {SyntaxError} 当 JSON 字符串格式非法时抛出
 * @author liangzai927
 */
export function parseWorkbookData(json: string): WorkbookData {
  const parsed: unknown = JSON.parse(json);
  return deserializeWorkbookData(parsed as SerializedWorkbookData);
}

/**
 * 序列化单个工作簿工作表条目。
 *
 * @param sheet - 需要序列化的工作簿工作表条目
 * @returns JSON 安全的工作簿工作表对象
 * @author liangzai927
 */
function serializeWorkbookSheet(sheet: WorkbookSheet): SerializedWorkbookSheet {
  return {
    id: sheet.id,
    name: sheet.name,
    sheet: serializeSheetData(sheet.sheet),
  };
}

/**
 * 复制工作表配置对象。
 *
 * @param config - 需要复制的工作表配置
 * @returns 复制后的工作表配置
 * @author liangzai927
 */
function copySheetConfig(config: SheetConfig): SheetConfig {
  return { ...config };
}

/**
 * 复制单元格数据对象。
 *
 * @param cell - 需要复制的单元格数据
 * @returns 复制后的单元格数据对象
 * @author liangzai927
 */
function copyCellData(cell: CellData): CellData {
  return {
    value: cell.value,
    displayValue: cell.displayValue,
    formula: cell.formula,
    style: cell.style ? copyCellStyle(cell.style) : undefined,
  };
}

/**
 * 复制单元格样式对象。
 *
 * @param style - 需要复制的单元格样式
 * @returns 复制后的单元格样式对象
 * @author liangzai927
 */
function copyCellStyle(style: CellStyle): CellStyle {
  return { ...style };
}

/**
 * 复制列数据对象。
 *
 * @param column - 需要复制的列数据
 * @returns 复制后的列数据对象
 * @author liangzai927
 */
function copyColumnData(column: ColumnData): ColumnData {
  return { ...column };
}

/**
 * 复制行数据对象。
 *
 * @param row - 需要复制的行数据
 * @returns 复制后的行数据对象
 * @author liangzai927
 */
function copyRowData(row: RowData): RowData {
  return { ...row };
}

/**
 * 复制单元格范围对象。
 *
 * @param range - 需要复制的单元格范围
 * @returns 复制后的单元格范围对象
 * @author liangzai927
 */
function copyCellRange(range: CellRange): CellRange {
  return { ...range };
}
