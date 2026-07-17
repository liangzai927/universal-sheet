import { createSheetData } from './sheet-model';
import type { SheetConfig, SheetData } from './types';

/** 工作簿内工作表的稳定唯一标识。 */
export type SheetId = string;

/** 工作簿中保存的工作表条目。 */
export interface WorkbookSheet {
  readonly id: SheetId;
  readonly name: string;
  readonly sheet: SheetData;
}

/** 包含一个或多个工作表的完整工作簿状态。 */
export interface WorkbookData {
  readonly sheets: ReadonlyArray<WorkbookSheet>;
  readonly activeSheetId: SheetId;
}

/** 创建初始工作表时使用的输入数据。 */
export interface WorkbookSheetInput {
  readonly id?: SheetId;
  readonly name?: string;
  readonly sheet?: SheetData;
}

/** 创建工作簿时使用的配置项。 */
export interface CreateWorkbookOptions {
  readonly sheets?: ReadonlyArray<WorkbookSheetInput>;
  readonly activeSheetId?: SheetId;
  readonly sheetConfig?: Partial<SheetConfig>;
}

/** 向已有工作簿添加工作表时使用的配置项。 */
export interface AddWorkbookSheetOptions {
  readonly id?: SheetId;
  readonly name?: string;
  readonly sheet?: SheetData;
  readonly makeActive?: boolean;
}

const DEFAULT_SHEET_ID = 'sheet-1';
const DEFAULT_SHEET_NAME = 'Sheet1';

/**
 * 创建至少包含一个工作表的工作簿。
 *
 * @param options - 初始工作表列表、激活工作表 ID 和默认工作表配置
 * @returns 新的工作簿状态
 * @author liangzai927
 */
export function createWorkbookData(options?: CreateWorkbookOptions): WorkbookData {
  const sourceSheets =
    options?.sheets && options.sheets.length > 0
      ? options.sheets
      : [{ id: DEFAULT_SHEET_ID, name: DEFAULT_SHEET_NAME }];

  const usedIds = new Set<SheetId>();
  const usedNames = new Set<string>();
  const sheets = sourceSheets.map((input, index) =>
    normalizeSheetInput(input, index, usedIds, usedNames, options?.sheetConfig),
  );

  const requestedActiveSheetId = options?.activeSheetId;
  const hasRequestedActiveSheet =
    requestedActiveSheetId !== undefined &&
    sheets.some((sheet) => sheet.id === requestedActiveSheetId);
  const firstSheetId = sheets[0]?.id ?? DEFAULT_SHEET_ID;

  return {
    sheets,
    activeSheetId: hasRequestedActiveSheet ? requestedActiveSheetId : firstSheetId,
  };
}

/**
 * 获取当前激活的工作表条目。
 *
 * @param workbook - 需要读取的工作簿
 * @returns 激活的工作表条目；当工作簿状态不一致时返回 null
 * @author liangzai927
 */
export function getActiveWorkbookSheet(workbook: WorkbookData): WorkbookSheet | null {
  return getWorkbookSheet(workbook, workbook.activeSheetId);
}

/**
 * 根据工作表 ID 查找工作表。
 *
 * @param workbook - 需要读取的工作簿
 * @param sheetId - 需要查找的稳定工作表 ID
 * @returns 匹配的工作表条目；不存在时返回 null
 * @author liangzai927
 */
export function getWorkbookSheet(workbook: WorkbookData, sheetId: SheetId): WorkbookSheet | null {
  return workbook.sheets.find((sheet) => sheet.id === sheetId) ?? null;
}

/**
 * 根据工作表名称查找工作表。
 *
 * @param workbook - 需要读取的工作簿
 * @param name - 需要查找的工作表名称
 * @returns 匹配的工作表条目；不存在时返回 null
 * @author liangzai927
 */
export function getWorkbookSheetByName(workbook: WorkbookData, name: string): WorkbookSheet | null {
  return workbook.sheets.find((sheet) => sheet.name === name) ?? null;
}

/**
 * 根据工作表 ID 切换当前激活工作表。
 *
 * @param workbook - 需要更新的工作簿
 * @param sheetId - 需要激活的工作表 ID
 * @returns 当工作表存在时返回新的工作簿；否则返回原工作簿
 * @author liangzai927
 */
export function setActiveWorkbookSheet(workbook: WorkbookData, sheetId: SheetId): WorkbookData {
  if (!workbook.sheets.some((sheet) => sheet.id === sheetId)) return workbook;
  if (workbook.activeSheetId === sheetId) return workbook;
  return { ...workbook, activeSheetId: sheetId };
}

/**
 * 替换指定工作表的数据内容。
 *
 * @param workbook - 需要更新的工作簿
 * @param sheetId - 需要替换数据的工作表 ID
 * @param sheetData - 新的工作表数据
 * @returns 当工作表存在时返回新的工作簿；否则返回原工作簿
 * @author liangzai927
 */
export function updateWorkbookSheet(
  workbook: WorkbookData,
  sheetId: SheetId,
  sheetData: SheetData,
): WorkbookData {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (sheetIndex < 0) return workbook;

  const sheets = [...workbook.sheets];
  const existing = sheets[sheetIndex];
  if (!existing) return workbook;
  sheets[sheetIndex] = { ...existing, sheet: sheetData };

  return { ...workbook, sheets };
}

/**
 * 向工作簿追加一个新的工作表。
 *
 * @param workbook - 需要更新的工作簿
 * @param options - 新工作表的 ID、名称、数据和是否激活
 * @returns 包含新增工作表的新工作簿
 * @author liangzai927
 */
export function addWorkbookSheet(
  workbook: WorkbookData,
  options?: AddWorkbookSheetOptions,
): WorkbookData {
  const usedIds = new Set(workbook.sheets.map((sheet) => sheet.id));
  const usedNames = new Set(workbook.sheets.map((sheet) => sheet.name));
  const index = workbook.sheets.length;
  const sheet: WorkbookSheet = {
    id: uniqueSheetId(options?.id ?? `sheet-${index + 1}`, usedIds),
    name: uniqueSheetName(options?.name ?? `Sheet${index + 1}`, usedNames),
    sheet: options?.sheet ?? createSheetData(),
  };
  const sheets = [...workbook.sheets, sheet];
  const makeActive = options?.makeActive ?? true;

  return {
    sheets,
    activeSheetId: makeActive ? sheet.id : workbook.activeSheetId,
  };
}

/**
 * 从工作簿中移除指定工作表。
 *
 * @param workbook - 需要更新的工作簿
 * @param sheetId - 需要移除的工作表 ID
 * @returns 当工作表可移除时返回新的工作簿；否则返回原工作簿
 * @author liangzai927
 */
export function removeWorkbookSheet(workbook: WorkbookData, sheetId: SheetId): WorkbookData {
  if (workbook.sheets.length <= 1) return workbook;

  const removeIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (removeIndex < 0) return workbook;

  const sheets = workbook.sheets.filter((sheet) => sheet.id !== sheetId);
  const replacementIndex = Math.min(removeIndex, sheets.length - 1);
  const replacementSheetId = sheets[replacementIndex]?.id ?? workbook.activeSheetId;

  return {
    sheets,
    activeSheetId: workbook.activeSheetId === sheetId ? replacementSheetId : workbook.activeSheetId,
  };
}

/**
 * 重命名已有工作表。
 *
 * @param workbook - 需要更新的工作簿
 * @param sheetId - 需要重命名的工作表 ID
 * @param name - 新工作表名称
 * @returns 当名称有效时返回新的工作簿；否则返回原工作簿
 * @author liangzai927
 */
export function renameWorkbookSheet(
  workbook: WorkbookData,
  sheetId: SheetId,
  name: string,
): WorkbookData {
  const nextName = name.trim();
  if (!nextName || hasSheetName(workbook, nextName, sheetId)) return workbook;

  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (sheetIndex < 0) return workbook;

  const sheets = [...workbook.sheets];
  const existing = sheets[sheetIndex];
  if (!existing) return workbook;
  sheets[sheetIndex] = { ...existing, name: nextName };

  return { ...workbook, sheets };
}

/**
 * 将工作表移动到目标位置。
 *
 * @param workbook - 需要更新的工作簿
 * @param sheetId - 需要移动的工作表 ID
 * @param targetIndex - 目标位置索引，超出范围时会被裁剪
 * @returns 当工作表存在时返回新的工作簿；否则返回原工作簿
 * @author liangzai927
 */
export function moveWorkbookSheet(
  workbook: WorkbookData,
  sheetId: SheetId,
  targetIndex: number,
): WorkbookData {
  const currentIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (currentIndex < 0) return workbook;

  const nextSheets = [...workbook.sheets];
  const [sheet] = nextSheets.splice(currentIndex, 1);
  if (!sheet) return workbook;

  const clampedIndex = Math.max(0, Math.min(targetIndex, nextSheets.length));
  nextSheets.splice(clampedIndex, 0, sheet);

  return { ...workbook, sheets: nextSheets };
}

/**
 * 将调用方传入的工作表数据规范化为工作簿工作表条目。
 *
 * @param input - 局部工作表输入
 * @param index - 从 0 开始的工作表位置
 * @param usedIds - 当前工作簿中已经分配的 ID 集合
 * @param usedNames - 当前工作簿中已经分配的名称集合
 * @param sheetConfig - 新建 SheetData 时使用的默认配置
 * @returns 规范化后的工作表条目
 * @author liangzai927
 */
function normalizeSheetInput(
  input: WorkbookSheetInput,
  index: number,
  usedIds: Set<SheetId>,
  usedNames: Set<string>,
  sheetConfig?: Partial<SheetConfig>,
): WorkbookSheet {
  const id = uniqueSheetId(input.id ?? `sheet-${index + 1}`, usedIds);
  const name = uniqueSheetName(input.name ?? `Sheet${index + 1}`, usedNames);
  const sheet = input.sheet ?? createSheetData(sheetConfig);

  return { id, name, sheet };
}

/**
 * 根据期望 ID 生成工作簿内唯一的工作表 ID。
 *
 * @param preferredId - 期望使用的工作表 ID
 * @param usedIds - 当前工作簿中已经分配的 ID 集合
 * @returns 唯一的工作表 ID
 * @author liangzai927
 */
function uniqueSheetId(preferredId: SheetId, usedIds: Set<SheetId>): SheetId {
  const baseId = preferredId.trim() || DEFAULT_SHEET_ID;
  let candidate = baseId;
  let suffix = 2;
  while (usedIds.has(candidate)) {
    candidate = `${baseId}-${suffix}`;
    suffix += 1;
  }
  usedIds.add(candidate);
  return candidate;
}

/**
 * 根据期望名称生成工作簿内唯一的工作表名称。
 *
 * @param preferredName - 期望使用的显示名称
 * @param usedNames - 当前工作簿中已经分配的名称集合
 * @returns 唯一的工作表名称
 * @author liangzai927
 */
function uniqueSheetName(preferredName: string, usedNames: Set<string>): string {
  const baseName = preferredName.trim() || DEFAULT_SHEET_NAME;
  let candidate = baseName;
  let suffix = 2;
  while (usedNames.has(candidate)) {
    candidate = `${baseName} (${suffix})`;
    suffix += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

/**
 * 检查工作簿中是否已经存在指定工作表名称。
 *
 * @param workbook - 需要检查的工作簿
 * @param name - 候选工作表名称
 * @param exceptSheetId - 允许保留同名的工作表 ID
 * @returns 当其他工作表已使用该名称时返回 true
 * @author liangzai927
 */
function hasSheetName(workbook: WorkbookData, name: string, exceptSheetId: SheetId): boolean {
  return workbook.sheets.some((sheet) => sheet.id !== exceptSheetId && sheet.name === name);
}
