import type { FormulaNode, FormulaPrimitive } from './ast';
import { createFormulaParseError } from './errors';
import { parseFormula } from './parser';

/** 创建公式记录时的配置。 */
export interface CreateFormulaRecordOptions {
  readonly cachedValue?: FormulaPrimitive;
  readonly parse?: boolean;
}

/** 可用于导入导出的公式记录。 */
export interface FormulaRecord {
  readonly formula: string;
  readonly excelFormula: string;
  readonly ast?: FormulaNode;
  readonly cachedValue?: FormulaPrimitive;
}

/**
 * 创建公式记录。
 *
 * @param input - 原始公式文本，可以带等号也可以不带等号
 * @param options - 创建公式记录时的配置
 * @returns 公式记录，包含应用内公式文本和 Excel 公式文本
 * @throws {FormulaError} 当公式为空或需要解析但语法非法时抛出
 * @author liangzai927
 */
export function createFormulaRecord(
  input: string,
  options: CreateFormulaRecordOptions = {},
): FormulaRecord {
  const formula = normalizeFormulaText(input);
  const record: FormulaRecord = {
    formula,
    excelFormula: toExcelFormulaText(formula),
  };

  const shouldParse = options.parse ?? true;
  const withAst = shouldParse ? { ...record, ast: parseFormula(formula) } : record;
  if (!Object.hasOwn(options, 'cachedValue')) return withAst;

  return {
    ...withAst,
    cachedValue: options.cachedValue,
  };
}

/**
 * 将任意公式文本归一化为应用内公式格式。
 *
 * @param input - 原始公式文本，可以带等号也可以不带等号
 * @returns 带单个等号前缀的公式文本
 * @throws {FormulaError} 当公式为空时抛出
 * @author liangzai927
 */
export function normalizeFormulaText(input: string): string {
  const trimmed = input.trim();
  const formula = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
  if (formula.length <= 1) {
    throw createFormulaParseError('Formula text cannot be empty');
  }
  return formula;
}

/**
 * 将 Excel 存储的公式文本转换为应用内公式文本。
 *
 * @param input - Excel 公式文本，通常不带等号
 * @returns 带等号前缀的应用内公式文本
 * @throws {FormulaError} 当公式为空时抛出
 * @author liangzai927
 */
export function fromExcelFormulaText(input: string): string {
  return normalizeFormulaText(input);
}

/**
 * 将应用内公式文本转换为 Excel 存储公式文本。
 *
 * @param input - 应用内公式文本，可以带等号也可以不带等号
 * @returns 不带等号前缀的 Excel 公式文本
 * @throws {FormulaError} 当公式为空时抛出
 * @author liangzai927
 */
export function toExcelFormulaText(input: string): string {
  return normalizeFormulaText(input).slice(1);
}
