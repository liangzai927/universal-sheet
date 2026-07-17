import { formatCellAddress, parseCellAddress } from './address';
import type { FormulaPrimitive } from './ast';
import { collectFormulaDependencies } from './dependencies';
import { FormulaError } from './errors';
import { evaluateFormula } from './evaluator';

/** 可参与公式重算的单元格输入。 */
export interface FormulaCellInput {
  readonly ref: string;
  readonly value?: FormulaPrimitive;
  readonly formula?: string;
}

/** 批量公式重算结果。 */
export interface FormulaCalculationResult {
  readonly values: ReadonlyMap<string, FormulaPrimitive>;
  readonly errors: ReadonlyMap<string, FormulaError>;
  readonly order: ReadonlyArray<string>;
}

/**
 * 批量计算一组公式单元格。
 *
 * @param cells - 参与计算的单元格列表
 * @returns 单元格计算值、错误和计算顺序
 * @author liangzai927
 */
export function calculateFormulaCells(
  cells: ReadonlyArray<FormulaCellInput>,
): FormulaCalculationResult {
  const context = createCalculationContext(cells);

  for (const ref of context.inputs.keys()) {
    tryResolveCellValue(ref, context);
  }

  return {
    values: context.values,
    errors: context.errors,
    order: context.order,
  };
}

/**
 * 尝试解析单元格值，失败时只记录错误。
 *
 * @param ref - A1 单元格引用
 * @param context - 公式重算上下文
 * @author liangzai927
 */
function tryResolveCellValue(ref: string, context: CalculationContext): void {
  try {
    resolveCellValue(ref, context, []);
  } catch (error) {
    const normalizedRef = normalizeCellRef(ref);
    context.errors.set(normalizedRef, normalizeCalculationError(error));
  }
}

interface CalculationContext {
  readonly inputs: Map<string, FormulaCellInput>;
  readonly values: Map<string, FormulaPrimitive>;
  readonly errors: Map<string, FormulaError>;
  readonly resolving: Set<string>;
  readonly order: Array<string>;
}

/**
 * 创建公式重算上下文。
 *
 * @param cells - 参与计算的单元格列表
 * @returns 公式重算上下文
 * @author liangzai927
 */
function createCalculationContext(cells: ReadonlyArray<FormulaCellInput>): CalculationContext {
  const inputs = new Map<string, FormulaCellInput>();
  const values = new Map<string, FormulaPrimitive>();

  for (const cell of cells) {
    const ref = normalizeCellRef(cell.ref);
    const normalized = { ...cell, ref };
    inputs.set(ref, normalized);
    if (!cell.formula) values.set(ref, cell.value ?? null);
  }

  return {
    inputs,
    values,
    errors: new Map<string, FormulaError>(),
    resolving: new Set<string>(),
    order: [],
  };
}

/**
 * 递归解析单元格值。
 *
 * @param ref - A1 单元格引用
 * @param context - 公式重算上下文
 * @param stack - 当前递归调用栈
 * @returns 单元格计算值
 * @throws {FormulaError} 当公式计算失败或发生循环引用时抛出
 * @author liangzai927
 */
function resolveCellValue(
  ref: string,
  context: CalculationContext,
  stack: ReadonlyArray<string>,
): FormulaPrimitive {
  const normalizedRef = normalizeCellRef(ref);
  const existingError = context.errors.get(normalizedRef);
  if (existingError) throw existingError;
  if (context.values.has(normalizedRef)) return context.values.get(normalizedRef) ?? null;

  const input = context.inputs.get(normalizedRef);
  if (!input?.formula) return null;

  if (context.resolving.has(normalizedRef)) {
    throw createCircularReferenceError([...stack, normalizedRef]);
  }

  try {
    context.resolving.add(normalizedRef);
    resolveFormulaDependencies(input.formula, context, [...stack, normalizedRef]);
    const value = evaluateFormula(input.formula, {
      getCellValue: (cellRef) => resolveCellValue(cellRef, context, [...stack, normalizedRef]),
    });
    context.values.set(normalizedRef, value);
    context.order.push(normalizedRef);
    return value;
  } catch (error) {
    const formulaError = normalizeCalculationError(error);
    context.errors.set(normalizedRef, formulaError);
    throw formulaError;
  } finally {
    context.resolving.delete(normalizedRef);
  }
}

/**
 * 预解析公式依赖，保证依赖单元格先计算。
 *
 * @param formula - 原始公式文本
 * @param context - 公式重算上下文
 * @param stack - 当前递归调用栈
 * @author liangzai927
 */
function resolveFormulaDependencies(
  formula: string,
  context: CalculationContext,
  stack: ReadonlyArray<string>,
): void {
  for (const ref of collectFormulaDependencies(formula).expandedCells) {
    resolveCellValue(ref, context, stack);
  }
}

/**
 * 归一化单元格引用文本。
 *
 * @param ref - 原始 A1 单元格引用
 * @returns 归一化后的 A1 单元格引用
 * @throws {FormulaError} 当单元格引用非法时抛出
 * @author liangzai927
 */
function normalizeCellRef(ref: string): string {
  return formatCellAddress(parseCellAddress(ref));
}

/**
 * 创建循环引用错误。
 *
 * @param stack - 循环引用调用栈
 * @returns 循环引用公式错误
 * @author liangzai927
 */
function createCircularReferenceError(stack: ReadonlyArray<string>): FormulaError {
  return new FormulaError('#VALUE!', `Circular formula reference: ${stack.join(' -> ')}`);
}

/**
 * 将未知错误归一化为 FormulaError。
 *
 * @param error - 捕获到的未知错误
 * @returns 公式错误
 * @author liangzai927
 */
function normalizeCalculationError(error: unknown): FormulaError {
  if (error instanceof FormulaError) return error;
  return new FormulaError('#VALUE!', 'Formula calculation failed');
}
