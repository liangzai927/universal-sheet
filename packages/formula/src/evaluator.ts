import { enumerateRangeAddress, formatCellAddress, parseRangeAddress } from './address';
import type { FormulaEvaluationContext, FormulaNode, FormulaPrimitive, FormulaValue } from './ast';
import { createFormulaNameError, createFormulaValueError, FormulaError } from './errors';
import { parseFormula } from './parser';

type FormulaFunction = (args: ReadonlyArray<FormulaValue>) => FormulaPrimitive;

const FUNCTIONS: ReadonlyMap<string, FormulaFunction> = new Map([
  ['SUM', sumFunction],
  ['AVERAGE', averageFunction],
  ['MIN', minFunction],
  ['MAX', maxFunction],
  ['COUNT', countFunction],
  ['IF', ifFunction],
]);

/**
 * 解析并计算公式字符串。
 *
 * @param input - 原始公式字符串
 * @param context - 单元格和区域引用的求值上下文
 * @returns 公式计算结果
 * @throws {FormulaError} 当公式解析或计算失败时抛出
 * @author liangzai927
 */
export function evaluateFormula(
  input: string,
  context: FormulaEvaluationContext = {},
): FormulaPrimitive {
  const value = evaluateFormulaNode(parseFormula(input), context);
  if (isFormulaPrimitiveArray(value)) {
    throw createFormulaValueError('Formula result cannot be a range');
  }
  return value;
}

/**
 * 计算公式 AST 节点。
 *
 * @param node - 需要计算的 AST 节点
 * @param context - 单元格和区域引用的求值上下文
 * @returns 节点计算结果
 * @throws {FormulaError} 当节点计算失败时抛出
 * @author liangzai927
 */
export function evaluateFormulaNode(
  node: FormulaNode,
  context: FormulaEvaluationContext = {},
): FormulaValue {
  switch (node.type) {
    case 'number':
    case 'string':
    case 'boolean':
      return node.value;
    case 'error':
      throw new FormulaError(node.code, `Formula contains error literal ${node.code}`);
    case 'cell':
      return context.getCellValue?.(node.ref) ?? null;
    case 'range':
      return resolveRangeValues(node.start, node.end, context);
    case 'unary':
      return evaluateUnary(node.operator, evaluateFormulaNode(node.argument, context));
    case 'binary':
      return evaluateBinary(
        node.operator,
        evaluateFormulaNode(node.left, context),
        evaluateFormulaNode(node.right, context),
      );
    case 'function':
      return evaluateFunction(
        node.name,
        node.args.map((arg) => evaluateFormulaNode(arg, context)),
      );
  }
}

/**
 * 解析并读取区域引用的值。
 *
 * @param start - 区域起始单元格引用
 * @param end - 区域结束单元格引用
 * @param context - 单元格和区域引用的求值上下文
 * @returns 区域内的单元格值列表
 * @author liangzai927
 */
function resolveRangeValues(
  start: string,
  end: string,
  context: FormulaEvaluationContext,
): ReadonlyArray<FormulaPrimitive> {
  const rangeValues = context.getRangeValues?.(start, end);
  if (rangeValues) return rangeValues;
  if (!context.getCellValue) return [];

  return enumerateRangeAddress(parseRangeAddress(`${start}:${end}`)).map(
    (address) => context.getCellValue?.(formatCellAddress(address)) ?? null,
  );
}

/**
 * 计算一元表达式。
 *
 * @param operator - 一元操作符
 * @param value - 被操作的值
 * @returns 一元表达式结果
 * @author liangzai927
 */
function evaluateUnary(operator: '+' | '-', value: FormulaValue): FormulaPrimitive {
  const number = toNumber(value);
  return operator === '-' ? -number : number;
}

/**
 * 计算二元表达式。
 *
 * @param operator - 二元操作符
 * @param left - 左侧值
 * @param right - 右侧值
 * @returns 二元表达式结果
 * @throws {FormulaError} 当除数为 0 或值类型非法时抛出
 * @author liangzai927
 */
function evaluateBinary(
  operator: string,
  left: FormulaValue,
  right: FormulaValue,
): FormulaPrimitive {
  if (operator === '&') return `${toScalar(left) ?? ''}${toScalar(right) ?? ''}`;
  if (operator === '=') return toScalar(left) === toScalar(right);
  if (operator === '<>') return toScalar(left) !== toScalar(right);
  if (operator === '<') return toNumber(left) < toNumber(right);
  if (operator === '<=') return toNumber(left) <= toNumber(right);
  if (operator === '>') return toNumber(left) > toNumber(right);
  if (operator === '>=') return toNumber(left) >= toNumber(right);

  const a = toNumber(left);
  const b = toNumber(right);
  if (operator === '+') return a + b;
  if (operator === '-') return a - b;
  if (operator === '*') return a * b;
  if (operator === '^') return a ** b;
  if (operator === '/') {
    if (b === 0) throw new FormulaError('#DIV/0!', 'Division by zero');
    return a / b;
  }
  throw createFormulaValueError(`Unsupported operator: ${operator}`);
}

/**
 * 计算函数调用。
 *
 * @param name - 函数名称
 * @param args - 已计算的函数参数
 * @returns 函数计算结果
 * @throws {FormulaError} 当函数不存在时抛出
 * @author liangzai927
 */
function evaluateFunction(name: string, args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  const fn = FUNCTIONS.get(name);
  if (!fn) throw createFormulaNameError(name);
  return fn(args);
}

/**
 * 计算 SUM 函数。
 *
 * @param args - 函数参数
 * @returns 求和结果
 * @author liangzai927
 */
function sumFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  return collectNumbers(args).reduce((total, value) => total + value, 0);
}

/**
 * 计算 AVERAGE 函数。
 *
 * @param args - 函数参数
 * @returns 平均值；没有数字时返回 0
 * @author liangzai927
 */
function averageFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  const numbers = collectNumbers(args);
  if (numbers.length === 0) return 0;
  return numbers.reduce((total, value) => total + value, 0) / numbers.length;
}

/**
 * 计算 MIN 函数。
 *
 * @param args - 函数参数
 * @returns 最小值；没有数字时返回 0
 * @author liangzai927
 */
function minFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  const numbers = collectNumbers(args);
  return numbers.length > 0 ? Math.min(...numbers) : 0;
}

/**
 * 计算 MAX 函数。
 *
 * @param args - 函数参数
 * @returns 最大值；没有数字时返回 0
 * @author liangzai927
 */
function maxFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  const numbers = collectNumbers(args);
  return numbers.length > 0 ? Math.max(...numbers) : 0;
}

/**
 * 计算 COUNT 函数。
 *
 * @param args - 函数参数
 * @returns 数字数量
 * @author liangzai927
 */
function countFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  return collectNumbers(args).length;
}

/**
 * 计算 IF 函数。
 *
 * @param args - 函数参数
 * @returns 条件分支结果
 * @author liangzai927
 */
function ifFunction(args: ReadonlyArray<FormulaValue>): FormulaPrimitive {
  const condition = Boolean(toScalar(args[0] ?? false));
  return toScalar(condition ? (args[1] ?? true) : (args[2] ?? false));
}

/**
 * 收集参数中的数字值。
 *
 * @param values - 公式值列表
 * @returns 数字值列表
 * @author liangzai927
 */
function collectNumbers(values: ReadonlyArray<FormulaValue>): Array<number> {
  const numbers: Array<number> = [];
  for (const value of values) {
    const flattened = isFormulaPrimitiveArray(value) ? value : [value];
    for (const item of flattened) {
      const number = maybeNumber(item);
      if (number !== null) numbers.push(number);
    }
  }
  return numbers;
}

/**
 * 将公式值转换为数字。
 *
 * @param value - 需要转换的公式值
 * @returns 数字值
 * @throws {FormulaError} 当值不能转换为数字时抛出
 * @author liangzai927
 */
function toNumber(value: FormulaValue): number {
  const number = maybeNumber(toScalar(value));
  if (number === null) throw createFormulaValueError('Expected numeric value');
  return number;
}

/**
 * 尝试将原始值转换为数字。
 *
 * @param value - 需要转换的原始值
 * @returns 转换后的数字；无法转换时返回 null
 * @author liangzai927
 */
function maybeNumber(value: FormulaPrimitive): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (value === null || value === '') return 0;
  if (typeof value === 'string') {
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }
  return null;
}

/**
 * 将区域值压缩为单个标量值。
 *
 * @param value - 需要压缩的公式值
 * @returns 标量值
 * @author liangzai927
 */
function toScalar(value: FormulaValue): FormulaPrimitive {
  return isFormulaPrimitiveArray(value) ? (value[0] ?? null) : value;
}

/**
 * 判断公式值是否为区域值列表。
 *
 * @param value - 需要检查的公式值
 * @returns 是区域值列表时返回 true
 * @author liangzai927
 */
function isFormulaPrimitiveArray(value: FormulaValue): value is ReadonlyArray<FormulaPrimitive> {
  return Array.isArray(value);
}
