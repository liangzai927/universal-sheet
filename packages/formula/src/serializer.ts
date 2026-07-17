import type { FormulaNode } from './ast';
import { normalizeFormulaText } from './formula-record';

const BINARY_PRECEDENCE: ReadonlyMap<string, number> = new Map([
  ['=', 1],
  ['<>', 1],
  ['<', 1],
  ['<=', 1],
  ['>', 1],
  ['>=', 1],
  ['&', 2],
  ['+', 3],
  ['-', 3],
  ['*', 4],
  ['/', 4],
  ['^', 5],
]);

/**
 * 将公式 AST 序列化为应用内公式文本。
 *
 * @param node - 需要序列化的公式 AST
 * @returns 带等号前缀的公式文本
 * @author liangzai927
 */
export function serializeFormula(node: FormulaNode): string {
  return normalizeFormulaText(serializeFormulaNode(node));
}

/**
 * 将公式 AST 节点序列化为公式表达式片段。
 *
 * @param node - 需要序列化的公式 AST 节点
 * @returns 不带等号前缀的公式表达式片段
 * @author liangzai927
 */
export function serializeFormulaNode(node: FormulaNode): string {
  return serializeNode(node, 0, false);
}

/**
 * 根据父节点优先级序列化公式 AST 节点。
 *
 * @param node - 需要序列化的公式 AST 节点
 * @param parentPrecedence - 父表达式优先级
 * @param isRightChild - 当前节点是否是父表达式右子节点
 * @returns 不带等号前缀的公式表达式片段
 * @author liangzai927
 */
function serializeNode(node: FormulaNode, parentPrecedence: number, isRightChild: boolean): string {
  switch (node.type) {
    case 'number':
      return String(node.value);
    case 'string':
      return `"${node.value.replaceAll('"', '""')}"`;
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE';
    case 'error':
      return node.code;
    case 'cell':
      return node.ref;
    case 'range':
      return `${node.start}:${node.end}`;
    case 'unary':
      return `${node.operator}${serializeNode(node.argument, 6, false)}`;
    case 'binary': {
      const precedence = getBinaryPrecedence(node.operator);
      const value = `${serializeNode(node.left, precedence, false)}${node.operator}${serializeNode(
        node.right,
        precedence,
        true,
      )}`;
      if (shouldWrapBinaryNode(parentPrecedence, precedence, isRightChild, node.operator)) {
        return `(${value})`;
      }
      return value;
    }
    case 'function':
      return `${node.name}(${node.args.map((arg) => serializeNode(arg, 0, false)).join(',')})`;
  }
}

/**
 * 获取二元运算符优先级。
 *
 * @param operator - 二元运算符
 * @returns 运算符优先级
 * @author liangzai927
 */
function getBinaryPrecedence(operator: string): number {
  return BINARY_PRECEDENCE.get(operator) ?? 0;
}

/**
 * 判断二元表达式是否需要括号。
 *
 * @param parentPrecedence - 父表达式优先级
 * @param currentPrecedence - 当前表达式优先级
 * @param isRightChild - 当前表达式是否是右子节点
 * @param operator - 当前表达式运算符
 * @returns 需要括号时返回 true
 * @author liangzai927
 */
function shouldWrapBinaryNode(
  parentPrecedence: number,
  currentPrecedence: number,
  isRightChild: boolean,
  operator: string,
): boolean {
  if (currentPrecedence < parentPrecedence) return true;
  if (currentPrecedence > parentPrecedence) return false;
  return isRightChild && ['-', '/', '^', '&', '=', '<>', '<', '<=', '>', '>='].includes(operator);
}
