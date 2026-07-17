import type { CellAddress } from './address';
import { formatCellAddress, parseCellAddress } from './address';
import type { FormulaNode } from './ast';
import { createFormulaParseError } from './errors';
import { parseFormula } from './parser';
import { serializeFormula } from './serializer';

/** 公式引用偏移配置。 */
export interface FormulaReferenceOffsetOptions {
  readonly rowOffset?: number;
  readonly colOffset?: number;
  readonly respectAbsolute?: boolean;
}

/** 公式引用结构变更类型。 */
export type FormulaReferenceChangeType =
  | 'insertRows'
  | 'deleteRows'
  | 'insertColumns'
  | 'deleteColumns';

/** 公式引用结构变更配置。 */
export interface FormulaReferenceChange {
  readonly type: FormulaReferenceChangeType;
  readonly startIndex: number;
  readonly count: number;
}

/**
 * 将公式中的单元格和区域引用按行列偏移。
 *
 * @param input - 原始公式文本
 * @param options - 引用偏移配置
 * @returns 偏移后的公式文本
 * @throws {FormulaError} 当公式非法或偏移后地址越界时抛出
 * @author liangzai927
 */
export function offsetFormulaReferences(
  input: string,
  options: FormulaReferenceOffsetOptions,
): string {
  return serializeFormula(offsetFormulaNode(parseFormula(input), options));
}

/**
 * 偏移公式 AST 节点中的引用。
 *
 * @param node - 需要偏移的公式 AST 节点
 * @param options - 引用偏移配置
 * @returns 偏移后的公式 AST 节点
 * @author liangzai927
 */
export function offsetFormulaNodeReferences(
  node: FormulaNode,
  options: FormulaReferenceOffsetOptions,
): FormulaNode {
  return offsetFormulaNode(node, options);
}

/**
 * 按行列结构变更更新公式引用。
 *
 * @param input - 原始公式文本
 * @param change - 行列结构变更配置
 * @returns 更新后的公式文本
 * @throws {FormulaError} 当公式非法或变更配置非法时抛出
 * @author liangzai927
 */
export function applyFormulaReferenceChange(input: string, change: FormulaReferenceChange): string {
  return serializeFormula(applyFormulaNodeReferenceChange(parseFormula(input), change));
}

/**
 * 按行列结构变更更新公式 AST 中的引用。
 *
 * @param node - 需要更新的公式 AST 节点
 * @param change - 行列结构变更配置
 * @returns 更新后的公式 AST 节点
 * @throws {FormulaError} 当变更配置非法时抛出
 * @author liangzai927
 */
export function applyFormulaNodeReferenceChange(
  node: FormulaNode,
  change: FormulaReferenceChange,
): FormulaNode {
  validateReferenceChange(change);
  return applyReferenceChangeToNode(node, change);
}

/**
 * 递归偏移公式 AST 节点。
 *
 * @param node - 需要偏移的公式 AST 节点
 * @param options - 引用偏移配置
 * @returns 偏移后的公式 AST 节点
 * @author liangzai927
 */
function offsetFormulaNode(node: FormulaNode, options: FormulaReferenceOffsetOptions): FormulaNode {
  switch (node.type) {
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return node;
    case 'cell':
      return { type: 'cell', ref: offsetCellReference(node.ref, options) };
    case 'range':
      return {
        type: 'range',
        start: offsetCellReference(node.start, options),
        end: offsetCellReference(node.end, options),
      };
    case 'unary':
      return { ...node, argument: offsetFormulaNode(node.argument, options) };
    case 'binary':
      return {
        ...node,
        left: offsetFormulaNode(node.left, options),
        right: offsetFormulaNode(node.right, options),
      };
    case 'function':
      return { ...node, args: node.args.map((arg) => offsetFormulaNode(arg, options)) };
  }
}

/**
 * 递归应用行列结构变更。
 *
 * @param node - 需要更新的公式 AST 节点
 * @param change - 行列结构变更配置
 * @returns 更新后的公式 AST 节点
 * @author liangzai927
 */
function applyReferenceChangeToNode(
  node: FormulaNode,
  change: FormulaReferenceChange,
): FormulaNode {
  switch (node.type) {
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return node;
    case 'cell':
      return applyReferenceChangeToCellNode(node.ref, change);
    case 'range':
      return applyReferenceChangeToRangeNode(node.start, node.end, change);
    case 'unary':
      return { ...node, argument: applyReferenceChangeToNode(node.argument, change) };
    case 'binary':
      return {
        ...node,
        left: applyReferenceChangeToNode(node.left, change),
        right: applyReferenceChangeToNode(node.right, change),
      };
    case 'function':
      return { ...node, args: node.args.map((arg) => applyReferenceChangeToNode(arg, change)) };
  }
}

/**
 * 对单元格引用应用行列结构变更。
 *
 * @param ref - 原始 A1 引用
 * @param change - 行列结构变更配置
 * @returns 更新后的单元格节点或错误节点
 * @author liangzai927
 */
function applyReferenceChangeToCellNode(ref: string, change: FormulaReferenceChange): FormulaNode {
  const address = applyReferenceChangeToAddress(parseCellAddress(ref), change);
  if (!address) return { type: 'error', code: '#REF!' };
  return { type: 'cell', ref: formatCellAddress(address) };
}

/**
 * 对区域引用应用行列结构变更。
 *
 * @param start - 区域起始 A1 引用
 * @param end - 区域结束 A1 引用
 * @param change - 行列结构变更配置
 * @returns 更新后的区域节点或错误节点
 * @author liangzai927
 */
function applyReferenceChangeToRangeNode(
  start: string,
  end: string,
  change: FormulaReferenceChange,
): FormulaNode {
  const nextStart = applyReferenceChangeToAddress(parseCellAddress(start), change);
  const nextEnd = applyReferenceChangeToAddress(parseCellAddress(end), change);
  if (!nextStart || !nextEnd) return { type: 'error', code: '#REF!' };
  return { type: 'range', start: formatCellAddress(nextStart), end: formatCellAddress(nextEnd) };
}

/**
 * 偏移单个 A1 引用。
 *
 * @param ref - 原始 A1 引用
 * @param options - 引用偏移配置
 * @returns 偏移后的 A1 引用
 * @author liangzai927
 */
function offsetCellReference(ref: string, options: FormulaReferenceOffsetOptions): string {
  return formatCellAddress(offsetCellAddress(parseCellAddress(ref), options));
}

/**
 * 偏移单元格地址。
 *
 * @param address - 原始单元格地址
 * @param options - 引用偏移配置
 * @returns 偏移后的单元格地址
 * @author liangzai927
 */
function offsetCellAddress(
  address: CellAddress,
  options: FormulaReferenceOffsetOptions,
): CellAddress {
  const respectAbsolute = options.respectAbsolute ?? true;
  const rowOffset = respectAbsolute && address.absoluteRow ? 0 : (options.rowOffset ?? 0);
  const colOffset = respectAbsolute && address.absoluteCol ? 0 : (options.colOffset ?? 0);

  return {
    ...address,
    row: address.row + rowOffset,
    col: address.col + colOffset,
  };
}

/**
 * 对单元格地址应用行列结构变更。
 *
 * @param address - 原始单元格地址
 * @param change - 行列结构变更配置
 * @returns 更新后的单元格地址；引用被删除时返回 null
 * @author liangzai927
 */
function applyReferenceChangeToAddress(
  address: CellAddress,
  change: FormulaReferenceChange,
): CellAddress | null {
  if (change.type === 'insertRows') {
    return address.row >= change.startIndex
      ? { ...address, row: address.row + change.count }
      : address;
  }
  if (change.type === 'deleteRows') {
    return applyDeletionToAddress(address, 'row', change.startIndex, change.count);
  }
  if (change.type === 'insertColumns') {
    return address.col >= change.startIndex
      ? { ...address, col: address.col + change.count }
      : address;
  }
  return applyDeletionToAddress(address, 'col', change.startIndex, change.count);
}

/**
 * 对单元格地址应用删除变更。
 *
 * @param address - 原始单元格地址
 * @param axis - 删除作用的坐标轴
 * @param startIndex - 删除起始索引
 * @param count - 删除数量
 * @returns 更新后的单元格地址；引用被删除时返回 null
 * @author liangzai927
 */
function applyDeletionToAddress(
  address: CellAddress,
  axis: 'row' | 'col',
  startIndex: number,
  count: number,
): CellAddress | null {
  const value = address[axis];
  const endIndex = startIndex + count;
  if (value >= startIndex && value < endIndex) return null;
  if (value < endIndex) return address;
  return { ...address, [axis]: value - count };
}

/**
 * 校验行列结构变更配置。
 *
 * @param change - 行列结构变更配置
 * @throws {FormulaError} 当配置非法时抛出
 * @author liangzai927
 */
function validateReferenceChange(change: FormulaReferenceChange): void {
  if (!Number.isInteger(change.startIndex) || change.startIndex < 0) {
    throw createFormulaParseError(`Invalid start index "${change.startIndex}"`);
  }
  if (!Number.isInteger(change.count) || change.count <= 0) {
    throw createFormulaParseError(`Invalid change count "${change.count}"`);
  }
}
