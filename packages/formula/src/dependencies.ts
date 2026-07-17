import type { CellAddress, RangeAddress } from './address';
import { enumerateRangeAddress, formatCellAddress, parseRangeAddress } from './address';
import type { FormulaNode } from './ast';
import { parseFormula } from './parser';

/** 公式引用依赖收集结果。 */
export interface FormulaDependencies {
  readonly cells: ReadonlyArray<string>;
  readonly ranges: ReadonlyArray<string>;
  readonly expandedCells: ReadonlyArray<string>;
}

/**
 * 收集公式字符串中的单元格和区域依赖。
 *
 * @param input - 原始公式文本
 * @returns 去重后的公式依赖信息
 * @throws {FormulaError} 当公式语法非法时抛出
 * @author liangzai927
 */
export function collectFormulaDependencies(input: string): FormulaDependencies {
  return collectFormulaNodeDependencies(parseFormula(input));
}

/**
 * 收集公式 AST 中的单元格和区域依赖。
 *
 * @param node - 公式 AST 根节点
 * @returns 去重后的公式依赖信息
 * @author liangzai927
 */
export function collectFormulaNodeDependencies(node: FormulaNode): FormulaDependencies {
  const collector = createDependencyCollector();
  collectNodeDependencies(node, collector);
  return {
    cells: Array.from(collector.cells),
    ranges: Array.from(collector.ranges),
    expandedCells: Array.from(collector.expandedCells),
  };
}

interface DependencyCollector {
  readonly cells: Set<string>;
  readonly ranges: Set<string>;
  readonly expandedCells: Set<string>;
}

/**
 * 创建公式依赖收集器。
 *
 * @returns 公式依赖收集器
 * @author liangzai927
 */
function createDependencyCollector(): DependencyCollector {
  return {
    cells: new Set<string>(),
    ranges: new Set<string>(),
    expandedCells: new Set<string>(),
  };
}

/**
 * 递归收集公式 AST 节点依赖。
 *
 * @param node - 当前公式 AST 节点
 * @param collector - 公式依赖收集器
 * @author liangzai927
 */
function collectNodeDependencies(node: FormulaNode, collector: DependencyCollector): void {
  switch (node.type) {
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return;
    case 'cell':
      collectCellDependency(node.ref, collector);
      return;
    case 'range':
      collectRangeDependency(node.start, node.end, collector);
      return;
    case 'unary':
      collectNodeDependencies(node.argument, collector);
      return;
    case 'binary':
      collectNodeDependencies(node.left, collector);
      collectNodeDependencies(node.right, collector);
      return;
    case 'function':
      for (const arg of node.args) collectNodeDependencies(arg, collector);
  }
}

/**
 * 收集单元格引用依赖。
 *
 * @param ref - A1 单元格引用
 * @param collector - 公式依赖收集器
 * @author liangzai927
 */
function collectCellDependency(ref: string, collector: DependencyCollector): void {
  collector.cells.add(ref);
  collector.expandedCells.add(ref);
}

/**
 * 收集区域引用依赖。
 *
 * @param start - 区域起始 A1 引用
 * @param end - 区域结束 A1 引用
 * @param collector - 公式依赖收集器
 * @author liangzai927
 */
function collectRangeDependency(start: string, end: string, collector: DependencyCollector): void {
  const range = parseRangeAddress(`${start}:${end}`);
  collector.ranges.add(formatRangeDependency(range));
  for (const cell of enumerateRangeAddress(range)) {
    collector.expandedCells.add(formatCellDependency(cell));
  }
}

/**
 * 格式化区域依赖文本。
 *
 * @param range - A1 区域地址
 * @returns 归一化后的区域依赖文本
 * @author liangzai927
 */
function formatRangeDependency(range: RangeAddress): string {
  return `${formatCellDependency(range.start)}:${formatCellDependency(range.end)}`;
}

/**
 * 格式化单元格依赖文本。
 *
 * @param address - A1 单元格地址
 * @returns 归一化后的单元格依赖文本
 * @author liangzai927
 */
function formatCellDependency(address: CellAddress): string {
  return formatCellAddress(address);
}
