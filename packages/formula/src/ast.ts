import type { FormulaErrorCode } from './errors';

/** 公式 AST 节点类型。 */
export type FormulaNode =
  | NumberLiteralNode
  | StringLiteralNode
  | BooleanLiteralNode
  | ErrorLiteralNode
  | CellReferenceNode
  | RangeReferenceNode
  | UnaryExpressionNode
  | BinaryExpressionNode
  | FunctionCallNode;

/** 数字字面量节点。 */
export interface NumberLiteralNode {
  readonly type: 'number';
  readonly value: number;
}

/** 字符串字面量节点。 */
export interface StringLiteralNode {
  readonly type: 'string';
  readonly value: string;
}

/** 布尔字面量节点。 */
export interface BooleanLiteralNode {
  readonly type: 'boolean';
  readonly value: boolean;
}

/** 错误字面量节点。 */
export interface ErrorLiteralNode {
  readonly type: 'error';
  readonly code: FormulaErrorCode;
}

/** 单元格引用节点。 */
export interface CellReferenceNode {
  readonly type: 'cell';
  readonly ref: string;
}

/** 区域引用节点。 */
export interface RangeReferenceNode {
  readonly type: 'range';
  readonly start: string;
  readonly end: string;
}

/** 一元表达式节点。 */
export interface UnaryExpressionNode {
  readonly type: 'unary';
  readonly operator: '+' | '-';
  readonly argument: FormulaNode;
}

/** 二元表达式节点。 */
export interface BinaryExpressionNode {
  readonly type: 'binary';
  readonly operator: FormulaBinaryOperator;
  readonly left: FormulaNode;
  readonly right: FormulaNode;
}

/** 函数调用节点。 */
export interface FunctionCallNode {
  readonly type: 'function';
  readonly name: string;
  readonly args: Array<FormulaNode>;
}

/** 公式二元运算符。 */
export type FormulaBinaryOperator =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '<='
  | '>'
  | '>=';

/** 公式原始值。 */
export type FormulaPrimitive = number | string | boolean | null;

/** 公式求值结果，区域引用会返回一组原始值。 */
export type FormulaValue = FormulaPrimitive | ReadonlyArray<FormulaPrimitive>;

/** 公式求值上下文。 */
export interface FormulaEvaluationContext {
  readonly getCellValue?: (ref: string) => FormulaPrimitive;
  readonly getRangeValues?: (start: string, end: string) => ReadonlyArray<FormulaPrimitive>;
}
