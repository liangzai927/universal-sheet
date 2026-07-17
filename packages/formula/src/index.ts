// @universal-sheet/formula - 公式解析与计算引擎。

export type { CellAddress, RangeAddress } from './address';
export {
  columnIndexToLabel,
  columnLabelToIndex,
  enumerateRangeAddress,
  formatCellAddress,
  formatRangeAddress,
  isCellAddress,
  parseCellAddress,
  parseRangeAddress,
} from './address';
export type {
  BinaryExpressionNode,
  BooleanLiteralNode,
  CellReferenceNode,
  ErrorLiteralNode,
  FormulaBinaryOperator,
  FormulaEvaluationContext,
  FormulaNode,
  FormulaPrimitive,
  FormulaValue,
  FunctionCallNode,
  NumberLiteralNode,
  RangeReferenceNode,
  StringLiteralNode,
  UnaryExpressionNode,
} from './ast';
export type { FormulaCalculationResult, FormulaCellInput } from './calculation';
export { calculateFormulaCells } from './calculation';
export type { FormulaDependencies } from './dependencies';
export { collectFormulaDependencies, collectFormulaNodeDependencies } from './dependencies';
export type { FormulaErrorCode } from './errors';
export {
  createFormulaNameError,
  createFormulaParseError,
  createFormulaValueError,
  FormulaError,
} from './errors';
export { evaluateFormula, evaluateFormulaNode } from './evaluator';
export type { CreateFormulaRecordOptions, FormulaRecord } from './formula-record';
export {
  createFormulaRecord,
  fromExcelFormulaText,
  normalizeFormulaText,
  toExcelFormulaText,
} from './formula-record';
export { parseFormula } from './parser';
export type {
  FormulaReferenceChange,
  FormulaReferenceChangeType,
  FormulaReferenceOffsetOptions,
} from './references';
export {
  applyFormulaNodeReferenceChange,
  applyFormulaReferenceChange,
  offsetFormulaNodeReferences,
  offsetFormulaReferences,
} from './references';
export { serializeFormula, serializeFormulaNode } from './serializer';
export type { FormulaToken, FormulaTokenType } from './tokenizer';
export { tokenizeFormula } from './tokenizer';

export const VERSION = '1.0.0';
