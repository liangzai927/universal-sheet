/** 公式错误类型。 */
export type FormulaErrorCode = '#DIV/0!' | '#REF!' | '#VALUE!' | '#NAME?' | '#N/A';

/** 公式解析或求值错误。 */
export class FormulaError extends Error {
  readonly code: FormulaErrorCode;

  /**
   * 创建公式错误。
   *
   * @param code - Excel 风格错误码
   * @param message - 错误描述
   * @author liangzai927
   */
  constructor(code: FormulaErrorCode, message: string) {
    super(message);
    this.name = 'FormulaError';
    this.code = code;
  }
}

/**
 * 创建公式解析错误。
 *
 * @param message - 错误描述
 * @returns 公式解析错误
 * @author liangzai927
 */
export function createFormulaParseError(message: string): FormulaError {
  return new FormulaError('#VALUE!', message);
}

/**
 * 创建公式名称错误。
 *
 * @param name - 未识别的函数名或引用名
 * @returns 公式名称错误
 * @author liangzai927
 */
export function createFormulaNameError(name: string): FormulaError {
  return new FormulaError('#NAME?', `Unknown formula name: ${name}`);
}

/**
 * 创建公式值错误。
 *
 * @param message - 错误描述
 * @returns 公式值错误
 * @author liangzai927
 */
export function createFormulaValueError(message: string): FormulaError {
  return new FormulaError('#VALUE!', message);
}
