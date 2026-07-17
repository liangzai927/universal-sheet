import type { FormulaErrorCode } from './errors';
import { createFormulaParseError } from './errors';

/** 公式 Token 类型。 */
export type FormulaTokenType =
  | 'number'
  | 'string'
  | 'error'
  | 'identifier'
  | 'operator'
  | 'leftParen'
  | 'rightParen'
  | 'comma'
  | 'colon'
  | 'eof';

/** 公式 Token。 */
export interface FormulaToken {
  readonly type: FormulaTokenType;
  readonly value: string;
  readonly position: number;
}

const TWO_CHAR_OPERATORS = new Set(['>=', '<=', '<>']);
const ONE_CHAR_OPERATORS = new Set(['+', '-', '*', '/', '^', '&', '=', '<', '>']);
const ERROR_LITERALS: ReadonlyArray<FormulaErrorCode> = [
  '#DIV/0!',
  '#VALUE!',
  '#NAME?',
  '#REF!',
  '#N/A',
];

/**
 * 将公式字符串切分为 Token。
 *
 * @param input - 原始公式字符串
 * @returns Token 列表
 * @throws {FormulaError} 当公式中存在无法识别的字符时抛出
 * @author liangzai927
 */
export function tokenizeFormula(input: string): Array<FormulaToken> {
  const source = input.startsWith('=') ? input.slice(1) : input;
  const tokens: Array<FormulaToken> = [];
  let index = 0;

  while (index < source.length) {
    const char = source[index];
    if (char === undefined) break;

    if (isWhitespace(char)) {
      index += 1;
      continue;
    }
    if (isDigit(char) || char === '.') {
      const result = readNumber(source, index);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }
    if (char === '"') {
      const result = readString(source, index);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }
    if (char === '#') {
      const result = readError(source, index);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }
    if (isIdentifierStart(char) || char === '$') {
      const result = readIdentifier(source, index);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }
    if (char === '(') {
      tokens.push({ type: 'leftParen', value: char, position: index });
      index += 1;
      continue;
    }
    if (char === ')') {
      tokens.push({ type: 'rightParen', value: char, position: index });
      index += 1;
      continue;
    }
    if (char === ',') {
      tokens.push({ type: 'comma', value: char, position: index });
      index += 1;
      continue;
    }
    if (char === ':') {
      tokens.push({ type: 'colon', value: char, position: index });
      index += 1;
      continue;
    }

    const twoChar = source.slice(index, index + 2);
    if (TWO_CHAR_OPERATORS.has(twoChar)) {
      tokens.push({ type: 'operator', value: twoChar, position: index });
      index += 2;
      continue;
    }
    if (ONE_CHAR_OPERATORS.has(char)) {
      tokens.push({ type: 'operator', value: char, position: index });
      index += 1;
      continue;
    }

    throw createFormulaParseError(`Unexpected character "${char}" at ${index}`);
  }

  tokens.push({ type: 'eof', value: '', position: source.length });
  return tokens;
}

/**
 * 读取 Excel 错误字面量 Token。
 *
 * @param source - 公式源码
 * @param start - 读取起始位置
 * @returns 错误 Token 和下一个读取位置
 * @throws {FormulaError} 当错误字面量非法时抛出
 * @author liangzai927
 */
function readError(
  source: string,
  start: number,
): { readonly token: FormulaToken; readonly nextIndex: number } {
  const code = ERROR_LITERALS.find((item) => source.startsWith(item, start));
  if (!code) throw createFormulaParseError(`Invalid error literal at ${start}`);

  return {
    token: { type: 'error', value: code, position: start },
    nextIndex: start + code.length,
  };
}

/**
 * 判断字符是否为空白字符。
 *
 * @param char - 需要检查的字符
 * @returns 为空白字符时返回 true
 * @author liangzai927
 */
function isWhitespace(char: string): boolean {
  return /\s/.test(char);
}

/**
 * 判断字符是否为数字。
 *
 * @param char - 需要检查的字符
 * @returns 为数字时返回 true
 * @author liangzai927
 */
function isDigit(char: string): boolean {
  return char >= '0' && char <= '9';
}

/**
 * 判断字符是否可作为标识符开头。
 *
 * @param char - 需要检查的字符
 * @returns 可作为标识符开头时返回 true
 * @author liangzai927
 */
function isIdentifierStart(char: string): boolean {
  return /[A-Za-z_]/.test(char);
}

/**
 * 判断字符是否可作为标识符内容。
 *
 * @param char - 需要检查的字符
 * @returns 可作为标识符内容时返回 true
 * @author liangzai927
 */
function isIdentifierPart(char: string): boolean {
  return /[A-Za-z0-9_$]/.test(char);
}

/**
 * 读取数字 Token。
 *
 * @param source - 公式源码
 * @param start - 读取起始位置
 * @returns 数字 Token 和下一个读取位置
 * @throws {FormulaError} 当数字格式非法时抛出
 * @author liangzai927
 */
function readNumber(
  source: string,
  start: number,
): { readonly token: FormulaToken; readonly nextIndex: number } {
  let index = start;
  let dotCount = 0;
  while (index < source.length) {
    const char = source[index];
    if (char === '.') dotCount += 1;
    if (char === undefined || (!isDigit(char) && char !== '.')) break;
    if (dotCount > 1) {
      throw createFormulaParseError(`Invalid number at ${start}`);
    }
    index += 1;
  }
  const value = source.slice(start, index);
  if (value === '.') {
    throw createFormulaParseError(`Invalid number at ${start}`);
  }
  return {
    token: { type: 'number', value, position: start },
    nextIndex: index,
  };
}

/**
 * 读取字符串 Token。
 *
 * @param source - 公式源码
 * @param start - 读取起始位置
 * @returns 字符串 Token 和下一个读取位置
 * @throws {FormulaError} 当字符串未闭合时抛出
 * @author liangzai927
 */
function readString(
  source: string,
  start: number,
): { readonly token: FormulaToken; readonly nextIndex: number } {
  let index = start + 1;
  let value = '';
  while (index < source.length) {
    const char = source[index];
    if (char === undefined) break;
    if (char === '"') {
      return {
        token: { type: 'string', value, position: start },
        nextIndex: index + 1,
      };
    }
    value += char;
    index += 1;
  }
  throw createFormulaParseError('Unterminated string literal');
}

/**
 * 读取标识符 Token。
 *
 * @param source - 公式源码
 * @param start - 读取起始位置
 * @returns 标识符 Token 和下一个读取位置
 * @author liangzai927
 */
function readIdentifier(
  source: string,
  start: number,
): { readonly token: FormulaToken; readonly nextIndex: number } {
  let index = start;
  while (index < source.length) {
    const char = source[index];
    if (char === undefined || !isIdentifierPart(char)) break;
    index += 1;
  }
  return {
    token: { type: 'identifier', value: source.slice(start, index), position: start },
    nextIndex: index,
  };
}
