import { isCellAddress } from './address';
import type { FormulaBinaryOperator, FormulaNode } from './ast';
import type { FormulaErrorCode } from './errors';
import { createFormulaParseError } from './errors';
import type { FormulaToken, FormulaTokenType } from './tokenizer';
import { tokenizeFormula } from './tokenizer';

/**
 * 将公式字符串解析为 AST。
 *
 * @param input - 原始公式字符串
 * @returns 公式 AST 根节点
 * @throws {FormulaError} 当公式语法非法时抛出
 * @author liangzai927
 */
export function parseFormula(input: string): FormulaNode {
  const parser = new FormulaParser(tokenizeFormula(input));
  return parser.parse();
}

class FormulaParser {
  private readonly tokens: Array<FormulaToken>;
  private current = 0;

  /**
   * 创建公式解析器。
   *
   * @param tokens - 公式 Token 列表
   * @author liangzai927
   */
  constructor(tokens: Array<FormulaToken>) {
    this.tokens = tokens;
  }

  /**
   * 解析完整公式表达式。
   *
   * @returns 公式 AST 根节点
   * @throws {FormulaError} 当存在多余 Token 时抛出
   * @author liangzai927
   */
  parse(): FormulaNode {
    const expression = this.parseComparison();
    if (!this.check('eof')) {
      throw createFormulaParseError(`Unexpected token "${this.peek().value}"`);
    }
    return expression;
  }

  /**
   * 解析比较表达式。
   *
   * @returns 比较表达式节点
   * @author liangzai927
   */
  private parseComparison(): FormulaNode {
    let node = this.parseConcat();
    while (this.matchOperator('=', '<>', '<', '<=', '>', '>=')) {
      const operator = this.previous().value as FormulaBinaryOperator;
      node = { type: 'binary', operator, left: node, right: this.parseConcat() };
    }
    return node;
  }

  /**
   * 解析文本连接表达式。
   *
   * @returns 文本连接表达式节点
   * @author liangzai927
   */
  private parseConcat(): FormulaNode {
    let node = this.parseAdditive();
    while (this.matchOperator('&')) {
      node = { type: 'binary', operator: '&', left: node, right: this.parseAdditive() };
    }
    return node;
  }

  /**
   * 解析加减表达式。
   *
   * @returns 加减表达式节点
   * @author liangzai927
   */
  private parseAdditive(): FormulaNode {
    let node = this.parseMultiplicative();
    while (this.matchOperator('+', '-')) {
      const operator = this.previous().value as FormulaBinaryOperator;
      node = { type: 'binary', operator, left: node, right: this.parseMultiplicative() };
    }
    return node;
  }

  /**
   * 解析乘除表达式。
   *
   * @returns 乘除表达式节点
   * @author liangzai927
   */
  private parseMultiplicative(): FormulaNode {
    let node = this.parsePower();
    while (this.matchOperator('*', '/')) {
      const operator = this.previous().value as FormulaBinaryOperator;
      node = { type: 'binary', operator, left: node, right: this.parsePower() };
    }
    return node;
  }

  /**
   * 解析幂运算表达式。
   *
   * @returns 幂运算表达式节点
   * @author liangzai927
   */
  private parsePower(): FormulaNode {
    let node = this.parseUnary();
    while (this.matchOperator('^')) {
      node = { type: 'binary', operator: '^', left: node, right: this.parseUnary() };
    }
    return node;
  }

  /**
   * 解析一元表达式。
   *
   * @returns 一元表达式节点
   * @author liangzai927
   */
  private parseUnary(): FormulaNode {
    if (this.matchOperator('+', '-')) {
      const operator = this.previous().value as '+' | '-';
      return { type: 'unary', operator, argument: this.parseUnary() };
    }
    return this.parsePrimary();
  }

  /**
   * 解析基础表达式。
   *
   * @returns 基础表达式节点
   * @throws {FormulaError} 当 Token 不能组成表达式时抛出
   * @author liangzai927
   */
  private parsePrimary(): FormulaNode {
    if (this.match('number')) return { type: 'number', value: Number(this.previous().value) };
    if (this.match('string')) return { type: 'string', value: this.previous().value };
    if (this.match('error'))
      return { type: 'error', code: this.previous().value as FormulaErrorCode };
    if (this.match('identifier')) return this.parseIdentifier(this.previous().value);
    if (this.match('leftParen')) {
      const node = this.parseComparison();
      this.consume('rightParen', 'Expected ")" after expression');
      return node;
    }
    throw createFormulaParseError(`Unexpected token "${this.peek().value}"`);
  }

  /**
   * 解析标识符表达式。
   *
   * @param value - 标识符文本
   * @returns 标识符对应的 AST 节点
   * @throws {FormulaError} 当标识符不是函数或单元格引用时抛出
   * @author liangzai927
   */
  private parseIdentifier(value: string): FormulaNode {
    if (this.match('leftParen')) return this.parseFunctionCall(value);
    if (value.toUpperCase() === 'TRUE') return { type: 'boolean', value: true };
    if (value.toUpperCase() === 'FALSE') return { type: 'boolean', value: false };
    if (!isCellAddress(value)) throw createFormulaParseError(`Invalid reference "${value}"`);

    const cellNode: FormulaNode = { type: 'cell', ref: value.toUpperCase() };
    if (!this.match('colon')) return cellNode;

    const end = this.consume('identifier', 'Expected cell reference after ":"').value;
    if (!isCellAddress(end)) throw createFormulaParseError(`Invalid range end "${end}"`);
    return { type: 'range', start: value.toUpperCase(), end: end.toUpperCase() };
  }

  /**
   * 解析函数调用。
   *
   * @param name - 函数名称
   * @returns 函数调用节点
   * @author liangzai927
   */
  private parseFunctionCall(name: string): FormulaNode {
    const args: Array<FormulaNode> = [];
    if (!this.check('rightParen')) {
      do {
        args.push(this.parseComparison());
      } while (this.match('comma'));
    }
    this.consume('rightParen', 'Expected ")" after function arguments');
    return { type: 'function', name: name.toUpperCase(), args };
  }

  /**
   * 匹配指定类型 Token。
   *
   * @param type - 目标 Token 类型
   * @returns 匹配成功时返回 true
   * @author liangzai927
   */
  private match(type: FormulaTokenType): boolean {
    if (!this.check(type)) return false;
    this.advance();
    return true;
  }

  /**
   * 匹配指定操作符。
   *
   * @param operators - 目标操作符列表
   * @returns 匹配成功时返回 true
   * @author liangzai927
   */
  private matchOperator(...operators: Array<string>): boolean {
    if (!this.check('operator') || !operators.includes(this.peek().value)) return false;
    this.advance();
    return true;
  }

  /**
   * 消费指定类型 Token。
   *
   * @param type - 目标 Token 类型
   * @param message - 匹配失败时的错误描述
   * @returns 被消费的 Token
   * @throws {FormulaError} 当当前 Token 类型不匹配时抛出
   * @author liangzai927
   */
  private consume(type: FormulaTokenType, message: string): FormulaToken {
    if (this.check(type)) return this.advance();
    throw createFormulaParseError(message);
  }

  /**
   * 检查当前 Token 类型。
   *
   * @param type - 目标 Token 类型
   * @returns 当前 Token 类型匹配时返回 true
   * @author liangzai927
   */
  private check(type: FormulaTokenType): boolean {
    return this.peek().type === type;
  }

  /**
   * 前进到下一个 Token。
   *
   * @returns 前进前的当前 Token
   * @author liangzai927
   */
  private advance(): FormulaToken {
    if (!this.check('eof')) this.current += 1;
    return this.previous();
  }

  /**
   * 获取当前 Token。
   *
   * @returns 当前 Token
   * @author liangzai927
   */
  private peek(): FormulaToken {
    return this.tokens[this.current] ?? this.lastToken();
  }

  /**
   * 获取前一个 Token。
   *
   * @returns 前一个 Token
   * @author liangzai927
   */
  private previous(): FormulaToken {
    return this.tokens[this.current - 1] ?? this.firstToken();
  }

  /**
   * 获取第一个 Token，空列表时返回 EOF 兜底 Token。
   *
   * @returns 第一个 Token 或 EOF Token
   * @author liangzai927
   */
  private firstToken(): FormulaToken {
    return this.tokens[0] ?? createEofToken();
  }

  /**
   * 获取最后一个 Token，空列表时返回 EOF 兜底 Token。
   *
   * @returns 最后一个 Token 或 EOF Token
   * @author liangzai927
   */
  private lastToken(): FormulaToken {
    return this.tokens[this.tokens.length - 1] ?? createEofToken();
  }
}

/**
 * 创建 EOF 兜底 Token。
 *
 * @returns EOF Token
 * @author liangzai927
 */
function createEofToken(): FormulaToken {
  return { type: 'eof', value: '', position: 0 };
}
