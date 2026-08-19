/**
 * Розширена система обробки формул для Discord AI Assistant Bot
 * Безпечна обробка математичних виразів та формул
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import type { LogMeta } from '@/types';
import { handleError } from './errorHandler';
import logger from './logger';
import { validateInput } from './security';

// Константи для обробки формул
const FORMULA_PROCESSOR_CONSTANTS = {
  MAX_FORMULA_LENGTH: 1000,
  MAX_NESTED_LEVELS: 10,
  MAX_ITERATIONS: 1000,
  ALLOWED_FUNCTIONS: [
    'sin',
    'cos',
    'tan',
    'asin',
    'acos',
    'atan',
    'sqrt',
    'pow',
    'exp',
    'log',
    'ln',
    'abs',
    'floor',
    'ceil',
    'round',
    'min',
    'max',
    'sum',
    'avg',
    'count',
    'if',
    'case',
  ] as readonly string[],
  ALLOWED_OPERATORS: [
    '+',
    '-',
    '*',
    '/',
    '^',
    '(',
    ')',
    '=',
    '<',
    '>',
    '<=',
    '>=',
    '!=',
    '==',
  ] as readonly string[],
  ALLOWED_VARIABLES: /^[a-zA-Z_][a-zA-Z0-9_]*$/,
  MAX_VARIABLES: 50,
  PRECISION: 10,
  TIMEOUT: 5000, // 5 секунд
} as const;

export interface FormulaResult {
  success: boolean;
  result?: number;
  error?: string;
  warnings: string[];
  variables: Record<string, number>;
  executionTime: number;
  complexity: number;
}

export interface FormulaValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  sanitizedFormula: string;
  variables: string[];
  functions: string[];
  complexity: number;
}

export interface FormulaProcessorStats {
  totalFormulas: number;
  successfulFormulas: number;
  failedFormulas: number;
  averageExecutionTime: number;
  totalExecutionTime: number;
  complexityDistribution: Record<string, number>;
  errorTypes: Record<string, number>;
  lastFormula?: {
    formula: string;
    result: number;
    executionTime: number;
    success: boolean;
  };
}

export class FormulaProcessor {
  private static instance: FormulaProcessor | null = null;
  private stats!: FormulaProcessorStats;
  private variableCache = new Map<string, number>();
  private functionCache = new Map<string, (...args: any[]) => any>();
  private _isInitialized = false;

  constructor() {
    if (FormulaProcessor.instance) {
      return FormulaProcessor.instance;
    }
    FormulaProcessor.instance = this;

    this.stats = {
      totalFormulas: 0,
      successfulFormulas: 0,
      failedFormulas: 0,
      averageExecutionTime: 0,
      totalExecutionTime: 0,
      complexityDistribution: {},
      errorTypes: {},
    };

    this.initialize();
  }

  /**
   * Ініціалізація обробника формул
   */
  private initialize(): void {
    try {
      logger.info('🧮 Ініціалізація FormulaProcessor...');

      // Ініціалізація математичних функцій
      this.initializeMathFunctions();

      // Ініціалізація кешу змінних
      this.initializeVariableCache();

      this._isInitialized = true;
      logger.info('✅ FormulaProcessor успішно ініціалізовано');
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'initialize' },
      });
      throw new Error('Помилка ініціалізації FormulaProcessor');
    }
  }

  /**
   * Ініціалізація математичних функцій
   */
  private initializeMathFunctions(): void {
    try {
      const mathFunctions: Record<string, (...args: any[]) => any> = {
        // Тригонометричні функції
        sin: Math.sin,
        cos: Math.cos,
        tan: Math.tan,
        asin: Math.asin,
        acos: Math.acos,
        atan: Math.atan,

        // Логарифмічні функції
        log: Math.log10,
        ln: Math.log,
        exp: Math.exp,

        // Степеневі функції
        sqrt: Math.sqrt,
        pow: Math.pow,
        abs: Math.abs,

        // Округлення
        floor: Math.floor,
        ceil: Math.ceil,
        round: Math.round,

        // Статистичні функції
        min: Math.min,
        max: Math.max,

        // Умовні функції
        if: (condition: boolean, trueValue: number, falseValue: number) =>
          condition ? trueValue : falseValue,
      };

      for (const [name, func] of Object.entries(mathFunctions)) {
        this.functionCache.set(name, func);
      }

      logger.debug(`📚 Ініціалізовано ${this.functionCache.size} математичних функцій`);
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'initializeMathFunctions' },
      });
    }
  }

  /**
   * Ініціалізація кешу змінних
   */
  private initializeVariableCache(): void {
    try {
      // Додавання констант
      this.variableCache.set('PI', Math.PI);
      this.variableCache.set('E', Math.E);
      this.variableCache.set('INFINITY', Infinity);
      this.variableCache.set('NAN', NaN);

      logger.debug(`📊 Ініціалізовано ${this.variableCache.size} констант`);
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'initializeVariableCache' },
      });
    }
  }

  /**
   * Валідація формули
   */
  public validateFormula(formula: string): FormulaValidationResult {
    const startTime = performance.now();

    try {
      logger.debug('🔍 Валідація формули...', {
        formula: formula.substring(0, 100),
        length: formula.length,
      } as LogMeta);

      const errors: string[] = [];
      const warnings: string[] = [];
      let sanitizedFormula = formula;

      // Перевірка довжини
      if (formula.length > FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH) {
        errors.push(
          `Формула занадто довга (${formula.length} символів, максимум ${FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH})`
        );
        sanitizedFormula = formula.substring(0, FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH);
      }

      // Валідація введення
      const validation = validateInput(formula, { inputType: 'command' });
      if (!validation.isValid) {
        errors.push(...validation.errors);
      }

      // Перевірка на небезпечні патерни
      const dangerousPatterns = [
        /eval\s*\(/i,
        /Function\s*\(/i,
        /setTimeout\s*\(/i,
        /setInterval\s*\(/i,
        /process\./i,
        /require\s*\(/i,
        /import\s*\(/i,
      ];

      for (const pattern of dangerousPatterns) {
        if (pattern.test(formula)) {
          errors.push('Формула містить небезпечні патерни');
          break;
        }
      }

      // Аналіз змінних та функцій
      const variables = this.extractVariables(formula);
      const functions = this.extractFunctions(formula);

      // Перевірка змінних
      for (const variable of variables) {
        if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_VARIABLES.test(variable)) {
          errors.push(`Недозволена змінна: ${variable}`);
        }
      }

      if (variables.length > FORMULA_PROCESSOR_CONSTANTS.MAX_VARIABLES) {
        errors.push(
          `Занадто багато змінних (${variables.length}, максимум ${FORMULA_PROCESSOR_CONSTANTS.MAX_VARIABLES})`
        );
      }

      // Перевірка функцій
      for (const func of functions) {
        if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_FUNCTIONS.includes(func)) {
          errors.push(`Недозволена функція: ${func}`);
        }
      }

      // Перевірка складності
      const complexity = this.calculateComplexity(formula);
      if (complexity > 100) {
        warnings.push(`Висока складність формули: ${complexity}`);
      }

      const duration = performance.now() - startTime;

      const result: FormulaValidationResult = {
        isValid: errors.length === 0,
        errors,
        warnings,
        sanitizedFormula,
        variables,
        functions,
        complexity,
      };

      if (errors.length > 0) {
        logger.warn('❌ Валідація формули невдала', {
          errors,
          warnings,
          formula: formula.substring(0, 100),
        } as LogMeta);
      } else {
        logger.debug('✅ Валідація формули успішна', {
          variables: variables.length,
          functions: functions.length,
          complexity,
          duration: `${duration.toFixed(2)}ms`,
        } as LogMeta);
      }

      return result;
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'validateFormula', formula: formula.substring(0, 100) },
      });

      return {
        isValid: false,
        errors: ['Помилка валідації формули'],
        warnings: [],
        sanitizedFormula: '',
        variables: [],
        functions: [],
        complexity: 0,
      };
    }
  }

  /**
   * Виконання формули
   */
  public async evaluateFormula(
    formula: string,
    variables: Record<string, number> = {}
  ): Promise<FormulaResult> {
    const startTime = performance.now();
    const operationId = this.generateOperationId(formula);

    try {
      logger.debug('🧮 Початок обчислення формули...', {
        formula: formula.substring(0, 100),
        variablesCount: Object.keys(variables).length,
        operationId,
      } as LogMeta);

      // Валідація формули
      const validation = this.validateFormula(formula);
      if (!validation.isValid) {
        throw new Error(`Формула не валідна: ${validation.errors.join(', ')}`);
      }

      // Об'єднання змінних
      const allVariables: Record<string, number> = {
        ...Object.fromEntries(this.variableCache),
        ...variables,
      };

      // Створення безпечного контексту виконання
      const result = await this.executeFormula(formula, allVariables);

      const duration = performance.now() - startTime;
      const complexity = validation.complexity;

      const formulaResult: FormulaResult = {
        success: true,
        result,
        warnings: validation.warnings,
        variables: allVariables,
        executionTime: duration,
        complexity,
      };

      this.updateStats(true, duration, complexity);
      this.stats.lastFormula = {
        formula: formula.substring(0, 100),
        result,
        executionTime: duration,
        success: true,
      };

      logger.info('✅ Формула успішно обчислена', {
        formula: formula.substring(0, 100),
        result,
        duration: `${duration.toFixed(2)}ms`,
        complexity,
        operationId,
      } as LogMeta);

      return formulaResult;
    } catch (error) {
      const duration = performance.now() - startTime;
      this.updateStats(false, duration, 0);

      const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

      logger.error('❌ Помилка обчислення формули', {
        formula: formula.substring(0, 100),
        error: errorMessage,
        duration: `${duration.toFixed(2)}ms`,
        operationId,
      } as LogMeta);

      return {
        success: false,
        error: errorMessage,
        warnings: [],
        variables: {},
        executionTime: duration,
        complexity: 0,
      };
    }
  }

  /**
   * Безпечне виконання формули
   */
  private async executeFormula(
    formula: string,
    variables: Record<string, number>
  ): Promise<number> {
    // Безпечний обчислювач на основі алгоритму shunting-yard (RPN)
    // Підтримує лише дозволені оператори/функції/змінні
    const tokenize = (expr: string): Array<{ t: 'num'|'id'|'op'|'lp'|'rp'|'comma'; v: string }> => {
      const out: Array<{ t: 'num'|'id'|'op'|'lp'|'rp'|'comma'; v: string }> = [];
      let i = 0;
      while (i < expr.length) {
        const ch = expr.charAt(i);
        if (/\s/.test(ch)) { i++; continue; }
        if (/[0-9.]/.test(ch)) {
          let j = i + 1;
          while (j < expr.length && /[0-9.]/.test(expr.charAt(j))) j++;
          out.push({ t: 'num', v: expr.slice(i, j) });
          i = j; continue;
        }
        if (/[a-zA-Z_]/.test(ch)) {
          let j = i + 1;
          while (j < expr.length && /[a-zA-Z0-9_]/.test(expr.charAt(j))) j++;
          out.push({ t: 'id', v: expr.slice(i, j) });
          i = j; continue;
        }
        if (ch === ',') { out.push({ t: 'comma', v: ',' }); i++; continue; }
        if (ch === '(') { out.push({ t: 'lp', v: '(' }); i++; continue; }
        if (ch === ')') { out.push({ t: 'rp', v: ')' }); i++; continue; }
        if (['+','-','*','/','^','=','<','>','!'].includes(ch)) {
          // обробка двосимвольних операторів
          const two = expr.slice(i, i+2);
          if (["<=",">=","!=","=="].includes(two)) { out.push({ t: 'op', v: two }); i += 2; continue; }
          out.push({ t: 'op', v: ch }); i++; continue;
        }
        throw new Error(`Недопустимий символ у формулі: '${ch}'`);
      }
      return out;
    };

    const precedence: Record<string, number> = { '==':1,'!=':1,'<':1,'>':1,'<=':1,'>=':1,'+':2,'-':2,'*':3,'/':3,'^':4 };
    const rightAssoc = new Set(['^']);

    const toRpn = (tokens: ReturnType<typeof tokenize>): Array<{ t: 'num'|'id'|'op'|'func'; v: string; argc?: number }> => {
      const out: Array<{ t: 'num'|'id'|'op'|'func'; v: string; argc?: number }> = [];
      const ops: Array<{ kind:'op'|'func'|'lp'; v: string; argc?: number }> = [];
      // функції розпізнаємо як id, за якими йде lp
      for (let i = 0; i < tokens.length; i++) {
        const tok = tokens[i]!;
        if (tok.t === 'num' || tok.t === 'id') {
          // lookahead: if id followed by lp -> function
          if (tok.t === 'id' && tokens[i+1]?.t === 'lp') {
            if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_FUNCTIONS.includes(tok.v)) {
              throw new Error(`Недозволена функція: ${tok.v}`);
            }
            ops.push({ kind: 'func', v: tok.v, argc: 0 });
          } else {
            out.push({ t: tok.t, v: tok.v });
          }
        } else if (tok.t === 'comma') {
          // виводимо оператори до найближчої lp; інкрементуємо argc у поточній функції
          while (ops.length) {
            const last = ops[ops.length - 1]!;
            if (last.kind === 'lp') break;
            const top = ops.pop()!;
            if (top.kind === 'op') out.push({ t: 'op', v: top.v });
            else if (top.kind === 'func') out.push({ t: 'func', v: top.v, argc: top.argc ?? 0 });
          }
          const funcIdx = [...ops].reverse().findIndex(x => x.kind === 'func');
          if (funcIdx >= 0) {
            const idx = ops.length - 1 - funcIdx;
            const target = ops[idx];
            if (target) target.argc = (target.argc ?? 0) + 1;
          }
        } else if (tok.t === 'op') {
          while (ops.length && ops[ops.length-1]!.kind === 'op') {
            const top = ops[ops.length-1]!;
            const p1 = precedence[tok.v] ?? 0;
            const p2 = precedence[top.v] ?? 0;
            if ((rightAssoc.has(tok.v) ? p1 < p2 : p1 <= p2)) {
              out.push({ t: 'op', v: ops.pop()!.v });
            } else break;
          }
          ops.push({ kind: 'op', v: tok.v });
        } else if (tok.t === 'lp') {
          ops.push({ kind: 'lp', v: '(' });
        } else if (tok.t === 'rp') {
          while (ops.length) {
            const last = ops[ops.length - 1]!;
            if (last.kind === 'lp') break;
            const top = ops.pop()!;
            if (top.kind === 'op') out.push({ t: 'op', v: top.v });
            else if (top.kind === 'func') out.push({ t: 'func', v: top.v, argc: top.argc ?? 0 });
          }
          if (!ops.length) throw new Error('Невірна дужкова структура');
          ops.pop(); // remove lp
          // якщо зверху функція — вона завершилась
          if (ops.length) {
            const last = ops[ops.length - 1]!;
            if (last.kind === 'func') {
              const fn = ops.pop()!;
              out.push({ t: 'func', v: fn.v, argc: (fn.argc ?? 0) + 1 });
            }
          }
        }
      }
      while (ops.length) {
        const top = ops.pop()!;
        if (top.kind === 'lp') throw new Error('Невірна дужкова структура');
        if (top.kind === 'op') out.push({ t: 'op', v: top.v });
        else out.push({ t: 'func', v: top.v, argc: top.argc ?? 0 });
      }
      return out;
    };

    const evaluateRpn = (rpn: Array<{ t: 'num'|'id'|'op'|'func'; v: string; argc?: number }>): number => {
      const stack: number[] = [];
      const funcs = Object.fromEntries(this.functionCache);
      const getVal = (id: string): number => {
        if (id in variables) return variables[id]!;
        if (['PI','E','INFINITY','NAN'].includes(id)) return this.variableCache.get(id)!;
        throw new Error(`Невідома змінна: ${id}`);
      };
      for (const tok of rpn) {
        if (tok.t === 'num') {
          const n = Number(tok.v);
          if (!isFinite(n)) throw new Error('Недійсне число');
          stack.push(n);
        } else if (tok.t === 'id') {
          stack.push(getVal(tok.v));
        } else if (tok.t === 'op') {
          const b = stack.pop();
          const a = stack.pop();
          if (a === undefined || b === undefined) throw new Error('Невірний вираз');
          let r: number | boolean;
          switch (tok.v) {
            case '+': r = a + b; break;
            case '-': r = a - b; break;
            case '*': r = a * b; break;
            case '/': r = a / b; break;
            case '^': r = Math.pow(a, b); break;
            case '==': r = a === b; break;
            case '!=': r = a !== b; break;
            case '<': r = a < b; break;
            case '>': r = a > b; break;
            case '<=': r = a <= b; break;
            case '>=': r = a >= b; break;
            default: throw new Error(`Оператор не підтримується: ${tok.v}`);
          }
          stack.push(typeof r === 'boolean' ? (r ? 1 : 0) : r);
        } else if (tok.t === 'func') {
          const argc = tok.argc ?? 0;
          const args: number[] = [];
          for (let i = 0; i < argc; i++) {
            const v = stack.pop();
            if (v === undefined) throw new Error('Невірна кількість аргументів функції');
            args.push(v);
          }
          args.reverse();
          const fn = funcs[tok.v] as unknown as (...xs: number[]) => number;
          if (typeof fn !== 'function') throw new Error(`Функція не знайдена: ${tok.v}`);
          const res = fn(...args);
          if (typeof res !== 'number' || !isFinite(res)) throw new Error('Невірний результат функції');
          stack.push(res);
        }
      }
      if (stack.length !== 1) throw new Error('Невірний вираз');
      return stack[0]!;
    };

    return new Promise((resolve, reject) => {
      const controller = setTimeout(() => reject(new Error('Таймаут виконання формули')), FORMULA_PROCESSOR_CONSTANTS.TIMEOUT);
      try {
        const tokens = tokenize(formula);
        const rpn = toRpn(tokens);
        const result = evaluateRpn(rpn);
        clearTimeout(controller);
        const roundedResult = Math.round(result * Math.pow(10, FORMULA_PROCESSOR_CONSTANTS.PRECISION)) / Math.pow(10, FORMULA_PROCESSOR_CONSTANTS.PRECISION);
        resolve(roundedResult);
      } catch (e) {
        clearTimeout(controller);
        reject(e);
      }
    });
  }

  /**
   * Витяг змінних з формули
   */
  private extractVariables(formula: string): string[] {
    const variables = new Set<string>();
    const variablePattern = /[a-zA-Z_][a-zA-Z0-9_]*/g;
    const matches = formula.match(variablePattern) || [];

    for (const match of matches) {
      // Виключення функцій та констант
      if (
        !FORMULA_PROCESSOR_CONSTANTS.ALLOWED_FUNCTIONS.includes(match) &&
        !['PI', 'E', 'INFINITY', 'NAN'].includes(match)
      ) {
        variables.add(match);
      }
    }

    return Array.from(variables);
  }

  /**
   * Витяг функцій з формули
   */
  private extractFunctions(formula: string): string[] {
    const functions = new Set<string>();
    const functionPattern = /[a-zA-Z_][a-zA-Z0-9_]*\s*\(/g;
    const matches = formula.match(functionPattern) || [];

    for (const match of matches) {
      const functionName = match.replace(/\s*\($/, '');
      functions.add(functionName);
    }

    return Array.from(functions);
  }

  /**
   * Розрахунок складності формули
   */
  private calculateComplexity(formula: string): number {
    let complexity = 0;

    // Базова складність за довжину
    complexity += formula.length * 0.1;

    // Складність за операції
    const operators = ['+', '-', '*', '/', '^', '(', ')'];
    for (const op of operators) {
      const count = (formula.match(new RegExp(`\\${op}`, 'g')) || []).length;
      complexity += count * 2;
    }

    // Складність за функції
    const functions = this.extractFunctions(formula);
    complexity += functions.length * 5;

    // Складність за змінні
    const variables = this.extractVariables(formula);
    complexity += variables.length * 3;

    // Складність за вкладеність
    const nestedLevel = this.calculateNestedLevel(formula);
    complexity += nestedLevel * 10;

    return Math.round(complexity);
  }

  /**
   * Розрахунок рівня вкладеності
   */
  private calculateNestedLevel(formula: string): number {
    let maxLevel = 0;
    let currentLevel = 0;

    for (const char of formula) {
      if (char === '(') {
        currentLevel++;
        maxLevel = Math.max(maxLevel, currentLevel);
      } else if (char === ')') {
        currentLevel--;
        if (currentLevel < 0) {
          return FORMULA_PROCESSOR_CONSTANTS.MAX_NESTED_LEVELS + 1; // Помилка
        }
      }
    }

    return maxLevel;
  }

  /**
   * Генерація ID операції
   */
  private generateOperationId(formula: string): string {
    const timestamp = Date.now();
    const hash = require('crypto')
      .createHash('md5')
      .update(`${formula}:${timestamp}`)
      .digest('hex');
    return `formula_${hash.substring(0, 8)}`;
  }

  /**
   * Оновлення статистики
   */
  private updateStats(success: boolean, duration: number, complexity: number): void {
    try {
      this.stats.totalFormulas++;
      this.stats.totalExecutionTime += duration;
      this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.totalFormulas;

      if (success) {
        this.stats.successfulFormulas++;
      } else {
        this.stats.failedFormulas++;
      }

      // Розподіл складності
      const complexityLevel =
        complexity < 10
          ? 'low'
          : complexity < 50
            ? 'medium'
            : complexity < 100
              ? 'high'
              : 'very_high';

      this.stats.complexityDistribution[complexityLevel] =
        (this.stats.complexityDistribution[complexityLevel] || 0) + 1;
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'updateStats' },
      });
    }
  }

  /**
   * Встановлення змінної
   */
  public setVariable(name: string, value: number): void {
    try {
      if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_VARIABLES.test(name)) {
        throw new Error(`Недозволена змінна: ${name}`);
      }

      this.variableCache.set(name, value);
      logger.debug(`📊 Встановлено змінну: ${name} = ${value}`);
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'setVariable', name, value },
      });
    }
  }

  /**
   * Отримання значення змінної
   */
  public getVariable(name: string): number | undefined {
    return this.variableCache.get(name);
  }

  /**
   * Очищення змінних
   */
  public clearVariables(): void {
    try {
      this.variableCache.clear();
      this.initializeVariableCache(); // Відновлення констант
      logger.info('🧹 Змінні очищено');
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'clearVariables' },
      });
    }
  }

  /**
   * Отримання статистики
   */
  public getStats(): FormulaProcessorStats {
    return { ...this.stats };
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    try {
      this.variableCache.clear();
      this.functionCache.clear();

      logger.info('🧹 Ресурси FormulaProcessor очищено');
    } catch (error) {
      handleError(error, {
        serviceName: 'FormulaProcessor',
        additionalContext: { operation: 'cleanup' },
      });
    }
  }

  /**
   * Перевірка стану ініціалізації
   */
  public isInitialized(): boolean {
    return this._isInitialized;
  }
}

// Експорт єдиного екземпляра
export const formulaProcessor = new FormulaProcessor();

// Експорт функцій для зручності
export const validateFormula = (formula: string) => formulaProcessor.validateFormula(formula);
export const evaluateFormula = async (formula: string, variables?: Record<string, number>) => {
  const res = await formulaProcessor.evaluateFormula(formula, variables);
  // Підтримка зворотної сумісності з тестами: повертати лише числовий результат
  return res.result!;
};
export const setVariable = (name: string, value: number) =>
  formulaProcessor.setVariable(name, value);
export const getVariable = (name: string) => formulaProcessor.getVariable(name);
export const clearVariables = () => formulaProcessor.clearVariables();
export const getFormulaProcessorStats = () => formulaProcessor.getStats();
export const cleanupFormulaProcessor = () => formulaProcessor.cleanup();
