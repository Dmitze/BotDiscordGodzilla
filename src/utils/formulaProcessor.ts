/**
 * Утиліта для обробки складних формул Google Sheets
 * Підтримує ARRAY_CONSTRAIN, SUMPRODUCT, MOD, COLUMN та інші функції
 * TypeScript версія 3.0.0
 */

import logger from './logger';
import { GoogleService } from '@/services/GoogleService';
import { AIService } from '@/services/AIService';
import { sanitizeInput } from './security';

// Константи для обробки формул
const FORMULA_CONFIG = {
  MAX_FORMULA_LENGTH: 50000,
  MAX_SHEETS_PER_FORMULA: 50,
  MAX_RANGE_SIZE: 1000,
  TIMEOUT: 30000, // 30 секунд
  CACHE_TTL: 5 * 60 * 1000, // 5 хвилин
} as const;

interface FormulaToken {
  type: 'function' | 'operator' | 'range' | 'value' | 'sheet' | 'cell';
  value: string;
  position: number;
  parameters?: FormulaToken[];
}

interface ParsedFormula {
  tokens: FormulaToken[];
  sheets: string[];
  ranges: string[];
  functions: string[];
  complexity: number;
}

interface FormulaResult {
  value: number;
  breakdown: Record<string, number>;
  processingTime: number;
  cache: boolean;
  error?: string;
}

interface FormulaCache {
  [key: string]: {
    result: FormulaResult;
    timestamp: number;
  };
}

class FormulaProcessor {
  private googleService: GoogleService;
  private aiService: AIService;
  private cache: FormulaCache = {};
  private processingFormulas: Set<string> = new Set();

  constructor() {
    this.googleService = new GoogleService();
    this.aiService = new AIService();
  }

  /**
   * Основний метод обробки формули
   */
  public async processFormula(formula: string): Promise<FormulaResult> {
    const startTime = performance.now();
    const formulaHash = this.hashFormula(formula);

    try {
      logger.info('Початок обробки формули', { 
        formulaLength: formula.length,
        hash: formulaHash.substring(0, 8)
      });

      // Перевірка кешу
      const cached = this.getCachedResult(formulaHash);
      if (cached) {
        logger.debug('Результат знайдено в кеші');
        return { ...cached, cache: true };
      }

      // Перевірка на рекурсію
      if (this.processingFormulas.has(formulaHash)) {
        throw new Error('Виявлено рекурсивну формулу');
      }

      this.processingFormulas.add(formulaHash);

      // Парсинг формули
      const parsed = this.parseFormula(formula);
      
      // Валідація
      this.validateFormula(parsed);
      
      // Виконання
      const result = await this.executeFormula(parsed);
      
      // Кешування результату
      this.cacheResult(formulaHash, result);

      const processingTime = performance.now() - startTime;
      
      logger.info('Формула оброблена успішно', {
        processingTime: `${processingTime.toFixed(2)}ms`,
        complexity: parsed.complexity,
        sheets: parsed.sheets.length,
      });

      return {
        ...result,
        processingTime,
        cache: false,
      };

    } catch (error) {
      const processingTime = performance.now() - startTime;
      logger.error('Помилка обробки формули:', error);
      
      return {
        value: 0,
        breakdown: {},
        processingTime,
        cache: false,
        error: error instanceof Error ? error.message : 'Невідома помилка',
      };
    } finally {
      this.processingFormulas.delete(formulaHash);
    }
  }

  /**
   * Парсинг формули на токени
   */
  private parseFormula(formula: string): ParsedFormula {
    const tokens: FormulaToken[] = [];
    const sheets = new Set<string>();
    const ranges = new Set<string>();
    const functions = new Set<string>();
    let complexity = 0;

    try {
      // Очищення формули
      const cleanFormula = this.cleanFormula(formula);
      
      // Розбиття на частини
      const parts = this.splitFormula(cleanFormula);
      
      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        const token = this.parseToken(part, i);
        
        if (token) {
          tokens.push(token);
          
          // Збір статистики
          if (token.type === 'function') {
            functions.add(token.value);
            complexity += this.getFunctionComplexity(token.value);
          } else if (token.type === 'sheet') {
            sheets.add(token.value);
          } else if (token.type === 'range') {
            ranges.add(token.value);
          }
        }
      }

      return {
        tokens,
        sheets: Array.from(sheets),
        ranges: Array.from(ranges),
        functions: Array.from(functions),
        complexity,
      };

    } catch (error) {
      logger.error('Помилка парсингу формули:', error);
      throw new Error('Неможливо розібрати формулу');
    }
  }

  /**
   * Очищення формули
   */
  private cleanFormula(formula: string): string {
    // Видалення зайвих пробілів
    let cleaned = formula.replace(/\s+/g, ' ').trim();
    
    // Видалення коментарів
    cleaned = cleaned.replace(/\/\*.*?\*\//g, '');
    
    // Нормалізація лапок
    cleaned = cleaned.replace(/[""]/g, '"');
    
    return cleaned;
  }

  /**
   * Розбиття формули на частини
   */
  private splitFormula(formula: string): string[] {
    const parts: string[] = [];
    let current = '';
    let parentheses = 0;
    let inQuotes = false;
    
    for (let i = 0; i < formula.length; i++) {
      const char = formula[i];
      
      if (char === '"' && (i === 0 || formula[i - 1] !== '\\')) {
        inQuotes = !inQuotes;
      }
      
      if (!inQuotes) {
        if (char === '(') parentheses++;
        if (char === ')') parentheses--;
        
        if (char === '+' && parentheses === 0) {
          if (current.trim()) {
            parts.push(current.trim());
            current = '';
          }
          continue;
        }
      }
      
      current += char;
    }
    
    if (current.trim()) {
      parts.push(current.trim());
    }
    
    return parts;
  }

  /**
   * Парсинг окремого токена
   */
  private parseToken(part: string, position: number): FormulaToken | null {
    // Функції
    if (part.startsWith('SUMPRODUCT(')) {
      return this.parseSUMPRODUCT(part, position);
    }
    
    if (part.startsWith('ARRAY_CONSTRAIN(')) {
      return this.parseARRAY_CONSTRAIN(part, position);
    }
    
    if (part.startsWith('MOD(')) {
      return this.parseMOD(part, position);
    }
    
    if (part.startsWith('COLUMN(')) {
      return this.parseCOLUMN(part, position);
    }
    
    // Аркуші
    const sheetMatch = part.match(/'([^']+)'!/);
    if (sheetMatch) {
      return {
        type: 'sheet',
        value: sheetMatch[1],
        position,
      };
    }
    
    // Діапазони
    const rangeMatch = part.match(/[A-Z]+\d+:[A-Z]+\d+/);
    if (rangeMatch) {
      return {
        type: 'range',
        value: rangeMatch[0],
        position,
      };
    }
    
    // Оператори
    if (['+', '-', '*', '/', '='].includes(part)) {
      return {
        type: 'operator',
        value: part,
        position,
      };
    }
    
    // Значення
    if (/^\d+(\.\d+)?$/.test(part)) {
      return {
        type: 'value',
        value: part,
        position,
      };
    }
    
    return null;
  }

  /**
   * Парсинг SUMPRODUCT
   */
  private parseSUMPRODUCT(part: string, position: number): FormulaToken {
    const content = part.slice(12, -1); // Видаляємо SUMPRODUCT( і )
    
    return {
      type: 'function',
      value: 'SUMPRODUCT',
      position,
      parameters: this.parseParameters(content),
    };
  }

  /**
   * Парсинг ARRAY_CONSTRAIN
   */
  private parseARRAY_CONSTRAIN(part: string, position: number): FormulaToken {
    const content = part.slice(16, -1); // Видаляємо ARRAY_CONSTRAIN( і )
    
    return {
      type: 'function',
      value: 'ARRAY_CONSTRAIN',
      position,
      parameters: this.parseParameters(content),
    };
  }

  /**
   * Парсинг MOD
   */
  private parseMOD(part: string, position: number): FormulaToken {
    const content = part.slice(4, -1); // Видаляємо MOD( і )
    
    return {
      type: 'function',
      value: 'MOD',
      position,
      parameters: this.parseParameters(content),
    };
  }

  /**
   * Парсинг COLUMN
   */
  private parseCOLUMN(part: string, position: number): FormulaToken {
    const content = part.slice(7, -1); // Видаляємо COLUMN( і )
    
    return {
      type: 'function',
      value: 'COLUMN',
      position,
      parameters: this.parseParameters(content),
    };
  }

  /**
   * Парсинг параметрів
   */
  private parseParameters(content: string): FormulaToken[] {
    const parameters: FormulaToken[] = [];
    let current = '';
    let parentheses = 0;
    
    for (let i = 0; i < content.length; i++) {
      const char = content[i];
      
      if (char === '(') parentheses++;
      if (char === ')') parentheses--;
      
      if (char === ',' && parentheses === 0) {
        if (current.trim()) {
          const token = this.parseToken(current.trim(), parameters.length);
          if (token) parameters.push(token);
        }
        current = '';
      } else {
        current += char;
      }
    }
    
    if (current.trim()) {
      const token = this.parseToken(current.trim(), parameters.length);
      if (token) parameters.push(token);
    }
    
    return parameters;
  }

  /**
   * Валідація формули
   */
  private validateFormula(parsed: ParsedFormula): void {
    // Перевірка довжини
    if (parsed.tokens.length > FORMULA_CONFIG.MAX_FORMULA_LENGTH) {
      throw new Error('Формула занадто довга');
    }
    
    // Перевірка кількості аркушів
    if (parsed.sheets.length > FORMULA_CONFIG.MAX_SHEETS_PER_FORMULA) {
      throw new Error('Занадто багато аркушів у формулі');
    }
    
    // Перевірка складності
    if (parsed.complexity > 1000) {
      throw new Error('Формула занадто складна');
    }
    
    // Перевірка підтримуваних функцій
    const supportedFunctions = ['SUMPRODUCT', 'ARRAY_CONSTRAIN', 'MOD', 'COLUMN', 'ARRAYFORMULA'];
    const unsupported = parsed.functions.filter(f => !supportedFunctions.includes(f));
    
    if (unsupported.length > 0) {
      throw new Error(`Непідтримувані функції: ${unsupported.join(', ')}`);
    }
  }

  /**
   * Виконання формули
   */
  private async executeFormula(parsed: ParsedFormula): Promise<{ value: number; breakdown: Record<string, number> }> {
    let total = 0;
    const breakdown: Record<string, number> = {};
    
    try {
      // Обробка кожного токена
      for (const token of parsed.tokens) {
        switch (token.type) {
          case 'function':
            const result = await this.executeFunction(token);
            total += result;
            breakdown[token.value] = result;
            break;
            
          case 'value':
            total += parseFloat(token.value) || 0;
            break;
            
          case 'operator':
            // Обробка операторів буде в наступній ітерації
            break;
        }
      }
      
      return { value: total, breakdown };
      
    } catch (error) {
      logger.error('Помилка виконання формули:', error);
      throw error;
    }
  }

  /**
   * Виконання функції
   */
  private async executeFunction(token: FormulaToken): Promise<number> {
    switch (token.value) {
      case 'SUMPRODUCT':
        return await this.executeSUMPRODUCT(token);
      case 'ARRAY_CONSTRAIN':
        return await this.executeARRAY_CONSTRAIN(token);
      case 'MOD':
        return await this.executeMOD(token);
      case 'COLUMN':
        return await this.executeCOLUMN(token);
      default:
        throw new Error(`Непідтримувана функція: ${token.value}`);
    }
  }

  /**
   * Виконання SUMPRODUCT
   */
  private async executeSUMPRODUCT(token: FormulaToken): Promise<number> {
    if (!token.parameters || token.parameters.length < 2) {
      throw new Error('SUMPRODUCT потребує мінімум 2 параметри');
    }
    
    let result = 0;
    
    // Отримання даних з аркушів
    for (const param of token.parameters) {
      if (param.type === 'range') {
        const sheetData = await this.getSheetDataForRange(param.value);
        result += this.calculateSUMPRODUCT(sheetData);
      }
    }
    
    return result;
  }

  /**
   * Виконання ARRAY_CONSTRAIN
   */
  private async executeARRAY_CONSTRAIN(token: FormulaToken): Promise<number> {
    if (!token.parameters || token.parameters.length < 3) {
      throw new Error('ARRAY_CONSTRAIN потребує 3 параметри');
    }
    
    // ARRAY_CONSTRAIN(array, rows, cols)
    const arrayParam = token.parameters[0];
    const rowsParam = token.parameters[1];
    const colsParam = token.parameters[2];
    
    if (arrayParam.type !== 'function') {
      throw new Error('Перший параметр ARRAY_CONSTRAIN має бути функцією');
    }
    
    const arrayResult = await this.executeFunction(arrayParam);
    const rows = parseInt(rowsParam.value) || 1;
    const cols = parseInt(colsParam.value) || 1;
    
    return arrayResult * rows * cols;
  }

  /**
   * Виконання MOD
   */
  private async executeMOD(token: FormulaToken): Promise<number> {
    if (!token.parameters || token.parameters.length < 2) {
      throw new Error('MOD потребує 2 параметри');
    }
    
    const dividend = parseFloat(token.parameters[0].value) || 0;
    const divisor = parseFloat(token.parameters[1].value) || 1;
    
    return dividend % divisor;
  }

  /**
   * Виконання COLUMN
   */
  private async executeCOLUMN(token: FormulaToken): Promise<number> {
    if (!token.parameters || token.parameters.length === 0) {
      throw new Error('COLUMN потребує параметр');
    }
    
    const range = token.parameters[0].value;
    return this.getColumnIndex(range);
  }

  /**
   * Отримання даних аркуша для діапазону
   */
  private async getSheetDataForRange(range: string): Promise<number[][]> {
    // Тут буде логіка отримання даних з Google Sheets
    // Поки що повертаємо заглушку
    return [[1, 2, 3], [4, 5, 6]];
  }

  /**
   * Розрахунок SUMPRODUCT
   */
  private calculateSUMPRODUCT(data: number[][]): number {
    let result = 0;
    
    for (const row of data) {
      for (const cell of row) {
        result += cell;
      }
    }
    
    return result;
  }

  /**
   * Отримання індексу стовпця
   */
  private getColumnIndex(range: string): number {
    const match = range.match(/^([A-Z]+)/);
    if (!match) return 0;
    
    const column = match[1];
    let index = 0;
    
    for (let i = 0; i < column.length; i++) {
      index = index * 26 + (column.charCodeAt(i) - 64);
    }
    
    return index;
  }

  /**
   * Отримання складності функції
   */
  private getFunctionComplexity(functionName: string): number {
    const complexityMap: Record<string, number> = {
      'SUMPRODUCT': 10,
      'ARRAY_CONSTRAIN': 5,
      'MOD': 1,
      'COLUMN': 1,
      'ARRAYFORMULA': 15,
    };
    
    return complexityMap[functionName] || 1;
  }

  /**
   * Хешування формули
   */
  private hashFormula(formula: string): string {
    // Простий хеш для кешування
    let hash = 0;
    for (let i = 0; i < formula.length; i++) {
      const char = formula.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Конвертація в 32-бітне число
    }
    return hash.toString(36);
  }

  /**
   * Отримання результату з кешу
   */
  private getCachedResult(hash: string): FormulaResult | null {
    const cached = this.cache[hash];
    if (!cached) return null;
    
    if (Date.now() - cached.timestamp > FORMULA_CONFIG.CACHE_TTL) {
      delete this.cache[hash];
      return null;
    }
    
    return cached.result;
  }

  /**
   * Кешування результату
   */
  private cacheResult(hash: string, result: FormulaResult): void {
    this.cache[hash] = {
      result,
      timestamp: Date.now(),
    };
    
    // Очищення старого кешу
    this.cleanupCache();
  }

  /**
   * Очищення кешу
   */
  private cleanupCache(): void {
    const now = Date.now();
    const keysToDelete = Object.keys(this.cache).filter(key => 
      now - this.cache[key].timestamp > FORMULA_CONFIG.CACHE_TTL
    );
    
    keysToDelete.forEach(key => delete this.cache[key]);
    
    if (keysToDelete.length > 0) {
      logger.debug(`Очищено ${keysToDelete.length} записів кешу формул`);
    }
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    this.cache = {};
    this.processingFormulas.clear();
    logger.info('Кеш формул очищено');
  }

  /**
   * Отримання статистики
   */
  public getStats(): {
    cacheSize: number;
    processingFormulas: number;
    cacheHits: number;
    cacheMisses: number;
  } {
    return {
      cacheSize: Object.keys(this.cache).length,
      processingFormulas: this.processingFormulas.size,
      cacheHits: 0, // Буде оновлюватися в реальному часі
      cacheMisses: 0, // Буде оновлюватися в реальному часі
    };
  }
}

export default FormulaProcessor; 