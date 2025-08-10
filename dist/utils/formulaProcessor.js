"use strict";
/**
 * Розширена система обробки формул для Discord AI Assistant Bot
 * Безпечна обробка математичних виразів та формул
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.cleanupFormulaProcessor = exports.getFormulaProcessorStats = exports.clearVariables = exports.getVariable = exports.setVariable = exports.evaluateFormula = exports.validateFormula = exports.formulaProcessor = exports.FormulaProcessor = void 0;
const errorHandler_1 = require("./errorHandler");
const logger_1 = __importDefault(require("./logger"));
const security_1 = require("./security");
// Константи для обробки формул
const FORMULA_PROCESSOR_CONSTANTS = {
    MAX_FORMULA_LENGTH: 1000,
    MAX_NESTED_LEVELS: 10,
    MAX_ITERATIONS: 1000,
    ALLOWED_FUNCTIONS: [
        'sin', 'cos', 'tan', 'asin', 'acos', 'atan',
        'sqrt', 'pow', 'exp', 'log', 'ln', 'abs',
        'floor', 'ceil', 'round', 'min', 'max',
        'sum', 'avg', 'count', 'if', 'case',
    ],
    ALLOWED_OPERATORS: ['+', '-', '*', '/', '^', '(', ')', '=', '<', '>', '<=', '>=', '!=', '=='],
    ALLOWED_VARIABLES: /^[a-zA-Z_][a-zA-Z0-9_]*$/,
    MAX_VARIABLES: 50,
    PRECISION: 10,
    TIMEOUT: 5000, // 5 секунд
};
class FormulaProcessor {
    constructor() {
        this.variableCache = new Map();
        this.functionCache = new Map();
        this._isInitialized = false;
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
    initialize() {
        try {
            logger_1.default.info('🧮 Ініціалізація FormulaProcessor...');
            // Ініціалізація математичних функцій
            this.initializeMathFunctions();
            // Ініціалізація кешу змінних
            this.initializeVariableCache();
            this._isInitialized = true;
            logger_1.default.info('✅ FormulaProcessor успішно ініціалізовано');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'initialize' },
            });
            throw new Error('Помилка ініціалізації FormulaProcessor');
        }
    }
    /**
     * Ініціалізація математичних функцій
     */
    initializeMathFunctions() {
        try {
            const mathFunctions = {
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
                if: (condition, trueValue, falseValue) => condition ? trueValue : falseValue,
            };
            for (const [name, func] of Object.entries(mathFunctions)) {
                this.functionCache.set(name, func);
            }
            logger_1.default.debug(`📚 Ініціалізовано ${this.functionCache.size} математичних функцій`);
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'initializeMathFunctions' },
            });
        }
    }
    /**
     * Ініціалізація кешу змінних
     */
    initializeVariableCache() {
        try {
            // Додавання констант
            this.variableCache.set('PI', Math.PI);
            this.variableCache.set('E', Math.E);
            this.variableCache.set('INFINITY', Infinity);
            this.variableCache.set('NAN', NaN);
            logger_1.default.debug(`📊 Ініціалізовано ${this.variableCache.size} констант`);
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'initializeVariableCache' },
            });
        }
    }
    /**
     * Валідація формули
     */
    validateFormula(formula) {
        const startTime = performance.now();
        try {
            logger_1.default.debug('🔍 Валідація формули...', {
                formula: formula.substring(0, 100),
                length: formula.length,
            });
            const errors = [];
            const warnings = [];
            let sanitizedFormula = formula;
            // Перевірка довжини
            if (formula.length > FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH) {
                errors.push(`Формула занадто довга (${formula.length} символів, максимум ${FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH})`);
                sanitizedFormula = formula.substring(0, FORMULA_PROCESSOR_CONSTANTS.MAX_FORMULA_LENGTH);
            }
            // Валідація введення
            const validation = (0, security_1.validateInput)(formula, { inputType: 'command' });
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
                errors.push(`Занадто багато змінних (${variables.length}, максимум ${FORMULA_PROCESSOR_CONSTANTS.MAX_VARIABLES})`);
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
            const result = {
                isValid: errors.length === 0,
                errors,
                warnings,
                sanitizedFormula,
                variables,
                functions,
                complexity,
            };
            if (errors.length > 0) {
                logger_1.default.warn('❌ Валідація формули невдала', {
                    errors,
                    warnings,
                    formula: formula.substring(0, 100),
                });
            }
            else {
                logger_1.default.debug('✅ Валідація формули успішна', {
                    variables: variables.length,
                    functions: functions.length,
                    complexity,
                    duration: `${duration.toFixed(2)}ms`,
                });
            }
            return result;
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
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
    async evaluateFormula(formula, variables = {}) {
        const startTime = performance.now();
        const operationId = this.generateOperationId(formula);
        try {
            logger_1.default.debug('🧮 Початок обчислення формули...', {
                formula: formula.substring(0, 100),
                variablesCount: Object.keys(variables).length,
                operationId,
            });
            // Валідація формули
            const validation = this.validateFormula(formula);
            if (!validation.isValid) {
                throw new Error(`Формула не валідна: ${validation.errors.join(', ')}`);
            }
            // Об'єднання змінних
            const allVariables = {
                ...Object.fromEntries(this.variableCache),
                ...variables,
            };
            // Створення безпечного контексту виконання
            const result = await this.executeFormula(formula, allVariables);
            const duration = performance.now() - startTime;
            const complexity = validation.complexity;
            const formulaResult = {
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
            logger_1.default.info('✅ Формула успішно обчислена', {
                formula: formula.substring(0, 100),
                result,
                duration: `${duration.toFixed(2)}ms`,
                complexity,
                operationId,
            });
            return formulaResult;
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.updateStats(false, duration, 0);
            const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
            logger_1.default.error('❌ Помилка обчислення формули', {
                formula: formula.substring(0, 100),
                error: errorMessage,
                duration: `${duration.toFixed(2)}ms`,
                operationId,
            });
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
    async executeFormula(formula, variables) {
        return new Promise((resolve, reject) => {
            try {
                // Створення безпечного контексту
                const safeContext = {
                    ...variables,
                    ...Object.fromEntries(this.functionCache),
                };
                // Обмеження часу виконання
                const timeout = setTimeout(() => {
                    reject(new Error('Таймаут виконання формули'));
                }, FORMULA_PROCESSOR_CONSTANTS.TIMEOUT);
                // Створення функції з обмеженим контекстом
                const safeFunction = new Function(...Object.keys(safeContext), `"use strict"; return (${formula});`);
                // Виконання з обмеженим контекстом
                const result = safeFunction(...Object.values(safeContext));
                clearTimeout(timeout);
                // Перевірка результату
                if (typeof result !== 'number' || !isFinite(result)) {
                    reject(new Error('Невірний результат формули'));
                }
                // Округлення до заданої точності
                const roundedResult = Math.round(result * Math.pow(10, FORMULA_PROCESSOR_CONSTANTS.PRECISION)) /
                    Math.pow(10, FORMULA_PROCESSOR_CONSTANTS.PRECISION);
                resolve(roundedResult);
            }
            catch (error) {
                reject(error);
            }
        });
    }
    /**
     * Витяг змінних з формули
     */
    extractVariables(formula) {
        const variables = new Set();
        const variablePattern = /[a-zA-Z_][a-zA-Z0-9_]*/g;
        const matches = formula.match(variablePattern) || [];
        for (const match of matches) {
            // Виключення функцій та констант
            if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_FUNCTIONS.includes(match) &&
                !['PI', 'E', 'INFINITY', 'NAN'].includes(match)) {
                variables.add(match);
            }
        }
        return Array.from(variables);
    }
    /**
     * Витяг функцій з формули
     */
    extractFunctions(formula) {
        const functions = new Set();
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
    calculateComplexity(formula) {
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
    calculateNestedLevel(formula) {
        let maxLevel = 0;
        let currentLevel = 0;
        for (const char of formula) {
            if (char === '(') {
                currentLevel++;
                maxLevel = Math.max(maxLevel, currentLevel);
            }
            else if (char === ')') {
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
    generateOperationId(formula) {
        const timestamp = Date.now();
        const hash = require('crypto').createHash('md5').update(`${formula}:${timestamp}`).digest('hex');
        return `formula_${hash.substring(0, 8)}`;
    }
    /**
     * Оновлення статистики
     */
    updateStats(success, duration, complexity) {
        try {
            this.stats.totalFormulas++;
            this.stats.totalExecutionTime += duration;
            this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.totalFormulas;
            if (success) {
                this.stats.successfulFormulas++;
            }
            else {
                this.stats.failedFormulas++;
            }
            // Розподіл складності
            const complexityLevel = complexity < 10 ? 'low' :
                complexity < 50 ? 'medium' :
                    complexity < 100 ? 'high' : 'very_high';
            this.stats.complexityDistribution[complexityLevel] =
                (this.stats.complexityDistribution[complexityLevel] || 0) + 1;
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'updateStats' },
            });
        }
    }
    /**
     * Встановлення змінної
     */
    setVariable(name, value) {
        try {
            if (!FORMULA_PROCESSOR_CONSTANTS.ALLOWED_VARIABLES.test(name)) {
                throw new Error(`Недозволена змінна: ${name}`);
            }
            this.variableCache.set(name, value);
            logger_1.default.debug(`📊 Встановлено змінну: ${name} = ${value}`);
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'setVariable', name, value },
            });
        }
    }
    /**
     * Отримання значення змінної
     */
    getVariable(name) {
        return this.variableCache.get(name);
    }
    /**
     * Очищення змінних
     */
    clearVariables() {
        try {
            this.variableCache.clear();
            this.initializeVariableCache(); // Відновлення констант
            logger_1.default.info('🧹 Змінні очищено');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'clearVariables' },
            });
        }
    }
    /**
     * Отримання статистики
     */
    getStats() {
        return { ...this.stats };
    }
    /**
     * Очищення ресурсів
     */
    cleanup() {
        try {
            this.variableCache.clear();
            this.functionCache.clear();
            logger_1.default.info('🧹 Ресурси FormulaProcessor очищено');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FormulaProcessor',
                additionalContext: { operation: 'cleanup' },
            });
        }
    }
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized() {
        return this._isInitialized;
    }
}
exports.FormulaProcessor = FormulaProcessor;
FormulaProcessor.instance = null;
// Експорт єдиного екземпляра
exports.formulaProcessor = new FormulaProcessor();
// Експорт функцій для зручності
const validateFormula = (formula) => exports.formulaProcessor.validateFormula(formula);
exports.validateFormula = validateFormula;
const evaluateFormula = (formula, variables) => exports.formulaProcessor.evaluateFormula(formula, variables);
exports.evaluateFormula = evaluateFormula;
const setVariable = (name, value) => exports.formulaProcessor.setVariable(name, value);
exports.setVariable = setVariable;
const getVariable = (name) => exports.formulaProcessor.getVariable(name);
exports.getVariable = getVariable;
const clearVariables = () => exports.formulaProcessor.clearVariables();
exports.clearVariables = clearVariables;
const getFormulaProcessorStats = () => exports.formulaProcessor.getStats();
exports.getFormulaProcessorStats = getFormulaProcessorStats;
const cleanupFormulaProcessor = () => exports.formulaProcessor.cleanup();
exports.cleanupFormulaProcessor = cleanupFormulaProcessor;
//# sourceMappingURL=formulaProcessor.js.map