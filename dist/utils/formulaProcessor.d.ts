/**
 * Розширена система обробки формул для Discord AI Assistant Bot
 * Безпечна обробка математичних виразів та формул
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
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
export declare class FormulaProcessor {
    private static instance;
    private stats;
    private variableCache;
    private functionCache;
    private _isInitialized;
    constructor();
    /**
     * Ініціалізація обробника формул
     */
    private initialize;
    /**
     * Ініціалізація математичних функцій
     */
    private initializeMathFunctions;
    /**
     * Ініціалізація кешу змінних
     */
    private initializeVariableCache;
    /**
     * Валідація формули
     */
    validateFormula(formula: string): FormulaValidationResult;
    /**
     * Виконання формули
     */
    evaluateFormula(formula: string, variables?: Record<string, number>): Promise<FormulaResult>;
    /**
     * Безпечне виконання формули
     */
    private executeFormula;
    /**
     * Витяг змінних з формули
     */
    private extractVariables;
    /**
     * Витяг функцій з формули
     */
    private extractFunctions;
    /**
     * Розрахунок складності формули
     */
    private calculateComplexity;
    /**
     * Розрахунок рівня вкладеності
     */
    private calculateNestedLevel;
    /**
     * Генерація ID операції
     */
    private generateOperationId;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Встановлення змінної
     */
    setVariable(name: string, value: number): void;
    /**
     * Отримання значення змінної
     */
    getVariable(name: string): number | undefined;
    /**
     * Очищення змінних
     */
    clearVariables(): void;
    /**
     * Отримання статистики
     */
    getStats(): FormulaProcessorStats;
    /**
     * Очищення ресурсів
     */
    cleanup(): void;
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized(): boolean;
}
export declare const formulaProcessor: FormulaProcessor;
export declare const validateFormula: (formula: string) => FormulaValidationResult;
export declare const evaluateFormula: (formula: string, variables?: Record<string, number>) => Promise<FormulaResult>;
export declare const setVariable: (name: string, value: number) => void;
export declare const getVariable: (name: string) => number | undefined;
export declare const clearVariables: () => void;
export declare const getFormulaProcessorStats: () => FormulaProcessorStats;
export declare const cleanupFormulaProcessor: () => void;
//# sourceMappingURL=formulaProcessor.d.ts.map