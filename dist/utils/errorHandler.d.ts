/**
 * Розширений обробник помилок для Discord AI Assistant Bot
 * Централізована обробка та логування помилок
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
export interface ErrorDetails {
    name: string;
    message: string;
    stack?: string;
    code?: string;
    cause?: Error;
    timestamp: Date;
    category: string;
    severity: string;
    context?: Record<string, unknown>;
    userId?: string;
    guildId?: string;
    channelId?: string;
    commandName?: string;
    serviceName?: string;
    requestId?: string;
    correlationId?: string;
}
export interface ErrorHandlerStats {
    totalErrors: number;
    errorsByCategory: Record<string, number>;
    errorsBySeverity: Record<string, number>;
    errorsByService: Record<string, number>;
    recentErrors: ErrorDetails[];
    lastError?: ErrorDetails;
    averageErrorRate: number;
    criticalErrors: number;
}
export declare class ErrorHandler {
    private static instance;
    private errorStats;
    private errorHistory;
    private readonly maxErrorHistory;
    private _isInitialized;
    constructor();
    /**
     * Ініціалізація обробника помилок
     */
    private initialize;
    /**
     * Налаштування глобальних обробників помилок
     */
    private setupGlobalErrorHandlers;
    /**
     * Обробка необробленої помилки
     */
    private handleUncaughtException;
    /**
     * Обробка необробленого rejection
     */
    private handleUnhandledRejection;
    /**
     * Обробка попередження
     */
    private handleWarning;
    /**
     * Основний метод обробки помилок
     */
    handleError(error: unknown, context?: {
        userId?: string;
        guildId?: string;
        channelId?: string;
        commandName?: string;
        serviceName?: string;
        requestId?: string;
        correlationId?: string;
        additionalContext?: Record<string, unknown>;
    }): ErrorDetails;
    /**
     * Створення деталей помилки
     */
    private createErrorDetails;
    /**
     * Категоризація помилки
     */
    private categorizeError;
    /**
     * Визначення серйозності помилки
     */
    private determineSeverity;
    /**
     * Логування помилки
     */
    private logError;
    /**
     * Оновлення статистики помилок
     */
    private updateStats;
    /**
     * Обрізання stack trace
     */
    private truncateStackTrace;
    /**
     * Створення fallback обробника помилок
     */
    private createFallbackErrorHandler;
    /**
     * Створення fallback деталей помилки
     */
    private createFallbackErrorDetails;
    /**
     * Отримання статистики помилок
     */
    getStats(): ErrorHandlerStats;
    /**
     * Отримання історії помилок
     */
    getErrorHistory(): ErrorDetails[];
    /**
     * Очищення історії помилок
     */
    clearErrorHistory(): void;
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized(): boolean;
}
export declare const errorHandler: ErrorHandler;
export declare const handleError: (error: unknown, context?: {
    userId?: string;
    guildId?: string;
    channelId?: string;
    commandName?: string;
    serviceName?: string;
    requestId?: string;
    correlationId?: string;
    additionalContext?: Record<string, unknown>;
}) => ErrorDetails;
export declare const getErrorStats: () => ErrorHandlerStats;
export declare const getErrorHistory: () => ErrorDetails[];
export declare const clearErrorHistory: () => void;
//# sourceMappingURL=errorHandler.d.ts.map