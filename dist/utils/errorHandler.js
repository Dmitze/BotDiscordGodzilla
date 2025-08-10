"use strict";
/**
 * Розширений обробник помилок для Discord AI Assistant Bot
 * Централізована обробка та логування помилок
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.clearErrorHistory = exports.getErrorHistory = exports.getErrorStats = exports.handleError = exports.errorHandler = exports.ErrorHandler = void 0;
const logger_1 = __importDefault(require("./logger"));
// Константи для обробки помилок
const ERROR_HANDLER_CONSTANTS = {
    MAX_ERROR_DETAILS: 1000,
    MAX_STACK_TRACE_LINES: 20,
    ERROR_CATEGORIES: {
        VALIDATION: 'validation',
        NETWORK: 'network',
        DATABASE: 'database',
        AUTHENTICATION: 'authentication',
        AUTHORIZATION: 'authorization',
        RATE_LIMIT: 'rate_limit',
        TIMEOUT: 'timeout',
        RESOURCE: 'resource',
        SYSTEM: 'system',
        UNKNOWN: 'unknown',
    },
    SEVERITY_LEVELS: {
        LOW: 'low',
        MEDIUM: 'medium',
        HIGH: 'high',
        CRITICAL: 'critical',
    },
};
class ErrorHandler {
    constructor() {
        this.errorHistory = [];
        this.maxErrorHistory = 1000;
        this._isInitialized = false;
        if (ErrorHandler.instance) {
            return ErrorHandler.instance;
        }
        ErrorHandler.instance = this;
        this.errorStats = {
            totalErrors: 0,
            errorsByCategory: {},
            errorsBySeverity: {},
            errorsByService: {},
            recentErrors: [],
            averageErrorRate: 0,
            criticalErrors: 0,
        };
        this.initialize();
    }
    /**
     * Ініціалізація обробника помилок
     */
    initialize() {
        try {
            logger_1.default.info('🔧 Ініціалізація ErrorHandler...');
            // Налаштування глобальних обробників помилок
            this.setupGlobalErrorHandlers();
            this._isInitialized = true;
            logger_1.default.info('✅ ErrorHandler успішно ініціалізовано');
        }
        catch (error) {
            console.error('❌ Помилка ініціалізації ErrorHandler:', error);
            this.createFallbackErrorHandler();
        }
    }
    /**
     * Налаштування глобальних обробників помилок
     */
    setupGlobalErrorHandlers() {
        // Обробка необроблених помилок
        process.on('uncaughtException', (error) => {
            this.handleUncaughtException(error);
        });
        // Обробка необроблених rejections
        process.on('unhandledRejection', (reason, promise) => {
            this.handleUnhandledRejection(reason, promise);
        });
        // Обробка попереджень
        process.on('warning', (warning) => {
            this.handleWarning(warning);
        });
        logger_1.default.info('🛡️ Глобальні обробники помилок налаштовано');
    }
    /**
     * Обробка необробленої помилки
     */
    handleUncaughtException(error) {
        const errorDetails = {
            name: error.name,
            message: error.message,
            ...(error.stack ? { stack: error.stack } : {}),
            timestamp: new Date(),
            category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
            severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL,
            context: {
                type: 'uncaught_exception',
                processId: process.pid,
                uptime: process.uptime(),
                memory: process.memoryUsage(),
            },
        };
        this.logError(errorDetails);
        // Логування критичної помилки
        logger_1.default.error('💥 Критична необроблена помилка:', {
            name: error.name,
            message: error.message,
            stack: this.truncateStackTrace(error.stack),
            type: 'system',
            eventType: 'uncaught_exception',
            severity: 'critical',
        });
        // Зупинка процесу при критичній помилці
        logger_1.default.error('🛑 Зупинка процесу через критичну помилку');
        process.exit(1);
    }
    /**
     * Обробка необробленого rejection
     */
    handleUnhandledRejection(reason, promise) {
        const errorDetails = {
            name: 'UnhandledRejection',
            message: reason instanceof Error ? reason.message : String(reason),
            ...(reason instanceof Error && reason.stack ? { stack: reason.stack } : {}),
            timestamp: new Date(),
            category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
            severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH,
            context: {
                type: 'unhandled_rejection',
                promise: promise.toString(),
                processId: process.pid,
                uptime: process.uptime(),
            },
        };
        this.logError(errorDetails);
        logger_1.default.error('💥 Необроблений rejection:', {
            reason: reason instanceof Error ? reason.message : String(reason),
            promise: promise.toString(),
            type: 'system',
            eventType: 'unhandled_rejection',
            severity: 'high',
        });
    }
    /**
     * Обробка попередження
     */
    handleWarning(warning) {
        const errorDetails = {
            name: warning.name,
            message: warning.message,
            ...(warning.stack ? { stack: warning.stack } : {}),
            timestamp: new Date(),
            category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
            severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW,
            context: {
                type: 'warning',
                processId: process.pid,
                uptime: process.uptime(),
            },
        };
        this.logError(errorDetails);
        logger_1.default.warn('⚠️ Попередження системи:', {
            name: warning.name,
            message: warning.message,
            type: 'system',
            eventType: 'warning',
            severity: 'low',
        });
    }
    /**
     * Основний метод обробки помилок
     */
    handleError(error, context = {}) {
        try {
            const errorDetails = this.createErrorDetails(error, context);
            this.logError(errorDetails);
            this.updateStats(errorDetails);
            return errorDetails;
        }
        catch (handlerError) {
            console.error('❌ Помилка в ErrorHandler:', handlerError);
            return this.createFallbackErrorDetails(error);
        }
    }
    /**
     * Створення деталей помилки
     */
    createErrorDetails(error, context) {
        const errorObj = error instanceof Error ? error : new Error(String(error));
        return {
            name: errorObj.name,
            message: errorObj.message,
            ...(errorObj.stack ? { stack: errorObj.stack } : {}),
            code: error?.code,
            ...(('cause' in errorObj) && errorObj.cause !== undefined
                ? { cause: errorObj.cause }
                : {}),
            timestamp: new Date(),
            category: this.categorizeError(errorObj),
            severity: this.determineSeverity(errorObj),
            context: {
                ...context.additionalContext,
                errorType: errorObj.constructor.name,
                hasStack: !!errorObj.stack,
            },
            ...(context.userId ? { userId: context.userId } : {}),
            ...(context.guildId ? { guildId: context.guildId } : {}),
            ...(context.channelId ? { channelId: context.channelId } : {}),
            ...(context.commandName ? { commandName: context.commandName } : {}),
            ...(context.serviceName ? { serviceName: context.serviceName } : {}),
            ...(context.requestId ? { requestId: context.requestId } : {}),
            ...(context.correlationId ? { correlationId: context.correlationId } : {}),
        };
    }
    /**
     * Категоризація помилки
     */
    categorizeError(error) {
        const message = error.message.toLowerCase();
        const name = error.name.toLowerCase();
        if (message.includes('validation') || name.includes('validation')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.VALIDATION;
        }
        if (message.includes('network') || message.includes('connection') || message.includes('timeout')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.NETWORK;
        }
        if (message.includes('database') || message.includes('sql') || message.includes('query')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.DATABASE;
        }
        if (message.includes('auth') || message.includes('token') || message.includes('permission')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.AUTHENTICATION;
        }
        if (message.includes('rate limit') || message.includes('too many requests')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.RATE_LIMIT;
        }
        if (message.includes('timeout') || message.includes('timed out')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.TIMEOUT;
        }
        if (message.includes('resource') || message.includes('memory') || message.includes('disk')) {
            return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.RESOURCE;
        }
        return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.UNKNOWN;
    }
    /**
     * Визначення серйозності помилки
     */
    determineSeverity(error) {
        const message = error.message.toLowerCase();
        const name = error.name.toLowerCase();
        if (name.includes('critical') || message.includes('critical')) {
            return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL;
        }
        if (name.includes('fatal') || message.includes('fatal')) {
            return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL;
        }
        if (message.includes('timeout') || message.includes('connection failed')) {
            return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH;
        }
        if (message.includes('validation') || message.includes('invalid')) {
            return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM;
        }
        if (message.includes('warning') || name.includes('warning')) {
            return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW;
        }
        return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM;
    }
    /**
     * Логування помилки
     */
    logError(errorDetails) {
        try {
            const logMeta = {
                errorName: errorDetails.name,
                errorMessage: errorDetails.message,
                errorCategory: errorDetails.category,
                errorSeverity: errorDetails.severity,
                errorCode: errorDetails.code,
                ...(errorDetails.userId ? { userId: errorDetails.userId } : {}),
                ...(errorDetails.guildId ? { guildId: errorDetails.guildId } : {}),
                ...(errorDetails.channelId ? { channelId: errorDetails.channelId } : {}),
                ...(errorDetails.commandName ? { commandName: errorDetails.commandName } : {}),
                ...(errorDetails.serviceName ? { serviceName: errorDetails.serviceName } : {}),
                ...(errorDetails.requestId ? { requestId: errorDetails.requestId } : {}),
                ...(errorDetails.correlationId ? { correlationId: errorDetails.correlationId } : {}),
                timestamp: errorDetails.timestamp.toISOString(),
                type: 'system',
                severity: errorDetails.severity,
            };
            // Логування в залежності від серйозності
            switch (errorDetails.severity) {
                case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL:
                    logger_1.default.error(`💥 Критична помилка: ${errorDetails.message}`, logMeta);
                    break;
                case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH:
                    logger_1.default.error(`❌ Серйозна помилка: ${errorDetails.message}`, logMeta);
                    break;
                case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM:
                    logger_1.default.warn(`⚠️ Помилка: ${errorDetails.message}`, logMeta);
                    break;
                case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW:
                    logger_1.default.debug(`ℹ️ Попередження: ${errorDetails.message}`, logMeta);
                    break;
                default:
                    logger_1.default.error(`❌ Помилка: ${errorDetails.message}`, logMeta);
            }
            // Логування stack trace для серйозних помилок
            if (errorDetails.severity !== ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW && errorDetails.stack) {
                logger_1.default.debug('📋 Stack trace:', {
                    stack: this.truncateStackTrace(errorDetails.stack),
                    errorName: errorDetails.name,
                });
            }
        }
        catch (logError) {
            console.error('❌ Помилка логування помилки:', logError);
        }
    }
    /**
     * Оновлення статистики помилок
     */
    updateStats(errorDetails) {
        try {
            this.errorStats.totalErrors++;
            this.errorStats.lastError = errorDetails;
            // Оновлення статистики по категоріях
            this.errorStats.errorsByCategory[errorDetails.category] =
                (this.errorStats.errorsByCategory[errorDetails.category] || 0) + 1;
            // Оновлення статистики по серйозності
            this.errorStats.errorsBySeverity[errorDetails.severity] =
                (this.errorStats.errorsBySeverity[errorDetails.severity] || 0) + 1;
            // Оновлення статистики по сервісах
            if (errorDetails.serviceName) {
                this.errorStats.errorsByService[errorDetails.serviceName] =
                    (this.errorStats.errorsByService[errorDetails.serviceName] || 0) + 1;
            }
            // Оновлення критичних помилок
            if (errorDetails.severity === ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL) {
                this.errorStats.criticalErrors++;
            }
            // Додавання до історії
            this.errorHistory.push(errorDetails);
            if (this.errorHistory.length > this.maxErrorHistory) {
                this.errorHistory.shift();
            }
            // Оновлення середньої частоти помилок
            const uptime = process.uptime();
            this.errorStats.averageErrorRate = uptime > 0 ? this.errorStats.totalErrors / uptime : 0;
        }
        catch (statsError) {
            console.error('❌ Помилка оновлення статистики помилок:', statsError);
        }
    }
    /**
     * Обрізання stack trace
     */
    truncateStackTrace(stack) {
        if (!stack)
            return '';
        const lines = stack.split('\n');
        const truncatedLines = lines.slice(0, ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES);
        if (lines.length > ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES) {
            truncatedLines.push(`... (${lines.length - ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES} more lines)`);
        }
        return truncatedLines.join('\n');
    }
    /**
     * Створення fallback обробника помилок
     */
    createFallbackErrorHandler() {
        console.error('🔧 Створення fallback обробника помилок...');
        process.on('uncaughtException', (error) => {
            console.error('💥 Критична помилка (fallback):', error);
            process.exit(1);
        });
        process.on('unhandledRejection', (reason) => {
            console.error('💥 Необроблений rejection (fallback):', reason);
        });
    }
    /**
     * Створення fallback деталей помилки
     */
    createFallbackErrorDetails(error) {
        return {
            name: 'UnknownError',
            message: error instanceof Error ? error.message : String(error),
            timestamp: new Date(),
            category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.UNKNOWN,
            severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM,
        };
    }
    /**
     * Отримання статистики помилок
     */
    getStats() {
        return { ...this.errorStats };
    }
    /**
     * Отримання історії помилок
     */
    getErrorHistory() {
        return [...this.errorHistory];
    }
    /**
     * Очищення історії помилок
     */
    clearErrorHistory() {
        this.errorHistory = [];
        logger_1.default.info('🧹 Історія помилок очищено');
    }
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized() {
        return this._isInitialized;
    }
}
exports.ErrorHandler = ErrorHandler;
ErrorHandler.instance = null;
// Експорт єдиного екземпляра
exports.errorHandler = new ErrorHandler();
// Експорт функцій для зручності
const handleError = (error, context) => exports.errorHandler.handleError(error, context);
exports.handleError = handleError;
const getErrorStats = () => exports.errorHandler.getStats();
exports.getErrorStats = getErrorStats;
const getErrorHistory = () => exports.errorHandler.getErrorHistory();
exports.getErrorHistory = getErrorHistory;
const clearErrorHistory = () => exports.errorHandler.clearErrorHistory();
exports.clearErrorHistory = clearErrorHistory;
//# sourceMappingURL=errorHandler.js.map