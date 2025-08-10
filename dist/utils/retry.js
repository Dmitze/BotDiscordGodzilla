"use strict";
/**
 * Утиліта для повторних спроб операцій
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.RetryManager = void 0;
const logger_1 = __importDefault(require("./logger"));
class RetryManager {
    /**
     * Виконання операції з повторними спробами
     */
    static async execute(operation, options = {}) {
        const config = { ...this.defaultOptions, ...options };
        const startTime = Date.now();
        let lastError;
        for (let attempt = 1; attempt <= config.maxAttempts; attempt++) {
            try {
                // Створюємо timeout promise
                const timeoutPromise = new Promise((_, reject) => {
                    setTimeout(() => reject(new Error('Operation timeout')), config.timeout);
                });
                // Виконуємо операцію з timeout
                const result = await Promise.race([operation(), timeoutPromise]);
                return {
                    success: true,
                    data: result,
                    attempts: attempt,
                    totalTime: Date.now() - startTime,
                };
            }
            catch (error) {
                lastError = error instanceof Error ? error : new Error(String(error));
                // Перевіряємо чи потрібно повторювати
                if (!config.shouldRetry(lastError)) {
                    return {
                        success: false,
                        error: lastError,
                        attempts: attempt,
                        totalTime: Date.now() - startTime,
                    };
                }
                // Остання спроба
                if (attempt === config.maxAttempts) {
                    logger_1.default.error(`❌ Операція невдала після ${attempt} спроб: ${lastError.message || String(lastError)}`);
                    return {
                        success: false,
                        error: lastError,
                        attempts: attempt,
                        totalTime: Date.now() - startTime,
                    };
                }
                // Викликаємо callback
                config.onRetry(attempt, lastError);
                // Розраховуємо затримку
                const delay = this.calculateDelay(attempt, config);
                logger_1.default.warn(`⚠️ Спроба ${attempt} невдала, повтор через ${delay}мс: ${lastError.message}`);
                // Чекаємо перед наступною спробою
                await this.sleep(delay);
            }
        }
        return {
            success: false,
            error: lastError,
            attempts: config.maxAttempts,
            totalTime: Date.now() - startTime,
        };
    }
    /**
     * Розрахунок затримки між спробами
     */
    static calculateDelay(attempt, config) {
        let delay;
        switch (config.backoff) {
            case 'fixed':
                delay = config.delay;
                break;
            case 'linear':
                delay = config.delay * attempt;
                break;
            case 'exponential':
                delay = config.delay * Math.pow(config.factor, attempt - 1);
                break;
            default:
                delay = config.delay;
        }
        return Math.min(delay, config.maxDelay);
    }
    /**
     * Затримка виконання
     */
    static sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
    /**
     * Створення функції з повторними спробами
     */
    static createRetryFunction(fn, options = {}) {
        return async (...args) => {
            return this.execute(() => fn(...args), options);
        };
    }
    /**
     * Retry для HTTP запитів
     */
    static async httpRequest(requestFn, options = {}) {
        const httpOptions = {
            shouldRetry: (error) => {
                // Повторюємо для 5xx помилок та мережевих помилок
                const statusCode = error.status || error.code;
                return statusCode >= 500 || statusCode === 'ECONNRESET' || statusCode === 'ETIMEDOUT';
            },
            ...options,
        };
        return this.execute(requestFn, httpOptions);
    }
    /**
     * Retry для операцій з базою даних
     */
    static async databaseOperation(operation, options = {}) {
        const dbOptions = {
            shouldRetry: (error) => {
                // Повторюємо для тимчасових помилок БД
                const errorMessage = error.message.toLowerCase();
                return errorMessage.includes('connection') ||
                    errorMessage.includes('timeout') ||
                    errorMessage.includes('deadlock') ||
                    errorMessage.includes('temporary');
            },
            maxAttempts: 5,
            delay: 2000,
            ...options,
        };
        return this.execute(operation, dbOptions);
    }
    /**
     * Retry для файлових операцій
     */
    static async fileOperation(operation, options = {}) {
        const fileOptions = {
            shouldRetry: (error) => {
                // Повторюємо для тимчасових помилок файлової системи
                const errorCode = error.code;
                return errorCode === 'EBUSY' ||
                    errorCode === 'EACCES' ||
                    errorCode === 'ENOENT' ||
                    errorCode === 'EAGAIN';
            },
            maxAttempts: 3,
            delay: 1000,
            ...options,
        };
        return this.execute(operation, fileOptions);
    }
    /**
     * Retry для Discord API операцій
     */
    static async discordOperation(operation, options = {}) {
        const discordOptions = {
            shouldRetry: (error) => {
                // Повторюємо для rate limits та тимчасових помилок Discord
                const statusCode = error.status;
                return statusCode === 429 || statusCode >= 500;
            },
            maxAttempts: 3,
            delay: 1000,
            backoff: 'exponential',
            ...options,
        };
        return this.execute(operation, discordOptions);
    }
}
exports.RetryManager = RetryManager;
RetryManager.defaultOptions = {
    maxAttempts: 3,
    delay: 1000,
    backoff: 'exponential',
    factor: 2,
    maxDelay: 30000,
    timeout: 30000,
    onRetry: () => { },
    shouldRetry: () => true,
};
exports.default = RetryManager;
//# sourceMappingURL=retry.js.map