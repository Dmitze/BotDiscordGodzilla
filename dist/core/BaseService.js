"use strict";
/**
 * Базовий клас для всіх сервісів
 * Надає спільну функціональність та інтерфейс
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.BaseService = void 0;
const logger_1 = __importDefault(require("@/utils/logger"));
// Константи для базового сервісу
const BASE_SERVICE_CONSTANTS = {
    INITIALIZATION_TIMEOUT: 30000, // 30 секунд
    SHUTDOWN_TIMEOUT: 10000, // 10 секунд
    HEALTH_CHECK_TIMEOUT: 5000, // 5 секунд
    MAX_RETRY_ATTEMPTS: 3,
    RETRY_DELAY: 1000, // 1 секунда
};
class BaseService {
    constructor(name, config) {
        this.isInitialized = false;
        this.isShuttingDown = false;
        this.retryCount = 0;
        this.initializationTimeout = null;
        this.shutdownTimeout = null;
        this.name = name;
        this.config = config;
        this.startTime = Date.now();
        logger_1.default.debug(`🔧 Створено базовий сервіс: ${this.name}`);
    }
    /**
     * Ініціалізація сервісу з детальним логуванням
     */
    async initialize() {
        if (this.isInitialized) {
            logger_1.default.warn(`⚠️ Сервіс ${this.name} вже ініціалізовано`);
            return;
        }
        if (this.isShuttingDown) {
            throw new Error(`Неможливо ініціалізувати сервіс ${this.name} під час зупинки`);
        }
        const startTime = Date.now();
        try {
            logger_1.default.info(`🚀 Ініціалізація сервісу ${this.name}...`);
            // Встановлення таймауту для ініціалізації
            this.initializationTimeout = setTimeout(() => {
                logger_1.default.error(`⏰ Таймаут ініціалізації сервісу ${this.name}`);
                throw new Error(`Таймаут ініціалізації сервісу ${this.name}`);
            }, BASE_SERVICE_CONSTANTS.INITIALIZATION_TIMEOUT);
            await this.onInitialize();
            // Очищення таймауту
            if (this.initializationTimeout) {
                clearTimeout(this.initializationTimeout);
                this.initializationTimeout = null;
            }
            this.isInitialized = true;
            this.retryCount = 0;
            const duration = Date.now() - startTime;
            logger_1.default.info(`✅ Сервіс ${this.name} успішно ініціалізовано за ${duration}ms`);
        }
        catch (error) {
            const duration = Date.now() - startTime;
            logger_1.default.error(`❌ Помилка ініціалізації сервісу ${this.name} після ${duration}ms:`, error);
            // Очищення таймауту
            if (this.initializationTimeout) {
                clearTimeout(this.initializationTimeout);
                this.initializationTimeout = null;
            }
            // Спроба повторної ініціалізації
            if (this.retryCount < BASE_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS) {
                this.retryCount++;
                logger_1.default.info(`🔄 Спроба повторної ініціалізації ${this.retryCount}/${BASE_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS} для сервісу ${this.name}...`);
                await new Promise(resolve => setTimeout(resolve, BASE_SERVICE_CONSTANTS.RETRY_DELAY));
                return this.initialize();
            }
            throw new Error(`Помилка ініціалізації сервісу ${this.name}: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Завершення роботи сервісу з детальним логуванням
     */
    async shutdown() {
        if (!this.isInitialized) {
            logger_1.default.debug(`ℹ️ Сервіс ${this.name} не ініціалізовано, пропускаю зупинку`);
            return;
        }
        if (this.isShuttingDown) {
            logger_1.default.warn(`⚠️ Сервіс ${this.name} вже зупиняється`);
            return;
        }
        this.isShuttingDown = true;
        const shutdownStartTime = Date.now();
        try {
            logger_1.default.info(`🛑 Завершення роботи сервісу ${this.name}...`);
            // Встановлення таймауту для зупинки
            this.shutdownTimeout = setTimeout(() => {
                logger_1.default.error(`⏰ Таймаут зупинки сервісу ${this.name}`);
                throw new Error(`Таймаут зупинки сервісу ${this.name}`);
            }, BASE_SERVICE_CONSTANTS.SHUTDOWN_TIMEOUT);
            await this.onShutdown();
            // Очищення таймауту
            if (this.shutdownTimeout) {
                clearTimeout(this.shutdownTimeout);
                this.shutdownTimeout = null;
            }
            this.isInitialized = false;
            this.isShuttingDown = false;
            const duration = Date.now() - shutdownStartTime;
            logger_1.default.info(`✅ Сервіс ${this.name} успішно зупинено за ${duration}ms`);
        }
        catch (error) {
            const duration = Date.now() - shutdownStartTime;
            logger_1.default.error(`❌ Помилка зупинки сервісу ${this.name} після ${duration}ms:`, error);
            // Очищення таймауту
            if (this.shutdownTimeout) {
                clearTimeout(this.shutdownTimeout);
                this.shutdownTimeout = null;
            }
            this.isInitialized = false;
            this.isShuttingDown = false;
            throw new Error(`Помилка зупинки сервісу ${this.name}: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Перевірка здоров'я сервісу з детальним логуванням
     */
    async healthCheck() {
        if (!this.isInitialized) {
            return {
                healthy: false,
                service: this.name,
                error: 'Сервіс не ініціалізовано',
            };
        }
        const startTime = Date.now();
        try {
            logger_1.default.debug(`🏥 Health check сервісу ${this.name}...`);
            const health = await this.onHealthCheck();
            const duration = Date.now() - startTime;
            if (!health.healthy) {
                logger_1.default.warn(`⚠️ Health check сервісу ${this.name} виявив проблеми за ${duration}ms:`, health);
            }
            else {
                logger_1.default.debug(`✅ Health check сервісу ${this.name} пройшов успішно за ${duration}ms`);
            }
            return {
                healthy: health.healthy,
                service: this.name,
                ...(health.error && { error: health.error }),
                ...(health.details && { details: health.details }),
            };
        }
        catch (error) {
            const duration = Date.now() - startTime;
            logger_1.default.error(`❌ Помилка health check сервісу ${this.name} після ${duration}ms:`, error);
            return {
                healthy: false,
                service: this.name,
                error: `Помилка health check: ${error instanceof Error ? error.message : 'Невідома помилка'}`,
            };
        }
    }
    /**
     * Отримання статистики сервісу з детальним логуванням
     */
    getStats() {
        try {
            const baseStats = {
                service: this.name,
                uptime: Date.now() - this.startTime,
                requests: 0,
                errors: 0,
                isInitialized: this.isInitialized,
                isShuttingDown: this.isShuttingDown,
                retryCount: this.retryCount,
            };
            const serviceStats = this.onGetStats();
            const combinedStats = { ...baseStats, ...serviceStats };
            logger_1.default.debug(`📊 Статистика сервісу ${this.name}:`, {
                uptime: `${Math.round(combinedStats.uptime / 1000)}s`,
                requests: combinedStats.requests,
                errors: combinedStats.errors,
                isInitialized: combinedStats.isInitialized,
            });
            return combinedStats;
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка отримання статистики сервісу ${this.name}:`, error);
            return {
                service: this.name,
                uptime: Date.now() - this.startTime,
                requests: 0,
                errors: 1,
                isInitialized: this.isInitialized,
                isShuttingDown: this.isShuttingDown,
                retryCount: this.retryCount,
                error: error instanceof Error ? error.message : 'Невідома помилка',
            };
        }
    }
    /**
     * Перевірка чи сервіс ініціалізовано
     */
    checkInitialized() {
        if (!this.isInitialized) {
            const error = `Сервіс ${this.name} не ініціалізовано`;
            logger_1.default.error(`❌ ${error}`);
            throw new Error(error);
        }
    }
    /**
     * Перевірка чи сервіс не зупиняється
     */
    checkNotShuttingDown() {
        if (this.isShuttingDown) {
            const error = `Сервіс ${this.name} зупиняється`;
            logger_1.default.warn(`⚠️ ${error}`);
            throw new Error(error);
        }
    }
    /**
     * Безпечне виконання операції з обробкою помилок
     */
    async safeExecute(operation, operationName, fallback) {
        const startTime = Date.now();
        try {
            logger_1.default.debug(`🔄 Виконання операції ${operationName} в сервісі ${this.name}...`);
            const result = await operation();
            const duration = Date.now() - startTime;
            logger_1.default.debug(`✅ Операція ${operationName} в сервісі ${this.name} завершена за ${duration}ms`);
            return result;
        }
        catch (error) {
            const duration = Date.now() - startTime;
            logger_1.default.error(`❌ Помилка операції ${operationName} в сервісі ${this.name} після ${duration}ms:`, error);
            if (fallback !== undefined) {
                logger_1.default.warn(`🔄 Використання fallback значення для операції ${operationName} в сервісі ${this.name}`);
                return fallback;
            }
            throw error;
        }
    }
    /**
     * Очищення ресурсів сервісу
     */
    async cleanup() {
        try {
            logger_1.default.info(`🧹 Очищення ресурсів сервісу ${this.name}...`);
            // Очищення таймаутів
            if (this.initializationTimeout) {
                clearTimeout(this.initializationTimeout);
                this.initializationTimeout = null;
            }
            if (this.shutdownTimeout) {
                clearTimeout(this.shutdownTimeout);
                this.shutdownTimeout = null;
            }
            logger_1.default.info(`✅ Ресурси сервісу ${this.name} очищено`);
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка очищення ресурсів сервісу ${this.name}:`, error);
        }
    }
}
exports.BaseService = BaseService;
//# sourceMappingURL=BaseService.js.map