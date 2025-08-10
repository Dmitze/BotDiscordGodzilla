"use strict";
/**
 * Основний файл Discord AI Assistant Bot
 * Точка входу в додаток
 * Версія 3.0.0 - Повністю рефакторовано з TypeScript
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.APP_CONFIG = void 0;
exports.getStats = getStats;
exports.restart = restart;
exports.shutdown = shutdown;
exports.getApp = getApp;
exports.main = main;
const dotenv_1 = require("dotenv");
const path_1 = require("path");
const fs_1 = require("fs");
const Bot_1 = require("@/core/Bot");
const Config_1 = require("@/config/Config");
const logger_1 = __importDefault(require("@/utils/logger"));
// Константи для конфігурації
const APP_CONFIG = {
    VERSION: '3.0.0',
    NAME: 'Discord AI Assistant Bot',
    STARTUP_TIMEOUT: 30000, // 30 секунд
    SHUTDOWN_TIMEOUT: 10000, // 10 секунд
    RESTART_DELAY: 5000, // 5 секунд
    MAX_MEMORY_USAGE: 1024 * 1024 * 1024, // 1GB
    HEALTH_CHECK_INTERVAL: 30000, // 30 секунд
};
exports.APP_CONFIG = APP_CONFIG;
// Глобальні обробники помилок
process.on('uncaughtException', (error) => {
    logger_1.default.error('💥 Необроблена помилка:', {
        name: error.name,
        message: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString(),
        memory: process.memoryUsage(),
        uptime: process.uptime(),
    });
    // Логування в файл для аналізу
    logger_1.default.error('Критична помилка додатку', {
        error: error.message,
        stack: error.stack,
        type: 'uncaught_exception',
    });
    process.exit(1);
});
process.on('unhandledRejection', (reason, promise) => {
    logger_1.default.error('💥 Необроблений rejection:', {
        reason: reason instanceof Error ? reason.message : String(reason),
        promise: promise.toString(),
        timestamp: new Date().toISOString(),
        memory: process.memoryUsage(),
        uptime: process.uptime(),
    });
    // Логування в файл для аналізу
    logger_1.default.error('Критичний rejection додатку', {
        reason: reason instanceof Error ? reason.message : String(reason),
        type: 'unhandled_rejection',
    });
});
// Завантаження змінних середовища
try {
    const envPath = (0, path_1.join)(process.cwd(), '.env');
    if ((0, fs_1.existsSync)(envPath)) {
        (0, dotenv_1.config)({ path: envPath });
        logger_1.default.info('✅ Змінні середовища завантажено з .env файлу', {});
    }
    else {
        (0, dotenv_1.config)();
        logger_1.default.warn('⚠️ .env файл не знайдено, використовую системні змінні', {});
    }
}
catch (error) {
    logger_1.default.error('❌ Помилка завантаження змінних середовища:', { error });
    throw new Error('Неможливо завантажити змінні середовища');
}
class Application {
    constructor() {
        this.bot = null;
        this.isStarting = false;
        this.isShuttingDown = false;
        this.startupTime = 0;
        this.restartCount = 0;
        this.maxRestarts = 5;
        this.healthCheckInterval = null;
        this.memoryCheckInterval = null;
        try {
            logger_1.default.info(`🚀 Ініціалізація ${APP_CONFIG.NAME} v${APP_CONFIG.VERSION}`, {});
            this.config = Config_1.Config.load();
            logger_1.default.info('✅ Конфігурація завантажена успішно', {});
        }
        catch (error) {
            logger_1.default.error('❌ Критична помилка ініціалізації додатку:', { error });
            throw new Error(`Помилка ініціалізації: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Запуск додатку з детальним логуванням
     */
    async start() {
        if (this.isStarting) {
            logger_1.default.warn('⚠️ Додаток вже запускається', {});
            return;
        }
        if (this.isShuttingDown) {
            logger_1.default.warn('⚠️ Неможливо запустити додаток під час зупинки', {});
            return;
        }
        this.isStarting = true;
        this.startupTime = Date.now();
        try {
            logger_1.default.info('🔄 Початок запуску додатку...', {});
            // Валідація конфігурації
            await this.validateConfiguration();
            // Перевірка системних ресурсів
            await this.checkSystemResources();
            // Створення та ініціалізація бота
            logger_1.default.info('🤖 Створення екземпляру бота...', {});
            this.bot = new Bot_1.Bot(this.config);
            logger_1.default.info('⚙️ Ініціалізація бота...', {});
            await this.bot.initialize();
            // Запуск моніторингу
            this.startMonitoring();
            const startupDuration = Date.now() - this.startupTime;
            logger_1.default.info(`✅ Додаток успішно запущено за ${startupDuration}ms`, {});
            // Скидання лічильника перезапусків при успішному запуску
            this.restartCount = 0;
            // Обробка сигналів завершення
            this.setupGracefulShutdown();
            // Логування статистики запуску
            this.logStartupStats();
        }
        catch (error) {
            const startupDuration = Date.now() - this.startupTime;
            logger_1.default.error(`❌ Помилка запуску додатку після ${startupDuration}ms:`, { error });
            // Спроба очищення ресурсів
            await this.cleanupOnError();
            throw new Error(`Помилка запуску: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
        finally {
            this.isStarting = false;
        }
    }
    /**
     * Зупинка додатку з детальним логуванням
     */
    async stop() {
        if (this.isShuttingDown) {
            logger_1.default.warn('⚠️ Додаток вже зупиняється', {});
            return;
        }
        this.isShuttingDown = true;
        const shutdownStartTime = Date.now();
        try {
            logger_1.default.info('🛑 Початок зупинки додатку...', {});
            // Зупинка моніторингу
            this.stopMonitoring();
            if (this.bot) {
                logger_1.default.info('🤖 Зупинка бота...', {});
                await this.bot.shutdown();
                this.bot = null;
            }
            const shutdownDuration = Date.now() - shutdownStartTime;
            logger_1.default.info(`✅ Додаток успішно зупинено за ${shutdownDuration}ms`, {});
        }
        catch (error) {
            const shutdownDuration = Date.now() - shutdownStartTime;
            logger_1.default.error(`❌ Помилка зупинки додатку після ${shutdownDuration}ms:`, { error });
            // Примусова зупинка при помилці
            logger_1.default.warn('🔄 Примусова зупинка процесу...', {});
            process.exit(1);
        }
        finally {
            this.isShuttingDown = false;
        }
    }
    /**
     * Отримання детальної статистики
     */
    getStats() {
        try {
            if (!this.bot) {
                return {
                    status: 'not_initialized',
                    uptime: process.uptime(),
                    memory: process.memoryUsage(),
                    version: APP_CONFIG.VERSION,
                };
            }
            const botStats = this.bot.getStats();
            const memoryUsage = process.memoryUsage();
            return {
                status: 'running',
                bot: botStats,
                uptime: process.uptime(),
                memory: {
                    rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`,
                    heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
                    heapTotal: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
                    external: `${Math.round(memoryUsage.external / 1024 / 1024)}MB`,
                },
                version: APP_CONFIG.VERSION,
                restartCount: this.restartCount,
                startupTime: this.startupTime,
                isStarting: this.isStarting,
                isShuttingDown: this.isShuttingDown,
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка отримання статистики:', { error });
            return {
                status: 'error',
                error: error instanceof Error ? error.message : 'Невідома помилка',
                uptime: process.uptime(),
                version: APP_CONFIG.VERSION,
            };
        }
    }
    /**
     * Перезапуск додатку з обмеженнями
     */
    async restart() {
        if (this.restartCount >= this.maxRestarts) {
            const error = `Досягнуто максимальну кількість перезапусків (${this.maxRestarts})`;
            logger_1.default.error(`❌ ${error}`, {});
            throw new Error(error);
        }
        this.restartCount++;
        logger_1.default.info(`🔄 Перезапуск додатку (спроба ${this.restartCount}/${this.maxRestarts})...`, {});
        try {
            // Зупинка поточного екземпляру
            if (this.bot) {
                logger_1.default.info('🛑 Зупинка поточного екземпляру...', {});
                await this.bot.shutdown();
                this.bot = null;
            }
            // Затримка перед перезапуском
            logger_1.default.info(`⏳ Затримка ${APP_CONFIG.RESTART_DELAY}ms перед перезапуском...`, {});
            await new Promise(resolve => setTimeout(resolve, APP_CONFIG.RESTART_DELAY));
            // Запуск нового екземпляру
            logger_1.default.info('🚀 Запуск нового екземпляру...', {});
            await this.start();
            logger_1.default.info('✅ Додаток успішно перезапущено', {});
        }
        catch (error) {
            logger_1.default.error('❌ Помилка при перезапуску:', { error });
            throw error;
        }
    }
    /**
     * Валідація конфігурації
     */
    async validateConfiguration() {
        try {
            logger_1.default.info('🔍 Валідація конфігурації...', {});
            // Перевірка обов'язкових полів
            const requiredFields = [
                'discord.token',
                'discord.clientId',
                'discord.guildId',
                'google.apiKey',
                'google.appScriptUrl',
                'ai.openai.apiKey'
            ];
            for (const field of requiredFields) {
                const value = this.getNestedValue(this.config, field);
                if (!value) {
                    throw new Error(`Відсутнє обов'язкове поле конфігурації: ${field}`);
                }
            }
            logger_1.default.info('✅ Конфігурація валідна', {});
        }
        catch (error) {
            logger_1.default.error('❌ Помилка валідації конфігурації:', { error });
            throw error;
        }
    }
    /**
     * Перевірка системних ресурсів
     */
    async checkSystemResources() {
        try {
            logger_1.default.info('🔍 Перевірка системних ресурсів...', {});
            const memoryUsage = process.memoryUsage();
            const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
            if (heapUsedMB > 500) {
                logger_1.default.warn(`⚠️ Високе використання пам'яті: ${Math.round(heapUsedMB)}MB`, {});
            }
            // Перевірка доступності файлової системи
            const testPath = (0, path_1.join)(process.cwd(), 'test_write');
            try {
                require('fs').writeFileSync(testPath, 'test');
                require('fs').unlinkSync(testPath);
            }
            catch (fsError) {
                logger_1.default.warn('⚠️ Проблеми з файловою системою:', { error: fsError });
            }
            logger_1.default.info('✅ Системні ресурси перевірено', {});
        }
        catch (error) {
            logger_1.default.error('❌ Помилка перевірки системних ресурсів:', { error });
            throw error;
        }
    }
    /**
     * Запуск моніторингу
     */
    startMonitoring() {
        // Health check
        this.healthCheckInterval = setInterval(async () => {
            try {
                if (this.bot) {
                    const health = await this.bot.healthCheck();
                    if (!health.healthy) {
                        logger_1.default.warn('⚠️ Health check виявив проблеми:', health);
                    }
                }
            }
            catch (error) {
                logger_1.default.error('❌ Помилка health check:', { error });
            }
        }, APP_CONFIG.HEALTH_CHECK_INTERVAL);
        // Memory monitoring
        this.memoryCheckInterval = setInterval(() => {
            try {
                const memoryUsage = process.memoryUsage();
                const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
                if (heapUsedMB > 800) {
                    logger_1.default.warn(`⚠️ Критичне використання пам'яті: ${Math.round(heapUsedMB)}MB`, {});
                }
                if (memoryUsage.heapUsed > APP_CONFIG.MAX_MEMORY_USAGE) {
                    logger_1.default.error("💥 Перевищено ліміт пам'яті, перезапуск...", {});
                    this.restart().catch(error => {
                        logger_1.default.error("❌ Помилка перезапуску через перевищення пам'яті:", { error });
                    });
                }
            }
            catch (error) {
                logger_1.default.error("❌ Помилка моніторингу пам'яті:", { error });
            }
        }, 60000); // Кожну хвилину
        logger_1.default.info('📊 Моніторинг запущено', {});
    }
    /**
     * Зупинка моніторингу
     */
    stopMonitoring() {
        if (this.healthCheckInterval) {
            clearInterval(this.healthCheckInterval);
            this.healthCheckInterval = null;
        }
        if (this.memoryCheckInterval) {
            clearInterval(this.memoryCheckInterval);
            this.memoryCheckInterval = null;
        }
        logger_1.default.info('📊 Моніторинг зупинено', {});
    }
    /**
     * Отримання вкладених значень об'єкта
     */
    getNestedValue(obj, path) {
        return path.split('.').reduce((current, key) => current?.[key], obj);
    }
    /**
     * Очищення ресурсів при помилці
     */
    async cleanupOnError() {
        try {
            logger_1.default.info('🧹 Очищення ресурсів при помилці...', {});
            this.stopMonitoring();
            if (this.bot) {
                try {
                    await this.bot.shutdown();
                }
                catch (shutdownError) {
                    logger_1.default.error('❌ Помилка при очищенні бота:', { error: shutdownError });
                }
                this.bot = null;
            }
            logger_1.default.info('✅ Ресурси очищено', {});
        }
        catch (error) {
            logger_1.default.error('❌ Помилка очищення ресурсів:', { error });
        }
    }
    /**
     * Логування статистики запуску
     */
    logStartupStats() {
        try {
            const stats = this.getStats();
            logger_1.default.info('📊 Статистика запуску:', {
                version: stats.version,
                uptime: `${Math.round(stats.uptime)}s`,
                memory: stats.memory,
                restartCount: stats.restartCount,
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка логування статистики запуску:', { error });
        }
    }
    /**
     * Налаштування graceful shutdown з покращеною обробкою
     */
    setupGracefulShutdown() {
        const shutdown = async (signal) => {
            logger_1.default.info(`📡 Отримано сигнал ${signal}, початок graceful shutdown...`, {});
            try {
                // Встановлення таймауту для shutdown
                const shutdownTimeout = setTimeout(() => {
                    logger_1.default.error('⏰ Таймаут graceful shutdown, примусова зупинка', {});
                    process.exit(1);
                }, APP_CONFIG.SHUTDOWN_TIMEOUT);
                await this.stop();
                clearTimeout(shutdownTimeout);
                logger_1.default.info('✅ Graceful shutdown завершено успішно', {});
                process.exit(0);
            }
            catch (error) {
                logger_1.default.error('❌ Помилка graceful shutdown:', { error });
                process.exit(1);
            }
        };
        // Обробка сигналів завершення
        process.on('SIGINT', () => shutdown('SIGINT'));
        process.on('SIGTERM', () => shutdown('SIGTERM'));
        process.on('SIGQUIT', () => shutdown('SIGQUIT'));
        logger_1.default.info('🛡️ Graceful shutdown налаштовано', {});
    }
}
// Глобальний екземпляр додатку
let app = null;
/**
 * Головна функція запуску з покращеною обробкою помилок
 */
async function main() {
    const startTime = Date.now();
    try {
        logger_1.default.info(`🎯 Запуск ${APP_CONFIG.NAME} v${APP_CONFIG.VERSION}`, {});
        app = new Application();
        await app.start();
        const totalStartupTime = Date.now() - startTime;
        logger_1.default.info(`🎉 Додаток повністю запущено за ${totalStartupTime}ms`, {});
    }
    catch (error) {
        const totalStartupTime = Date.now() - startTime;
        logger_1.default.error(`💥 Критична помилка при запуску після ${totalStartupTime}ms:`, { error });
        // Детальне логування помилки
        if (error instanceof Error) {
            logger_1.default.error('Деталі помилки:', {
                name: error.name,
                message: error.message,
                stack: error.stack,
            });
        }
        process.exit(1);
    }
}
/**
 * Функції для зовнішнього використання з покращеною обробкою помилок
 */
function getStats() {
    try {
        return app?.getStats() || { status: 'not_initialized' };
    }
    catch (error) {
        logger_1.default.error('❌ Помилка отримання статистики:', { error });
        return { status: 'error', error: error instanceof Error ? error.message : 'Невідома помилка' };
    }
}
async function restart() {
    try {
        if (!app) {
            throw new Error('Додаток не ініціалізовано');
        }
        return await app.restart();
    }
    catch (error) {
        logger_1.default.error('❌ Помилка перезапуску:', { error });
        throw error;
    }
}
async function shutdown() {
    try {
        if (!app) {
            logger_1.default.warn('⚠️ Додаток не ініціалізовано для зупинки', {});
            return;
        }
        return await app.stop();
    }
    catch (error) {
        logger_1.default.error('❌ Помилка зупинки:', { error });
        throw error;
    }
}
function getApp() {
    return app;
}
// Запуск додатку, якщо файл виконано напряму
if (require.main === module) {
    main().catch((error) => {
        logger_1.default.error('💥 Фатальна помилка в головній функції:', { error });
        process.exit(1);
    });
}
//# sourceMappingURL=index.js.map