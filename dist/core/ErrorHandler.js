"use strict";
/**
 * Error Handler для Discord бота
 * Централізована обробка помилок з покращеною архітектурою
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ErrorHandler = void 0;
const logger_1 = __importDefault(require("../utils/logger"));
class ErrorHandler {
    constructor(serviceContainer) {
        this.serviceContainer = serviceContainer;
        this.errorTypes = new Map();
        this.errorCounts = new Map();
        this.isActive = false;
        this.notificationQueue = [];
        this.maxQueueSize = 100;
    }
    /**
     * Ініціалізація обробника помилок
     */
    async initialize() {
        try {
            logger_1.default.info('🛡️ Ініціалізація обробника помилок...');
            // Реєстрація типів помилок
            this.registerErrorTypes();
            // Запуск обробки черги сповіщень
            this.startNotificationProcessor();
            this.isActive = true;
            logger_1.default.info('✅ Обробник помилок ініціалізовано');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка ініціалізації обробника помилок:', error);
            throw error;
        }
    }
    /**
     * Реєстрація типів помилок
     */
    registerErrorTypes() {
        // Discord API помилки
        this.errorTypes.set('DiscordAPIError', {
            severity: 'error',
            category: 'discord',
            retryable: true,
            maxRetries: 3,
            notificationThreshold: 5,
        });
        // Google API помилки
        this.errorTypes.set('GoogleAPIError', {
            severity: 'error',
            category: 'google',
            retryable: true,
            maxRetries: 3,
            notificationThreshold: 3,
        });
        // AI API помилки
        this.errorTypes.set('AIAPIError', {
            severity: 'error',
            category: 'ai',
            retryable: true,
            maxRetries: 2,
            notificationThreshold: 2,
        });
        // Валідація помилки
        this.errorTypes.set('ValidationError', {
            severity: 'warn',
            category: 'validation',
            retryable: false,
            maxRetries: 0,
            notificationThreshold: 10,
        });
        // Помилки кешу
        this.errorTypes.set('CacheError', {
            severity: 'warn',
            category: 'cache',
            retryable: true,
            maxRetries: 2,
            notificationThreshold: 5,
        });
        // Помилки файлів
        this.errorTypes.set('FileError', {
            severity: 'error',
            category: 'file',
            retryable: true,
            maxRetries: 2,
            notificationThreshold: 3,
        });
        // Помилки мережі
        this.errorTypes.set('NetworkError', {
            severity: 'error',
            category: 'network',
            retryable: true,
            maxRetries: 5,
            notificationThreshold: 2,
        });
        // Помилки бази даних
        this.errorTypes.set('DatabaseError', {
            severity: 'error',
            category: 'database',
            retryable: true,
            maxRetries: 3,
            notificationThreshold: 1,
        });
        logger_1.default.debug('✅ Типи помилок зареєстровано');
    }
    /**
     * Обробка помилки
     */
    async handle(error, context = {}) {
        try {
            // Класифікація помилки
            const errorInfo = this.classifyError(error);
            // Логування помилки
            this.logError(error, errorInfo, context);
            // Підрахунок помилок
            this.incrementErrorCount(errorInfo.type);
            // Перевірка чи потрібно відправити сповіщення
            if (this.shouldSendNotification(errorInfo.type)) {
                await this.queueNotification(error, errorInfo, context);
            }
            // Повернення інформації про помилку
            return {
                handled: true,
                errorInfo,
                retryable: errorInfo.retryable,
                maxRetries: errorInfo.maxRetries,
            };
        }
        catch (handleError) {
            logger_1.default.error('❌ Помилка в обробнику помилок:', handleError);
            return {
                handled: false,
                error: handleError,
            };
        }
    }
    /**
     * Обробка необроблених помилок
     */
    handleUncaughtException(error) {
        logger_1.default.error('🚨 КРИТИЧНА ПОМИЛКА - Uncaught Exception:', error);
        const errorInfo = {
            type: 'UncaughtException',
            severity: 'critical',
            category: 'system',
            retryable: false,
            maxRetries: 0,
            notificationThreshold: 0,
        };
        this.logError(error, errorInfo, { type: 'uncaughtException' });
        this.incrementErrorCount('UncaughtException');
        // Критичні помилки завжди потребують сповіщення
        this.queueNotification(error, errorInfo, { type: 'uncaughtException' });
        // Спроба graceful shutdown
        this.attemptGracefulShutdown();
    }
    /**
     * Обробка необроблених rejections
     */
    handleUnhandledRejection(reason, promise) {
        logger_1.default.error('🚨 КРИТИЧНА ПОМИЛКА - Unhandled Rejection:', reason);
        const errorInfo = {
            type: 'UnhandledRejection',
            severity: 'critical',
            category: 'system',
            retryable: false,
            maxRetries: 0,
            notificationThreshold: 0,
        };
        this.logError(reason, errorInfo, { type: 'unhandledRejection', promise });
        this.incrementErrorCount('UnhandledRejection');
        // Критичні помилки завжди потребують сповіщення
        this.queueNotification(reason, errorInfo, { type: 'unhandledRejection', promise });
        // Спроба graceful shutdown
        this.attemptGracefulShutdown();
    }
    /**
     * Спроба graceful shutdown
     */
    async attemptGracefulShutdown() {
        try {
            logger_1.default.warn('🛑 Спроба graceful shutdown через критичну помилку...');
            if (this.serviceContainer) {
                await this.serviceContainer.shutdown();
            }
            logger_1.default.info('✅ Graceful shutdown завершено');
            process.exit(1);
        }
        catch (error) {
            logger_1.default.error('❌ Помилка при graceful shutdown:', error);
            process.exit(1);
        }
    }
    /**
     * Додавання сповіщення до черги
     */
    async queueNotification(error, errorInfo, context) {
        const notification = {
            error,
            errorInfo,
            context,
            timestamp: new Date(),
            id: this.generateNotificationId(),
        };
        this.notificationQueue.push(notification);
        // Обмеження розміру черги
        if (this.notificationQueue.length > this.maxQueueSize) {
            this.notificationQueue.shift();
        }
        logger_1.default.debug(`📧 Сповіщення додано до черги: ${notification.id}`);
    }
    /**
     * Запуск обробника сповіщень
     */
    startNotificationProcessor() {
        setInterval(async () => {
            if (this.notificationQueue.length > 0) {
                const notification = this.notificationQueue.shift();
                await this.sendNotification(notification.error, notification.errorInfo, notification.context);
            }
        }, 5000); // Обробка кожні 5 секунд
    }
    /**
     * Генерація ID для сповіщення
     */
    generateNotificationId() {
        return `notif_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }
    /**
     * Класифікація помилки
     */
    classifyError(error) {
        const errorName = error.name || 'UnknownError';
        const defaultInfo = {
            type: errorName,
            severity: 'error',
            category: 'unknown',
            retryable: false,
            maxRetries: 0,
            notificationThreshold: 5,
        };
        const errorType = this.errorTypes.get(errorName);
        return errorType ? { ...errorType, type: errorName } : defaultInfo;
    }
    /**
     * Логування помилки
     */
    logError(error, errorInfo, context) {
        const logData = {
            error: {
                name: error.name,
                message: error.message,
                stack: error.stack,
            },
            errorInfo,
            context,
            timestamp: new Date().toISOString(),
        };
        switch (errorInfo.severity) {
            case 'critical':
                logger_1.default.error('🚨 КРИТИЧНА ПОМИЛКА:', logData);
                break;
            case 'error':
                logger_1.default.error('❌ ПОМИЛКА:', logData);
                break;
            case 'warn':
                logger_1.default.warn('⚠️ ПОПЕРЕДЖЕННЯ:', logData);
                break;
            default:
                logger_1.default.info('ℹ️ ІНФОРМАЦІЯ:', logData);
        }
    }
    /**
     * Підрахунок помилок
     */
    incrementErrorCount(type) {
        const currentCount = this.errorCounts.get(type) || 0;
        this.errorCounts.set(type, currentCount + 1);
    }
    /**
     * Відправка сповіщення
     */
    async sendNotification(error, errorInfo, context) {
        try {
            // Спроба відправити Discord сповіщення
            await this.sendDiscordNotification(error, errorInfo, context);
            // Можна додати інші канали сповіщень (email, Slack, etc.)
        }
        catch (notificationError) {
            logger_1.default.error('❌ Помилка відправки сповіщення:', notificationError);
        }
    }
    /**
     * Перевірка чи потрібно відправити сповіщення
     */
    shouldSendNotification(errorType) {
        const errorInfo = this.errorTypes.get(errorType);
        if (!errorInfo)
            return false;
        const currentCount = this.errorCounts.get(errorType) || 0;
        return currentCount >= errorInfo.notificationThreshold;
    }
    /**
     * Відправка Discord сповіщення
     */
    async sendDiscordNotification(error, errorInfo, context) {
        try {
            // Отримання Discord клієнта з Service Container
            const bot = this.serviceContainer?.get('bot');
            if (!bot || !bot.client) {
                logger_1.default.warn('Discord клієнт недоступний для сповіщення');
                return;
            }
            const channel = this.findNotificationChannel(bot.client);
            if (!channel) {
                logger_1.default.warn('Канал сповіщень не знайдено');
                return;
            }
            const embed = {
                color: this.getErrorColor(errorInfo.severity),
                title: `🚨 Помилка: ${errorInfo.type}`,
                description: this.getUserFriendlyMessage(error, errorInfo),
                fields: [
                    {
                        name: 'Категорія',
                        value: errorInfo.category,
                        inline: true,
                    },
                    {
                        name: 'Серйозність',
                        value: errorInfo.severity,
                        inline: true,
                    },
                    {
                        name: 'Контекст',
                        value: JSON.stringify(context, null, 2),
                        inline: false,
                    },
                ],
                timestamp: new Date(),
                footer: {
                    text: 'Discord AI Bot - Error Handler',
                },
            };
            await channel.send({ embeds: [embed] });
            logger_1.default.info('✅ Discord сповіщення відправлено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка відправки Discord сповіщення:', error);
        }
    }
    /**
     * Отримання кольору для помилки
     */
    getErrorColor(severity) {
        switch (severity) {
            case 'critical':
                return 0xFF0000; // Червоний
            case 'error':
                return 0xFF6B6B; // Світло-червоний
            case 'warn':
                return 0xFFA500; // Помаранчевий
            default:
                return 0x808080; // Сірий
        }
    }
    /**
     * Пошук каналу для сповіщень
     */
    findNotificationChannel(client) {
        try {
            // Пошук каналу з назвою "errors" або "logs"
            for (const guild of client.guilds.cache.values()) {
                const errorChannel = guild.channels.cache.find((channel) => channel.type === 0 && // Text channel
                    (channel.name.includes('error') ||
                        channel.name.includes('log') ||
                        channel.name.includes('admin')));
                if (errorChannel) {
                    return errorChannel;
                }
            }
            return null;
        }
        catch (error) {
            logger_1.default.error('Помилка пошуку каналу сповіщень:', error);
            return null;
        }
    }
    /**
     * Отримання зрозумілого повідомлення про помилку
     */
    getUserFriendlyMessage(error, errorInfo) {
        const baseMessage = error.message || 'Невідома помилка';
        switch (errorInfo.type) {
            case 'DiscordAPIError':
                return `Помилка Discord API: ${baseMessage}`;
            case 'GoogleAPIError':
                return `Помилка Google API: ${baseMessage}`;
            case 'AIAPIError':
                return `Помилка AI сервісу: ${baseMessage}`;
            case 'ValidationError':
                return `Помилка валідації: ${baseMessage}`;
            case 'CacheError':
                return `Помилка кешу: ${baseMessage}`;
            case 'FileError':
                return `Помилка файлу: ${baseMessage}`;
            case 'NetworkError':
                return `Мережева помилка: ${baseMessage}`;
            case 'DatabaseError':
                return `Помилка бази даних: ${baseMessage}`;
            default:
                return baseMessage;
        }
    }
    /**
     * Отримання статистики помилок
     */
    getStats() {
        return {
            totalErrors: Array.from(this.errorCounts.values()).reduce((a, b) => a + b, 0),
            errorCounts: Object.fromEntries(this.errorCounts),
            notificationQueueSize: this.notificationQueue.length,
            isActive: this.isActive,
        };
    }
    /**
     * Очищення статистики помилок
     */
    clearErrorStats() {
        this.errorCounts.clear();
        this.notificationQueue.length = 0;
        logger_1.default.info('✅ Статистика помилок очищена');
    }
    /**
     * Перевірка активності
     */
    isActive() {
        return this.isActive;
    }
    /**
     * Завершення роботи
     */
    async shutdown() {
        logger_1.default.info('🛑 Завершення роботи Error Handler...');
        this.isActive = false;
        // Обробка залишкових сповіщень
        while (this.notificationQueue.length > 0) {
            const notification = this.notificationQueue.shift();
            await this.sendNotification(notification.error, notification.errorInfo, notification.context);
        }
        logger_1.default.info('✅ Error Handler завершено');
    }
}
exports.ErrorHandler = ErrorHandler;
//# sourceMappingURL=ErrorHandler.js.map