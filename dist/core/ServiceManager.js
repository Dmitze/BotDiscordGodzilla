"use strict";
/**
 * Менеджер сервісів Discord бота
 * Централізоване управління всіма сервісами
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const logger_1 = __importDefault(require("../utils/logger"));
const AIService_1 = __importDefault(require("../services/AIService"));
const GoogleService_1 = __importDefault(require("../services/GoogleService"));
const CacheService_1 = __importDefault(require("../services/CacheService"));
const MetricsService_1 = __importDefault(require("../services/MetricsService"));
const SchedulerService_1 = __importDefault(require("../services/SchedulerService"));
class ServiceManager {
    constructor(bot) {
        this.bot = bot;
        this.services = new Map();
        this.isInitialized = false;
    }
    /**
     * Ініціалізація менеджера сервісів
     */
    async initialize() {
        try {
            logger_1.default.info('🔧 Ініціалізація менеджера сервісів...');
            // Створення сервісів
            await this.createServices();
            // Ініціалізація сервісів
            await this.initializeServices();
            this.isInitialized = true;
            logger_1.default.info('✅ Менеджер сервісів ініціалізовано');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка ініціалізації менеджера сервісів:', error);
            throw error;
        }
    }
    /**
     * Створення сервісів
     */
    async createServices() {
        // AI Service
        this.services.set('ai', new AIService_1.default(this.bot));
        // Google Service
        this.services.set('google', new GoogleService_1.default(this.bot));
        // Cache Service (якщо Redis увімкнено)
        if (this.bot.config.redis.enabled) {
            this.services.set('cache', new CacheService_1.default(this.bot));
        }
        // Metrics Service (якщо метрики увімкнені)
        if (this.bot.config.isMetricsEnabled()) {
            this.services.set('metrics', new MetricsService_1.default(this.bot));
        }
        // Scheduler Service
        this.services.set('scheduler', new SchedulerService_1.default(this.bot));
    }
    /**
     * Ініціалізація сервісів
     */
    async initializeServices() {
        const initPromises = Array.from(this.services.entries()).map(async ([name, service]) => {
            try {
                if (service.initialize) {
                    await service.initialize();
                    logger_1.default.debug(`✅ Сервіс ${name} ініціалізовано`);
                }
            }
            catch (error) {
                logger_1.default.error(`❌ Помилка ініціалізації сервісу ${name}:`, error);
                // Видаляємо сервіс, який не вдалося ініціалізувати
                this.services.delete(name);
            }
        });
        await Promise.allSettled(initPromises);
    }
    /**
     * Запуск метрик
     */
    async startMetrics() {
        const metricsService = this.services.get('metrics');
        if (metricsService && metricsService.start) {
            await metricsService.start();
            logger_1.default.info('📊 Метрики запущено');
        }
    }
    /**
     * Запуск кешування
     */
    async startCache() {
        const cacheService = this.services.get('cache');
        if (cacheService && cacheService.start) {
            await cacheService.start();
            logger_1.default.info('💾 Кеш запущено');
        }
    }
    /**
     * Запуск планувальника
     */
    async startScheduler() {
        const schedulerService = this.services.get('scheduler');
        if (schedulerService && schedulerService.start) {
            await schedulerService.start();
            logger_1.default.info('⏰ Планувальник запущено');
        }
    }
    /**
     * Отримання сервісу за назвою
     */
    getService(name) {
        return this.services.get(name);
    }
    /**
     * Перевірка наявності сервісу
     */
    hasService(name) {
        return this.services.has(name);
    }
    /**
     * Отримання всіх сервісів
     */
    getAllServices() {
        return Array.from(this.services.values());
    }
    /**
     * Отримання назв всіх сервісів
     */
    getServiceNames() {
        return Array.from(this.services.keys());
    }
    /**
     * Виконання методу на всіх сервісах
     */
    async executeOnAllServices(methodName, ...args) {
        const promises = Array.from(this.services.values()).map(async (service) => {
            if (service[methodName] && typeof service[methodName] === 'function') {
                try {
                    return await service[methodName](...args);
                }
                catch (error) {
                    logger_1.default.error(`Помилка виконання ${methodName} на сервісі:`, error);
                    return null;
                }
            }
            return null;
        });
        return Promise.allSettled(promises);
    }
    /**
     * Отримання статусу сервісів
     */
    getServicesStatus() {
        const status = {};
        for (const [name, service] of this.services.entries()) {
            status[name] = {
                isActive: service.isActive ? service.isActive() : true,
                hasMethod: (method) => service[method] && typeof service[method] === 'function',
                stats: service.getStats ? service.getStats() : null,
            };
        }
        return status;
    }
    /**
     * Graceful shutdown всіх сервісів
     */
    async shutdown() {
        logger_1.default.info('🛑 Завершення роботи сервісів...');
        try {
            await this.executeOnAllServices('shutdown');
            logger_1.default.info('✅ Сервіси успішно завершено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка при завершенні сервісів:', error);
        }
    }
    /**
     * Статистика сервісів
     */
    getStats() {
        return {
            total: this.services.size,
            active: Array.from(this.services.values()).filter(service => service.isActive ? service.isActive() : true).length,
            services: this.getServiceNames(),
            status: this.getServicesStatus(),
        };
    }
}
exports.default = ServiceManager;
//# sourceMappingURL=ServiceManager.js.map