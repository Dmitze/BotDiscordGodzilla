"use strict";
/**
 * Контейнер сервісів з Dependency Injection
 * Централізоване управління всіма сервісами
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ServiceContainer = void 0;
class ServiceContainer {
    constructor(config) {
        this.services = new Map();
        this.config = config;
    }
    /**
     * Реєстрація сервісу
     */
    register(name, service) {
        if (this.services.has(name)) {
            throw new Error(`Сервіс ${name} вже зареєстрований`);
        }
        this.services.set(name, service);
    }
    /**
     * Отримання сервісу
     */
    get(name) {
        const service = this.services.get(name);
        if (!service) {
            throw new Error(`Сервіс ${name} не знайдено`);
        }
        return service;
    }
    /**
     * Перевірка чи сервіс існує
     */
    has(name) {
        return this.services.has(name);
    }
    /**
     * Отримання всіх сервісів
     */
    getAll() {
        return new Map(this.services);
    }
    /**
     * Ініціалізація всіх сервісів
     */
    async initialize() {
        const initPromises = [];
        for (const [name, service] of this.services.entries()) {
            try {
                initPromises.push(service.initialize());
            }
            catch (error) {
                throw new Error(`Помилка ініціалізації сервісу ${name}: ${error}`);
            }
        }
        await Promise.all(initPromises);
    }
    /**
     * Завершення роботи всіх сервісів
     */
    async shutdown() {
        const shutdownPromises = [];
        for (const [name, service] of this.services.entries()) {
            try {
                shutdownPromises.push(service.shutdown());
            }
            catch (error) {
                console.error(`Помилка завершення сервісу ${name}:`, error);
            }
        }
        await Promise.all(shutdownPromises);
    }
    /**
     * Health check всіх сервісів
     */
    async getHealthStatus() {
        const healthStatus = {};
        for (const [name, service] of this.services.entries()) {
            try {
                healthStatus[name] = await service.healthCheck();
            }
            catch (error) {
                healthStatus[name] = {
                    healthy: false,
                    service: name,
                    error: `Помилка health check: ${error}`,
                };
            }
        }
        return healthStatus;
    }
    /**
     * Отримання статистики всіх сервісів
     */
    getAllStats() {
        const stats = {};
        for (const [name, service] of this.services.entries()) {
            try {
                stats[name] = service.getStats();
            }
            catch (error) {
                stats[name] = {
                    error: `Помилка отримання статистики: ${error}`,
                };
            }
        }
        return stats;
    }
    /**
     * Видалення сервісу
     */
    remove(name) {
        return this.services.delete(name);
    }
    /**
     * Очищення всіх сервісів
     */
    clear() {
        this.services.clear();
    }
    /**
     * Отримання кількості сервісів
     */
    get size() {
        return this.services.size;
    }
}
exports.ServiceContainer = ServiceContainer;
//# sourceMappingURL=ServiceContainer.js.map