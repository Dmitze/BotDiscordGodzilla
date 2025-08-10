/**
 * Контейнер сервісів з Dependency Injection
 * Централізоване управління всіма сервісами
 */
import type { BaseService, BotConfig, HealthStatus } from '@/types';
export declare class ServiceContainer {
    private services;
    private readonly config;
    constructor(config: BotConfig);
    /**
     * Реєстрація сервісу
     */
    register<T extends BaseService>(name: string, service: T): void;
    /**
     * Отримання сервісу
     */
    get<T extends BaseService>(name: string): T;
    /**
     * Перевірка чи сервіс існує
     */
    has(name: string): boolean;
    /**
     * Отримання всіх сервісів
     */
    getAll(): Map<string, BaseService>;
    /**
     * Ініціалізація всіх сервісів
     */
    initialize(): Promise<void>;
    /**
     * Завершення роботи всіх сервісів
     */
    shutdown(): Promise<void>;
    /**
     * Health check всіх сервісів
     */
    getHealthStatus(): Promise<Record<string, HealthStatus>>;
    /**
     * Отримання статистики всіх сервісів
     */
    getAllStats(): Record<string, unknown>;
    /**
     * Видалення сервісу
     */
    remove(name: string): boolean;
    /**
     * Очищення всіх сервісів
     */
    clear(): void;
    /**
     * Отримання кількості сервісів
     */
    get size(): number;
}
//# sourceMappingURL=ServiceContainer.d.ts.map