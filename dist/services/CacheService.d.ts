/**
 * Redis Cache Service
 * Оптимізоване кешування з підтримкою різних стратегій
 */
import type { BotConfig, HealthStatus, ServiceStats, CacheStats, CacheOptions } from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
interface CacheServiceStats extends ServiceStats {
    hits: number;
    misses: number;
    sets: number;
    deletes: number;
    errors: number;
    hitRate: number;
    totalRequests: number;
}
interface CacheServiceOptions extends CacheOptions {
    compress?: boolean;
    serialize?: boolean;
}
export declare class CacheService extends BaseServiceClass {
    private client;
    private isConnected;
    private stats;
    private readonly defaultTTL;
    private readonly maxRetries;
    private readonly retryDelay;
    constructor(config: BotConfig);
    /**
     * Ініціалізація Redis клієнта
     */
    protected onInitialize(): Promise<void>;
    /**
     * Налаштування обробників подій Redis
     */
    private setupEventHandlers;
    /**
     * Підключення до Redis
     */
    private connect;
    /**
     * Валідація підключення
     */
    private validateConnection;
    /**
     * Отримання значення з кешу
     */
    get<T = unknown>(key: string, options?: CacheServiceOptions): Promise<T | null>;
    /**
     * Збереження значення в кеш
     */
    set<T = unknown>(key: string, value: T, ttl?: number, options?: CacheServiceOptions): Promise<boolean>;
    /**
     * Видалення ключа з кешу
     */
    delete(key: string): Promise<boolean>;
    /**
     * Видалення ключів за патерном
     */
    deletePattern(pattern: string): Promise<number>;
    /**
     * Перевірка існування ключа
     */
    exists(key: string): Promise<boolean>;
    /**
     * Встановлення TTL для ключа
     */
    expire(key: string, ttl: number): Promise<boolean>;
    /**
     * Отримання TTL ключа
     */
    ttl(key: string): Promise<number>;
    /**
     * Отримання або встановлення значення
     */
    getOrSet<T = unknown>(key: string, fallbackFn: () => Promise<T>, ttl?: number, options?: CacheServiceOptions): Promise<T>;
    /**
     * Отримання множинних значень
     */
    mget<T = unknown>(keys: string[], options?: CacheServiceOptions): Promise<(T | null)[]>;
    /**
     * Збереження множинних значень
     */
    mset<T = unknown>(keyValuePairs: Array<{
        key: string;
        value: T;
        ttl?: number;
    }>, defaultTTL?: number): Promise<boolean>;
    /**
     * Очищення всього кешу
     */
    clear(): Promise<boolean>;
    /**
     * Отримання статистики кешу
     */
    getCacheStats(): CacheStats;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Health check
     */
    protected onHealthCheck(): Promise<HealthStatus>;
    /**
     * Завершення роботи
     */
    protected onShutdown(): Promise<void>;
    /**
     * Отримання статистики
     */
    protected onGetStats(): Partial<CacheServiceStats>;
}
export {};
//# sourceMappingURL=CacheService.d.ts.map