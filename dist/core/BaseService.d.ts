/**
 * Базовий клас для всіх сервісів
 * Надає спільну функціональність та інтерфейс
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { BaseService as IBaseService, BotConfig, HealthStatus, ServiceStats } from '@/types';
export declare abstract class BaseService implements IBaseService {
    readonly name: string;
    readonly config: BotConfig;
    protected isInitialized: boolean;
    protected startTime: number;
    protected isShuttingDown: boolean;
    protected retryCount: number;
    private initializationTimeout;
    private shutdownTimeout;
    constructor(name: string, config: BotConfig);
    /**
     * Ініціалізація сервісу з детальним логуванням
     */
    initialize(): Promise<void>;
    /**
     * Завершення роботи сервісу з детальним логуванням
     */
    shutdown(): Promise<void>;
    /**
     * Перевірка здоров'я сервісу з детальним логуванням
     */
    healthCheck(): Promise<HealthStatus>;
    /**
     * Отримання статистики сервісу з детальним логуванням
     */
    getStats(): ServiceStats;
    /**
     * Перевірка чи сервіс ініціалізовано
     */
    protected checkInitialized(): void;
    /**
     * Перевірка чи сервіс не зупиняється
     */
    protected checkNotShuttingDown(): void;
    /**
     * Безпечне виконання операції з обробкою помилок
     */
    protected safeExecute<T>(operation: () => Promise<T>, operationName: string, fallback?: T): Promise<T>;
    /**
     * Очищення ресурсів сервісу
     */
    protected cleanup(): Promise<void>;
    /**
     * Абстрактні методи для реалізації в нащадках
     */
    protected abstract onInitialize(): Promise<void>;
    protected abstract onShutdown(): Promise<void>;
    protected abstract onHealthCheck(): Promise<HealthStatus>;
    protected abstract onGetStats(): Partial<ServiceStats>;
}
//# sourceMappingURL=BaseService.d.ts.map