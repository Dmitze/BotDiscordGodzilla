/**
 * Metrics Service для Discord бота
 * Централізоване управління метриками та моніторингом
 */
import type { BotConfig, HealthStatus, ServiceStats, CacheStats, QueueStats } from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
interface MetricsServiceStats extends ServiceStats {
    requests: number;
    errors: number;
    startTime: number;
    metricsCount: number;
}
export declare class MetricsService extends BaseServiceClass {
    private registry;
    private metrics;
    private server;
    private stats;
    private updateInterval;
    constructor(config: BotConfig);
    /**
     * Ініціалізація Metrics сервісу
     */
    protected onInitialize(): Promise<void>;
    /**
     * Створення Prometheus реєстру
     */
    private createRegistry;
    /**
     * Створення метрик
     */
    private createMetrics;
    /**
     * Запуск HTTP сервера
     */
    private startServer;
    /**
     * Інкремент лічильника команд
     */
    incrementCommand(command: string, status?: string): void;
    /**
     * Інкремент лічильника повідомлень
     */
    incrementMessage(type: string): void;
    /**
     * Інкремент лічильника помилок
     */
    incrementError(type: string, service?: string): void;
    /**
     * Встановлення кількості активних користувачів
     */
    setActiveUsers(count: number): void;
    /**
     * Встановлення кількості активних серверів
     */
    setActiveGuilds(count: number): void;
    /**
     * Оновлення використання пам'яті
     */
    updateMemoryUsage(): void;
    /**
     * Оновлення часу роботи
     */
    updateUptime(): void;
    /**
     * Вимірювання тривалості команди
     */
    measureCommandDuration(command: string, duration: number): void;
    /**
     * Вимірювання часу відповіді API
     */
    measureApiResponseTime(service: string, endpoint: string, duration: number): void;
    /**
     * Оновлення метрик кешу
     */
    updateCacheMetrics(cacheStats: CacheStats): void;
    /**
     * Оновлення метрик черг
     */
    updateQueueMetrics(queueStats: QueueStats): void;
    /**
     * Оновлення метрик connection pool
     */
    updateConnectionPoolMetrics(connectionStats: Record<string, unknown>): void;
    /**
     * Оновлення AI метрик
     */
    updateAIMetrics(provider: string, status: string, duration: number): void;
    /**
     * Оновлення Google API метрик
     */
    updateGoogleApiMetrics(service: string, endpoint: string, status: string, duration: number): void;
    /**
     * Оновлення всіх метрик
     */
    updateAllMetrics(): void;
    /**
     * Запуск періодичних оновлень
     */
    private startPeriodicUpdates;
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
    protected onGetStats(): Partial<MetricsServiceStats>;
}
export {};
//# sourceMappingURL=MetricsService.d.ts.map