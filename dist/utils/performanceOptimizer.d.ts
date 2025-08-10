/**
 * Утиліти для оптимізації продуктивності Discord AI Assistant Bot
 * Моніторинг та оптимізація продуктивності
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { SystemMetrics } from '@/types';
/**
 * Клас для оптимізації продуктивності
 */
export declare class PerformanceOptimizer {
    private static instance;
    private performanceData;
    private caches;
    private metricsHistory;
    private gcInterval;
    private metricsInterval;
    private cacheCleanupInterval;
    private isInitialized;
    private isShuttingDown;
    constructor();
    /**
     * Запуск моніторингу продуктивності
     */
    private startMonitoring;
    /**
     * Збір метрик продуктивності
     */
    private collectMetrics;
    /**
     * Отримання використання CPU
     */
    private getCPUUsage;
    /**
     * Перевірка порогів продуктивності
     */
    private checkThresholds;
    /**
     * Оптимізація пам'яті
     */
    private optimizeMemory;
    /**
     * Оптимізація CPU
     */
    private optimizeCPU;
    /**
     * Виконання garbage collection
     */
    private performGarbageCollection;
    /**
     * Очищення кешів
     */
    private cleanupCaches;
    /**
     * Вимірювання часу виконання операції
     */
    measureOperation<T>(operation: () => Promise<T>, operationName: string, category?: string): Promise<T>;
    /**
     * Запис операції
     */
    private recordOperation;
    /**
     * Кешування з автоматичним очищенням
     */
    cache<T>(cacheName: string, key: string, value: T, ttl?: number): void;
    /**
     * Отримання з кешу
     */
    getFromCache<T>(cacheName: string, key: string): T | null;
    /**
     * Отримання статистики продуктивності
     */
    getPerformanceStats(): Record<string, unknown>;
    /**
     * Отримання системних метрик
     */
    getSystemMetrics(): SystemMetrics;
    /**
     * Очищення ресурсів
     */
    cleanup(): void;
    /**
     * Перевірка стану ініціалізації
     */
    getInitializedState(): boolean;
    /**
     * Перевірка стану зупинки
     */
    getShuttingDownState(): boolean;
}
export declare const performanceOptimizer: PerformanceOptimizer;
export declare const measureOperation: <T>(operation: () => Promise<T>, operationName: string, category?: string) => Promise<T>;
export declare const cache: <T>(cacheName: string, key: string, value: T, ttl?: number) => void;
export declare const getFromCache: <T>(cacheName: string, key: string) => T | null;
export declare const getPerformanceStats: () => Record<string, unknown>;
export declare const getSystemMetrics: () => SystemMetrics;
export declare const cleanupPerformanceOptimizer: () => void;
//# sourceMappingURL=performanceOptimizer.d.ts.map