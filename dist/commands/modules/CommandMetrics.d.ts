/**
 * Система метрик та статистики для команд Discord бота
 * Збір, аналіз та звітність по використанню команд
 * Версія 1.0.0 - Виокремлено з BaseCommand
 */
export interface CommandMetrics {
    commandName: string;
    executionCount: number;
    successCount: number;
    errorCount: number;
    averageExecutionTime: number;
    totalExecutionTime: number;
    lastExecuted: number;
    slowExecutions: number;
    cacheHits: number;
    cacheMisses: number;
    userCount: number;
    retries: number;
    cooldownHits: number;
}
export interface ExecutionMetric {
    commandName: string;
    userId: string;
    executionTime: number;
    success: boolean;
    error?: string;
    timestamp: number;
    fromCache: boolean;
    retryCount: number;
}
export interface PerformanceThresholds {
    slowExecutionMs: number;
    verySlowExecutionMs: number;
    maxExecutionTime: number;
    warningErrorRate: number;
    criticalErrorRate: number;
}
export declare class CommandMetricsCollector {
    private static instance;
    private metrics;
    private executionHistory;
    private readonly maxHistorySize;
    private readonly thresholds;
    constructor(thresholds?: Partial<PerformanceThresholds>);
    /**
     * Записати метрику виконання команди
     */
    recordExecution(commandName: string, userId: string, executionTime: number, success: boolean, options?: {
        error?: string;
        fromCache?: boolean;
        retryCount?: number;
    }): void;
    /**
     * Створити порожні метрики для нової команди
     */
    private createEmptyMetrics;
    /**
     * Оновити метрики команди
     */
    private updateCommandMetrics;
    /**
     * Додати виконання в історію
     */
    private addToHistory;
    /**
     * Отримати метрики команди
     */
    getCommandMetrics(commandName: string): CommandMetrics | undefined;
    /**
     * Отримати всі метрики
     */
    getAllMetrics(): CommandMetrics[];
    /**
     * Отримати топ команд за використанням
     */
    getTopCommands(limit?: number): CommandMetrics[];
    /**
     * Отримати команди з найбільшою кількістю помилок
     */
    getCommandsWithMostErrors(limit?: number): CommandMetrics[];
    /**
     * Отримати найповільніші команди
     */
    getSlowestCommands(limit?: number): CommandMetrics[];
    /**
     * Аналіз трендів використання
     */
    analyzeTrends(timeframe?: number): {
        totalExecutions: number;
        errorRate: number;
        averageResponseTime: number;
        topCommands: string[];
        trendDirection: 'up' | 'down' | 'stable';
    };
    /**
     * Генерація звіту про продуктивність
     */
    generatePerformanceReport(): {
        summary: {
            totalCommands: number;
            totalExecutions: number;
            overallErrorRate: number;
            averageResponseTime: number;
        };
        alerts: Array<{
            level: 'warning' | 'critical';
            message: string;
            command?: string;
            metric?: string;
            value?: number;
        }>;
        recommendations: string[];
    };
    /**
     * Записати cooldown hit
     */
    recordCooldownHit(commandName: string): void;
    /**
     * Очистити метрики (для тестування)
     */
    clearMetrics(): void;
    /**
     * Періодичне звітування
     */
    private startPeriodicReporting;
    /**
     * Експорт метрик для Prometheus
     */
    exportPrometheusMetrics(): string;
}
export default CommandMetricsCollector;
//# sourceMappingURL=CommandMetrics.d.ts.map