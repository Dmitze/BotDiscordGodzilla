/**
 * Основний файл Discord AI Assistant Bot
 * Точка входу в додаток
 * Версія 3.0.0 - Повністю рефакторовано з TypeScript
 */
declare const APP_CONFIG: {
    readonly VERSION: "3.0.0";
    readonly NAME: "Discord AI Assistant Bot";
    readonly STARTUP_TIMEOUT: 30000;
    readonly SHUTDOWN_TIMEOUT: 10000;
    readonly RESTART_DELAY: 5000;
    readonly MAX_MEMORY_USAGE: number;
    readonly HEALTH_CHECK_INTERVAL: 30000;
};
declare class Application {
    private bot;
    private config;
    private isStarting;
    private isShuttingDown;
    private startupTime;
    private restartCount;
    private readonly maxRestarts;
    private healthCheckInterval;
    private memoryCheckInterval;
    constructor();
    /**
     * Запуск додатку з детальним логуванням
     */
    start(): Promise<void>;
    /**
     * Зупинка додатку з детальним логуванням
     */
    stop(): Promise<void>;
    /**
     * Отримання детальної статистики
     */
    getStats(): any;
    /**
     * Перезапуск додатку з обмеженнями
     */
    restart(): Promise<void>;
    /**
     * Валідація конфігурації
     */
    private validateConfiguration;
    /**
     * Перевірка системних ресурсів
     */
    private checkSystemResources;
    /**
     * Запуск моніторингу
     */
    private startMonitoring;
    /**
     * Зупинка моніторингу
     */
    private stopMonitoring;
    /**
     * Отримання вкладених значень об'єкта
     */
    private getNestedValue;
    /**
     * Очищення ресурсів при помилці
     */
    private cleanupOnError;
    /**
     * Логування статистики запуску
     */
    private logStartupStats;
    /**
     * Налаштування graceful shutdown з покращеною обробкою
     */
    private setupGracefulShutdown;
}
/**
 * Головна функція запуску з покращеною обробкою помилок
 */
declare function main(): Promise<void>;
/**
 * Функції для зовнішнього використання з покращеною обробкою помилок
 */
export declare function getStats(): any;
export declare function restart(): Promise<void>;
export declare function shutdown(): Promise<void>;
export declare function getApp(): Application | null;
export { main, APP_CONFIG };
//# sourceMappingURL=index.d.ts.map