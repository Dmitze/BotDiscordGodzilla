/**
 * Менеджер сервісів Discord бота
 * Централізоване управління всіма сервісами
 * TypeScript версія
 */
interface Bot {
    config: {
        redis: {
            enabled: boolean;
        };
        isMetricsEnabled(): boolean;
    };
}
interface Service {
    initialize?: () => Promise<void>;
    start?: () => Promise<void>;
    shutdown?: () => Promise<void>;
    isActive?: () => boolean;
    getStats?: () => any;
    [key: string]: any;
}
interface ServiceStatus {
    isActive: boolean;
    hasMethod: (method: string) => boolean;
    stats: any;
}
interface ServiceManagerStats {
    total: number;
    active: number;
    services: string[];
    status: Record<string, ServiceStatus>;
}
declare class ServiceManager {
    private bot;
    private services;
    private isInitialized;
    constructor(bot: Bot);
    /**
     * Ініціалізація менеджера сервісів
     */
    initialize(): Promise<void>;
    /**
     * Створення сервісів
     */
    private createServices;
    /**
     * Ініціалізація сервісів
     */
    private initializeServices;
    /**
     * Запуск метрик
     */
    startMetrics(): Promise<void>;
    /**
     * Запуск кешування
     */
    startCache(): Promise<void>;
    /**
     * Запуск планувальника
     */
    startScheduler(): Promise<void>;
    /**
     * Отримання сервісу за назвою
     */
    getService(name: string): Service | undefined;
    /**
     * Перевірка наявності сервісу
     */
    hasService(name: string): boolean;
    /**
     * Отримання всіх сервісів
     */
    getAllServices(): Service[];
    /**
     * Отримання назв всіх сервісів
     */
    getServiceNames(): string[];
    /**
     * Виконання методу на всіх сервісах
     */
    executeOnAllServices(methodName: string, ...args: any[]): Promise<PromiseSettledResult<any>[]>;
    /**
     * Отримання статусу сервісів
     */
    getServicesStatus(): Record<string, ServiceStatus>;
    /**
     * Graceful shutdown всіх сервісів
     */
    shutdown(): Promise<void>;
    /**
     * Статистика сервісів
     */
    getStats(): ServiceManagerStats;
}
export default ServiceManager;
//# sourceMappingURL=ServiceManager.d.ts.map