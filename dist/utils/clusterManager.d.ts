/**
 * Cluster Manager для масштабування Discord бота
 * Підтримка кластеризації та load balancing
 * TypeScript версія
 */
interface ClusterConfig {
    workers: number;
    restartDelay: number;
    maxRestarts: number;
    [key: string]: any;
}
interface ClusterStats {
    totalWorkers: number;
    activeWorkers: number;
    restarts: number;
    startTime: number;
    uptime: number;
}
interface WorkerMessage {
    type: string;
    data: any;
    timestamp: number;
}
declare class ClusterManager {
    private config;
    private workers;
    private restartCounts;
    private isMaster;
    private isActive;
    private stats;
    private static formatError;
    private static formatAny;
    constructor(config?: Partial<ClusterConfig>);
    /**
     * Запуск кластера
     */
    start(): Promise<void>;
    /**
     * Створення worker процесу
     */
    private createWorker;
    /**
     * Налаштування обробників подій
     */
    private setupEventHandlers;
    /**
     * Обробка повідомлень від workers
     */
    private handleWorkerMessage;
    /**
     * Обробка виходу worker
     */
    private handleWorkerExit;
    /**
     * Перевірка чи потрібно перезапустити worker
     */
    private shouldRestartWorker;
    /**
     * Перезапуск worker
     */
    private restartWorker;
    /**
     * Оновлення статистики worker
     */
    private updateWorkerStats;
    /**
     * Отримання оптимального worker для розподілу навантаження
     */
    getOptimalWorker(): number | null;
    /**
     * Відправка повідомлення конкретному worker
     */
    sendToWorker(workerId: number, message: WorkerMessage): boolean;
    /**
     * Розсилка повідомлення всім workers
     */
    broadcast(message: WorkerMessage): void;
    /**
     * Отримання статистики кластера
     */
    getClusterStats(): ClusterStats;
    /**
     * Зупинка кластера
     */
    stop(): Promise<void>;
    /**
     * Перезапуск кластера
     */
    restart(): Promise<void>;
    /**
     * Перевірка чи кластер активний
     */
    isClusterActive(): boolean;
    /**
     * Отримання кількості активних workers
     */
    getActiveWorkersCount(): number;
}
export default ClusterManager;
export { ClusterManager };
//# sourceMappingURL=clusterManager.d.ts.map