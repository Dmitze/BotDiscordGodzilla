/**
 * 📋 Queue Manager Module
 * Система черг для асинхронної обробки завдань
 * TypeScript версія
 *
 * Функції:
 * - Черги для різних типів завдань
 * - Пріоритизація завдань
 * - Обробка помилок
 * - Моніторинг черг
 */
import { EventEmitter } from 'events';
interface JobData {
    id: string;
    priority: 'high' | 'normal' | 'low';
    job: Function | TypedJob;
    timestamp: number;
    retries: number;
    maxRetries: number;
}
interface TypedJob {
    type: 'sheets_query' | 'ai_request' | 'file_operation' | 'export_data';
    data: any;
    handler?: Function;
}
interface QueueStats {
    processed: number;
    failed: number;
    pending: number;
    averageProcessingTime: number;
}
interface QueueInfo {
    pending: number;
    active: number;
    maxConcurrent: number;
}
interface QueueStatsResult {
    queues: {
        high: QueueInfo;
        normal: QueueInfo;
        low: QueueInfo;
    };
    stats: QueueStats & {
        averageProcessingTime: number;
    };
    totalPending: number;
    totalActive: number;
}
interface OptimizationRecommendation {
    type: 'queue' | 'performance' | 'reliability';
    priority: 'high' | 'medium' | 'low';
    message: string;
    action: string;
}
declare class QueueManager extends EventEmitter {
    private queues;
    private processing;
    private stats;
    private maxConcurrent;
    private activeJobs;
    constructor();
    /**
     * Додавання завдання в чергу
     */
    addJob(priority: 'high' | 'normal' | 'low', job: Function | TypedJob): string;
    /**
     * Генерація унікального ID завдання
     */
    private generateJobId;
    /**
     * Запуск обробки черг
     */
    private startProcessing;
    /**
     * Обробка черги
     */
    private processQueue;
    /**
     * Отримання статистики черг
     */
    getQueueStats(): QueueStatsResult;
    /**
     * Виконання завдання
     */
    private executeJob;
    /**
     * Виконання типізованого завдання
     */
    private executeTypedJob;
    /**
     * Виконання запиту до Google Sheets
     */
    private executeSheetsQuery;
    /**
     * Виконання AI запиту
     */
    private executeAIRequest;
    /**
     * Виконання файлової операції
     */
    private executeFileOperation;
    /**
     * Виконання експорту даних
     */
    private executeExportData;
    /**
     * Оновлення середнього часу обробки
     */
    private updateAverageProcessingTime;
    /**
     * Зміна пріоритету завдання
     */
    changeJobPriority(jobId: string, newPriority: 'high' | 'normal' | 'low'): boolean;
    /**
     * Отримання завдання за ID
     */
    getJob(jobId: string): (JobData & {
        priority: string;
    }) | null;
    /**
     * Видалення завдання
     */
    removeJob(jobId: string): boolean;
    /**
     * Налаштування максимальної кількості одночасних завдань
     */
    setMaxConcurrent(priority: 'high' | 'normal' | 'low', max: number): void;
    /**
     * Пауза обробки черги
     */
    pauseQueue(priority: 'high' | 'normal' | 'low'): void;
    /**
     * Відновлення обробки черги
     */
    resumeQueue(priority: 'high' | 'normal' | 'low'): void;
    /**
     * Отримання рекомендацій по оптимізації
     */
    getOptimizationRecommendations(): OptimizationRecommendation[];
    /**
     * Скидання статистики
     */
    resetStats(): void;
}
declare const _default: QueueManager;
export default _default;
//# sourceMappingURL=queueManager.d.ts.map