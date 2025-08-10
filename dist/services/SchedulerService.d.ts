/**
 * Scheduler Service для Discord бота
 * Централізоване управління плануваними завданнями
 * TypeScript версія
 */
import { Client } from 'discord.js';
interface Bot {
    getService(name: string): any;
    serviceManager?: any;
    client?: Client;
}
interface JobDetails {
    name: string;
    schedule: string;
    task: string;
    createdAt: Date;
    lastRun: Date | null;
    nextRun: Date;
    executions: number;
    errors: number;
    isActive: boolean;
}
interface SchedulerStats {
    jobsCreated: number;
    jobsExecuted: number;
    jobsFailed: number;
    activeJobs: number;
    jobs: JobDetails[];
    isActive: boolean;
}
declare class SchedulerService {
    private bot;
    private jobs;
    private scheduler;
    private stats;
    private _isActive;
    constructor(bot: Bot);
    /**
     * Ініціалізація Scheduler сервісу
     */
    initialize(): Promise<void>;
    /**
     * Створення планувальника
     */
    private createScheduler;
    /**
     * Реєстрація стандартних завдань
     */
    private registerDefaultJobs;
    /**
     * Планування завдання
     */
    scheduleJob(name: string, schedule: string, task: () => Promise<void> | void, options?: any): any;
    /**
     * Виконання завдання
     */
    private executeJob;
    /**
     * Зупинка завдання
     */
    stopJob(name: string): void;
    /**
     * Отримання інформації про завдання
     */
    getJobInfo(name: string): JobDetails | null;
    /**
     * Отримання всіх завдань
     */
    getAllJobs(): JobDetails[];
    /**
     * Очищення кешу
     */
    private cleanupCache;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Перевірка здоров'я
     */
    private healthCheck;
    /**
     * Створення резервної копії
     */
    private createBackup;
    /**
     * Сповіщення про помилку завдання
     */
    private notifyJobError;
    /**
     * Сповіщення про проблеми здоров'я
     */
    private notifyHealthIssue;
    /**
     * Отримання статистики
     */
    getStats(): SchedulerStats;
    /**
     * Перевірка активності
     */
    isActive(): boolean;
    /**
     * Завершення роботи
     */
    shutdown(): Promise<void>;
}
export default SchedulerService;
//# sourceMappingURL=SchedulerService.d.ts.map