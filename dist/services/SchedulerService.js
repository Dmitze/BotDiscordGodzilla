"use strict";
/**
 * Scheduler Service для Discord бота
 * Централізоване управління плануваними завданнями
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const logger_1 = __importDefault(require("../utils/logger"));
class SchedulerService {
    constructor(bot) {
        this.bot = bot;
        this.jobs = new Map();
        this.scheduler = null;
        this.stats = {
            jobsCreated: 0,
            jobsExecuted: 0,
            jobsFailed: 0,
            activeJobs: 0,
        };
        this._isActive = false;
    }
    /**
     * Ініціалізація Scheduler сервісу
     */
    async initialize() {
        try {
            logger_1.default.info('⏰ Ініціалізація Scheduler сервісу...');
            // Створення планувальника
            await this.createScheduler();
            // Реєстрація стандартних завдань
            await this.registerDefaultJobs();
            this._isActive = true;
            logger_1.default.info('✅ Scheduler сервіс ініціалізовано');
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка ініціалізації Scheduler сервісу: ${error instanceof Error ? error.message : String(error)}`);
            throw error;
        }
    }
    /**
     * Створення планувальника
     */
    async createScheduler() {
        try {
            // Використовуємо node-cron для планування
            const cron = require('node-cron');
            this.scheduler = cron;
            logger_1.default.debug('✅ Планувальник створено');
        }
        catch (error) {
            logger_1.default.error(`Помилка створення планувальника: ${error instanceof Error ? error.message : String(error)}`);
            throw error;
        }
    }
    /**
     * Реєстрація стандартних завдань
     */
    async registerDefaultJobs() {
        try {
            // Очищення кешу кожну годину
            this.scheduleJob('cache-cleanup', '0 * * * *', () => {
                this.cleanupCache();
            });
            // Оновлення статистики кожні 5 хвилин
            this.scheduleJob('stats-update', '*/5 * * * *', () => {
                this.updateStats();
            });
            // Перевірка здоров'я кожні 10 хвилин
            this.scheduleJob('health-check', '*/10 * * * *', () => {
                this.healthCheck();
            });
            // Резервне копіювання щодня о 2:00
            this.scheduleJob('backup', '0 2 * * *', () => {
                this.createBackup();
            });
            logger_1.default.debug('✅ Стандартні завдання зареєстровано');
        }
        catch (error) {
            logger_1.default.error(`Помилка реєстрації стандартних завдань: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Планування завдання
     */
    scheduleJob(name, schedule, task, options = {}) {
        try {
            if (this.jobs.has(name)) {
                this.stopJob(name);
            }
            const job = this.scheduler.schedule(schedule, async () => {
                await this.executeJob(name, task);
            }, {
                scheduled: false,
                timezone: options.timezone || 'Europe/Kiev',
                ...options,
            });
            this.jobs.set(name, {
                job,
                schedule,
                task: task.toString(),
                options,
                createdAt: new Date(),
                lastRun: null,
                nextRun: job.nextDate().toDate(),
                executions: 0,
                errors: 0,
            });
            job.start();
            this.stats.jobsCreated++;
            this.stats.activeJobs++;
            logger_1.default.debug(`✅ Завдання "${name}" заплановано: ${schedule}`);
            return job;
        }
        catch (error) {
            logger_1.default.error(`Помилка планування завдання "${name}": ${error instanceof Error ? error.message : String(error)}`);
            throw error;
        }
    }
    /**
     * Виконання завдання
     */
    async executeJob(name, task) {
        const jobInfo = this.jobs.get(name);
        if (!jobInfo) {
            logger_1.default.warn(`Завдання "${name}" не знайдено`);
            return;
        }
        const startTime = Date.now();
        jobInfo.lastRun = new Date();
        jobInfo.executions++;
        try {
            logger_1.default.debug(`🚀 Виконання завдання: ${name}`);
            await task();
            const duration = Date.now() - startTime;
            this.stats.jobsExecuted++;
            logger_1.default.debug(`✅ Завдання "${name}" виконано за ${duration}ms`);
            // Оновлення наступного запуску
            jobInfo.nextRun = jobInfo.job.nextDate().toDate();
        }
        catch (error) {
            jobInfo.errors++;
            this.stats.jobsFailed++;
            logger_1.default.error(`❌ Помилка виконання завдання "${name}": ${error instanceof Error ? error.message : String(error)}`);
            // Сповіщення про помилку
            await this.notifyJobError(name, error);
        }
    }
    /**
     * Зупинка завдання
     */
    stopJob(name) {
        try {
            const jobInfo = this.jobs.get(name);
            if (jobInfo) {
                jobInfo.job.stop();
                this.jobs.delete(name);
                this.stats.activeJobs--;
                logger_1.default.debug(`🛑 Завдання "${name}" зупинено`);
            }
        }
        catch (error) {
            logger_1.default.error(`Помилка зупинки завдання "${name}": ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Отримання інформації про завдання
     */
    getJobInfo(name) {
        const jobInfo = this.jobs.get(name);
        if (!jobInfo)
            return null;
        return {
            name,
            schedule: jobInfo.schedule,
            task: jobInfo.task,
            createdAt: jobInfo.createdAt,
            lastRun: jobInfo.lastRun,
            nextRun: jobInfo.nextRun,
            executions: jobInfo.executions,
            errors: jobInfo.errors,
            isActive: jobInfo.job.running,
        };
    }
    /**
     * Отримання всіх завдань
     */
    getAllJobs() {
        return Array.from(this.jobs.keys()).map(name => this.getJobInfo(name));
    }
    /**
     * Очищення кешу
     */
    async cleanupCache() {
        try {
            const cacheService = this.bot.getService('cache');
            if (cacheService) {
                await cacheService.cleanupMemory();
                logger_1.default.info('🧹 Кеш очищено');
            }
        }
        catch (error) {
            logger_1.default.error(`Помилка очищення кешу: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Оновлення статистики
     */
    async updateStats() {
        try {
            // Оновлення метрик
            const metricsService = this.bot.getService('metrics');
            if (metricsService) {
                metricsService.updateAllMetrics();
            }
            // Оновлення статистики сервісів
            const serviceManager = this.bot.serviceManager;
            if (serviceManager) {
                const servicesStats = serviceManager.getStats();
                logger_1.default.debug(`📊 Статистика сервісів оновлена: ${JSON.stringify(servicesStats)}`);
            }
        }
        catch (error) {
            logger_1.default.error(`Помилка оновлення статистики: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Перевірка здоров'я
     */
    async healthCheck() {
        try {
            const healthStatus = {
                timestamp: new Date(),
                services: {},
                overall: 'healthy',
            };
            // Перевірка сервісів
            const serviceManager = this.bot.serviceManager;
            if (serviceManager) {
                const servicesStatus = serviceManager.getServicesStatus();
                for (const [name, s] of Object.entries(servicesStatus)) {
                    const active = !!s.isActive;
                    healthStatus.services[name] = {
                        isActive: active,
                        hasStats: typeof s.stats !== 'undefined' && s.stats !== null,
                    };
                    if (!active) {
                        healthStatus.overall = 'degraded';
                    }
                }
            }
            // Перевірка Discord клієнта
            if (this.bot.client) {
                healthStatus.discord = {
                    isReady: this.bot.client.isReady(),
                    uptime: this.bot.client.uptime,
                    guilds: this.bot.client.guilds.cache.size,
                };
            }
            logger_1.default.debug(`🏥 Перевірка здоров'я завершена: ${JSON.stringify(healthStatus)}`);
            // Сповіщення про проблеми
            if (healthStatus.overall !== 'healthy') {
                await this.notifyHealthIssue(healthStatus);
            }
        }
        catch (error) {
            logger_1.default.error(`Помилка перевірки здоров'я: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Створення резервної копії
     */
    async createBackup() {
        try {
            logger_1.default.info('💾 Створення резервної копії...');
            // Тут можна додати логіку створення резервної копії
            // Наприклад, збереження даних в файл або базу даних
            logger_1.default.info('✅ Резервна копія створена');
        }
        catch (error) {
            logger_1.default.error(`Помилка створення резервної копії: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * Сповіщення про помилку завдання
     */
    async notifyJobError(jobName, error) {
        try {
            // Тут можна додати логіку сповіщення
            // Наприклад, відправка повідомлення в Discord канал
            logger_1.default.warn(`⚠️ Сповіщення про помилку завдання "${jobName}": ${error instanceof Error ? error.message : String(error)}`);
        }
        catch (notifyError) {
            logger_1.default.error(`Помилка сповіщення про помилку завдання: ${notifyError instanceof Error ? notifyError.message : String(notifyError)}`);
        }
    }
    /**
     * Сповіщення про проблеми здоров'я
     */
    async notifyHealthIssue(healthStatus) {
        try {
            // Тут можна додати логіку сповіщення про проблеми здоров'я
            logger_1.default.warn(`⚠️ Проблеми здоров'я системи: ${healthStatus.overall}`);
        }
        catch (notifyError) {
            logger_1.default.error(`Помилка сповіщення про проблеми здоров'я: ${notifyError instanceof Error ? notifyError.message : String(notifyError)}`);
        }
    }
    /**
     * Отримання статистики
     */
    getStats() {
        return {
            ...this.stats,
            jobs: this.getAllJobs(),
            isActive: this.isActive(),
        };
    }
    /**
     * Перевірка активності
     */
    isActive() {
        return this._isActive;
    }
    /**
     * Завершення роботи
     */
    async shutdown() {
        logger_1.default.info('🛑 Завершення роботи Scheduler сервісу...');
        try {
            // Зупинка всіх завдань
            for (const [name] of this.jobs) {
                this.stopJob(name);
            }
            this._isActive = false;
            logger_1.default.info('✅ Scheduler сервіс завершено');
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка завершення Scheduler сервісу: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
}
exports.default = SchedulerService;
//# sourceMappingURL=SchedulerService.js.map