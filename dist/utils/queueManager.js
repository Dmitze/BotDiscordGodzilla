"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
const events_1 = require("events");
class QueueManager extends events_1.EventEmitter {
    constructor() {
        super();
        this.queues = {
            high: [],
            normal: [],
            low: [],
        };
        this.processing = {
            high: false,
            normal: false,
            low: false,
        };
        this.stats = {
            processed: 0,
            failed: 0,
            pending: 0,
            averageProcessingTime: 0,
        };
        this.maxConcurrent = {
            high: 3,
            normal: 5,
            low: 2,
        };
        this.activeJobs = {
            high: 0,
            normal: 0,
            low: 0,
        };
        this.startProcessing();
    }
    /**
     * Додавання завдання в чергу
     */
    addJob(priority, job) {
        const jobId = this.generateJobId();
        const jobData = {
            id: jobId,
            priority,
            job,
            timestamp: Date.now(),
            retries: 0,
            maxRetries: 3,
        };
        this.queues[priority].push(jobData);
        this.stats.pending++;
        this.emit('jobAdded', { jobId, priority });
        console.log(`📋 Додано завдання ${jobId} в чергу ${priority}`);
        return jobId;
    }
    /**
     * Генерація унікального ID завдання
     */
    generateJobId() {
        return `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }
    /**
     * Запуск обробки черг
     */
    startProcessing() {
        setInterval(() => {
            this.processQueue('high');
            this.processQueue('normal');
            this.processQueue('low');
        }, 1000);
    }
    /**
     * Обробка черги
     */
    async processQueue(priority) {
        if (this.activeJobs[priority] >= this.maxConcurrent[priority]) {
            return;
        }
        if (this.queues[priority].length === 0) {
            return;
        }
        const jobData = this.queues[priority].shift();
        this.stats.pending--;
        this.activeJobs[priority]++;
        console.log(`⚡ Початок обробки завдання ${jobData.id} (${priority})`);
        try {
            const startTime = Date.now();
            const result = await this.executeJob(jobData);
            const processingTime = Date.now() - startTime;
            this.stats.processed++;
            this.updateAverageProcessingTime(processingTime);
            this.emit('jobCompleted', {
                jobId: jobData.id,
                priority,
                result,
                processingTime,
            });
            console.log(`✅ Завдання ${jobData.id} завершено за ${processingTime}мс`);
        }
        catch (error) {
            this.stats.failed++;
            if (jobData.retries < jobData.maxRetries) {
                jobData.retries++;
                this.queues[priority].unshift(jobData);
                this.stats.pending++;
                console.log(`🔄 Повторна спроба завдання ${jobData.id} (${jobData.retries}/${jobData.maxRetries})`);
            }
            else {
                this.emit('jobFailed', {
                    jobId: jobData.id,
                    priority,
                    error: error.message,
                    retries: jobData.retries,
                });
                console.log(`❌ Завдання ${jobData.id} не вдалося виконати після ${jobData.retries} спроб`);
            }
        }
        finally {
            this.activeJobs[priority]--;
        }
    }
    /**
     * Отримання статистики черг
     */
    getQueueStats() {
        return {
            queues: {
                high: {
                    pending: this.queues.high.length,
                    active: this.activeJobs.high,
                    maxConcurrent: this.maxConcurrent.high,
                },
                normal: {
                    pending: this.queues.normal.length,
                    active: this.activeJobs.normal,
                    maxConcurrent: this.maxConcurrent.normal,
                },
                low: {
                    pending: this.queues.low.length,
                    active: this.activeJobs.low,
                    maxConcurrent: this.maxConcurrent.low,
                },
            },
            stats: {
                ...this.stats,
                averageProcessingTime: Math.round(this.stats.averageProcessingTime),
            },
            totalPending: this.queues.high.length + this.queues.normal.length + this.queues.low.length,
            totalActive: this.activeJobs.high + this.activeJobs.normal + this.activeJobs.low,
        };
    }
    /**
     * Виконання завдання
     */
    async executeJob(jobData) {
        const { job } = jobData;
        if (typeof job === 'function') {
            return await job();
        }
        else if (job.type && job.handler) {
            return await this.executeTypedJob(job);
        }
        else {
            throw new Error('Невідомий тип завдання');
        }
    }
    /**
     * Виконання типізованого завдання
     */
    async executeTypedJob(job) {
        switch (job.type) {
            case 'sheets_query':
                return await this.executeSheetsQuery(job);
            case 'ai_request':
                return await this.executeAIRequest(job);
            case 'file_operation':
                return await this.executeFileOperation(job);
            case 'export_data':
                return await this.executeExportData(job);
            default:
                throw new Error(`Невідомий тип завдання: ${job.type}`);
        }
    }
    /**
     * Виконання запиту до Google Sheets
     */
    async executeSheetsQuery(job) {
        const { query, range, options: _options = {} } = job.data;
        // Імітація запиту до Google Sheets
        await new Promise(resolve => setTimeout(resolve, 1000 + Math.random() * 2000));
        return {
            type: 'sheets_query',
            data: `Результат запиту: ${query}`,
            range,
            timestamp: Date.now(),
        };
    }
    /**
     * Виконання AI запиту
     */
    async executeAIRequest(job) {
        const { prompt, context, options: _options = {} } = job.data;
        // Імітація AI запиту
        await new Promise(resolve => setTimeout(resolve, 2000 + Math.random() * 3000));
        return {
            type: 'ai_request',
            response: `AI відповідь на: ${prompt}`,
            context,
            timestamp: Date.now(),
        };
    }
    /**
     * Виконання файлової операції
     */
    async executeFileOperation(job) {
        const { operation, filePath, options: _options = {} } = job.data;
        // Імітація файлової операції
        await new Promise(resolve => setTimeout(resolve, 500 + Math.random() * 1000));
        return {
            type: 'file_operation',
            operation,
            filePath,
            result: `Операція ${operation} виконана`,
            timestamp: Date.now(),
        };
    }
    /**
     * Виконання експорту даних
     */
    async executeExportData(job) {
        const { data: _data, format, options: _options = {} } = job.data;
        // Імітація експорту
        await new Promise(resolve => setTimeout(resolve, 1500 + Math.random() * 2500));
        return {
            type: 'export_data',
            format,
            filePath: `/tmp/export_${Date.now()}.${format}`,
            timestamp: Date.now(),
        };
    }
    /**
     * Оновлення середнього часу обробки
     */
    updateAverageProcessingTime(newTime) {
        const { processed, averageProcessingTime } = this.stats;
        if (processed === 1) {
            this.stats.averageProcessingTime = newTime;
        }
        else {
            this.stats.averageProcessingTime =
                (averageProcessingTime * (processed - 1) + newTime) / processed;
        }
    }
    /**
     * Зміна пріоритету завдання
     */
    changeJobPriority(jobId, newPriority) {
        for (const priority of ['high', 'normal', 'low']) {
            const jobIndex = this.queues[priority].findIndex(job => job.id === jobId);
            if (jobIndex !== -1) {
                const [moved] = this.queues[priority].splice(jobIndex, 1);
                if (!moved) {
                    return false;
                }
                moved.priority = newPriority;
                this.queues[newPriority].push(moved);
                console.log(`🔄 Змінено пріоритет завдання ${jobId} з ${priority} на ${newPriority}`);
                return true;
            }
        }
        return false;
    }
    /**
     * Отримання завдання за ID
     */
    getJob(jobId) {
        for (const priority of ['high', 'normal', 'low']) {
            const job = this.queues[priority].find(job => job.id === jobId);
            if (job) {
                return { ...job, priority };
            }
        }
        return null;
    }
    /**
     * Видалення завдання
     */
    removeJob(jobId) {
        for (const priority of ['high', 'normal', 'low']) {
            const jobIndex = this.queues[priority].findIndex(job => job.id === jobId);
            if (jobIndex !== -1) {
                this.queues[priority].splice(jobIndex, 1);
                this.stats.pending--;
                console.log(`🗑️ Видалено завдання ${jobId} з черги ${priority}`);
                return true;
            }
        }
        return false;
    }
    /**
     * Налаштування максимальної кількості одночасних завдань
     */
    setMaxConcurrent(priority, max) {
        this.maxConcurrent[priority] = max;
        console.log(`⚙️ Встановлено максимум ${max} одночасних завдань для черги ${priority}`);
    }
    /**
     * Пауза обробки черги
     */
    pauseQueue(priority) {
        this.processing[priority] = false;
        console.log(`⏸️ Пауза обробки черги ${priority}`);
    }
    /**
     * Відновлення обробки черги
     */
    resumeQueue(priority) {
        this.processing[priority] = true;
        console.log(`▶️ Відновлено обробку черги ${priority}`);
    }
    /**
     * Отримання рекомендацій по оптимізації
     */
    getOptimizationRecommendations() {
        const stats = this.getQueueStats();
        const recommendations = [];
        // Рекомендації по чергах
        if (stats.queues.high.pending > 10) {
            recommendations.push({
                type: 'queue',
                priority: 'high',
                message: 'Високий пріоритет: багато завдань в черзі (>10)',
                action: 'Збільшити кількість одночасних завдань або додати більше ресурсів',
            });
        }
        if (stats.queues.normal.pending > 20) {
            recommendations.push({
                type: 'queue',
                priority: 'medium',
                message: 'Звичайний пріоритет: багато завдань в черзі (>20)',
                action: 'Оптимізувати обробку або збільшити пріоритет важливих завдань',
            });
        }
        // Рекомендації по часу обробки
        if (stats.stats.averageProcessingTime > 5000) {
            recommendations.push({
                type: 'performance',
                priority: 'high',
                message: 'Середній час обробки занадто високий (>5с)',
                action: 'Оптимізувати завдання або додати кешування',
            });
        }
        // Рекомендації по помилках
        const errorRate = (stats.stats.failed / (stats.stats.processed + stats.stats.failed)) * 100;
        if (errorRate > 10) {
            recommendations.push({
                type: 'reliability',
                priority: 'high',
                message: `Високий відсоток помилок (${errorRate.toFixed(1)}%)`,
                action: 'Перевірити логіку завдань та додати обробку помилок',
            });
        }
        return recommendations;
    }
    /**
     * Скидання статистики
     */
    resetStats() {
        this.stats = {
            processed: 0,
            failed: 0,
            pending: 0,
            averageProcessingTime: 0,
        };
        console.log('🔄 Статистика черг скинута');
    }
}
// Експорт синглтона
exports.default = new QueueManager();
//# sourceMappingURL=queueManager.js.map