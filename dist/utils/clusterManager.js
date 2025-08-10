"use strict";
/**
 * Cluster Manager для масштабування Discord бота
 * Підтримка кластеризації та load balancing
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ClusterManager = void 0;
const cluster_1 = __importDefault(require("cluster"));
const os_1 = __importDefault(require("os"));
const logger_1 = __importDefault(require("./logger"));
class ClusterManager {
    static formatError(error) {
        if (error instanceof Error)
            return error.message;
        try {
            return JSON.stringify(error);
        }
        catch {
            return String(error);
        }
    }
    static formatAny(value) {
        if (typeof value === 'string')
            return value;
        try {
            return JSON.stringify(value);
        }
        catch {
            return String(value);
        }
    }
    constructor(config = {}) {
        this.config = {
            workers: config.workers || os_1.default.cpus().length,
            restartDelay: config.restartDelay || 5000,
            maxRestarts: config.maxRestarts || 10,
            ...config
        };
        this.workers = new Map();
        this.restartCounts = new Map();
        const isPrimary = cluster_1.default.isPrimary ?? cluster_1.default.isMaster ?? false;
        this.isMaster = isPrimary;
        this.isActive = false;
        this.stats = {
            totalWorkers: 0,
            activeWorkers: 0,
            restarts: 0,
            startTime: Date.now(),
            uptime: 0,
        };
    }
    /**
     * Запуск кластера
     */
    async start() {
        if (!this.isMaster) {
            logger_1.default.info('🔧 Worker процес запущено');
            return;
        }
        try {
            logger_1.default.info(`🚀 Запуск кластера з ${this.config.workers} workers...`);
            // Створення workers
            for (let i = 0; i < this.config.workers; i++) {
                await this.createWorker();
            }
            // Налаштування обробників подій
            this.setupEventHandlers();
            this.isActive = true;
            logger_1.default.info(`✅ Кластер запущено: ${this.workers.size} workers`);
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка запуску кластера: ${ClusterManager.formatError(error)}`);
            throw error;
        }
    }
    /**
     * Створення worker процесу
     */
    async createWorker() {
        try {
            const worker = cluster_1.default.fork();
            this.workers.set(worker.id, {
                id: worker.id,
                pid: worker.process.pid ?? 0,
                status: 'starting',
                startTime: Date.now(),
                restarts: 0,
            });
            this.stats.totalWorkers++;
            this.stats.activeWorkers++;
            logger_1.default.info(`🔧 Worker ${worker.id} створено (PID: ${worker.process.pid})`);
            return worker;
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка створення worker: ${ClusterManager.formatError(error)}`);
            throw error;
        }
    }
    /**
     * Налаштування обробників подій
     */
    setupEventHandlers() {
        // Worker online
        cluster_1.default.on('online', (worker) => {
            const workerInfo = this.workers.get(worker.id);
            if (workerInfo) {
                workerInfo.status = 'online';
                logger_1.default.info(`✅ Worker ${worker.id} онлайн`);
            }
        });
        // Worker message
        cluster_1.default.on('message', (worker, message) => {
            this.handleWorkerMessage(worker, message);
        });
        // Worker exit
        cluster_1.default.on('exit', (worker, code, signal) => {
            this.handleWorkerExit(worker, code, signal);
        });
        // Worker disconnect
        cluster_1.default.on('disconnect', (worker) => {
            const workerInfo = this.workers.get(worker.id);
            if (workerInfo) {
                workerInfo.status = 'offline';
                logger_1.default.warn(`⚠️ Worker ${worker.id} відключено`);
            }
        });
    }
    /**
     * Обробка повідомлень від workers
     */
    handleWorkerMessage(worker, message) {
        try {
            const workerInfo = this.workers.get(worker.id);
            if (!workerInfo)
                return;
            switch (message.type) {
                case 'stats':
                    this.updateWorkerStats(worker.id, message.data);
                    break;
                case 'error':
                    logger_1.default.error(`❌ Worker ${worker.id} помилка: ${ClusterManager.formatAny(message.data)}`);
                    break;
                case 'ready':
                    logger_1.default.info(`✅ Worker ${worker.id} готовий`);
                    break;
                default:
                    logger_1.default.debug(`📨 Повідомлення від worker ${worker.id}: ${ClusterManager.formatAny(message)}`);
            }
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка обробки повідомлення worker: ${ClusterManager.formatError(error)}`);
        }
    }
    /**
     * Обробка виходу worker
     */
    async handleWorkerExit(worker, code, signal) {
        try {
            const workerInfo = this.workers.get(worker.id);
            if (!workerInfo)
                return;
            this.stats.activeWorkers--;
            this.workers.delete(worker.id);
            logger_1.default.warn(`⚠️ Worker ${worker.id} завершився (код: ${code}, сигнал: ${signal})`);
            // Перезапуск worker якщо потрібно
            if (this.shouldRestartWorker(worker.id)) {
                await this.restartWorker(worker.id);
            }
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка обробки виходу worker: ${ClusterManager.formatError(error)}`);
        }
    }
    /**
     * Перевірка чи потрібно перезапустити worker
     */
    shouldRestartWorker(workerId) {
        const restartCount = this.restartCounts.get(workerId) || 0;
        return restartCount < this.config.maxRestarts && this.isActive;
    }
    /**
     * Перезапуск worker
     */
    async restartWorker(workerId) {
        try {
            const restartCount = this.restartCounts.get(workerId) || 0;
            this.restartCounts.set(workerId, restartCount + 1);
            this.stats.restarts++;
            logger_1.default.info(`🔄 Перезапуск worker ${workerId} (спроба ${restartCount + 1}/${this.config.maxRestarts})`);
            // Затримка перед перезапуском
            await new Promise(resolve => setTimeout(resolve, this.config.restartDelay));
            const newWorker = await this.createWorker();
            const workerInfo = this.workers.get(newWorker.id);
            if (workerInfo) {
                workerInfo.restarts = restartCount + 1;
            }
            logger_1.default.info(`✅ Worker ${workerId} перезапущено як ${newWorker.id}`);
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка перезапуску worker ${workerId}: ${ClusterManager.formatError(error)}`);
        }
    }
    /**
     * Оновлення статистики worker
     */
    updateWorkerStats(workerId, stats) {
        const workerInfo = this.workers.get(workerId);
        if (workerInfo) {
            workerInfo.stats = stats;
        }
    }
    /**
     * Отримання оптимального worker для розподілу навантаження
     */
    getOptimalWorker() {
        if (this.workers.size === 0)
            return null;
        let optimalWorker = null;
        let minLoad = Infinity;
        for (const [workerId, workerInfo] of this.workers.entries()) {
            if (workerInfo.status === 'online') {
                const load = workerInfo.stats?.load || 0;
                if (load < minLoad) {
                    minLoad = load;
                    optimalWorker = workerId;
                }
            }
        }
        return optimalWorker;
    }
    /**
     * Відправка повідомлення конкретному worker
     */
    sendToWorker(workerId, message) {
        // cluster.workers має string keys; знаходимо по id без порушення індекс-підписів
        const workersRec = cluster_1.default.workers;
        const workerList = Object.values(workersRec ?? {});
        const worker = workerList.find(w => w && w.id === workerId);
        if (!worker) {
            logger_1.default.warn(`⚠️ Worker ${workerId} не знайдено`);
            return false;
        }
        try {
            worker.send(message);
            return true;
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка відправки повідомлення worker ${workerId}: ${ClusterManager.formatError(error)}`);
            return false;
        }
    }
    /**
     * Розсилка повідомлення всім workers
     */
    broadcast(message) {
        try {
            for (const [workerId] of this.workers.entries()) {
                this.sendToWorker(workerId, message);
            }
            logger_1.default.debug(`📢 Розіслано повідомлення ${this.workers.size} workers`);
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка розсилки повідомлень: ${ClusterManager.formatError(error)}`);
        }
    }
    /**
     * Отримання статистики кластера
     */
    getClusterStats() {
        this.stats.uptime = Date.now() - this.stats.startTime;
        return { ...this.stats };
    }
    /**
     * Зупинка кластера
     */
    async stop() {
        if (!this.isMaster)
            return;
        try {
            logger_1.default.info('🛑 Зупинка кластера...');
            this.isActive = false;
            // Зупинка всіх workers
            for (const [workerId] of this.workers.entries()) {
                const workersRec = cluster_1.default.workers;
                const worker = Object.values(workersRec ?? {}).find(w => w && w.id === workerId);
                if (worker)
                    worker.kill();
            }
            // Очищення
            this.workers.clear();
            this.restartCounts.clear();
            logger_1.default.info('✅ Кластер зупинено');
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка зупинки кластера: ${ClusterManager.formatError(error)}`);
            throw error;
        }
    }
    /**
     * Перезапуск кластера
     */
    async restart() {
        try {
            logger_1.default.info('🔄 Перезапуск кластера...');
            await this.stop();
            await this.start();
            logger_1.default.info('✅ Кластер перезапущено');
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка перезапуску кластера: ${ClusterManager.formatError(error)}`);
            throw error;
        }
    }
    /**
     * Перевірка чи кластер активний
     */
    isClusterActive() {
        return this.isActive;
    }
    /**
     * Отримання кількості активних workers
     */
    getActiveWorkersCount() {
        return this.stats.activeWorkers;
    }
}
exports.ClusterManager = ClusterManager;
exports.default = ClusterManager;
//# sourceMappingURL=clusterManager.js.map