/**
 * Cluster Manager для масштабування Discord бота
 * Підтримка кластеризації та load balancing
 * TypeScript версія
 */

import cluster from 'cluster';
import type { Worker } from 'cluster';
import os from 'os';
import logger from './logger';

interface ClusterConfig {
  workers: number;
  restartDelay: number;
  maxRestarts: number;
  [key: string]: any;
}

interface WorkerInfo {
  id: number;
  pid: number;
  status: 'starting' | 'online' | 'offline' | 'restarting';
  startTime: number;
  restarts: number;
  stats?: any;
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

class ClusterManager {
  private config: ClusterConfig;
  private workers: Map<number, WorkerInfo>;
  private restartCounts: Map<number, number>;
  private isMaster: boolean;
  private isActive: boolean;
  private stats: ClusterStats;

  private static formatError(error: unknown): string {
    if (error instanceof Error) return error.message;
    try {
      return JSON.stringify(error);
    } catch {
      return String(error);
    }
  }
  private static formatAny(value: unknown): string {
    if (typeof value === 'string') return value;
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  constructor(config: Partial<ClusterConfig> = {}) {
    this.config = {
      workers: config.workers || os.cpus().length,
      restartDelay: config.restartDelay || 5000,
      maxRestarts: config.maxRestarts || 10,
      ...config,
    };

    this.workers = new Map();
    this.restartCounts = new Map();
    const isPrimary: boolean = (cluster as any).isPrimary ?? (cluster as any).isMaster ?? false;
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
  async start(): Promise<void> {
    if (!this.isMaster) {
      logger.info('🔧 Worker процес запущено');
      return;
    }

    try {
      logger.info(`🚀 Запуск кластера з ${this.config.workers} workers...`);

      // Створення workers
      for (let i = 0; i < this.config.workers; i++) {
        await this.createWorker();
      }

      // Налаштування обробників подій
      this.setupEventHandlers();

      this.isActive = true;
      logger.info(`✅ Кластер запущено: ${this.workers.size} workers`);
    } catch (error) {
      logger.error(`❌ Помилка запуску кластера: ${ClusterManager.formatError(error)}`);
      throw error;
    }
  }

  /**
   * Створення worker процесу
   */
  private async createWorker(): Promise<Worker> {
    try {
      const worker: Worker = (cluster as any).fork();

      this.workers.set(worker.id, {
        id: worker.id,
        pid: worker.process.pid ?? 0,
        status: 'starting',
        startTime: Date.now(),
        restarts: 0,
      });

      this.stats.totalWorkers++;
      this.stats.activeWorkers++;

      logger.info(`🔧 Worker ${worker.id} створено (PID: ${worker.process.pid})`);
      return worker;
    } catch (error) {
      logger.error(`❌ Помилка створення worker: ${ClusterManager.formatError(error)}`);
      throw error;
    }
  }

  /**
   * Налаштування обробників подій
   */
  private setupEventHandlers(): void {
    // Worker online
    (cluster as any).on('online', (worker: Worker) => {
      const workerInfo = this.workers.get(worker.id);
      if (workerInfo) {
        workerInfo.status = 'online';
        logger.info(`✅ Worker ${worker.id} онлайн`);
      }
    });

    // Worker message
    (cluster as any).on('message', (worker: Worker, message: any) => {
      this.handleWorkerMessage(worker, message);
    });

    // Worker exit
    (cluster as any).on('exit', (worker: Worker, code: number, signal: string) => {
      this.handleWorkerExit(worker, code, signal);
    });

    // Worker disconnect
    (cluster as any).on('disconnect', (worker: Worker) => {
      const workerInfo = this.workers.get(worker.id);
      if (workerInfo) {
        workerInfo.status = 'offline';
        logger.warn(`⚠️ Worker ${worker.id} відключено`);
      }
    });
  }

  /**
   * Обробка повідомлень від workers
   */
  private handleWorkerMessage(worker: Worker, message: any): void {
    try {
      const workerInfo = this.workers.get(worker.id);
      if (!workerInfo) return;

      switch (message.type) {
        case 'stats':
          this.updateWorkerStats(worker.id, message.data);
          break;
        case 'error':
          logger.error(`❌ Worker ${worker.id} помилка: ${ClusterManager.formatAny(message.data)}`);
          break;
        case 'ready':
          logger.info(`✅ Worker ${worker.id} готовий`);
          break;
        default:
          logger.debug(
            `📨 Повідомлення від worker ${worker.id}: ${ClusterManager.formatAny(message)}`
          );
      }
    } catch (error) {
      logger.error(`❌ Помилка обробки повідомлення worker: ${ClusterManager.formatError(error)}`);
    }
  }

  /**
   * Обробка виходу worker
   */
  private async handleWorkerExit(worker: Worker, code: number, signal: string): Promise<void> {
    try {
      const workerInfo = this.workers.get(worker.id);
      if (!workerInfo) return;

      this.stats.activeWorkers--;
      this.workers.delete(worker.id);

      logger.warn(`⚠️ Worker ${worker.id} завершився (код: ${code}, сигнал: ${signal})`);

      // Перезапуск worker якщо потрібно
      if (this.shouldRestartWorker(worker.id)) {
        await this.restartWorker(worker.id);
      }
    } catch (error) {
      logger.error(`❌ Помилка обробки виходу worker: ${ClusterManager.formatError(error)}`);
    }
  }

  /**
   * Перевірка чи потрібно перезапустити worker
   */
  private shouldRestartWorker(workerId: number): boolean {
    const restartCount = this.restartCounts.get(workerId) || 0;
    return restartCount < this.config.maxRestarts && this.isActive;
  }

  /**
   * Перезапуск worker
   */
  private async restartWorker(workerId: number): Promise<void> {
    try {
      const restartCount = this.restartCounts.get(workerId) || 0;
      this.restartCounts.set(workerId, restartCount + 1);
      this.stats.restarts++;

      logger.info(
        `🔄 Перезапуск worker ${workerId} (спроба ${restartCount + 1}/${this.config.maxRestarts})`
      );

      // Затримка перед перезапуском
      await new Promise(resolve => setTimeout(resolve, this.config.restartDelay));

      const newWorker = await this.createWorker();
      const workerInfo = this.workers.get(newWorker.id);
      if (workerInfo) {
        workerInfo.restarts = restartCount + 1;
      }

      logger.info(`✅ Worker ${workerId} перезапущено як ${newWorker.id}`);
    } catch (error) {
      logger.error(
        `❌ Помилка перезапуску worker ${workerId}: ${ClusterManager.formatError(error)}`
      );
    }
  }

  /**
   * Оновлення статистики worker
   */
  private updateWorkerStats(workerId: number, stats: any): void {
    const workerInfo = this.workers.get(workerId);
    if (workerInfo) {
      workerInfo.stats = stats;
    }
  }

  /**
   * Отримання оптимального worker для розподілу навантаження
   */
  getOptimalWorker(): number | null {
    if (this.workers.size === 0) return null;

    let optimalWorker: number | null = null;
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
  sendToWorker(workerId: number, message: WorkerMessage): boolean {
    // cluster.workers має string keys; знаходимо по id без порушення індекс-підписів
    const workersRec = (cluster as any).workers as Record<string, Worker | undefined> | undefined;
    const workerList = Object.values(workersRec ?? {});
    const worker = workerList.find(w => w && w.id === workerId);
    if (!worker) {
      logger.warn(`⚠️ Worker ${workerId} не знайдено`);
      return false;
    }

    try {
      worker.send(message);
      return true;
    } catch (error) {
      logger.error(
        `❌ Помилка відправки повідомлення worker ${workerId}: ${ClusterManager.formatError(error)}`
      );
      return false;
    }
  }

  /**
   * Розсилка повідомлення всім workers
   */
  broadcast(message: WorkerMessage): void {
    try {
      for (const [workerId] of this.workers.entries()) {
        this.sendToWorker(workerId, message);
      }
      logger.debug(`📢 Розіслано повідомлення ${this.workers.size} workers`);
    } catch (error) {
      logger.error(`❌ Помилка розсилки повідомлень: ${ClusterManager.formatError(error)}`);
    }
  }

  /**
   * Отримання статистики кластера
   */
  getClusterStats(): ClusterStats {
    this.stats.uptime = Date.now() - this.stats.startTime;
    return { ...this.stats };
  }

  /**
   * Зупинка кластера
   */
  async stop(): Promise<void> {
    if (!this.isMaster) return;

    try {
      logger.info('🛑 Зупинка кластера...');
      this.isActive = false;

      // Зупинка всіх workers
      for (const [workerId] of this.workers.entries()) {
        const workersRec = (cluster as any).workers as
          | Record<string, Worker | undefined>
          | undefined;
        const worker = Object.values(workersRec ?? {}).find(w => w && w.id === workerId);
        if (worker) worker.kill();
      }

      // Очищення
      this.workers.clear();
      this.restartCounts.clear();

      logger.info('✅ Кластер зупинено');
    } catch (error) {
      logger.error(`❌ Помилка зупинки кластера: ${ClusterManager.formatError(error)}`);
      throw error;
    }
  }

  /**
   * Перезапуск кластера
   */
  async restart(): Promise<void> {
    try {
      logger.info('🔄 Перезапуск кластера...');
      await this.stop();
      await this.start();
      logger.info('✅ Кластер перезапущено');
    } catch (error) {
      logger.error(`❌ Помилка перезапуску кластера: ${ClusterManager.formatError(error)}`);
      throw error;
    }
  }

  /**
   * Перевірка чи кластер активний
   */
  isClusterActive(): boolean {
    return this.isActive;
  }

  /**
   * Отримання кількості активних workers
   */
  getActiveWorkersCount(): number {
    return this.stats.activeWorkers;
  }
}

export default ClusterManager;
export { ClusterManager };
