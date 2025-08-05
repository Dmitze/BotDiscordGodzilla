/**
 * Cluster Manager для масштабування Discord бота
 * Підтримка кластеризації та load balancing
 * TypeScript версія
 */

import cluster from 'cluster';
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

  constructor(config: Partial<ClusterConfig> = {}) {
    this.config = {
      workers: config.workers || os.cpus().length,
      restartDelay: config.restartDelay || 5000,
      maxRestarts: config.maxRestarts || 10,
      ...config
    };
    
    this.workers = new Map();
    this.restartCounts = new Map();
    this.isMaster = cluster.isPrimary;
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
      logger.error('❌ Помилка запуску кластера:', error);
      throw error;
    }
  }

  /**
   * Створення worker процесу
   */
  private async createWorker(): Promise<cluster.Worker> {
    try {
      const worker = cluster.fork();
      
      this.workers.set(worker.id, {
        id: worker.id,
        pid: worker.process.pid,
        status: 'starting',
        startTime: Date.now(),
        restarts: 0,
      });

      this.stats.totalWorkers++;
      this.stats.activeWorkers++;

      logger.info(`🔧 Worker ${worker.id} створено (PID: ${worker.process.pid})`);
      return worker;
    } catch (error) {
      logger.error('❌ Помилка створення worker:', error);
      throw error;
    }
  }

  /**
   * Налаштування обробників подій
   */
  private setupEventHandlers(): void {
    // Worker online
    cluster.on('online', (worker) => {
      const workerInfo = this.workers.get(worker.id);
      if (workerInfo) {
        workerInfo.status = 'online';
        logger.info(`✅ Worker ${worker.id} онлайн`);
      }
    });

    // Worker message
    cluster.on('message', (worker, message) => {
      this.handleWorkerMessage(worker, message);
    });

    // Worker exit
    cluster.on('exit', (worker, code, signal) => {
      this.handleWorkerExit(worker, code, signal);
    });

    // Worker disconnect
    cluster.on('disconnect', (worker) => {
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
  private handleWorkerMessage(worker: cluster.Worker, message: any): void {
    try {
      const workerInfo = this.workers.get(worker.id);
      if (!workerInfo) return;

      switch (message.type) {
        case 'stats':
          this.updateWorkerStats(worker.id, message.data);
          break;
        case 'error':
          logger.error(`❌ Worker ${worker.id} помилка:`, message.data);
          break;
        case 'ready':
          logger.info(`✅ Worker ${worker.id} готовий`);
          break;
        default:
          logger.debug(`📨 Повідомлення від worker ${worker.id}:`, message);
      }
    } catch (error) {
      logger.error('❌ Помилка обробки повідомлення worker:', error);
    }
  }

  /**
   * Обробка виходу worker
   */
  private async handleWorkerExit(worker: cluster.Worker, code: number, signal: string): Promise<void> {
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
      logger.error('❌ Помилка обробки виходу worker:', error);
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

      logger.info(`🔄 Перезапуск worker ${workerId} (спроба ${restartCount + 1}/${this.config.maxRestarts})`);

      // Затримка перед перезапуском
      await new Promise(resolve => setTimeout(resolve, this.config.restartDelay));

      const newWorker = await this.createWorker();
      const workerInfo = this.workers.get(newWorker.id);
      if (workerInfo) {
        workerInfo.restarts = restartCount + 1;
      }

      logger.info(`✅ Worker ${workerId} перезапущено як ${newWorker.id}`);
    } catch (error) {
      logger.error(`❌ Помилка перезапуску worker ${workerId}:`, error);
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
    const worker = cluster.workers?.[workerId];
    if (!worker) {
      logger.warn(`⚠️ Worker ${workerId} не знайдено`);
      return false;
    }

    try {
      worker.send(message);
      return true;
    } catch (error) {
      logger.error(`❌ Помилка відправки повідомлення worker ${workerId}:`, error);
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
      logger.error('❌ Помилка розсилки повідомлень:', error);
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
        const worker = cluster.workers?.[workerId];
        if (worker) {
          worker.kill();
        }
      }

      // Очищення
      this.workers.clear();
      this.restartCounts.clear();

      logger.info('✅ Кластер зупинено');
    } catch (error) {
      logger.error('❌ Помилка зупинки кластера:', error);
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
      logger.error('❌ Помилка перезапуску кластера:', error);
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