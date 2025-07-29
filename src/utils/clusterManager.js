/**
 * Cluster Manager для масштабування Discord бота
 * Підтримка кластеризації та load balancing
 */

const cluster = require('cluster');
const os = require('os');
const logger = require('./logger');

class ClusterManager {
  constructor(config = {}) {
    this.config = {
      workers: config.workers || os.cpus().length,
      restartDelay: config.restartDelay || 5000,
      maxRestarts: config.maxRestarts || 10,
      ...config
    };
    
    this.workers = new Map();
    this.restartCounts = new Map();
    this.isMaster = cluster.isMaster;
    this.isActive = false;
    
    this.stats = {
      totalWorkers: 0,
      activeWorkers: 0,
      restarts: 0,
      startTime: Date.now(),
    };
  }

  /**
   * Запуск кластера
   */
  async start() {
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
  async createWorker() {
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
  setupEventHandlers() {
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
        workerInfo.status = 'disconnected';
        logger.warn(`⚠️ Worker ${worker.id} відключено`);
      }
    });
  }

  /**
   * Обробка повідомлень від workers
   */
  handleWorkerMessage(worker, message) {
    try {
      switch (message.type) {
        case 'stats':
          this.updateWorkerStats(worker.id, message.data);
          break;
        case 'error':
          logger.error(`❌ Worker ${worker.id} помилка:`, message.error);
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
  async handleWorkerExit(worker, code, signal) {
    const workerInfo = this.workers.get(worker.id);
    if (!workerInfo) return;

    workerInfo.status = 'exited';
    this.stats.activeWorkers--;
    this.stats.restarts++;

    logger.warn(`⚠️ Worker ${worker.id} завершився (код: ${code}, сигнал: ${signal})`);

    // Перевірка чи потрібно перезапустити
    if (this.shouldRestartWorker(worker.id)) {
      await this.restartWorker(worker.id);
    } else {
      logger.error(`❌ Worker ${worker.id} досяг максимальної кількості перезапусків`);
      this.workers.delete(worker.id);
    }
  }

  /**
   * Перевірка чи потрібно перезапустити worker
   */
  shouldRestartWorker(workerId) {
    const restartCount = this.restartCounts.get(workerId) || 0;
    return restartCount < this.config.maxRestarts;
  }

  /**
   * Перезапуск worker
   */
  async restartWorker(workerId) {
    try {
      const currentCount = this.restartCounts.get(workerId) || 0;
      this.restartCounts.set(workerId, currentCount + 1);

      logger.info(`🔄 Перезапуск worker ${workerId} (спроба ${currentCount + 1}/${this.config.maxRestarts})`);

      // Затримка перед перезапуском
      await new Promise(resolve => setTimeout(resolve, this.config.restartDelay));

      // Створення нового worker
      await this.createWorker();

      // Видалення старого worker з мапи
      this.workers.delete(workerId);
      this.restartCounts.delete(workerId);

    } catch (error) {
      logger.error(`❌ Помилка перезапуску worker ${workerId}:`, error);
    }
  }

  /**
   * Оновлення статистики worker
   */
  updateWorkerStats(workerId, stats) {
    const workerInfo = this.workers.get(workerId);
    if (workerInfo) {
      workerInfo.stats = stats;
      workerInfo.lastUpdate = Date.now();
    }
  }

  /**
   * Розподіл навантаження між workers
   */
  getOptimalWorker() {
    let optimalWorker = null;
    let minLoad = Infinity;

    for (const [workerId, workerInfo] of this.workers.entries()) {
      if (workerInfo.status === 'online' && workerInfo.stats) {
        const load = workerInfo.stats.load || 0;
        if (load < minLoad) {
          minLoad = load;
          optimalWorker = workerId;
        }
      }
    }

    return optimalWorker;
  }

  /**
   * Відправка повідомлення до worker
   */
  sendToWorker(workerId, message) {
    const worker = cluster.workers[workerId];
    if (worker && worker.isConnected()) {
      worker.send(message);
      return true;
    }
    return false;
  }

  /**
   * Відправка повідомлення до всіх workers
   */
  broadcast(message) {
    let sentCount = 0;
    
    for (const [workerId] of this.workers.entries()) {
      if (this.sendToWorker(workerId, message)) {
        sentCount++;
      }
    }

    logger.debug(`📢 Повідомлення відправлено до ${sentCount} workers`);
    return sentCount;
  }

  /**
   * Отримання статистики кластера
   */
  getClusterStats() {
    return {
      ...this.stats,
      workers: Array.from(this.workers.values()),
      restartCounts: Object.fromEntries(this.restartCounts),
      uptime: Date.now() - this.stats.startTime,
      isActive: this.isActive,
    };
  }

  /**
   * Зупинка кластера
   */
  async stop() {
    if (!this.isMaster) return;

    try {
      logger.info('🛑 Зупинка кластера...');

      // Відправка сигналу зупинки до всіх workers
      this.broadcast({ type: 'shutdown' });

      // Очікування завершення workers
      for (const [workerId, workerInfo] of this.workers.entries()) {
        const worker = cluster.workers[workerId];
        if (worker && worker.isConnected()) {
          worker.disconnect();
          
          // Примусове завершення через 10 секунд
          setTimeout(() => {
            if (worker.isConnected()) {
              worker.kill();
            }
          }, 10000);
        }
      }

      this.isActive = false;
      logger.info('✅ Кластер зупинено');
    } catch (error) {
      logger.error('❌ Помилка зупинки кластера:', error);
      throw error;
    }
  }

  /**
   * Перезапуск кластера
   */
  async restart() {
    logger.info('🔄 Перезапуск кластера...');
    await this.stop();
    await new Promise(resolve => setTimeout(resolve, 2000));
    await this.start();
  }
}

module.exports = ClusterManager; 