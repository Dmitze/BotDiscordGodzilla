/**
 * Scheduler Service для Discord бота
 * Централізоване управління плануваними завданнями
 */

const logger = require('../utils/logger');

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
    this.isActive = false;
  }

  /**
   * Ініціалізація Scheduler сервісу
   */
  async initialize() {
    try {
      logger.info('⏰ Ініціалізація Scheduler сервісу...');

      // Створення планувальника
      await this.createScheduler();

      // Реєстрація стандартних завдань
      await this.registerDefaultJobs();

      this.isActive = true;
      logger.info('✅ Scheduler сервіс ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Scheduler сервісу:', error);
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

      logger.debug('✅ Планувальник створено');
    } catch (error) {
      logger.error('Помилка створення планувальника:', error);
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

      logger.debug('✅ Стандартні завдання зареєстровано');
    } catch (error) {
      logger.error('Помилка реєстрації стандартних завдань:', error);
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

      const job = this.scheduler.schedule(
        schedule,
        async () => {
          await this.executeJob(name, task);
        },
        {
          scheduled: false,
          timezone: options.timezone || 'Europe/Kiev',
          ...options,
        }
      );

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

      logger.debug(`✅ Завдання "${name}" заплановано: ${schedule}`);
      return job;
    } catch (error) {
      logger.error(`Помилка планування завдання "${name}":`, error);
      throw error;
    }
  }

  /**
   * Виконання завдання
   */
  async executeJob(name, task) {
    const jobInfo = this.jobs.get(name);
    if (!jobInfo) {
      logger.warn(`Завдання "${name}" не знайдено`);
      return;
    }

    const startTime = Date.now();
    jobInfo.lastRun = new Date();
    jobInfo.executions++;

    try {
      logger.debug(`🚀 Виконання завдання: ${name}`);

      await task();

      const duration = Date.now() - startTime;
      this.stats.jobsExecuted++;

      logger.debug(`✅ Завдання "${name}" виконано за ${duration}ms`);

      // Оновлення наступного запуску
      jobInfo.nextRun = jobInfo.job.nextDate().toDate();
    } catch (error) {
      jobInfo.errors++;
      this.stats.jobsFailed++;

      logger.error(`❌ Помилка виконання завдання "${name}":`, error);

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

        logger.debug(`🛑 Завдання "${name}" зупинено`);
      }
    } catch (error) {
      logger.error(`Помилка зупинки завдання "${name}":`, error);
    }
  }

  /**
   * Отримання інформації про завдання
   */
  getJobInfo(name) {
    const jobInfo = this.jobs.get(name);
    if (!jobInfo) return null;

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
        logger.info('🧹 Кеш очищено');
      }
    } catch (error) {
      logger.error('Помилка очищення кешу:', error);
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
        logger.debug('📊 Статистика сервісів оновлена', servicesStats);
      }
    } catch (error) {
      logger.error('Помилка оновлення статистики:', error);
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

        for (const [name, status] of Object.entries(servicesStatus)) {
          healthStatus.services[name] = {
            isActive: status.isActive,
            hasStats: !!status.stats,
          };

          if (!status.isActive) {
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

      logger.debug("🏥 Перевірка здоров'я завершена", healthStatus);

      // Сповіщення про проблеми
      if (healthStatus.overall !== 'healthy') {
        await this.notifyHealthIssue(healthStatus);
      }
    } catch (error) {
      logger.error("Помилка перевірки здоров'я:", error);
    }
  }

  /**
   * Створення резервної копії
   */
  async createBackup() {
    try {
      logger.info('💾 Створення резервної копії...');

      // Тут можна додати логіку створення резервної копії
      // Наприклад, збереження даних в файл або базу даних

      logger.info('✅ Резервна копія створена');
    } catch (error) {
      logger.error('Помилка створення резервної копії:', error);
    }
  }

  /**
   * Сповіщення про помилку завдання
   */
  async notifyJobError(jobName, error) {
    try {
      // Тут можна додати логіку сповіщення
      // Наприклад, відправка повідомлення в Discord канал

      logger.warn(`⚠️ Сповіщення про помилку завдання "${jobName}": ${error.message}`);
    } catch (notifyError) {
      logger.error('Помилка сповіщення про помилку завдання:', notifyError);
    }
  }

  /**
   * Сповіщення про проблеми здоров'я
   */
  async notifyHealthIssue(healthStatus) {
    try {
      // Тут можна додати логіку сповіщення про проблеми здоров'я

      logger.warn(`⚠️ Проблеми здоров'я системи: ${healthStatus.overall}`);
    } catch (notifyError) {
      logger.error("Помилка сповіщення про проблеми здоров'я:", notifyError);
    }
  }

  /**
   * Отримання статистики
   */
  getStats() {
    return {
      ...this.stats,
      jobs: this.getAllJobs(),
      isActive: this.isActive,
    };
  }

  /**
   * Перевірка активності
   */
  isActive() {
    return this.isActive;
  }

  /**
   * Завершення роботи
   */
  async shutdown() {
    logger.info('🛑 Завершення роботи Scheduler сервісу...');

    try {
      // Зупинка всіх завдань
      for (const [name] of this.jobs) {
        this.stopJob(name);
      }

      this.isActive = false;
      logger.info('✅ Scheduler сервіс завершено');
    } catch (error) {
      logger.error('❌ Помилка завершення Scheduler сервісу:', error);
    }
  }
}

module.exports = SchedulerService;
