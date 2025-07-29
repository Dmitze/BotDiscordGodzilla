/**
 * Metrics Service для Discord бота
 * Централізоване управління метриками та моніторингом
 */

const logger = require('../utils/logger');

class MetricsService {
  constructor(bot) {
    this.bot = bot;
    this.config = bot.config.metrics;
    this.registry = null;
    this.metrics = {};
    this.server = null;
    this.stats = {
      requests: 0,
      errors: 0,
      startTime: Date.now(),
    };
    this.isActive = false;
  }

  /**
   * Ініціалізація Metrics сервісу
   */
  async initialize() {
    try {
      logger.info('📊 Ініціалізація Metrics сервісу...');

      // Створення Prometheus реєстру
      await this.createRegistry();

      // Створення метрик
      this.createMetrics();

      // Запуск HTTP сервера
      await this.startServer();

      this.isActive = true;
      logger.info('✅ Metrics сервіс ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Metrics сервісу:', error);
      throw error;
    }
  }

  /**
   * Створення Prometheus реєстру
   */
  async createRegistry() {
    try {
      const { Registry, collectDefaultMetrics } = require('prom-client');

      this.registry = new Registry();

      // Збір стандартних метрик Node.js
      collectDefaultMetrics({ register: this.registry });

      logger.debug('✅ Prometheus реєстр створено');
    } catch (error) {
      logger.error('Помилка створення Prometheus реєстру:', error);
      throw error;
    }
  }

  /**
   * Створення метрик
   */
  createMetrics() {
    try {
      const { Counter, Gauge, Histogram } = require('prom-client');

      // Лічильники
      this.metrics.commandsTotal = new Counter({
        name: 'discord_bot_commands_total',
        help: 'Загальна кількість виконаних команд',
        labelNames: ['command', 'status'],
      });

      this.metrics.messagesTotal = new Counter({
        name: 'discord_bot_messages_total',
        help: 'Загальна кількість повідомлень',
        labelNames: ['type'],
      });

      this.metrics.errorsTotal = new Counter({
        name: 'discord_bot_errors_total',
        help: 'Загальна кількість помилок',
        labelNames: ['type', 'service'],
      });

      // Гейджи
      this.metrics.activeUsers = new Gauge({
        name: 'discord_bot_active_users',
        help: 'Кількість активних користувачів',
      });

      this.metrics.activeGuilds = new Gauge({
        name: 'discord_bot_active_guilds',
        help: 'Кількість активних серверів',
      });

      // Метрики продуктивності
      this.metrics.cacheHitRate = new Gauge({
        name: 'discord_bot_cache_hit_rate',
        help: 'Відсоток попадань в кеш',
      });

      this.metrics.cacheSize = new Gauge({
        name: 'discord_bot_cache_size',
        help: 'Розмір кешу в байтах',
      });

      this.metrics.queueLength = new Gauge({
        name: 'discord_bot_queue_length',
        help: 'Довжина черги завдань',
        labelNames: ['priority'],
      });

      this.metrics.connectionPoolUsage = new Gauge({
        name: 'discord_bot_connection_pool_usage',
        help: 'Використання connection pool',
        labelNames: ['service'],
      });

      // Метрики AI
      this.metrics.aiRequestsTotal = new Counter({
        name: 'discord_bot_ai_requests_total',
        help: 'Загальна кількість AI запитів',
        labelNames: ['provider', 'status'],
      });

      this.metrics.aiResponseTime = new Histogram({
        name: 'discord_bot_ai_response_time_seconds',
        help: 'Час відповіді AI в секундах',
        labelNames: ['provider'],
        buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
      });

      // Метрики Google API
      this.metrics.googleApiRequestsTotal = new Counter({
        name: 'discord_bot_google_api_requests_total',
        help: 'Загальна кількість запитів до Google API',
        labelNames: ['service', 'endpoint', 'status'],
      });

      this.metrics.googleApiResponseTime = new Histogram({
        name: 'discord_bot_google_api_response_time_seconds',
        help: 'Час відповіді Google API в секундах',
        labelNames: ['service'],
        buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
      });

      this.metrics.memoryUsage = new Gauge({
        name: 'discord_bot_memory_usage_bytes',
        help: "Використання пам'яті в байтах",
      });

      this.metrics.uptime = new Gauge({
        name: 'discord_bot_uptime_seconds',
        help: 'Час роботи бота в секундах',
      });

      // Гістограми
      this.metrics.commandDuration = new Histogram({
        name: 'discord_bot_command_duration_seconds',
        help: 'Тривалість виконання команд',
        labelNames: ['command'],
        buckets: [0.1, 0.5, 1, 2, 5, 10],
      });

      this.metrics.apiResponseTime = new Histogram({
        name: 'discord_bot_api_response_time_seconds',
        help: 'Час відповіді API',
        labelNames: ['service', 'endpoint'],
        buckets: [0.1, 0.5, 1, 2, 5, 10],
      });

      // Реєстрація метрик
      Object.values(this.metrics).forEach(metric => {
        this.registry.registerMetric(metric);
      });

      logger.debug('✅ Метрики створено');
    } catch (error) {
      logger.error('Помилка створення метрик:', error);
      throw error;
    }
  }

  /**
   * Запуск HTTP сервера
   */
  async startServer() {
    try {
      const express = require('express');
      const app = express();
      const port = this.config.port || 9090;

      // Middleware для логування
      app.use((req, res, next) => {
        this.stats.requests++;
        logger.debug(`${req.method} ${req.path}`);
        next();
      });

      // Endpoint для метрик
      app.get('/metrics', async (req, res) => {
        try {
          const metrics = await this.registry.metrics();
          res.set('Content-Type', this.registry.contentType);
          res.end(metrics);
        } catch (error) {
          this.stats.errors++;
          logger.error('Помилка отримання метрик:', error);
          res.status(500).send('Помилка отримання метрик');
        }
      });

      // Endpoint для здоров'я
      app.get('/health', (req, res) => {
        res.json({
          status: 'ok',
          uptime: process.uptime(),
          timestamp: Date.now(),
        });
      });

      // Endpoint для статистики
      app.get('/stats', (req, res) => {
        res.json(this.getStats());
      });

      // Обробка помилок
      app.use((error, req, res, next) => {
        this.stats.errors++;
        logger.error('HTTP помилка:', error);
        res.status(500).send('Внутрішня помилка сервера');
      });

      this.server = app.listen(port, () => {
        logger.info(`📊 Metrics сервер запущено на порту ${port}`);
      });
    } catch (error) {
      logger.error('Помилка запуску Metrics сервера:', error);
      throw error;
    }
  }

  /**
   * Інкремент лічильника команд
   */
  incrementCommand(command, status = 'success') {
    try {
      this.metrics.commandsTotal.inc({ command, status });
    } catch (error) {
      logger.error('Помилка інкременту лічильника команд:', error);
    }
  }

  /**
   * Інкремент лічильника повідомлень
   */
  incrementMessage(type) {
    try {
      this.metrics.messagesTotal.inc({ type });
    } catch (error) {
      logger.error('Помилка інкременту лічильника повідомлень:', error);
    }
  }

  /**
   * Інкремент лічильника помилок
   */
  incrementError(type, service = 'unknown') {
    try {
      this.metrics.errorsTotal.inc({ type, service });
    } catch (error) {
      logger.error('Помилка інкременту лічильника помилок:', error);
    }
  }

  /**
   * Оновлення кількості активних користувачів
   */
  setActiveUsers(count) {
    try {
      this.metrics.activeUsers.set(count);
    } catch (error) {
      logger.error('Помилка оновлення активних користувачів:', error);
    }
  }

  /**
   * Оновлення кількості активних серверів
   */
  setActiveGuilds(count) {
    try {
      this.metrics.activeGuilds.set(count);
    } catch (error) {
      logger.error('Помилка оновлення активних серверів:', error);
    }
  }

  /**
   * Оновлення використання пам'яті
   */
  updateMemoryUsage() {
    try {
      const usage = process.memoryUsage();
      this.metrics.memoryUsage.set(usage.heapUsed);
    } catch (error) {
      logger.error("Помилка оновлення використання пам'яті:", error);
    }
  }

  /**
   * Оновлення часу роботи
   */
  updateUptime() {
    try {
      this.metrics.uptime.set(process.uptime());
    } catch (error) {
      logger.error('Помилка оновлення часу роботи:', error);
    }
  }

  /**
   * Вимірювання тривалості команди
   */
  measureCommandDuration(command, duration) {
    try {
      this.metrics.commandDuration.observe({ command }, duration);
    } catch (error) {
      logger.error('Помилка вимірювання тривалості команди:', error);
    }
  }

  /**
   * Вимірювання часу відповіді API
   */
  measureApiResponseTime(service, endpoint, duration) {
    try {
      this.metrics.apiResponseTime.observe({ service, endpoint }, duration);
    } catch (error) {
      logger.error('Помилка вимірювання часу відповіді API:', error);
    }
  }

  /**
   * Оновлення метрик кешу
   */
  updateCacheMetrics(cacheStats) {
    try {
      if (this.metrics.cacheHitRate && cacheStats) {
        const hitRate = cacheStats.hits / (cacheStats.hits + cacheStats.misses) * 100;
        this.metrics.cacheHitRate.set(hitRate);
      }
    } catch (error) {
      logger.error('Помилка оновлення метрик кешу:', error);
    }
  }

  /**
   * Оновлення метрик черги
   */
  updateQueueMetrics(queueStats) {
    try {
      if (this.metrics.queueLength && queueStats) {
        Object.keys(queueStats).forEach(priority => {
          this.metrics.queueLength.set({ priority }, queueStats[priority].length || 0);
        });
      }
    } catch (error) {
      logger.error('Помилка оновлення метрик черги:', error);
    }
  }

  /**
   * Оновлення метрик connection pool
   */
  updateConnectionPoolMetrics(connectionStats) {
    try {
      if (this.metrics.connectionPoolUsage && connectionStats) {
        Object.keys(connectionStats).forEach(service => {
          const usage = connectionStats[service].inUse ? 1 : 0;
          this.metrics.connectionPoolUsage.set({ service }, usage);
        });
      }
    } catch (error) {
      logger.error('Помилка оновлення метрик connection pool:', error);
    }
  }

  /**
   * Оновлення метрик AI
   */
  updateAIMetrics(provider, status, duration) {
    try {
      if (this.metrics.aiRequestsTotal) {
        this.metrics.aiRequestsTotal.inc({ provider, status });
      }
      if (this.metrics.aiResponseTime && duration) {
        this.metrics.aiResponseTime.observe({ provider }, duration / 1000);
      }
    } catch (error) {
      logger.error('Помилка оновлення метрик AI:', error);
    }
  }

  /**
   * Оновлення метрик Google API
   */
  updateGoogleApiMetrics(service, endpoint, status, duration) {
    try {
      if (this.metrics.googleApiRequestsTotal) {
        this.metrics.googleApiRequestsTotal.inc({ service, endpoint, status });
      }
      if (this.metrics.googleApiResponseTime && duration) {
        this.metrics.googleApiResponseTime.observe({ service }, duration / 1000);
      }
    } catch (error) {
      logger.error('Помилка оновлення метрик Google API:', error);
    }
  }

  /**
   * Оновлення всіх метрик
   */
  updateAllMetrics() {
    this.updateMemoryUsage();
    this.updateUptime();

    if (this.bot.client) {
      this.setActiveUsers(this.bot.client.users.cache.size);
      this.setActiveGuilds(this.bot.client.guilds.cache.size);
    }

    // Оновлення метрик сервісів
    if (this.bot.serviceContainer) {
      // Метрики кешу
      const cacheService = this.bot.serviceContainer.get('cache');
      if (cacheService) {
        this.updateCacheMetrics(cacheService.getCacheStats());
      }

      // Метрики черги
      const queueManager = this.bot.queueManager;
      if (queueManager) {
        this.updateQueueMetrics(queueManager.getQueueStats());
      }

      // Метрики Google API
      const googleService = this.bot.serviceContainer.get('google');
      if (googleService) {
        const connectionStats = googleService.getConnectionStats();
        this.updateConnectionPoolMetrics(connectionStats);
      }
    }
  }

  /**
   * Запуск періодичного оновлення метрик
   */
  startPeriodicUpdates() {
    setInterval(() => {
      this.updateAllMetrics();
    }, 30000); // Кожні 30 секунд
  }

  /**
   * Отримання статистики
   */
  getStats() {
    return {
      ...this.stats,
      uptime: process.uptime(),
      memoryUsage: process.memoryUsage(),
      isActive: this.isActive,
      metricsCount: Object.keys(this.metrics).length,
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
    logger.info('🛑 Завершення роботи Metrics сервісу...');

    try {
      if (this.server) {
        this.server.close();
      }

      this.isActive = false;
      logger.info('✅ Metrics сервіс завершено');
    } catch (error) {
      logger.error('❌ Помилка завершення Metrics сервісу:', error);
    }
  }
}

module.exports = MetricsService;
