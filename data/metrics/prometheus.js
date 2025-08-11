const prometheus = require('prom-client');
const express = require('express');
const os = require('os');

// Lightweight local logger wrapper to avoid direct console usage
const log = {
  info: (msg, meta) => console.log(msg, meta ? JSON.stringify(meta) : ''),
  warn: (msg, meta) => console.warn(msg, meta ? JSON.stringify(meta) : ''),
  error: (msg, meta) => console.error(msg, meta ? JSON.stringify(meta) : ''),
};

// Read config from env to avoid TS import in JS
const METRICS_ENABLED = process.env.METRICS_ENABLED !== 'false';
const METRICS_PORT = parseInt(process.env.METRICS_PORT || '9091', 10);
const METRICS_PATH = process.env.METRICS_PATH || '/metrics';

class MetricsCollector {
  constructor() {
    this.registry = new prometheus.Registry();
    this.initializeMetrics();
  }

  /**
   * Ініціалізація метрик
   */
  initializeMetrics() {
    this.initializeSystemGauges();
    // Метрики команд
    this.commandCounter = new prometheus.Counter({
      name: 'discord_bot_commands_total',
      help: 'Загальна кількість виконаних команд',
      labelNames: ['command', 'status', 'user_id'],
      registers: [this.registry]
    });

    this.commandDuration = new prometheus.Histogram({
      name: 'discord_bot_command_duration_seconds',
      help: 'Час виконання команд',
      labelNames: ['command'],
      buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
      registers: [this.registry]
    });

    // Метрики API запитів
    this.apiRequestCounter = new prometheus.Counter({
      name: 'discord_bot_api_requests_total',
      help: 'Загальна кількість API запитів',
      labelNames: ['service', 'method', 'status'],
      registers: [this.registry]
    });

    this.apiRequestDuration = new prometheus.Histogram({
      name: 'discord_bot_api_request_duration_seconds',
      help: 'Час виконання API запитів',
      labelNames: ['service'],
      buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
      registers: [this.registry]
    });

    // Метрики кешу
    this.cacheHits = new prometheus.Counter({
      name: 'discord_bot_cache_hits_total',
      help: 'Кількість попадань в кеш',
      labelNames: ['cache_type'],
      registers: [this.registry]
    });

    this.cacheMisses = new prometheus.Counter({
      name: 'discord_bot_cache_misses_total',
      help: 'Кількість промахів кешу',
      labelNames: ['cache_type'],
      registers: [this.registry]
    });

    // Метрики помилок
    this.errorCounter = new prometheus.Counter({
      name: 'discord_bot_errors_total',
      help: 'Загальна кількість помилок',
      labelNames: ['type', 'command'],
      registers: [this.registry]
    });

    // Метрики активності
    this.activeUsers = new prometheus.Gauge({
      name: 'discord_bot_active_users',
      help: 'Кількість активних користувачів',
      registers: [this.registry]
    });

    this.activeGuilds = new prometheus.Gauge({
      name: 'discord_bot_active_guilds',
      help: 'Кількість активних серверів',
      registers: [this.registry]
    });

    // Метрики пам'яті
    this.memoryUsage = new prometheus.Gauge({
      name: 'discord_bot_memory_usage_bytes',
      help: 'Використання пам\'яті',
      labelNames: ['type'],
      registers: [this.registry]
    });

    // Метрики часу роботи
    this.uptime = new prometheus.Gauge({
      name: 'discord_bot_uptime_seconds',
      help: 'Час роботи бота в секундах',
      registers: [this.registry]
    });

    // Метрики AI-запитів
    this.aiRequestCounter = new prometheus.Counter({
      name: 'discord_bot_ai_requests_total',
      help: 'Кількість AI-запитів',
      labelNames: ['model', 'status'],
      registers: [this.registry]
    });

    this.aiRequestDuration = new prometheus.Histogram({
      name: 'discord_bot_ai_request_duration_seconds',
      help: 'Час виконання AI-запитів',
      labelNames: ['model'],
      buckets: [1, 2, 5, 10, 30, 60],
      registers: [this.registry]
    });

    // Метрики пошуку
    this.searchCounter = new prometheus.Counter({
      name: 'discord_bot_searches_total',
      help: 'Кількість пошукових запитів',
      labelNames: ['type', 'results_count'],
      registers: [this.registry]
    });

    // Метрики експорту
    this.exportCounter = new prometheus.Counter({
      name: 'discord_bot_exports_total',
      help: 'Кількість експортів',
      labelNames: ['format', 'size_range'],
      registers: [this.registry]
    });
  }

  initializeSystemGauges() {
    this.cpuUsage = new prometheus.Gauge({
      name: 'discord_bot_cpu_usage_microseconds',
      help: 'CPU usage user/system in microseconds',
      labelNames: ['type'],
      registers: [this.registry]
    });
    this.diskSpace = new prometheus.Gauge({
      name: 'discord_bot_disk_space_bytes',
      help: 'Disk space info',
      labelNames: ['type'],
      registers: [this.registry]
    });
  }

  /**
   * Оновлення CPU та пам'яті (для тестів і ручного виклику)
   */
  updateCpuAndDisk() {
    try {
      const { user, system } = process.cpuUsage();
      this.cpuUsage.set({ type: 'user' }, user);
      this.cpuUsage.set({ type: 'system' }, system);
      const free = os.freemem();
      const total = os.totalmem();
      this.diskSpace.set({ type: 'mem_free' }, free);
      this.diskSpace.set({ type: 'mem_total' }, total);
    } catch (error) {
      log.warn('⚠️ Помилка updateCpuAndDisk', { type: 'metrics', event: 'update_cpu_disk_failed', error: String(error) });
    }
  }

  /**
   * Запис метрики команди
   */
  recordCommand(command, status, userId) {
    try { this.commandCounter.inc({ command, status, user_id: userId }); } catch {}
  }

  /**
   * Запис часу виконання команди
   */
  recordCommandDuration(command, duration) {
    try { this.commandDuration.observe({ command }, duration); } catch {}
  }

  /**
   * Запис API запиту
   */
  recordApiRequest(service, method, status) {
    try { this.apiRequestCounter.inc({ service, method, status }); } catch {}
  }

  /**
   * Запис часу API запиту
   */
  recordApiRequestDuration(service, duration) {
    try { this.apiRequestDuration.observe({ service }, duration); } catch {}
  }

  /**
   * Запис попадання в кеш
   */
  recordCacheHit(cacheType) {
    try { this.cacheHits.inc({ cache_type: cacheType }); } catch {}
  }

  /**
   * Запис промаху кешу
   */
  recordCacheMiss(cacheType) {
    try { this.cacheMisses.inc({ cache_type: cacheType }); } catch {}
  }

  /**
   * Запис помилки
   */
  recordError(type, command) {
    try { this.errorCounter.inc({ type, command }); } catch {}
  }

  /**
   * Оновлення кількості активних користувачів
   */
  setActiveUsers(count) {
    this.activeUsers.set(count);
  }

  /**
   * Оновлення кількості активних серверів
   */
  setActiveGuilds(count) {
    this.activeGuilds.set(count);
  }

  /**
   * Оновлення використання пам'яті
   */
  updateMemoryUsage() {
    const memUsage = process.memoryUsage();
    this.memoryUsage.set({ type: 'heap_used' }, memUsage.heapUsed);
    this.memoryUsage.set({ type: 'heap_total' }, memUsage.heapTotal);
    this.memoryUsage.set({ type: 'external' }, memUsage.external);
    this.memoryUsage.set({ type: 'rss' }, memUsage.rss);
  }

  /**
   * Оновлення часу роботи
   */
  updateUptime() {
    this.uptime.set(process.uptime());
  }

  /**
   * Запис AI-запиту
   */
  recordAiRequest(model, status) {
    this.aiRequestCounter.inc({ model, status });
  }

  /**
   * Запис часу AI-запиту
   */
  recordAiRequestDuration(model, duration) {
    this.aiRequestDuration.observe({ model }, duration);
  }

  /**
   * Запис пошукового запиту
   */
  recordSearch(type, resultsCount) {
    const sizeRange = this.getSizeRange(resultsCount);
    this.searchCounter.inc({ type, results_count: sizeRange });
  }

  /**
   * Запис експорту
   */
  recordExport(format, fileSize) {
    const sizeRange = this.getSizeRange(fileSize);
    this.exportCounter.inc({ format, size_range: sizeRange });
  }

  /**
   * Отримання діапазону розміру
   */
  getSizeRange(size) {
    if (size < 10) return '0-10';
    if (size < 50) return '10-50';
    if (size < 100) return '50-100';
    if (size < 500) return '100-500';
    if (size < 1000) return '500-1000';
    return '1000+';
  }

  /**
   * Отримання метрик у форматі Prometheus
   */
  async getMetrics() {
    return await this.registry.metrics();
  }

  /**
   * Запуск HTTP сервера для метрик
   */
  startMetricsServer() {
    if (!METRICS_ENABLED) {
      log.info('📊 Метрики вимкнені (ENV)', { type: 'metrics', event: 'disabled_env' });
      return;
    }

    const app = express();
    const port = METRICS_PORT;
    const path = METRICS_PATH;

    // простейший rate-limit
    const windowMs = 10_000; // 10s
    const max = 5; // запросов
    const hits = new Map();

    // Периодичне оновлення системних метрик (поза хендлером)
    if (!this._systemInterval) {
      this._systemInterval = setInterval(() => {
        try {
          const { user, system } = process.cpuUsage();
          this.cpuUsage.set({ type: 'user' }, user);
          this.cpuUsage.set({ type: 'system' }, system);
          const free = os.freemem();
          const total = os.totalmem();
          this.diskSpace.set({ type: 'mem_free' }, free);
          this.diskSpace.set({ type: 'mem_total' }, total);
        } catch (error) {
          log.warn('⚠️ Помилка оновлення системних метрик', { type: 'metrics', event: 'system_metrics_update_failed', error: String(error) });
        }
      }, 5000);
    }

    // Ендпоінт для метрик
    app.get(path, async (req, res) => {
      const ip = req.ip;
      const now = Date.now();

      const arr = hits.get(ip) || [];
      const pruned = arr.filter((t) => now - t < windowMs);
      pruned.push(now);
      hits.set(ip, pruned);
      if (pruned.length > max) {
        return res.status(429).end('Rate limited');
      }
      try {
        res.set('Content-Type', this.registry.contentType);
        res.end(await this.registry.metrics());
      } catch (error) {
        log.error('Помилка отримання метрик', { type: 'metrics', event: 'render_failed', error: String(error) });
        return res.status(500).end('Помилка отримання метрик');
      }
    });

    // Ендпоінт для перевірки здоров'я
    app.get('/health', (req, res) => {
      res.json({
        status: 'ok',
        uptime: process.uptime(),
        timestamp: new Date().toISOString()
      });
    });

    // Запуск сервера
    const server = app.listen(port, () => {
      log.info(`📊 Метрики доступні на http://localhost:${port}${path}`, { type: 'metrics', event: 'server_started', port, path });
      log.info(`🏥 Health check доступний на http://localhost:${port}/health`, { type: 'metrics', event: 'health_endpoint_ready', port });
    });

    return server;
  }

  /**
   * Оновлення всіх метрик
   */
  updateAllMetrics() {
    this.updateMemoryUsage();
    this.updateUptime();
  }

  /**
   * Створення звіту метрик
   */
  async generateMetricsReport() {
    const metrics = await this.registry.getMetricsAsJSON();

    const report = {
      timestamp: new Date().toISOString(),
      uptime: process.uptime(),
      memory: process.memoryUsage(),
      metrics: {}
    };

    for (const metric of metrics) {
      report.metrics[metric.name] = {
        help: metric.help,
        type: metric.type,
        values: metric.values
      };
    }

    return report;
  }
}

module.exports = MetricsCollector;