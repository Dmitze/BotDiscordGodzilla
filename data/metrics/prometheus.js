const prometheus = require('prom-client');
const express = require('express');
const os = require('os');
const { Config } = require('../../src/config/Config');

class MetricsCollector {
  constructor() {
    this.registry = new prometheus.Registry();
    this.initializeMetrics();
  }

  /**
   * Ініціалізація метрик та періодичних оновлень
   */
  initializeMetrics() {
    try {
      // Стандартні колектори процеса/Node.js
      this._stopDefaultMetrics = prometheus.collectDefaultMetrics({ register: this.registry });

      // Кастомні метрики
      this.commandCounter = new prometheus.Counter({
        name: 'bot_commands_total',
        help: 'Total bot commands by status and user',
        labelNames: ['command', 'status', 'user_id'],
        registers: [this.registry],
      });

      this.responseTime = new prometheus.Histogram({
        name: 'api_response_time_ms',
        help: 'API response time in ms',
        labelNames: ['route', 'status'],
        buckets: [50, 100, 200, 500, 1000, 2000, 5000],
        registers: [this.registry],
      });

      this.cpuUsage = new prometheus.Gauge({
        name: 'process_cpu_usage_microseconds_total',
        help: 'CPU usage from process.cpuUsage()',
        labelNames: ['type'],
        registers: [this.registry],
      });

      this.diskSpace = new prometheus.Gauge({
        name: 'system_memory_bytes',
        help: 'System memory stats from os module',
        labelNames: ['type'],
        registers: [this.registry],
      });

      this.memoryRss = new prometheus.Gauge({
        name: 'process_memory_rss_bytes',
        help: 'Resident Set Size memory usage of the process',
        registers: [this.registry],
      });

      this.processUptime = new prometheus.Gauge({
        name: 'process_uptime_seconds',
        help: 'Process uptime in seconds',
        registers: [this.registry],
      });

      // Додаткові метрики для API та кешу
      this.apiRequests = new prometheus.Counter({
        name: 'api_requests_total',
        help: 'API requests total by service/method/status',
        labelNames: ['service', 'method', 'status'],
        registers: [this.registry],
      });
      this.apiRequestTime = new prometheus.Histogram({
        name: 'api_request_time_ms',
        help: 'API request duration in ms',
        labelNames: ['service'],
        buckets: [10, 50, 100, 200, 500, 1000, 2000],
        registers: [this.registry],
      });
      this.cacheHits = new prometheus.Counter({
        name: 'cache_hits_total',
        help: 'Cache hits total by cache name',
        labelNames: ['cache'],
        registers: [this.registry],
      });
      this.cacheMisses = new prometheus.Counter({
        name: 'cache_misses_total',
        help: 'Cache misses total by cache name',
        labelNames: ['cache'],
        registers: [this.registry],
      });
      this.errorCounter = new prometheus.Counter({
        name: 'errors_total',
        help: 'Errors total by type and source',
        labelNames: ['type', 'source'],
        registers: [this.registry],
      });

      // Періодичні оновлення, зняття хендлів через unref
      this._systemInterval = setInterval(() => {
        try {
          this.updateAllMetrics();
        } catch {}
      }, 5000);
      if (this._systemInterval.unref) this._systemInterval.unref();
    } catch (error) {
      console.error('❌ initializeMetrics failed', { error: String(error) });
    }
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
      console.warn('⚠️ Помилка updateCpuAndDisk', { type: 'metrics', event: 'update_cpu_disk_failed', error: String(error) });
    }
  }

  /**
   * Запис метрики команди
   */
  recordCommand(command, status, userId) {
    try { this.commandCounter.inc({ command, status, user_id: userId }); } catch {}
  }

  /**
   * Запуск HTTP сервера для метрик
   */
  startMetricsServer() {
    const metricsConfig = (() => {
      try { return Config.get().metrics; } catch { return { enabled: false, port: 9091, path: '/metrics' }; }
    })();
    if (!metricsConfig.enabled) {
      console.log('📊 Метрики вимкнені в конфігурації');
      return;
    }

    // Idempotent start
    if (this._server && this._server.listening) {
      console.warn('⚠️ Сервер метрик вже запущено', { type: 'metrics', event: 'already_running' });
      return this._server;
    }

    const app = express();
    const port = metricsConfig.port;

    // Ендпоінт для метрик
    app.get(metricsConfig.path, async (req, res) => {
      try {
        res.set('Content-Type', this.registry.contentType);
        res.end(await this.registry.metrics());
      } catch (error) {
        console.error('Помилка отримання метрик:', error);
        res.status(500).end('Помилка отримання метрик');
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
      console.log(`📊 Метрики доступні на http://localhost:${port}${metricsConfig.path}`);
      console.log(`🏥 Health check доступний на http://localhost:${port}/health`);
    });

    // Save server reference for stop
    this._server = server;
    return server;
  }

  /**
   * Зупинка HTTP сервера для метрик і очищення ресурсів
   */
  async stopMetricsServer() {
    try {
      if (this._systemInterval) {
        clearInterval(this._systemInterval);
        this._systemInterval = null;
      }
      // Зупиняємо стандартні метрики Prometheus, щоб прибрати таймери
      if (typeof this._stopDefaultMetrics === 'function') {
        try { this._stopDefaultMetrics(); } catch {}
        this._stopDefaultMetrics = null;
      }
      if (this._server && this._server.listening) {
        await new Promise((resolve, reject) => {
          this._server.close((err) => (err ? reject(err) : resolve()));
        });
        this._server = null;
      }
      if (this._rateLimitHits) {
        this._rateLimitHits.clear();
        this._rateLimitHits = null;
      }
      console.info('🛑 Сервер метрик зупинено', { type: 'metrics', event: 'server_stopped' });
      return true;
    } catch (error) {
      console.error('❌ Помилка зупинки сервера метрик', { type: 'metrics', event: 'server_stop_failed', error: String(error) });
      return false;
    }
  }

  /**
   * Оновлення всіх метрик
   */
  updateAllMetrics() {
    this.updateMemoryUsage();
    this.updateUptime();
    this.updateCpuAndDisk();
  }

  /**
   * Оновлення памʼяті процеса
   */
  updateMemoryUsage() {
    try {
      const mem = process.memoryUsage();
      this.memoryRss.set(mem.rss || 0);
    } catch (error) {
      console.warn('⚠️ updateMemoryUsage failed', String(error));
    }
  }

  /**
   * Оновлення аптайму процеса
   */
  updateUptime() {
    try {
      this.processUptime.set(process.uptime());
    } catch (error) {
      console.warn('⚠️ updateUptime failed', String(error));
    }
  }

  // Публічні зручні методи для тестів та інтеграцій
  recordCommandDuration(command, seconds) {
    try { this.responseTime.observe({ route: `cmd:${command}`, status: 'ok' }, (seconds || 0) * 1000); } catch {}
  }
  recordApiRequest(service, method, status) {
    try { this.apiRequests.inc({ service, method, status }); } catch {}
  }
  recordApiRequestDuration(service, seconds) {
    try { this.apiRequestTime.observe({ service }, (seconds || 0) * 1000); } catch {}
  }
  recordCacheHit(cache) {
    try { this.cacheHits.inc({ cache }); } catch {}
  }
  recordCacheMiss(cache) {
    try { this.cacheMisses.inc({ cache }); } catch {}
  }
  recordError(type, source) {
    try { this.errorCounter.inc({ type, source }); } catch {}
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