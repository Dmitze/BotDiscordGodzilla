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