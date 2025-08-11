/**
 * Metrics Service для Discord бота
 * Централізоване управління метриками та моніторингом
 */

import { Registry, Counter, Gauge, Histogram, collectDefaultMetrics } from 'prom-client';
import type {
  BotConfig,
  ServiceStats,
  CacheStats,
  QueueStats,
  HealthStatus,
} from '@/types';

import { BaseService as BaseServiceClass } from '@/core/BaseService';
import logger from '@/utils/logger';

// Стандартизований проектный логгер используется вместо console

interface MetricsServiceStats extends ServiceStats {
  requests: number;
  errors: number;
  startTime: number;
  metricsCount: number;
}

interface MetricsCollection {
  commandsTotal: Counter<string>;
  messagesTotal: Counter<string>;
  errorsTotal: Counter<string>;
  activeUsers: Gauge<string>;
  activeGuilds: Gauge<string>;
  memoryUsage: Gauge<string>;
  uptime: Gauge<string>;
  commandDuration: Histogram<string>;
  apiResponseTime: Histogram<string>;
  cacheHitRate: Gauge<string>;
  cacheSize: Gauge<string>;
  queueLength: Gauge<string>;
  connectionPoolUsage: Gauge<string>;
  aiRequestsTotal: Counter<string>;
  aiResponseTime: Histogram<string>;
  googleApiRequestsTotal: Counter<string>;
  googleApiResponseTime: Histogram<string>;
}

export class MetricsService extends BaseServiceClass {
  private registry: Registry | null = null;
  private metrics: MetricsCollection | null = null;
  private server: any = null;
  private stats: MetricsServiceStats;
  private updateInterval: NodeJS.Timeout | null = null;

  constructor(config: BotConfig) {
    super('MetricsService', config);
    this.stats = {
      service: 'MetricsService',
      uptime: 0,
      requests: 0,
      errors: 0,
      startTime: Date.now(),
      metricsCount: 0,
    };
  }

  /**
   * Ініціалізація Metrics сервісу
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('📊 Ініціалізація Metrics сервісу...', { type: 'metrics_service', event: 'init', component: 'MetricsService' });

      // Створення Prometheus реєстру
      await this.createRegistry();

      // Створення метрик
      this.createMetrics();

      // Запуск HTTP сервера
      await this.startServer();

      // Запуск періодичних оновлень
      this.startPeriodicUpdates();

      logger.info('✅ Metrics сервіс ініціалізовано', { type: 'metrics_service', event: 'init_success', component: 'MetricsService' });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Metrics сервісу:', {
        type: 'metrics_service', event: 'init_failed', component: 'MetricsService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Створення Prometheus реєстру
   */
  private async createRegistry(): Promise<void> {
    try {
      this.registry = new Registry();

      // Збір стандартних метрик Node.js
      collectDefaultMetrics({ register: this.registry });

      logger.debug('✅ Prometheus реєстр створено', { type: 'metrics_service', event: 'registry_created', component: 'MetricsService' });
    } catch (error) {
      logger.error('Помилка створення Prometheus реєстру:', {
        type: 'metrics_service', event: 'registry_create_failed', component: 'MetricsService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Створення метрик
   */
  private createMetrics(): void {
    try {
      if (!this.registry) {
        throw new Error('Реєстр не ініціалізовано');
      }

      this.metrics = {
        // Лічильники
        commandsTotal: new Counter({
          name: 'discord_bot_commands_total',
          help: 'Загальна кількість виконаних команд',
          labelNames: ['command', 'status'],
          registers: [this.registry],
        }),

        messagesTotal: new Counter({
          name: 'discord_bot_messages_total',
          help: 'Загальна кількість повідомлень',
          labelNames: ['type'],
          registers: [this.registry],
        }),

        errorsTotal: new Counter({
          name: 'discord_bot_errors_total',
          help: 'Загальна кількість помилок',
          labelNames: ['type', 'service'],
          registers: [this.registry],
        }),

        // Гейджи
        activeUsers: new Gauge({
          name: 'discord_bot_active_users',
          help: 'Кількість активних користувачів',
          registers: [this.registry],
        }),

        activeGuilds: new Gauge({
          name: 'discord_bot_active_guilds',
          help: 'Кількість активних серверів',
          registers: [this.registry],
        }),

        memoryUsage: new Gauge({
          name: 'discord_bot_memory_usage_bytes',
          help: 'Використання пам\'яті в байтах',
          registers: [this.registry],
        }),

        uptime: new Gauge({
          name: 'discord_bot_uptime_seconds',
          help: 'Час роботи бота в секундах',
          registers: [this.registry],
        }),

        // Гістограми
        commandDuration: new Histogram({
          name: 'discord_bot_command_duration_seconds',
          help: 'Тривалість виконання команд',
          labelNames: ['command'],
          buckets: [0.1, 0.5, 1, 2, 5, 10],
          registers: [this.registry],
        }),

        apiResponseTime: new Histogram({
          name: 'discord_bot_api_response_time_seconds',
          help: 'Час відповіді API',
          labelNames: ['service'],
          buckets: [0.1, 0.5, 1, 2, 5, 10],
          registers: [this.registry],
        }),

        // Кеш метрики
        cacheHitRate: new Gauge({
          name: 'discord_bot_cache_hit_rate_percent',
          help: 'Відсоток попадань в кеш',
          registers: [this.registry],
        }),

        cacheSize: new Gauge({
          name: 'discord_bot_cache_size',
          help: 'Розмір кешу',
          registers: [this.registry],
        }),

        // Черги
        queueLength: new Gauge({
          name: 'discord_bot_queue_length',
          help: 'Довжина черги',
          labelNames: ['priority'],
          registers: [this.registry],
        }),

        // Connection Pool
        connectionPoolUsage: new Gauge({
          name: 'discord_bot_connection_pool_usage_percent',
          help: 'Використання connection pool',
          labelNames: ['service'],
          registers: [this.registry],
        }),

        // AI метрики
        aiRequestsTotal: new Counter({
          name: 'discord_bot_ai_requests_total',
          help: 'Загальна кількість AI запитів',
          labelNames: ['provider', 'status'],
          registers: [this.registry],
        }),

        aiResponseTime: new Histogram({
          name: 'discord_bot_ai_response_time_seconds',
          help: 'Час відповіді AI',
          labelNames: ['provider'],
          buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
          registers: [this.registry],
        }),

        // Google API метрики
        googleApiRequestsTotal: new Counter({
          name: 'discord_bot_google_api_requests_total',
          help: 'Загальна кількість Google API запитів',
          labelNames: ['service', 'endpoint', 'status'],
          registers: [this.registry],
        }),

        googleApiResponseTime: new Histogram({
          name: 'discord_bot_google_api_response_time_seconds',
          help: 'Час відповіді Google API',
          labelNames: ['service'],
          buckets: [0.1, 0.5, 1, 2, 5, 10],
          registers: [this.registry],
        }),
      };

      this.stats.metricsCount = Object.keys(this.metrics).length;
      logger.debug('✅ Метрики створено', { type: 'metrics_service', event: 'metrics_created', component: 'MetricsService', metricsCount: this.stats.metricsCount });
    } catch (error) {
      logger.error('Помилка створення метрик:', {
        type: 'metrics_service', event: 'metrics_create_failed', component: 'MetricsService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Запуск HTTP сервера
   */
  private async startServer(): Promise<void> {
    try {
      if (!this.config.metrics.enabled) {
        logger.info('Metrics сервер вимкнено', { type: 'metrics_service', event: 'disabled', component: 'MetricsService' });
        return;
      }

      const http = require('http');
      if (this.server && this.server.listening) {
        logger.warn('⚠️ Metrics сервер вже запущено', { type: 'metrics_service', event: 'already_running', component: 'MetricsService' });
        return;
      }
      
      this.server = http.createServer(async (req: any, res: any) => {
        try {
          if (req.url === this.config.metrics.path) {
            res.writeHead(200, { 'Content-Type': 'text/plain' });
            
            if (this.registry) {
              const metrics = await this.registry.metrics();
              res.end(metrics);
            } else {
              res.end('# Metrics not available');
            }
          } else {
            res.writeHead(404);
            res.end('Not Found');
          }
        } catch (error) {
          logger.error('Помилка обробки metrics запиту:', {
            type: 'metrics_service', event: 'request_failed', component: 'MetricsService',
            errorName: error instanceof Error ? error.name : undefined,
            errorMessage: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined,
          });
          res.writeHead(500);
          res.end('Internal Server Error');
        }
      });

      this.server.listen(this.config.metrics.port, () => {
        logger.info(`📊 Metrics сервер запущено на порту ${this.config.metrics.port}`, { type: 'metrics_service', event: 'server_started', component: 'MetricsService', port: this.config.metrics.port, path: this.config.metrics.path });
      });

      this.server.on('error', (error: Error) => {
        logger.error('Помилка metrics сервера:', {
          type: 'metrics_service', event: 'server_error', component: 'MetricsService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      });
    } catch (error) {
      logger.error('Помилка запуску metrics сервера:', {
        type: 'metrics_service', event: 'server_start_failed', component: 'MetricsService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Інкремент лічильника команд
   */
  public incrementCommand(command: string, status: string = 'success'): void {
    if (this.metrics) {
      this.metrics.commandsTotal.inc({ command, status });
    }
  }

  /**
   * Інкремент лічильника повідомлень
   */
  public incrementMessage(type: string): void {
    if (this.metrics) {
      this.metrics.messagesTotal.inc({ type });
    }
  }

  /**
   * Інкремент лічильника помилок
   */
  public incrementError(type: string, service: string = 'unknown'): void {
    if (this.metrics) {
      this.metrics.errorsTotal.inc({ type, service });
    }
  }

  /**
   * Встановлення кількості активних користувачів
   */
  public setActiveUsers(count: number): void {
    if (this.metrics) {
      this.metrics.activeUsers.set(count);
    }
  }

  /**
   * Встановлення кількості активних серверів
   */
  public setActiveGuilds(count: number): void {
    if (this.metrics) {
      this.metrics.activeGuilds.set(count);
    }
  }

  /**
   * Оновлення використання пам'яті
   */
  public updateMemoryUsage(): void {
    if (this.metrics) {
      const memUsage = process.memoryUsage();
      this.metrics.memoryUsage.set(memUsage.heapUsed);
    }
  }

  /**
   * Оновлення часу роботи
   */
  public updateUptime(): void {
    if (this.metrics) {
      const uptime = process.uptime();
      this.metrics.uptime.set(uptime);
    }
  }

  /**
   * Вимірювання тривалості команди
   */
  public measureCommandDuration(command: string, duration: number): void {
    if (this.metrics) {
      this.metrics.commandDuration.observe({ command }, duration / 1000);
    }
  }

  /**
   * Вимірювання часу відповіді API
   */
  public measureApiResponseTime(service: string, _endpoint: string, duration: number): void {
    if (this.metrics) {
      this.metrics.apiResponseTime.observe({ service }, duration / 1000);
    }
  }

  /**
   * Оновлення метрик кешу
   */
  public updateCacheMetrics(cacheStats: CacheStats): void {
    if (this.metrics) {
      const totalRequests = cacheStats.hits + cacheStats.misses;
      const hitRate = totalRequests > 0 ? (cacheStats.hits / totalRequests) * 100 : 0;
      
      this.metrics.cacheHitRate.set(hitRate);
      this.metrics.cacheSize.set(cacheStats.hits + cacheStats.misses);
    }
  }

  /**
   * Оновлення метрик черг
   */
  public updateQueueMetrics(queueStats: QueueStats): void {
    if (this.metrics) {
      this.metrics.queueLength.set({ priority: 'high' }, queueStats.high.length);
      this.metrics.queueLength.set({ priority: 'normal' }, queueStats.normal.length);
      this.metrics.queueLength.set({ priority: 'low' }, queueStats.low.length);
    }
  }

  /**
   * Оновлення метрик connection pool
   */
  public updateConnectionPoolMetrics(connectionStats: Record<string, unknown>): void {
    if (this.metrics) {
      for (const [service, stats] of Object.entries(connectionStats)) {
        const usage = (stats as any).inUse ? 100 : 0;
        this.metrics.connectionPoolUsage.set({ service }, usage);
      }
    }
  }

  /**
   * Оновлення AI метрик
   */
  public updateAIMetrics(provider: string, status: string, duration: number): void {
    if (this.metrics) {
      this.metrics.aiRequestsTotal.inc({ provider, status });
      this.metrics.aiResponseTime.observe({ provider }, duration / 1000);
    }
  }

  /**
   * Оновлення Google API метрик
   */
  public updateGoogleApiMetrics(service: string, endpoint: string, status: string, duration: number): void {
    if (this.metrics) {
      this.metrics.googleApiRequestsTotal.inc({ service, endpoint, status });
      this.metrics.googleApiResponseTime.observe({ service }, duration / 1000);
    }
  }

  /**
   * Оновлення всіх метрик
   */
  public updateAllMetrics(): void {
    try {
      this.updateMemoryUsage();
      this.updateUptime();
      
      // TODO: Отримати статистику з інших сервісів
      // const cacheStats = this.bot.serviceContainer.get('CacheService').getCacheStats();
      // this.updateCacheMetrics(cacheStats);
      
      // const queueStats = this.bot.queueManager.getQueueStats();
      // this.updateQueueMetrics(queueStats);
      
      // const connectionStats = this.bot.serviceContainer.get('GoogleService').getConnectionStats();
      // this.updateConnectionPoolMetrics(connectionStats);
      
    } catch (error) {
      logger.error('Помилка оновлення метрик:', { type: 'metrics_service', event: 'update_failed', component: 'MetricsService', errorName: error instanceof Error ? error.name : undefined, errorMessage: error instanceof Error ? error.message : String(error), stack: error instanceof Error ? error.stack : undefined });
    }
  }

  /**
   * Запуск періодичних оновлень
   */
  private startPeriodicUpdates(): void {
    this.updateInterval = setInterval(() => {
      this.updateAllMetrics();
    }, 30000); // Кожні 30 секунд
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    try {
      if (!this.config.metrics.enabled) {
        return {
          healthy: true,
          service: this.name,
          details: { enabled: false },
        };
      }

      if (!this.registry || !this.metrics) {
        return {
          healthy: false,
          service: this.name,
          error: 'Метрики не ініціалізовано',
        };
      }

      // Тестовий запит до metrics endpoint
      if (this.server) {
        try {
          const http = require('http');
          const response = await new Promise((resolve, reject) => {
            const req = http.get(`http://localhost:${this.config.metrics.port}${this.config.metrics.path}`, (res: any) => {
              let data = '';
              res.on('data', (chunk: string) => data += chunk);
              res.on('end', () => resolve({ statusCode: res.statusCode, data }));
            });
            req.on('error', reject);
            req.setTimeout(5000, () => reject(new Error('Timeout')));
          });

          if ((response as any).statusCode !== 200) {
            return {
              healthy: false,
              service: this.name,
              error: `Metrics endpoint returned ${(response as any).statusCode}`,
            };
          }
        } catch (error) {
          return {
            healthy: false,
            service: this.name,
            error: `Metrics endpoint test failed: ${error}`,
          };
        }
      }

      return {
        healthy: true,
        service: this.name,
        details: {
          metricsCount: this.stats.metricsCount,
          serverRunning: !!this.server,
          port: this.config.metrics.port,
        },
      };
    } catch (error) {
      return {
        healthy: false,
        service: this.name,
        error: `Health check failed: ${error}`,
      };
    }
  }

  /**
   * Завершення роботи
   */
  protected async onShutdown(): Promise<void> {
    try {
      if (this.updateInterval) {
        clearInterval(this.updateInterval);
        this.updateInterval = null;
      }

      if (this.server) {
        await new Promise<void>((resolve) => this.server.close(() => resolve()));
        this.server = null;
      }

      logger.info('✅ Metrics Service зупинено', { type: 'metrics_service', event: 'shutdown_success', component: 'MetricsService' });
    } catch (error) {
      logger.error('❌ Помилка зупинки Metrics Service:', { type: 'metrics_service', event: 'shutdown_failed', component: 'MetricsService', errorName: error instanceof Error ? error.name : undefined, errorMessage: error instanceof Error ? error.message : String(error), stack: error instanceof Error ? error.stack : undefined });
      throw error;
    }
  }

  /**
   * Отримання статистики
   */
  protected onGetStats(): Partial<MetricsServiceStats> {
    return this.stats;
  }
} 