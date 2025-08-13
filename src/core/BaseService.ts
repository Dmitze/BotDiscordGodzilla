/**
 * Базовий клас для всіх сервісів
 * Надає спільну функціональність та інтерфейс
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import type { BaseService as IBaseService, BotConfig, HealthStatus, ServiceStats } from '@/types';
import logger from '@/utils/logger';

// Константи для базового сервісу
const BASE_SERVICE_CONSTANTS = {
  INITIALIZATION_TIMEOUT: 180000, // 180 секунд (збільшено для повільних підключень)
  SHUTDOWN_TIMEOUT: 10000, // 10 секунд
  HEALTH_CHECK_TIMEOUT: 5000, // 5 секунд
  MAX_RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000, // 1 секунда
} as const;

export abstract class BaseService implements IBaseService {
  public readonly name: string;
  public readonly config: BotConfig;
  protected _initialized = false;
  protected startTime: number;
  protected isShuttingDown = false;
  protected retryCount = 0;
  private initializationTimeout: NodeJS.Timeout | null = null;
  private shutdownTimeout: NodeJS.Timeout | null = null;

  constructor(name: string, config: BotConfig) {
    this.name = name;
    this.config = config;
    this.startTime = Date.now();

    logger.debug(`🔧 Створено базовий сервіс: ${this.name}`, {
      type: 'service',
      event: 'created',
      service: this.name,
    });
  }

  /**
   * Повертає ім'я сервісу
   */
  public getName(): string {
    return this.name;
  }

  /**
   * Ознака, чи сервіс ініціалізовано
   */
  public isInitialized(): boolean {
    return this._initialized;
  }

  /**
   * Ініціалізація сервісу з детальним логуванням
   */
  public async initialize(): Promise<void> {
    if (this._initialized) {
      logger.warn(`⚠️ Сервіс ${this.name} вже ініціалізовано`, {
        type: 'service',
        event: 'already_initialized',
        service: this.name,
      });
      return;
    }

    if (this.isShuttingDown) {
      throw new Error(`Неможливо ініціалізувати сервіс ${this.name} під час зупинки`);
    }

    const startTime = Date.now();

    try {
      logger.info(`🚀 Ініціалізація сервісу ${this.name}...`, {
        type: 'service',
        event: 'initialize_start',
        service: this.name,
      });

      // Встановлення таймауту для ініціалізації (без throw з setTimeout)
      const initPromise = this.onInitialize();
      const timeoutPromise = new Promise<never>((_, reject) => {
        this.initializationTimeout = setTimeout(() => {
          logger.error(`⏰ Таймаут ініціалізації сервісу ${this.name}`, {
            type: 'service',
            event: 'initialize_timeout',
            service: this.name,
          });
          reject(new Error(`Таймаут ініціалізації сервісу ${this.name}`));
        }, BASE_SERVICE_CONSTANTS.INITIALIZATION_TIMEOUT);
      });

      await Promise.race([initPromise, timeoutPromise]);

      // Очищення таймауту
      if (this.initializationTimeout) {
        clearTimeout(this.initializationTimeout);
        this.initializationTimeout = null;
      }

      this._initialized = true;
      this.retryCount = 0;

      const duration = Date.now() - startTime;
      logger.info(`✅ Сервіс ${this.name} успішно ініціалізовано за ${duration}ms`, {
        type: 'service',
        event: 'initialized',
        service: this.name,
        durationMs: duration,
      });
    } catch (error) {
      const duration = Date.now() - startTime;
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'initialize_failed',
              service: this.name,
              durationMs: duration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'initialize_failed',
              service: this.name,
              durationMs: duration,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка ініціалізації сервісу ${this.name} після ${duration}ms`, meta);

      // Очищення таймауту
      if (this.initializationTimeout) {
        clearTimeout(this.initializationTimeout);
        this.initializationTimeout = null;
      }

      // Спроба повторної ініціалізації
      if (this.retryCount < BASE_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS) {
        this.retryCount++;
        logger.info(
          `🔄 Спроба повторної ініціалізації ${this.retryCount}/${BASE_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS} для сервісу ${this.name}...`,
          {
            type: 'service',
            event: 'initialize_retry',
            service: this.name,
            attempt: this.retryCount,
          }
        );

        await new Promise(resolve => setTimeout(resolve, BASE_SERVICE_CONSTANTS.RETRY_DELAY));
        return this.initialize();
      }

      throw new Error(
        `Помилка ініціалізації сервісу ${this.name}: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Завершення роботи сервісу з детальним логуванням
   */
  public async shutdown(): Promise<void> {
    if (!this._initialized) {
      logger.debug(`ℹ️ Сервіс ${this.name} не ініціалізовано, пропускаю зупинку`, {
        type: 'service',
        event: 'shutdown_skip_not_initialized',
        service: this.name,
      });
      return;
    }

    if (this.isShuttingDown) {
      logger.warn(`⚠️ Сервіс ${this.name} вже зупиняється`, {
        type: 'service',
        event: 'shutdown_already_in_progress',
        service: this.name,
      });
      return;
    }

    this.isShuttingDown = true;
    const shutdownStartTime = Date.now();

    try {
      logger.info(`🛑 Завершення роботи сервісу ${this.name}...`, {
        type: 'service',
        event: 'shutdown_start',
        service: this.name,
      });

      // Встановлення таймауту для зупинки
      this.shutdownTimeout = setTimeout(() => {
        logger.error(`⏰ Таймаут зупинки сервісу ${this.name}`, {
          type: 'service',
          event: 'shutdown_timeout',
          service: this.name,
        });
        throw new Error(`Таймаут зупинки сервісу ${this.name}`);
      }, BASE_SERVICE_CONSTANTS.SHUTDOWN_TIMEOUT);

      await this.onShutdown();

      // Очищення таймауту
      if (this.shutdownTimeout) {
        clearTimeout(this.shutdownTimeout);
        this.shutdownTimeout = null;
      }

      this._initialized = false;
      this.isShuttingDown = false;

      const duration = Date.now() - shutdownStartTime;
      logger.info(`✅ Сервіс ${this.name} успішно зупинено за ${duration}ms`, {
        type: 'service',
        event: 'shutdown',
        service: this.name,
        durationMs: duration,
      });
    } catch (error) {
      const duration = Date.now() - shutdownStartTime;
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'shutdown_failed',
              service: this.name,
              durationMs: duration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'shutdown_failed',
              service: this.name,
              durationMs: duration,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка зупинки сервісу ${this.name} після ${duration}ms`, meta);

      // Очищення таймауту
      if (this.shutdownTimeout) {
        clearTimeout(this.shutdownTimeout);
        this.shutdownTimeout = null;
      }

      this._initialized = false;
      this.isShuttingDown = false;

      throw new Error(
        `Помилка зупинки сервісу ${this.name}: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Перевірка здоров'я сервісу з детальним логуванням
   */
  public async healthCheck(): Promise<HealthStatus> {
    if (!this._initialized) {
      return {
        healthy: false,
        service: this.name,
        error: 'Сервіс не ініціалізовано',
      };
    }

    const startTime = Date.now();

    try {
      logger.debug(`🏥 Health check сервісу ${this.name}...`, {
        type: 'service',
        event: 'healthcheck_start',
        service: this.name,
      });

      const health = await this.onHealthCheck();
      const duration = Date.now() - startTime;

      if (!health.healthy) {
        logger.warn(`⚠️ Health check сервісу ${this.name} виявив проблеми за ${duration}ms`, {
          type: 'service',
          event: 'healthcheck_unhealthy',
          service: this.name,
          durationMs: duration,
          details: health,
        });
      } else {
        logger.debug(`✅ Health check сервісу ${this.name} пройшов успішно за ${duration}ms`, {
          type: 'service',
          event: 'healthcheck_ok',
          service: this.name,
          durationMs: duration,
        });
      }

      return {
        healthy: health.healthy,
        service: this.name,
        ...(health.error && { error: health.error }),
        ...(health.details && { details: health.details }),
      };
    } catch (error) {
      const duration = Date.now() - startTime;
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'healthcheck_failed',
              service: this.name,
              durationMs: duration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'healthcheck_failed',
              service: this.name,
              durationMs: duration,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка health check сервісу ${this.name} після ${duration}ms`, meta);

      return {
        healthy: false,
        service: this.name,
        error: `Помилка health check: ${error instanceof Error ? error.message : 'Невідома помилка'}`,
      };
    }
  }

  /**
   * Сумісний alias для healthCheck
   */
  public async getHealthStatus(): Promise<HealthStatus> {
    return this.healthCheck();
  }

  /**
   * Отримання статистики сервісу з детальним логуванням
   */
  public getStats(): ServiceStats {
    try {
      const baseStats: ServiceStats = {
        service: this.name,
        uptime: Date.now() - this.startTime,
        requests: 0,
        errors: 0,
        isInitialized: this._initialized,
        isShuttingDown: this.isShuttingDown,
        retryCount: this.retryCount,
      };

      const serviceStats = this.onGetStats();
      const combinedStats = {
        ...baseStats,
        ...serviceStats,
        type: 'service',
        event: 'get_stats',
        service: this.name,
      };

      logger.debug(`📊 Статистика сервісу ${this.name}:`, combinedStats);

      return combinedStats;
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'get_stats_failed',
              service: this.name,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'get_stats_failed',
              service: this.name,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка отримання статистики сервісу ${this.name}`, meta);

      return {
        service: this.name,
        uptime: Date.now() - this.startTime,
        requests: 0,
        errors: 1,
        isInitialized: this._initialized,
        isShuttingDown: this.isShuttingDown,
        retryCount: this.retryCount,
        error: error instanceof Error ? error.message : 'Невідома помилка',
      } as ServiceStats;
    }
  }

  /**
   * Перевірка чи сервіс ініціалізовано
   */
  protected checkInitialized(): void {
    if (!this._initialized) {
      const error = `Сервіс ${this.name} не ініціалізовано`;
      logger.error(`❌ ${error}`, {
        type: 'service',
        event: 'check_initialized_failed',
        service: this.name,
      });
      throw new Error(error);
    }
  }

  /**
   * Перевірка чи сервіс не зупиняється
   */
  protected checkNotShuttingDown(): void {
    if (this.isShuttingDown) {
      const error = `Сервіс ${this.name} зупиняється`;
      logger.warn(`⚠️ ${error}`, {
        type: 'service',
        event: 'check_shutting_down',
        service: this.name,
      });
      throw new Error(error);
    }
  }

  /**
   * Безпечне виконання операції з обробкою помилок
   */
  protected async safeExecute<T>(
    operation: () => Promise<T>,
    operationName: string,
    fallback?: T
  ): Promise<T> {
    const startTime = Date.now();

    try {
      logger.debug(`🔄 Виконання операції ${operationName} в сервісі ${this.name}...`, {
        type: 'service',
        event: 'operation_start',
        service: this.name,
        operation: operationName,
      });

      const result = await operation();

      const duration = Date.now() - startTime;
      logger.debug(
        `✅ Операція ${operationName} в сервісі ${this.name} завершена за ${duration}ms`,
        {
          type: 'service',
          event: 'operation_success',
          service: this.name,
          operation: operationName,
          durationMs: duration,
        }
      );

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'operation_failed',
              service: this.name,
              operation: operationName,
              durationMs: duration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'operation_failed',
              service: this.name,
              operation: operationName,
              durationMs: duration,
              errorMessage: String(error),
            };
      logger.error(
        `❌ Помилка операції ${operationName} в сервісі ${this.name} після ${duration}ms`,
        meta
      );

      if (fallback !== undefined) {
        logger.warn(
          `🔄 Використання fallback значення для операції ${operationName} в сервісі ${this.name}`,
          {
            type: 'service',
            event: 'operation_fallback',
            service: this.name,
            operation: operationName,
          }
        );
        return fallback;
      }

      throw error;
    }
  }

  /**
   * Очищення ресурсів сервісу
   */
  protected async cleanup(): Promise<void> {
    try {
      logger.info(`🧹 Очищення ресурсів сервісу ${this.name}...`, {
        type: 'service',
        event: 'cleanup_start',
        service: this.name,
      });

      // Очищення таймаутів
      if (this.initializationTimeout) {
        clearTimeout(this.initializationTimeout);
        this.initializationTimeout = null;
      }

      if (this.shutdownTimeout) {
        clearTimeout(this.shutdownTimeout);
        this.shutdownTimeout = null;
      }

      logger.info(`✅ Ресурси сервісу ${this.name} очищено`, {
        type: 'service',
        event: 'cleanup_success',
        service: this.name,
      });
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'service',
              event: 'cleanup_failed',
              service: this.name,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'service',
              event: 'cleanup_failed',
              service: this.name,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка очищення ресурсів сервісу ${this.name}`, meta);
    }
  }

  /**
   * Абстрактні методи для реалізації в нащадках
   */
  protected abstract onInitialize(): Promise<void>;
  protected abstract onShutdown(): Promise<void>;
  protected abstract onHealthCheck(): Promise<HealthStatus>;
  protected abstract onGetStats(): Partial<ServiceStats>;
}
