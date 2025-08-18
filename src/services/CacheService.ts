/**
 * Redis Cache Service
 * Оптимізоване кешування з підтримкою різних стратегій
 */

import { createClient } from 'redis';
import type { BotConfig, HealthStatus, ServiceStats, CacheStats, CacheOptions } from '@/types';

import { BaseService as BaseServiceClass } from '@/core/BaseService';
import logger from '@/utils/logger';

// logger: використовуємо стандартний alias-імпорт із '@/utils/logger'

interface CacheServiceStats extends ServiceStats {
  hits: number;
  misses: number;
  sets: number;
  deletes: number;
  errors: number;
  hitRate: number;
  totalRequests: number;
}

interface CacheServiceOptions extends CacheOptions {
  compress?: boolean;
  serialize?: boolean;
}

export class CacheService extends BaseServiceClass {
  private client: ReturnType<typeof createClient> | null = null;
  private isConnected = false;
  private stats: CacheServiceStats;
  private readonly defaultTTL = 3600; // 1 година
  private readonly maxRetries = 3;
  private readonly retryDelay = 1000; // 1 секунда

  constructor(config: BotConfig) {
    super('CacheService', config);
    this.stats = {
      service: 'CacheService',
      uptime: 0,
      requests: 0,
      errors: 0,
      hits: 0,
      misses: 0,
      sets: 0,
      deletes: 0,
      hitRate: 0,
      totalRequests: 0,
    };
  }

  /**
   * Ініціалізація Redis клієнта
   */
  protected async onInitialize(): Promise<void> {
    try {
      if (!this.isEnabled()) {
        logger.info('Redis кешування вимкнено', { component: 'CacheService' });
        return;
      }

      // Створення Redis клієнта з оптимізацією
      const redisOptions: Parameters<typeof createClient>[0] = {
        socket: {
          connectTimeout: 10000,
          reconnectStrategy: retries => {
            if (retries > this.maxRetries) {
              logger.error('Redis: Максимальна кількість спроб підключення досягнута', {
                component: 'CacheService',
              });
              return false;
            }
            return Math.min(retries * this.retryDelay, 30000);
          },
        },
      };

      // Додаємо url лише якщо визначено, щоб уникнути undefined при exactOptionalPropertyTypes
      if (this.config.redis?.url) {
        (redisOptions as { url: string }).url = this.config.redis.url;
      }

      this.client = createClient(redisOptions);

      // Обробники подій
      this.setupEventHandlers();

      // Підключення до Redis
      await this.connect();

      // Валідація підключення
      await this.validateConnection();

      logger.info('✅ Redis Cache Service ініціалізовано', { component: 'CacheService' });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Redis:', {
        type: 'cache_service',
        event: 'init_failed',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Перевірка чи ввімкнено Redis (з урахуванням legacy-прапора cache.enabled)
   */
  private isEnabled(): boolean {
    const anyCfg = this.config as any;
    return Boolean(this.config.redis?.enabled ?? anyCfg.cache?.enabled ?? false);
  }

  /**
   * Налаштування обробників подій Redis
   */
  private setupEventHandlers(): void {
    if (!this.client) return;

    this.client.on('connect', () => {
      logger.info('🔗 Redis: Підключено', { component: 'CacheService' });
      this.isConnected = true;
    });

    this.client.on('ready', () => {
      logger.info('✅ Redis: Готовий до роботи', { component: 'CacheService' });
    });

    this.client.on('error', error => {
      logger.error('❌ Redis помилка:', {
        type: 'cache_service',
        event: 'redis_error',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      this.isConnected = false;
      this.stats.errors++;
    });

    this.client.on('end', () => {
      logger.warn("🔌 Redis: З'єднання закрито", { component: 'CacheService' });
      this.isConnected = false;
    });

    this.client.on('reconnecting', () => {
      logger.info('🔄 Redis: Перепідключення...', { component: 'CacheService' });
    });
  }

  /**
   * Підключення до Redis
   */
  private async connect(): Promise<void> {
    if (!this.client) {
      throw new Error('Redis клієнт не ініціалізовано');
    }

    try {
      await this.client.connect();
      logger.info('✅ Підключення до Redis успішне', { component: 'CacheService' });
    } catch (error) {
      logger.error('❌ Помилка підключення до Redis:', {
        type: 'cache_service',
        event: 'connect_failed',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Валідація підключення
   */
  private async validateConnection(): Promise<void> {
    if (!this.client) {
      throw new Error('Redis клієнт не ініціалізовано');
    }

    try {
      await this.client.ping();
      logger.info('✅ Redis підключення валідне', { component: 'CacheService' });
    } catch (error) {
      logger.error('❌ Помилка валідації Redis підключення:', {
        type: 'cache_service',
        event: 'ping_failed',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Отримання значення з кешу
   */
  public async get<T = unknown>(key: string, options: CacheServiceOptions = {}): Promise<T | null> {
    if (!this.client || !this.isConnected) {
      this.stats.misses++;
      return null;
    }

    try {
      const value = await this.client.get(key);

      if (value === null) {
        this.stats.misses++;
        this.updateStats();
        return null;
      }

      this.stats.hits++;
      this.updateStats();

      // Десеріалізація
      if (options.serialize !== false) {
        try {
          return JSON.parse(value) as T;
        } catch {
          return value as T;
        }
      }

      return value as T;
    } catch (error) {
      this.stats.errors++;
      this.stats.misses++;
      this.updateStats();
      logger.error('❌ Помилка отримання з кешу:', {
        type: 'cache_service',
        event: 'get_failed',
        service: this.name,
        key,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error instanceof Error ? error : new Error(String(error));
    }
  }

  /**
   * Збереження значення в кеш
   */
  public async set<T = unknown>(
    key: string,
    value: T,
    ttl: number = this.defaultTTL,
    options: CacheServiceOptions = {}
  ): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      let serializedValue: string;

      // Серіалізація
      if (options.serialize !== false) {
        serializedValue = JSON.stringify(value);
      } else {
        serializedValue = String(value);
      }

      // Підтримка як setEx, так і set з EX
      if (typeof (this.client as any).setEx === 'function') {
        await (this.client as any).setEx(key, ttl, serializedValue);
      } else if (typeof (this.client as any).set === 'function') {
        await (this.client as any).set(key, serializedValue, { EX: ttl });
      } else {
        throw new Error('Redis client does not support set/setEx');
      }
      this.stats.sets++;
      this.updateStats();

      return true;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка збереження в кеш:', {
        type: 'cache_service',
        event: 'set_failed',
        service: this.name,
        key,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error instanceof Error ? error : new Error(String(error));
    }
  }

  /**
   * Видалення ключа з кешу
   */
  public async delete(key: string): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      const result = await this.client.del(key);
      this.stats.deletes++;
      this.updateStats();
      return result > 0;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка видалення з кешу:', {
        type: 'cache_service',
        event: 'delete_failed',
        service: this.name,
        key,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return false;
    }
  }

  /**
   * Видалення ключів за патерном
   */
  public async deletePattern(pattern: string): Promise<number> {
    if (!this.client || !this.isConnected) {
      return 0;
    }

    try {
      const keys = await this.client.keys(pattern);
      if (keys.length === 0) {
        return 0;
      }

      const result = await this.client.del(keys);
      this.stats.deletes += result;
      this.updateStats();
      return result;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка видалення за патерном:', {
        type: 'cache_service',
        event: 'delete_pattern_failed',
        service: this.name,
        pattern,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return 0;
    }
  }

  /**
   * Перевірка існування ключа
   */
  public async exists(key: string): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      const result = await this.client.exists(key);
      return result > 0;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка перевірки існування ключа:', {
        type: 'cache_service',
        event: 'exists_failed',
        service: this.name,
        key,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return false;
    }
  }

  /**
   * Встановлення TTL для ключа
   */
  public async expire(key: string, ttl: number): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      const result = await this.client.expire(key, ttl);
      return Boolean(result);
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка встановлення TTL:', {
        type: 'cache_service',
        event: 'expire_failed',
        service: this.name,
        key,
        ttl,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return false;
    }
  }

  /**
   * Отримання TTL ключа
   */
  public async ttl(key: string): Promise<number> {
    if (!this.client || !this.isConnected) {
      return -2; // Ключ не існує
    }

    try {
      const result = await this.client.ttl(key);
      return result;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка отримання TTL:', {
        type: 'cache_service',
        event: 'ttl_failed',
        service: this.name,
        key,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return -2;
    }
  }

  /**
   * Отримання або встановлення значення
   */
  public async getOrSet<T = unknown>(
    key: string,
    fallbackFn: () => Promise<T>,
    ttl: number = this.defaultTTL,
    options: CacheServiceOptions = {}
  ): Promise<T> {
    // Спробувати отримати з кешу
    const cached = await this.get<T>(key, options);
    if (cached !== null) {
      return cached;
    }

    // Виконати fallback функцію
    const value = await fallbackFn();

    // Зберегти в кеш
    await this.set(key, value, ttl, options);

    return value;
  }

  /**
   * Отримання множинних значень
   */
  public async mget<T = unknown>(
    keys: string[],
    options: CacheServiceOptions = {}
  ): Promise<(T | null)[]> {
    if (!this.client || !this.isConnected) {
      return keys.map(() => null);
    }

    try {
      const values = await this.client.mGet(keys);

      const result = values.map(value => {
        if (value === null) {
          this.stats.misses++;
          return null;
        }

        this.stats.hits++;

        // Десеріалізація
        if (options.serialize !== false) {
          try {
            return JSON.parse(value) as T;
          } catch {
            return value as T;
          }
        }

        return value as T;
      });
      this.updateStats();
      return result;
    } catch (error) {
      this.stats.errors++;
      this.stats.misses += keys.length;
      this.updateStats();
      logger.error('❌ Помилка множинного отримання:', {
        type: 'cache_service',
        event: 'mget_failed',
        service: this.name,
        keysCount: keys.length,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return keys.map(() => null);
    }
  }

  /**
   * Збереження множинних значень
   */
  public async mset<T = unknown>(
    keyValuePairs: Array<{ key: string; value: T; ttl?: number }>,
    defaultTTL: number = this.defaultTTL
  ): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      const pipeline = this.client.multi();

      for (const { key, value, ttl } of keyValuePairs) {
        const serializedValue = JSON.stringify(value);
        const finalTTL = ttl || defaultTTL;
        pipeline.setEx(key, finalTTL, serializedValue);
      }

      await pipeline.exec();
      this.stats.sets += keyValuePairs.length;
      this.updateStats();

      return true;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка множинного збереження:', {
        type: 'cache_service',
        event: 'mset_failed',
        service: this.name,
        items: keyValuePairs.length,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return false;
    }
  }

  /**
   * Очищення всього кешу
   */
  public async clear(): Promise<boolean> {
    if (!this.client || !this.isConnected) {
      return false;
    }

    try {
      await this.client.flushDb();
      logger.info('🧹 Кеш очищено', { component: 'CacheService' });
      return true;
    } catch (error) {
      this.stats.errors++;
      this.updateStats();
      logger.error('❌ Помилка очищення кешу:', {
        type: 'cache_service',
        event: 'clear_failed',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return false;
    }
  }

  /**
   * Отримання статистики кешу
   */
  public getCacheStats(): CacheStats {
    return {
      hits: this.stats.hits,
      misses: this.stats.misses,
      sets: this.stats.sets,
      deletes: this.stats.deletes,
      errors: this.stats.errors,
    };
  }

  /**
   * Оновлення статистики
   */
  private updateStats(): void {
    this.stats.totalRequests = this.stats.hits + this.stats.misses;
    this.stats.hitRate =
      this.stats.totalRequests > 0 ? (this.stats.hits / this.stats.totalRequests) * 100 : 0;
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    try {
      if (!this.isEnabled()) {
        return {
          healthy: true,
          service: this.name,
          details: { enabled: false },
        };
      }

      if (!this.client || !this.isConnected) {
        return {
          healthy: false,
          service: this.name,
          error: 'Redis клієнт не підключено',
        };
      }

      // Тестовий запит
      try {
        await this.client.ping();
      } catch (error) {
        return {
          healthy: false,
          service: this.name,
          error: `Redis ping failed: ${error}`,
        };
      }

      return {
        healthy: true,
        service: this.name,
        details: {
          connected: this.isConnected,
          hitRate: this.stats.hitRate,
          totalRequests: this.stats.totalRequests,
          errors: this.stats.errors,
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
      if (this.client && this.isConnected) {
        if (typeof (this.client as any).quit === 'function') {
          await (this.client as any).quit();
        } else if (typeof (this.client as any).disconnect === 'function') {
          await (this.client as any).disconnect();
        }
        this.isConnected = false;
      }

      logger.info('✅ Cache Service зупинено', { component: 'CacheService' });
    } catch (error) {
      logger.error('❌ Помилка зупинки Cache Service:', {
        type: 'cache_service',
        event: 'shutdown_failed',
        service: this.name,
        component: 'CacheService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Отримання статистики
   */
  protected onGetStats(): Partial<CacheServiceStats> {
    return this.stats;
  }

  /**
   * Розмір кешу (кількість ключів)
   */
  public async getSize(): Promise<number> {
    if (!this.client || !this.isConnected) return 0;
    const keys = await this.client.keys('*');
    return keys.length;
  }

  /**
   * Спрощена статистика для юніт-тестів
   */
  public async getCacheStatsSimple(): Promise<{ hits: number; misses: number; size: number; hitRate: number }> {
    const size = await this.getSize();
    // Повертаємо hitRate у діапазоні [0..1] для сумісності зі старими тестами
    const hitRate = this.stats.totalRequests > 0 ? this.stats.hits / this.stats.totalRequests : 0;
    return {
      hits: this.stats.hits,
      misses: this.stats.misses,
      size,
      hitRate,
    };
  }
}
