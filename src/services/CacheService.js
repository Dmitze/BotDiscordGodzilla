/**
 * Redis Cache Service
 * Оптимізоване кешування з підтримкою різних стратегій
 */

const BaseService = require('../core/BaseService');
const logger = require('../utils/logger');
const redis = require('redis');

class CacheService extends BaseService {
  constructor(config) {
    super('CacheService', config);
    this.client = null;
    this.isConnected = false;
    this.cacheStats = {
      hits: 0,
      misses: 0,
      sets: 0,
      deletes: 0,
      errors: 0,
    };
    this.defaultTTL = 3600; // 1 година
    this.maxRetries = 3;
    this.retryDelay = 1000; // 1 секунда
  }

  /**
   * Ініціалізація Redis клієнта
   */
  async onInitialize() {
    try {
      if (!this.config.redis.enabled) {
        logger.info('Redis кешування вимкнено');
        return;
      }

      // Створення Redis клієнта з оптимізацією
      this.client = redis.createClient({
        url: this.config.redis.url,
        socket: {
          connectTimeout: 10000,
          lazyConnect: true,
          reconnectStrategy: (retries) => {
            if (retries > this.maxRetries) {
              logger.error('Redis: Максимальна кількість спроб підключення досягнута');
              return false;
            }
            return Math.min(retries * this.retryDelay, 30000);
          },
        },
        retry_strategy: (options) => {
          if (options.total_retry_time > 1000 * 60 * 60) {
            return new Error('Retry time exhausted');
          }
          if (options.attempt > this.maxRetries) {
            return undefined;
          }
          return Math.min(options.attempt * this.retryDelay, 30000);
        },
      });

      // Обробники подій
      this.setupEventHandlers();

      // Підключення до Redis
      await this.connect();

      // Валідація підключення
      await this.validateConnection();

      logger.info('✅ Redis Cache Service ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Redis:', error);
      throw error;
    }
  }

  /**
   * Налаштування обробників подій Redis
   */
  setupEventHandlers() {
    this.client.on('connect', () => {
      logger.info('🔗 Redis: Підключено');
      this.isConnected = true;
    });

    this.client.on('ready', () => {
      logger.info('✅ Redis: Готовий до роботи');
    });

    this.client.on('error', (error) => {
      logger.error('❌ Redis помилка:', error);
      this.isConnected = false;
      this.cacheStats.errors++;
    });

         this.client.on('end', () => {
       logger.warn('🔌 Redis: З\'єднання закрито');
       this.isConnected = false;
     });

    this.client.on('reconnecting', () => {
      logger.info('🔄 Redis: Перепідключення...');
    });
  }

  /**
   * Підключення до Redis
   */
  async connect() {
    try {
      await this.client.connect();
    } catch (error) {
      logger.error('❌ Помилка підключення до Redis:', error);
      throw error;
    }
  }

  /**
   * Валідація підключення
   */
  async validateConnection() {
    try {
      await this.client.ping();
      logger.info('✅ Redis підключення валідовано');
    } catch (error) {
      logger.error('❌ Помилка валідації Redis підключення:', error);
      throw error;
    }
  }

  /**
   * Отримання значення з кешу
   */
  async get(key, options = {}) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected) {
        this.cacheStats.misses++;
        return null;
      }

      const value = await this.client.get(key);
      
      if (value) {
        this.cacheStats.hits++;
        this.updateStats(true, Date.now() - startTime);
        
        // Парсинг JSON якщо потрібно
        if (options.parseJSON !== false) {
          try {
            return JSON.parse(value);
          } catch {
            return value;
          }
        }
        return value;
      } else {
        this.cacheStats.misses++;
        this.updateStats(true, Date.now() - startTime);
        return null;
      }
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error(`❌ Помилка отримання з кешу (${key}):`, error);
      return null;
    }
  }

  /**
   * Збереження значення в кеш
   */
  async set(key, value, ttl = this.defaultTTL, options = {}) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected) {
        return false;
      }

      // Серіалізація значення
      const serializedValue = typeof value === 'string' 
        ? value 
        : JSON.stringify(value);

      // Встановлення з TTL
      const result = await this.client.setEx(key, ttl, serializedValue);
      
      if (result === 'OK') {
        this.cacheStats.sets++;
        this.updateStats(true, Date.now() - startTime);
        
        // Логування якщо увімкнено
        if (options.log) {
          logger.debug(`💾 Кеш: збережено ${key} (TTL: ${ttl}s)`);
        }
        
        return true;
      }
      
      return false;
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error(`❌ Помилка збереження в кеш (${key}):`, error);
      return false;
    }
  }

  /**
   * Видалення з кешу
   */
  async delete(key) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected) {
        return false;
      }

      const result = await this.client.del(key);
      this.cacheStats.deletes++;
      this.updateStats(true, Date.now() - startTime);
      
      return result > 0;
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error(`❌ Помилка видалення з кешу (${key}):`, error);
      return false;
    }
  }

  /**
   * Видалення за патерном
   */
  async deletePattern(pattern) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected) {
        return 0;
      }

      const keys = await this.client.keys(pattern);
      if (keys.length === 0) {
        return 0;
      }

      const result = await this.client.del(keys);
      this.cacheStats.deletes += keys.length;
      this.updateStats(true, Date.now() - startTime);
      
      logger.info(`🗑️ Кеш: видалено ${result} ключів за патерном ${pattern}`);
      return result;
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error(`❌ Помилка видалення за патерном (${pattern}):`, error);
      return 0;
    }
  }

  /**
   * Перевірка наявності ключа
   */
  async exists(key) {
    try {
      if (!this.isConnected) {
        return false;
      }

      const result = await this.client.exists(key);
      return result === 1;
    } catch (error) {
      logger.error(`❌ Помилка перевірки ключа (${key}):`, error);
      return false;
    }
  }

  /**
   * Встановлення TTL для ключа
   */
  async expire(key, ttl) {
    try {
      if (!this.isConnected) {
        return false;
      }

      const result = await this.client.expire(key, ttl);
      return result === 1;
    } catch (error) {
      logger.error(`❌ Помилка встановлення TTL (${key}):`, error);
      return false;
    }
  }

  /**
   * Отримання TTL ключа
   */
  async ttl(key) {
    try {
      if (!this.isConnected) {
        return -1;
      }

      return await this.client.ttl(key);
    } catch (error) {
      logger.error(`❌ Помилка отримання TTL (${key}):`, error);
      return -1;
    }
  }

  /**
   * Кешування з fallback функцією
   */
  async getOrSet(key, fallbackFn, ttl = this.defaultTTL, options = {}) {
    try {
      // Спроба отримати з кешу
      let value = await this.get(key, options);
      
      if (value !== null) {
        return value;
      }

      // Якщо немає в кеші, викликаємо fallback функцію
      if (typeof fallbackFn === 'function') {
        value = await fallbackFn();
        
        if (value !== null && value !== undefined) {
          await this.set(key, value, ttl, options);
        }
        
        return value;
      }
      
      return null;
    } catch (error) {
      logger.error(`❌ Помилка getOrSet (${key}):`, error);
      return null;
    }
  }

  /**
   * Batch операції
   */
  async mget(keys) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected || !Array.isArray(keys)) {
        return keys.map(() => null);
      }

      const values = await this.client.mGet(keys);
      this.updateStats(true, Date.now() - startTime);
      
      return values.map(value => {
        if (value) {
          this.cacheStats.hits++;
          try {
            return JSON.parse(value);
          } catch {
            return value;
          }
        } else {
          this.cacheStats.misses++;
          return null;
        }
      });
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка batch отримання:', error);
      return keys.map(() => null);
    }
  }

  /**
   * Batch збереження
   */
  async mset(keyValuePairs, ttl = this.defaultTTL) {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected || !Array.isArray(keyValuePairs)) {
        return false;
      }

      const pipeline = this.client.multi();
      
      keyValuePairs.forEach(([key, value]) => {
        const serializedValue = typeof value === 'string' 
          ? value 
          : JSON.stringify(value);
        pipeline.setEx(key, ttl, serializedValue);
      });

      await pipeline.exec();
      this.cacheStats.sets += keyValuePairs.length;
      this.updateStats(true, Date.now() - startTime);
      
      return true;
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка batch збереження:', error);
      return false;
    }
  }

  /**
   * Очищення всього кешу
   */
  async clear() {
    const startTime = Date.now();
    
    try {
      if (!this.isConnected) {
        return false;
      }

      await this.client.flushDb();
      this.updateStats(true, Date.now() - startTime);
      
      logger.info('🗑️ Кеш: повністю очищено');
      return true;
    } catch (error) {
      this.cacheStats.errors++;
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка очищення кешу:', error);
      return false;
    }
  }

  /**
   * Отримання статистики кешу
   */
  getCacheStats() {
    const hitRate = this.cacheStats.hits + this.cacheStats.misses > 0
      ? (this.cacheStats.hits / (this.cacheStats.hits + this.cacheStats.misses)) * 100
      : 0;

    return {
      ...this.cacheStats,
      hitRate: Math.round(hitRate * 100) / 100,
      isConnected: this.isConnected,
      totalOperations: this.cacheStats.hits + this.cacheStats.misses + this.cacheStats.sets + this.cacheStats.deletes,
    };
  }

  /**
   * Health check
   */
  async onHealthCheck() {
    try {
      if (!this.isConnected) {
        return {
          healthy: false,
          error: 'Redis не підключено',
          service: this.name,
        };
      }

      await this.client.ping();
      
      return {
        healthy: true,
        service: this.name,
        stats: this.getCacheStats(),
      };
    } catch (error) {
      return {
        healthy: false,
        error: error.message,
        service: this.name,
      };
    }
  }

  /**
   * Завершення роботи
   */
  async onShutdown() {
    try {
             if (this.client && this.isConnected) {
         await this.client.quit();
         logger.info('✅ Redis з\'єднання закрито');
       }
         } catch (error) {
       logger.error('❌ Помилка закриття Redis з\'єднання:', error);
     }
  }

  /**
   * Отримання розширеної статистики
   */
  getStats() {
    return {
      ...super.getStats(),
      cache: this.getCacheStats(),
      isConnected: this.isConnected,
    };
  }
}

module.exports = CacheService;
