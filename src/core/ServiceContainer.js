/**
 * Service Container для Dependency Injection
 * Централізоване управління всіма сервісами додатку
 */

const logger = require('../utils/logger');
const AIService = require('../services/AIService');
const GoogleService = require('../services/GoogleService');
const CacheService = require('../services/CacheService');
const MetricsService = require('../services/MetricsService');
const SchedulerService = require('../services/SchedulerService');

class ServiceContainer {
  constructor(config) {
    this.config = config;
    this.services = new Map();
    this.singletons = new Map();
    this.isInitialized = false;
  }

  /**
   * Ініціалізація Service Container
   */
  async initialize() {
    try {
      logger.info('🔧 Ініціалізація Service Container...');

      // Реєстрація базових сервісів
      await this.registerCoreServices();

      // Ініціалізація всіх сервісів
      await this.initializeServices();

      this.isInitialized = true;
      logger.info('✅ Service Container ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Service Container:', error);
      throw error;
    }
  }

  /**
   * Реєстрація основних сервісів
   */
  async registerCoreServices() {
    // AI Service
    this.register('ai', () => new AIService(this));

    // Google Service
    this.register('google', () => new GoogleService(this.config));

    // Cache Service
    this.register('cache', () => new CacheService(this.config));

    // Metrics Service
    this.register('metrics', () => new MetricsService(this.config));

    // Scheduler Service
    this.register('scheduler', () => new SchedulerService(this));

    logger.info('✅ Основні сервіси зареєстровані');
  }

  /**
   * Реєстрація сервісу
   */
  register(name, factory, options = {}) {
    if (this.services.has(name)) {
      logger.warn(`Сервіс ${name} вже зареєстрований, перезаписую...`);
    }

    this.services.set(name, {
      factory,
      options: {
        singleton: true,
        ...options,
      },
    });

    logger.debug(`Сервіс ${name} зареєстрований`);
  }

  /**
   * Отримання сервісу
   */
  get(name) {
    if (!this.services.has(name)) {
      throw new Error(`Сервіс ${name} не знайдено`);
    }

    const serviceInfo = this.services.get(name);
    const { factory, options } = serviceInfo;

    // Якщо це singleton і вже створений
    if (options.singleton && this.singletons.has(name)) {
      return this.singletons.get(name);
    }

    // Створення нового екземпляру
    const service = factory();

    // Збереження singleton
    if (options.singleton) {
      this.singletons.set(name, service);
    }

    return service;
  }

  /**
   * Перевірка наявності сервісу
   */
  has(name) {
    return this.services.has(name);
  }

  /**
   * Ініціалізація всіх сервісів
   */
  async initializeServices() {
    logger.info('🔄 Ініціалізація сервісів...');

    const serviceNames = Array.from(this.services.keys());
    const initPromises = serviceNames.map(async (name) => {
      try {
        const service = this.get(name);
        if (service && typeof service.initialize === 'function') {
          await service.initialize();
          logger.debug(`✅ Сервіс ${name} ініціалізовано`);
        }
      } catch (error) {
        logger.error(`❌ Помилка ініціалізації сервісу ${name}:`, error);
        throw error;
      }
    });

    await Promise.all(initPromises);
    logger.info('✅ Всі сервіси ініціалізовано');
  }

  /**
   * Отримання конфігурації
   */
  getConfig() {
    return this.config;
  }

  /**
   * Отримання статистики
   */
  getStats() {
    const stats = {
      services: {},
      singletons: this.singletons.size,
      total: this.services.size,
    };

    // Збір статистики з кожного сервісу
    for (const [name, service] of this.singletons) {
      if (service && typeof service.getStats === 'function') {
        stats.services[name] = service.getStats();
      }
    }

    return stats;
  }

  /**
   * Завершення роботи Service Container
   */
  async shutdown() {
    logger.info('🛑 Завершення роботи Service Container...');

    const shutdownPromises = Array.from(this.singletons.entries()).map(
      async ([name, service]) => {
        try {
          if (service && typeof service.shutdown === 'function') {
            await service.shutdown();
            logger.debug(`✅ Сервіс ${name} завершено`);
          }
        } catch (error) {
          logger.error(`❌ Помилка завершення сервісу ${name}:`, error);
        }
      }
    );

    await Promise.all(shutdownPromises);

    // Очищення
    this.singletons.clear();
    this.services.clear();
    this.isInitialized = false;

    logger.info('✅ Service Container завершено');
  }

  /**
   * Отримання списку всіх сервісів
   */
  getServiceList() {
    return Array.from(this.services.keys());
  }

  /**
   * Перевірка стану сервісу
   */
  isServiceHealthy(name) {
    try {
      const service = this.get(name);
      if (service && typeof service.isHealthy === 'function') {
        return service.isHealthy();
      }
      return true; // Якщо немає методу перевірки, вважаємо здоровим
    } catch (error) {
      logger.error(`Помилка перевірки стану сервісу ${name}:`, error);
      return false;
    }
  }

  /**
   * Отримання health check для всіх сервісів
   */
  getHealthStatus() {
    const health = {
      overall: true,
      services: {},
      timestamp: new Date().toISOString(),
    };

    for (const name of this.services.keys()) {
      const isHealthy = this.isServiceHealthy(name);
      health.services[name] = {
        healthy: isHealthy,
        status: isHealthy ? 'ok' : 'error',
      };

      if (!isHealthy) {
        health.overall = false;
      }
    }

    return health;
  }
}

module.exports = { ServiceContainer }; 