/**
 * Менеджер сервісів Discord бота
 * Централізоване управління всіма сервісами
 * TypeScript версія
 */

import logger from '../utils/logger';
import AIService from '../services/AIService';
import GoogleService from '../services/GoogleService';
import CacheService from '../services/CacheService';
import MetricsService from '../services/MetricsService';
import SchedulerService from '../services/SchedulerService';

interface Bot {
  config: {
    redis: {
      enabled: boolean;
    };
    isMetricsEnabled(): boolean;
  };
}

interface Service {
  initialize?: () => Promise<void>;
  start?: () => Promise<void>;
  shutdown?: () => Promise<void>;
  isActive?: () => boolean;
  getStats?: () => any;
  [key: string]: any;
}

interface ServiceStatus {
  isActive: boolean;
  hasMethod: (method: string) => boolean;
  stats: any;
}

interface ServiceManagerStats {
  total: number;
  active: number;
  services: string[];
  status: Record<string, ServiceStatus>;
}

class ServiceManager {
  private bot: Bot;
  private services: Map<string, Service>;
  private isInitialized: boolean;

  constructor(bot: Bot) {
    this.bot = bot;
    this.services = new Map();
    this.isInitialized = false;
  }

  /**
   * Ініціалізація менеджера сервісів
   */
  async initialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація менеджера сервісів...');

      // Створення сервісів
      await this.createServices();

      // Ініціалізація сервісів
      await this.initializeServices();

      this.isInitialized = true;
      logger.info('✅ Менеджер сервісів ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера сервісів:', error);
      throw error;
    }
  }

  /**
   * Створення сервісів
   */
  private async createServices(): Promise<void> {
    // AI Service
    this.services.set('ai', new AIService(this.bot));

    // Google Service
    this.services.set('google', new GoogleService(this.bot));

    // Cache Service (якщо Redis увімкнено)
    if (this.bot.config.redis.enabled) {
      this.services.set('cache', new CacheService(this.bot));
    }

    // Metrics Service (якщо метрики увімкнені)
    if (this.bot.config.isMetricsEnabled()) {
      this.services.set('metrics', new MetricsService(this.bot));
    }

    // Scheduler Service
    this.services.set('scheduler', new SchedulerService(this.bot));
  }

  /**
   * Ініціалізація сервісів
   */
  private async initializeServices(): Promise<void> {
    const initPromises = Array.from(this.services.entries()).map(async ([name, service]) => {
      try {
        if (service.initialize) {
          await service.initialize();
          logger.debug(`✅ Сервіс ${name} ініціалізовано`);
        }
      } catch (error) {
        logger.error(`❌ Помилка ініціалізації сервісу ${name}:`, error);
        // Видаляємо сервіс, який не вдалося ініціалізувати
        this.services.delete(name);
      }
    });

    await Promise.allSettled(initPromises);
  }

  /**
   * Запуск метрик
   */
  async startMetrics(): Promise<void> {
    const metricsService = this.services.get('metrics');
    if (metricsService && metricsService.start) {
      await metricsService.start();
      logger.info('📊 Метрики запущено');
    }
  }

  /**
   * Запуск кешування
   */
  async startCache(): Promise<void> {
    const cacheService = this.services.get('cache');
    if (cacheService && cacheService.start) {
      await cacheService.start();
      logger.info('💾 Кеш запущено');
    }
  }

  /**
   * Запуск планувальника
   */
  async startScheduler(): Promise<void> {
    const schedulerService = this.services.get('scheduler');
    if (schedulerService && schedulerService.start) {
      await schedulerService.start();
      logger.info('⏰ Планувальник запущено');
    }
  }

  /**
   * Отримання сервісу за назвою
   */
  getService(name: string): Service | undefined {
    return this.services.get(name);
  }

  /**
   * Перевірка наявності сервісу
   */
  hasService(name: string): boolean {
    return this.services.has(name);
  }

  /**
   * Отримання всіх сервісів
   */
  getAllServices(): Service[] {
    return Array.from(this.services.values());
  }

  /**
   * Отримання назв всіх сервісів
   */
  getServiceNames(): string[] {
    return Array.from(this.services.keys());
  }

  /**
   * Виконання методу на всіх сервісах
   */
  async executeOnAllServices(methodName: string, ...args: any[]): Promise<PromiseSettledResult<any>[]> {
    const promises = Array.from(this.services.values()).map(async service => {
      if (service[methodName] && typeof service[methodName] === 'function') {
        try {
          return await service[methodName](...args);
        } catch (error) {
          logger.error(`Помилка виконання ${methodName} на сервісі:`, error);
          return null;
        }
      }
      return null;
    });

    return Promise.allSettled(promises);
  }

  /**
   * Отримання статусу сервісів
   */
  getServicesStatus(): Record<string, ServiceStatus> {
    const status: Record<string, ServiceStatus> = {};

    for (const [name, service] of this.services.entries()) {
      status[name] = {
        isActive: service.isActive ? service.isActive() : true,
        hasMethod: (method: string) => service[method] && typeof service[method] === 'function',
        stats: service.getStats ? service.getStats() : null,
      };
    }

    return status;
  }

  /**
   * Graceful shutdown всіх сервісів
   */
  async shutdown(): Promise<void> {
    logger.info('🛑 Завершення роботи сервісів...');

    try {
      await this.executeOnAllServices('shutdown');
      logger.info('✅ Сервіси успішно завершено');
    } catch (error) {
      logger.error('❌ Помилка при завершенні сервісів:', error);
    }
  }

  /**
   * Статистика сервісів
   */
  getStats(): ServiceManagerStats {
    return {
      total: this.services.size,
      active: Array.from(this.services.values()).filter(service =>
        service.isActive ? service.isActive() : true
      ).length,
      services: this.getServiceNames(),
      status: this.getServicesStatus(),
    };
  }
}

export default ServiceManager; 