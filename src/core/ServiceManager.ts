/**
 * Менеджер сервісів Discord бота
 * Централізоване управління всіма сервісами
 * TypeScript версія
 */

import logger from '@/utils/logger';

import { AIService } from '../services/AIService';
import { GoogleService } from '../services/GoogleService';
import { CacheService } from '../services/CacheService';
import { MemoryCacheService } from '@/services/MemoryCacheService';
import { MetricsService } from '../services/MetricsService';
import { SheetsContextService } from '../services/SheetsContextService';
import SchedulerService from '../services/SchedulerService';
import { DriveIndexerService } from '../services/DriveIndexerService';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';
import type { SearchIndex } from '@/search/SearchIndex';
import { DriveChangesService } from '../services/DriveChangesService';
import type { BotConfig } from '@/types';
import { WorkspaceDbService } from '@/services/WorkspaceDbService';

interface Bot {
  config: BotConfig;
  getService: (name: string) => any;
  serviceManager?: any;
  client?: any;
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
  //

  constructor(bot: Bot) {
    this.bot = bot;
    this.services = new Map();
    //
  }

  /**
   * Ініціалізація менеджера сервісів
   */
  async initialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація менеджера сервісів...', {
        type: 'service_manager',
        event: 'init',
        component: 'ServiceManager',
      });

      // Створення сервісів
      await this.createServices();

      // Ініціалізація сервісів
      await this.initializeServices();

      logger.info('✅ Менеджер сервісів ініціалізовано', {
        type: 'service_manager',
        event: 'init_success',
        component: 'ServiceManager',
      });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера сервісів', {
        type: 'service_manager',
        event: 'init_failed',
        component: 'ServiceManager',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }

  /**
   * Створення сервісів
   */
  private async createServices(): Promise<void> {
    // AI Service
    this.services.set('ai', new AIService(this.bot.config));

    // Google Service
    this.services.set('google', new GoogleService(this.bot.config));

    // Cache Service: завжди доступний у контейнері
    // - Якщо Redis увімкнено — використовуємо Redis CacheService
    // - Інакше або у тестах — легкий MemoryCacheService
    const useRedis = Boolean((this.bot.config as any).redis?.enabled);
    if (useRedis) {
      this.services.set('cache', new CacheService(this.bot.config));
    } else {
      this.services.set('cache', new MemoryCacheService(this.bot.config));
    }

    // Metrics Service (якщо метрики увімкнені)
    const metricsEnabled = Boolean((this.bot.config as any).metrics?.enabled);
    if (metricsEnabled) {
      this.services.set('metrics', new MetricsService(this.bot.config));
    }

    // Scheduler Service
    this.services.set('scheduler', new SchedulerService(this.bot));

    // Sheets Context Service
    this.services.set('sheetsContext', new SheetsContextService(this.bot.config));

    // Drive Changes Service (polling changes)
    this.services.set('driveChanges', new DriveChangesService(this.bot.config as BotConfig));

    // Drive Indexer Service (потребує доступу до інших сервісів через getService)
    this.services.set(
      'driveIndexer',
      new DriveIndexerService({
        config: this.bot.config as BotConfig,
        getService: (name: string) => this.getService(name),
      } as any)
    );

    // Workspace (персональний простір користувача) на SQLite
    try {
      const workspace = new WorkspaceDbService(this.bot.config as BotConfig);
      this.services.set('workspace', workspace as unknown as Service);
      logger.info('🗂️ WorkspaceDbService зареєстровано', {
        type: 'service_manager',
        event: 'workspace_registered',
        component: 'ServiceManager',
      });
    } catch (e) {
      logger.error('❌ Не вдалося створити WorkspaceDbService', {
        type: 'service_manager',
        event: 'workspace_register_failed',
        component: 'ServiceManager',
        errorMessage: e instanceof Error ? e.message : String(e),
      });
    }

    // Persistent Search Index (SQLite FTS5)
    try {
      const dbPath = process.env['BOT_INDEX_DB_PATH'] || './data/search-index.db';
      const searchIndex: SearchIndex = new SqliteSearchIndex({ dbPath });
      this.services.set('searchIndex', searchIndex as unknown as Service);
      logger.info('🗂️ SqliteSearchIndex зареєстровано', {
        type: 'service_manager',
        event: 'search_index_registered',
        component: 'ServiceManager',
        dbPath,
      });
    } catch (e) {
      logger.error('❌ Не вдалося створити SqliteSearchIndex', {
        type: 'service_manager',
        event: 'search_index_register_failed',
        component: 'ServiceManager',
        errorMessage: e instanceof Error ? e.message : String(e),
      });
    }

    // Зв'язуємо MetricsService з GoogleService (якщо обидва доступні)
    try {
      const google = this.services.get('google');
      const metrics = this.services.get('metrics');
      if (google && metrics && typeof google['setMetricsService'] === 'function') {
        google['setMetricsService'](metrics);
        logger.debug('🔗 Підключено MetricsService до GoogleService');
      }
    } catch (e) {
      logger.warn('Не вдалося підключити MetricsService до GoogleService', { error: (e as Error).message });
    }
  }

  /**
   * Ініціалізація сервісів
   */
  private async initializeServices(): Promise<void> {
    const initPromises = Array.from(this.services.entries()).map(async ([name, service]) => {
      try {
        if (service.initialize) {
          await service.initialize();
          logger.debug('✅ Сервіс ініціалізовано', {
            type: 'service_manager',
            event: 'service_initialized',
            component: 'ServiceManager',
            service: name,
          });
        }
      } catch (error) {
        logger.error('❌ Помилка ініціалізації сервісу', {
          type: 'service_manager',
          event: 'service_init_failed',
          component: 'ServiceManager',
          service: name,
          errorName: error instanceof Error ? error.name : undefined,
          errorMessage: error instanceof Error ? error.message : String(error),
          stack: error instanceof Error ? error.stack : undefined,
        });

        // Видаляємо сервіс, який не вдалося ініціалізувати
        this.services.delete(name);
      }
    });

    await Promise.allSettled(initPromises);
  }

  /** Повертає сервіс за назвою */
  public getService<T = any>(name: string): T | undefined {
    return this.services.get(name) as unknown as T | undefined;
  }

  /** Список зареєстрованих сервісів */
  public getServiceNames(): string[] {
    return Array.from(this.services.keys());
  }

  /**
   * Запуск метрик
   */
  async startMetrics(): Promise<void> {
    const metricsService = this.services.get('metrics');
    if (metricsService && metricsService.start) {
      await metricsService.start();
      logger.info('📊 Метрики запущено', {
        type: 'service_manager',
        event: 'metrics_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Запуск кешування
   */
  async startCache(): Promise<void> {
    const cacheService = this.services.get('cache');
    if (cacheService && cacheService.start) {
      await cacheService.start();
      logger.info('💾 Кеш запущено', {
        type: 'service_manager',
        event: 'cache_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Запуск планувальника
   */
  async startScheduler(): Promise<void> {
    const schedulerService = this.services.get('scheduler');
    if (schedulerService && schedulerService.start) {
      await schedulerService.start();
      logger.info('⏰ Планувальник запущено', {
        type: 'service_manager',
        event: 'scheduler_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Отримання сервісу за назвою
   */
  // getService(name: string): Service | undefined { // removed duplicate; prefer generic variant above
  //   return this.services.get(name);
  // }

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
  // getServiceNames(): string[] { // removed duplicate; prefer public generic variant above
  //   return Array.from(this.services.keys());
  // }

  /**
   * Виконання методу на всіх сервісах
   */
  async executeOnAllServices(
    methodName: string,
    ...args: any[]
  ): Promise<PromiseSettledResult<any>[]> {
    const promises = Array.from(this.services.values()).map(async service => {
      if (service[methodName] && typeof service[methodName] === 'function') {
        try {
          return await service[methodName](...args);
        } catch (error) {
          logger.error('Помилка виконання методу на сервісі', {
            type: 'service_manager',
            event: 'method_execution_failed',
            component: 'ServiceManager',
            methodName: methodName,
            errorName: error instanceof Error ? error.name : undefined,
            errorMessage: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined,
          });
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
    logger.info('🛑 Завершення роботи сервісів...', {
      type: 'service_manager',
      event: 'shutdown',
      component: 'ServiceManager',
    });

    try {
      await this.executeOnAllServices('shutdown');
      logger.info('✅ Сервіси успішно завершено', {
        type: 'service_manager',
        event: 'shutdown_success',
        component: 'ServiceManager',
      });
    } catch (error) {
      logger.error('❌ Помилка при завершенні сервісів', {
        type: 'service_manager',
        event: 'shutdown_failed',
        component: 'ServiceManager',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
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
