/**
 * Утиліти для оптимізації продуктивності Discord AI Assistant Bot
 * Моніторинг та оптимізація продуктивності
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import type { LogMeta, PerformanceMetrics, SystemMetrics } from '@/types';
import logger from './logger';

// Константи для оптимізації продуктивності
const PERFORMANCE_CONSTANTS = {
  MEMORY_THRESHOLD: 80, // 80% використання пам'яті
  CPU_THRESHOLD: 70, // 70% використання CPU
  GC_INTERVAL: 300000, // 5 хвилин
  METRICS_INTERVAL: 60000, // 1 хвилина
  SLOW_OPERATION_THRESHOLD: 5000, // 5 секунд
  CACHE_CLEANUP_INTERVAL: 600000, // 10 хвилин
  MAX_CACHE_SIZE: 1000,
  MAX_METRICS_HISTORY: 100,
  CPU_MEASUREMENT_DELAY: 100, // 100ms для вимірювання CPU
  MAX_SLOW_OPERATIONS: 50,
  SLOW_OPERATIONS_RETAIN: 25,
} as const;

// Інтерфейси для метрик
interface PerformanceData {
  timestamp: Date;
  memory: {
    rss: number;
    heapUsed: number;
    heapTotal: number;
    external: number;
  };
  cpu: {
    usage: number;
    load: number;
  };
  operations: Map<string, PerformanceMetrics>;
  slowOperations: PerformanceMetrics[];
}

interface CacheEntry<T> {
  value: T;
  timestamp: number;
  accessCount: number;
  lastAccess: number;
}

/**
 * Клас для оптимізації продуктивності
 */
export class PerformanceOptimizer {
  private static instance: PerformanceOptimizer | null = null;
  private performanceData!: PerformanceData;
  private caches = new Map<string, Map<string, CacheEntry<unknown>>>();
  private metricsHistory: PerformanceData[] = [];
  private gcInterval: NodeJS.Timeout | null = null;
  private metricsInterval: NodeJS.Timeout | null = null;
  private cacheCleanupInterval: NodeJS.Timeout | null = null;
  private isInitialized = false;
  private isShuttingDown = false;

  constructor() {
    if (PerformanceOptimizer.instance) {
      logger.debug('🔄 Повернення існуючого екземпляру PerformanceOptimizer');
      return PerformanceOptimizer.instance;
    }

    logger.info('🔧 Ініціалізація PerformanceOptimizer...');
    PerformanceOptimizer.instance = this;

    this.performanceData = {
      timestamp: new Date(),
      memory: process.memoryUsage(),
      cpu: { usage: 0, load: 0 },
      operations: new Map(),
      slowOperations: [],
    };

    this.startMonitoring();
    this.isInitialized = true;
    logger.info('✅ PerformanceOptimizer успішно ініціалізовано');
  }

  /**
   * Запуск моніторингу продуктивності
   */
  private startMonitoring(): void {
    try {
      logger.info('📊 Запуск моніторингу продуктивності...');

      // Garbage collection
      this.gcInterval = setInterval(() => {
        this.performGarbageCollection();
      }, PERFORMANCE_CONSTANTS.GC_INTERVAL);

      // Метрики продуктивності
      this.metricsInterval = setInterval(() => {
        this.collectMetrics();
      }, PERFORMANCE_CONSTANTS.METRICS_INTERVAL);

      // Очищення кешу
      this.cacheCleanupInterval = setInterval(() => {
        this.cleanupCaches();
      }, PERFORMANCE_CONSTANTS.CACHE_CLEANUP_INTERVAL);

      logger.info('✅ Моніторинг продуктивності запущено');
    } catch (error) {
      logger.error('❌ Помилка запуску моніторингу продуктивності:', error as LogMeta);
      throw new Error(`Помилка запуску моніторингу: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
    }
  }

  /**
   * Збір метрик продуктивності
   */
  private collectMetrics(): void {
    try {
      const memoryUsage = process.memoryUsage();
      const cpuUsage = this.getCPUUsage();

      this.performanceData = {
        timestamp: new Date(),
        memory: memoryUsage,
        cpu: cpuUsage,
        operations: new Map(this.performanceData.operations),
        slowOperations: [...this.performanceData.slowOperations],
      };

      // Збереження в історію
      this.metricsHistory.push({ ...this.performanceData });
      if (this.metricsHistory.length > PERFORMANCE_CONSTANTS.MAX_METRICS_HISTORY) {
        this.metricsHistory.shift();
      }

      // Перевірка порогів
      this.checkThresholds();

      logger.debug('📊 Метрики продуктивності зібрано', {
        memory: {
          rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`,
          heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
          heapTotal: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
        },
        cpu: {
          usage: `${cpuUsage.usage.toFixed(2)}%`,
          load: cpuUsage.load.toFixed(2),
        },
      } as LogMeta);
    } catch (error) {
      logger.error('❌ Помилка збору метрик продуктивності:', error as LogMeta);
    }
  }

  /**
   * Отримання використання CPU
   */
  private getCPUUsage(): { usage: number; load: number } {
    try {
      const startUsage = process.cpuUsage();
      const startTime = process.hrtime.bigint();

      // Коротка затримка для вимірювання
      setTimeout(() => {
        try {
          const endUsage = process.cpuUsage(startUsage);
          const endTime = process.hrtime.bigint();

          const elapsed = Number(endTime - startTime) / 1000000; // в мілісекундах
          const usage = ((endUsage.user + endUsage.system) / 1000) / elapsed * 100;

          this.performanceData.cpu.usage = Math.min(usage, 100);
        } catch (cpuError) {
          logger.error('❌ Помилка вимірювання CPU:', cpuError as LogMeta);
          this.performanceData.cpu.usage = 0;
        }
      }, PERFORMANCE_CONSTANTS.CPU_MEASUREMENT_DELAY);

      return {
        usage: this.performanceData.cpu.usage,
        load: require('os').loadavg()[0] || 0,
      };
    } catch (error) {
      logger.error('❌ Помилка отримання використання CPU:', error as LogMeta);
      return { usage: 0, load: 0 };
    }
  }

  /**
   * Перевірка порогів продуктивності
   */
  private checkThresholds(): void {
    try {
      const memoryUsage = this.performanceData.memory;
      const cpuUsage = this.performanceData.cpu;
      const heapUsedPercent = (memoryUsage.heapUsed / memoryUsage.heapTotal) * 100;

      // Перевірка пам'яті
      if (heapUsedPercent > PERFORMANCE_CONSTANTS.MEMORY_THRESHOLD) {
        logger.warn('⚠️ Високе використання пам\'яті', {
          heapUsed: `${Math.round(heapUsedPercent)}%`,
          threshold: `${PERFORMANCE_CONSTANTS.MEMORY_THRESHOLD}%`,
        } as LogMeta);
        this.optimizeMemory();
      }

      // Перевірка CPU
      if (cpuUsage.usage > PERFORMANCE_CONSTANTS.CPU_THRESHOLD) {
        logger.warn('⚠️ Високе використання CPU', {
          usage: `${cpuUsage.usage.toFixed(2)}%`,
          threshold: `${PERFORMANCE_CONSTANTS.CPU_THRESHOLD}%`,
        } as LogMeta);
        this.optimizeCPU();
      }
    } catch (error) {
      logger.error('❌ Помилка перевірки порогів продуктивності:', error as LogMeta);
    }
  }

  /**
   * Оптимізація пам'яті
   */
  private optimizeMemory(): void {
    try {
      logger.info('🧹 Оптимізація пам\'яті...');

      // Примусова garbage collection
      if (global.gc) {
        global.gc();
        logger.debug('✅ Garbage collection виконано');
      }

      // Очищення кешів
      this.cleanupCaches();

      // Очищення метрик
      if (this.metricsHistory.length > PERFORMANCE_CONSTANTS.MAX_METRICS_HISTORY / 2) {
        this.metricsHistory = this.metricsHistory.slice(-PERFORMANCE_CONSTANTS.MAX_METRICS_HISTORY / 2);
        logger.debug('✅ Історія метрик очищено');
      }

      logger.info('✅ Оптимізація пам\'яті завершено');
    } catch (error) {
      logger.error('❌ Помилка оптимізації пам\'яті:', error as LogMeta);
    }
  }

  /**
   * Оптимізація CPU
   */
  private optimizeCPU(): void {
    try {
      logger.info('⚡ Оптимізація CPU...');

      // Зменшення інтервалів моніторингу
      if (this.metricsInterval) {
        clearInterval(this.metricsInterval);
        this.metricsInterval = setInterval(() => {
          this.collectMetrics();
        }, PERFORMANCE_CONSTANTS.METRICS_INTERVAL * 2);
      }

      // Очищення повільних операцій
      this.performanceData.slowOperations = [];

      logger.info('✅ Оптимізація CPU завершено');
    } catch (error) {
      logger.error('❌ Помилка оптимізації CPU:', error as LogMeta);
    }
  }

  /**
   * Виконання garbage collection
   */
  private performGarbageCollection(): void {
    try {
      if (global.gc) {
        const beforeMemory = process.memoryUsage();
        global.gc();
        const afterMemory = process.memoryUsage();

        const freedMemory = beforeMemory.heapUsed - afterMemory.heapUsed;

        logger.debug('🧹 Garbage collection виконано', {
          freedMemory: `${Math.round(freedMemory / 1024 / 1024)}MB`,
          beforeHeap: `${Math.round(beforeMemory.heapUsed / 1024 / 1024)}MB`,
          afterHeap: `${Math.round(afterMemory.heapUsed / 1024 / 1024)}MB`,
        } as LogMeta);
      }
    } catch (error) {
      logger.error('❌ Помилка garbage collection:', error as LogMeta);
    }
  }

  /**
   * Очищення кешів
   */
  private cleanupCaches(): void {
    try {
      let totalCleaned = 0;
      const now = Date.now();
      const maxAge = 30 * 60 * 1000; // 30 хвилин

      for (const [_cacheName, cache] of this.caches.entries()) {
        let cleanedCount = 0;
        const entriesToDelete: string[] = [];

        for (const [key, entry] of cache.entries()) {
          // Видалення застарілих записів
          if (now - entry.timestamp > maxAge) {
            entriesToDelete.push(key);
            cleanedCount++;
          }
        }

        entriesToDelete.forEach(key => cache.delete(key));
        totalCleaned += cleanedCount;

        // Обмеження розміру кешу
        if (cache.size > PERFORMANCE_CONSTANTS.MAX_CACHE_SIZE) {
          const entries = Array.from(cache.entries());
          entries.sort((a, b) => a[1].lastAccess - b[1].lastAccess);

          const toDelete = entries.slice(0, cache.size - PERFORMANCE_CONSTANTS.MAX_CACHE_SIZE);
          toDelete.forEach(([key]) => cache.delete(key));
          totalCleaned += toDelete.length;
        }
      }

      if (totalCleaned > 0) {
        logger.debug(`🧹 Очищено ${totalCleaned} записів з кешів`);
      }
    } catch (error) {
      logger.error('❌ Помилка очищення кешів:', error as LogMeta);
    }
  }

  /**
   * Вимірювання часу виконання операції
   */
  public async measureOperation<T>(
    operation: () => Promise<T>,
    operationName: string,
    category: string = 'general'
  ): Promise<T> {
    const startTime = performance.now();

    try {
      const result = await operation();
      const duration = performance.now() - startTime;

      this.recordOperation(operationName, duration, category);

      return result;
    } catch (error) {
      const duration = performance.now() - startTime;
      this.recordOperation(operationName, duration, category, error);
      throw error;
    }
  }

  /**
   * Запис операції
   */
  private recordOperation(
    operationName: string,
    duration: number,
    category: string,
    error?: unknown
  ): void {
    try {
      const metric: PerformanceMetrics = {
        operation: operationName,
        duration,
        category,
        metadata: {
          error: error ? (error instanceof Error ? error.message : String(error)) : undefined,
        },
      };

      // Збереження в операції
      this.performanceData.operations.set(operationName, metric);

      // Запис повільних операцій
      if (duration > PERFORMANCE_CONSTANTS.SLOW_OPERATION_THRESHOLD) {
        this.performanceData.slowOperations.push(metric);

        // Обмеження кількості повільних операцій
        if (this.performanceData.slowOperations.length > PERFORMANCE_CONSTANTS.MAX_SLOW_OPERATIONS) {
          this.performanceData.slowOperations = this.performanceData.slowOperations.slice(-PERFORMANCE_CONSTANTS.SLOW_OPERATIONS_RETAIN);
        }

        logger.warn(`🐌 Повільна операція: ${operationName}`, {
          duration: `${duration.toFixed(2)}ms`,
          category,
          threshold: `${PERFORMANCE_CONSTANTS.SLOW_OPERATION_THRESHOLD}ms`,
        } as LogMeta);
      }

      // Логування продуктивності
      logger.performance(operationName, duration, {
        category,
        error: error ? true : false,
      } as LogMeta);
    } catch (error) {
      logger.error('❌ Помилка запису операції:', error as LogMeta);
    }
  }

  /**
   * Кешування з автоматичним очищенням
   */
  public cache<T>(
    cacheName: string,
    key: string,
    value: T,
    _ttl: number = 300000 // 5 хвилин за замовчуванням (не використовується прямо, очищення планове)
  ): void {
    try {
      if (!this.caches.has(cacheName)) {
        this.caches.set(cacheName, new Map());
      }

      const cache = this.caches.get(cacheName)!;
      const now = Date.now();

      cache.set(key, {
        value,
        timestamp: now,
        accessCount: 0,
        lastAccess: now,
      });

      logger.debug(`💾 Значення кешовано: ${cacheName}:${key}`);
    } catch (error) {
      logger.error('❌ Помилка кешування:', error as LogMeta);
    }
  }

  /**
   * Отримання з кешу
   */
  public getFromCache<T>(cacheName: string, key: string): T | null {
    try {
      const cache = this.caches.get(cacheName);
      if (!cache) return null;

      const entry = cache.get(key) as CacheEntry<T> | undefined;
      if (!entry) return null;

      // Оновлення статистики доступу
      entry.accessCount++;
      entry.lastAccess = Date.now();

      logger.debug(`📖 Значення отримано з кешу: ${cacheName}:${key}`);
      return entry.value;
    } catch (error) {
      logger.error('❌ Помилка отримання з кешу:', error as LogMeta);
      return null;
    }
  }

  /**
   * Отримання статистики продуктивності
   */
  public getPerformanceStats(): Record<string, unknown> {
    try {
      const currentMemory = this.performanceData.memory;
      const currentCPU = this.performanceData.cpu;
      const heapUsedPercent = (currentMemory.heapUsed / currentMemory.heapTotal) * 100;

      return {
        memory: {
          rss: `${Math.round(currentMemory.rss / 1024 / 1024)}MB`,
          heapUsed: `${Math.round(currentMemory.heapUsed / 1024 / 1024)}MB`,
          heapTotal: `${Math.round(currentMemory.heapTotal / 1024 / 1024)}MB`,
          heapUsedPercent: `${heapUsedPercent.toFixed(2)}%`,
        },
        cpu: {
          usage: `${currentCPU.usage.toFixed(2)}%`,
          load: currentCPU.load.toFixed(2),
        },
        operations: {
          total: this.performanceData.operations.size,
          slow: this.performanceData.slowOperations.length,
        },
        caches: {
          total: this.caches.size,
          entries: Array.from(this.caches.values()).reduce((sum, cache) => sum + cache.size, 0),
        },
        metrics: {
          historySize: this.metricsHistory.length,
          lastUpdate: this.performanceData.timestamp,
        },
      };
    } catch (error) {
      logger.error('❌ Помилка отримання статистики продуктивності:', error as LogMeta);
      return {};
    }
  }

  /**
   * Отримання системних метрик
   */
  public getSystemMetrics(): SystemMetrics {
    try {
      const memoryUsage = process.memoryUsage();
      const cpuUsage = this.performanceData.cpu;

      return {
        memory: {
          rss: memoryUsage.rss,
          heapUsed: memoryUsage.heapUsed,
          heapTotal: memoryUsage.heapTotal,
          external: memoryUsage.external,
        },
        cpu: {
          usage: cpuUsage.usage,
          load: cpuUsage.load,
        },
        uptime: process.uptime(),
        processId: process.pid,
      };
    } catch (error) {
      logger.error('❌ Помилка отримання системних метрик:', error as LogMeta);
      return {
        memory: { rss: 0, heapUsed: 0, heapTotal: 0, external: 0 },
        cpu: { usage: 0, load: 0 },
        uptime: 0,
        processId: 0,
      };
    }
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    if (this.isShuttingDown) {
      logger.warn('⚠️ PerformanceOptimizer вже зупиняється');
      return;
    }

    this.isShuttingDown = true;

    try {
      logger.info('🧹 Очищення ресурсів PerformanceOptimizer...');

      // Зупинка інтервалів
      if (this.gcInterval) {
        clearInterval(this.gcInterval);
        this.gcInterval = null;
      }

      if (this.metricsInterval) {
        clearInterval(this.metricsInterval);
        this.metricsInterval = null;
      }

      if (this.cacheCleanupInterval) {
        clearInterval(this.cacheCleanupInterval);
        this.cacheCleanupInterval = null;
      }

      // Очищення кешів
      this.caches.clear();
      this.metricsHistory = [];
      this.performanceData.operations.clear();
      this.performanceData.slowOperations = [];

      logger.info('✅ Ресурси PerformanceOptimizer очищено');
    } catch (error) {
      logger.error('❌ Помилка очищення PerformanceOptimizer:', error as LogMeta);
    } finally {
      this.isShuttingDown = false;
    }
  }

  /**
   * Перевірка стану ініціалізації
   */
  public getInitializedState(): boolean {
    return this.isInitialized;
  }

  /**
   * Перевірка стану зупинки
   */
  public getShuttingDownState(): boolean {
    return this.isShuttingDown;
  }
}

// Експорт єдиного екземпляра
export const performanceOptimizer = new PerformanceOptimizer();

// Експорт функцій для зручності
export const measureOperation = <T>(
  operation: () => Promise<T>,
  operationName: string,
  category?: string
): Promise<T> => {
  return performanceOptimizer.measureOperation(operation, operationName, category);
};

export const cache = <T>(
  cacheName: string,
  key: string,
  value: T,
  ttl?: number
): void => {
  performanceOptimizer.cache(cacheName, key, value, ttl);
};

export const getFromCache = <T>(cacheName: string, key: string): T | null => {
  return performanceOptimizer.getFromCache<T>(cacheName, key);
};

export const getPerformanceStats = () => performanceOptimizer.getPerformanceStats();
export const getSystemMetrics = () => performanceOptimizer.getSystemMetrics();
export const cleanupPerformanceOptimizer = () => performanceOptimizer.cleanup(); 