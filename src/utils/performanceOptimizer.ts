/**
 * 🚀 Performance Optimizer Module
 * Оптимізація продуктивності Discord AI Assistant Bot
 * TypeScript версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 *
 * Функції:
 * - Кешування результатів
 * - Оптимізація запитів
 * - Моніторинг продуктивності
 * - Автоматична оптимізація
 */

import { performance } from 'perf_hooks';
import fs from 'fs-extra';
import path from 'path';
import logger from './logger';

// Константи для конфігурації оптимізації
const OPTIMIZER_CONFIG = {
  MAX_RESPONSE_TIMES: 1000,
  MAX_MEMORY_RECORDS: 100,
  MEMORY_CLEANUP_THRESHOLD: 500 * 1024 * 1024, // 500MB
  METRICS_SAVE_INTERVAL: 5 * 60 * 1000, // 5 хвилин
  MEMORY_MONITOR_INTERVAL: 30 * 1000, // 30 секунд
  CACHE_CLEANUP_INTERVAL: 10 * 60 * 1000, // 10 хвилин
  MAX_CACHE_SIZE: 1000,
  MAX_QUERY_CACHE_SIZE: 500,
  PERFORMANCE_THRESHOLD: 5000, // 5 секунд
  CACHE_HIT_THRESHOLD: 50, // 50%
  MEMORY_THRESHOLD: 200 * 1024 * 1024, // 200MB
} as const;

interface PerformanceMetrics {
  responseTimes: Array<{
    context: string;
    executionTime: number;
    memoryDelta?: number;
    timestamp: number;
    error?: string;
    userId?: string;
    operation?: string;
  }>;
  cacheHits: number;
  cacheMisses: number;
  queryOptimizations: number;
  memoryUsage: Array<{
    timestamp: number;
    rss: number;
    heapUsed: number;
    heapTotal: number;
    external: number;
  }>;
  errors: number;
  warnings: number;
  autoOptimizations: number;
}

interface OptimizationRules {
  maxBatchSize: number;
  cacheTTL: number;
  retryAttempts: number;
  timeout: number;
  maxConcurrent?: number;
  maxFileSize?: number;
}

interface CacheEntry {
  data: any;
  timestamp: number;
  ttl: number;
  accessCount: number;
  lastAccess: number;
  size: number;
}

interface PerformanceStats {
  averageResponseTime: number;
  maxResponseTime: number;
  minResponseTime: number;
  cacheHitRate: number;
  memoryUsage: {
    current: NodeJS.MemoryUsage;
    average: number;
    max: number;
    trend: 'increasing' | 'decreasing' | 'stable';
  };
  optimizations: number;
  cacheSize: number;
  queryCacheSize: number;
  errors: number;
  warnings: number;
  autoOptimizations: number;
}

interface OptimizationRecommendation {
  type: 'performance' | 'cache' | 'memory' | 'optimization' | 'error';
  priority: 'high' | 'medium' | 'low';
  message: string;
  action: string;
  impact: 'high' | 'medium' | 'low';
  estimatedImprovement: string;
}

interface PerformanceContext {
  userId?: string;
  operation?: string;
  service?: string;
  priority?: 'high' | 'medium' | 'low';
}

class PerformanceOptimizer {
  private metrics: PerformanceMetrics;
  private optimizationRules: Map<string, OptimizationRules>;
  private cache: Map<string, CacheEntry>;
  private queryCache: Map<string, CacheEntry>;
  private isMonitoring: boolean = false;
  private cleanupInterval: NodeJS.Timeout | null = null;
  private metricsInterval: NodeJS.Timeout | null = null;
  private memoryInterval: NodeJS.Timeout | null = null;

  constructor() {
    this.metrics = {
      responseTimes: [],
      cacheHits: 0,
      cacheMisses: 0,
      queryOptimizations: 0,
      memoryUsage: [],
      errors: 0,
      warnings: 0,
      autoOptimizations: 0,
    };

    this.optimizationRules = new Map();
    this.cache = new Map();
    this.queryCache = new Map();

    this.loadOptimizationRules();
    this.startMonitoring();
    
    logger.info('Performance Optimizer ініціалізовано');
  }

  /**
   * Завантаження правил оптимізації з детальним логуванням
   */
  private loadOptimizationRules(): void {
    try {
      // Правила для оптимізації запитів Google Sheets
      this.optimizationRules.set('sheets_query', {
        maxBatchSize: 1000,
        cacheTTL: 300000, // 5 хвилин
        retryAttempts: 3,
        timeout: 10000,
      });

      // Правила для AI запитів
      this.optimizationRules.set('ai_query', {
        maxConcurrent: 5,
        cacheTTL: 600000, // 10 хвилин
        retryAttempts: 2,
        timeout: 30000,
      });

      // Правила для файлових операцій
      this.optimizationRules.set('file_operation', {
        maxFileSize: 10 * 1024 * 1024, // 10MB
        cacheTTL: 1800000, // 30 хвилин
        retryAttempts: 3,
        timeout: 60000,
      });

      // Правила для команд Discord
      this.optimizationRules.set('discord_command', {
        maxBatchSize: 100,
        cacheTTL: 60000, // 1 хвилина
        retryAttempts: 1,
        timeout: 5000,
      });

      logger.info('Правила оптимізації завантажено', {
        rulesCount: this.optimizationRules.size,
        rules: Array.from(this.optimizationRules.keys()),
      });
    } catch (error) {
      logger.error('Помилка завантаження правил оптимізації:', error);
      throw new Error(`Помилка ініціалізації оптимізатора: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
    }
  }

  /**
   * Запуск моніторингу продуктивності з детальним логуванням
   */
  private startMonitoring(): void {
    if (this.isMonitoring) {
      logger.warn('Моніторинг вже запущено');
      return;
    }

    try {
      // Моніторинг пам'яті
      this.memoryInterval = setInterval(() => {
        this.monitorMemory();
      }, OPTIMIZER_CONFIG.MEMORY_MONITOR_INTERVAL);

      // Збереження метрик
      this.metricsInterval = setInterval(() => {
        this.saveMetrics();
      }, OPTIMIZER_CONFIG.METRICS_SAVE_INTERVAL);

      // Очищення кешу
      this.cleanupInterval = setInterval(() => {
        this.cleanupExpiredCache();
      }, OPTIMIZER_CONFIG.CACHE_CLEANUP_INTERVAL);

      this.isMonitoring = true;
      logger.info('Моніторинг продуктивності запущено');
    } catch (error) {
      logger.error('Помилка запуску моніторингу:', error);
      throw error;
    }
  }

  /**
   * Моніторинг пам'яті з детальним логуванням
   */
  private monitorMemory(): void {
    try {
      const memUsage = process.memoryUsage();
      this.metrics.memoryUsage.push({
        timestamp: Date.now(),
        rss: memUsage.rss,
        heapUsed: memUsage.heapUsed,
        heapTotal: memUsage.heapTotal,
        external: memUsage.external,
      });

      // Зберігаємо тільки останні записи
      if (this.metrics.memoryUsage.length > OPTIMIZER_CONFIG.MAX_MEMORY_RECORDS) {
        this.metrics.memoryUsage.shift();
      }

      // Автоматична очистка кешу при високому використанні пам'яті
      if (memUsage.heapUsed > OPTIMIZER_CONFIG.MEMORY_CLEANUP_THRESHOLD) {
        logger.warn('Високе використання пам\'яті, очищення кешу', {
          heapUsed: `${Math.round(memUsage.heapUsed / 1024 / 1024)}MB`,
          threshold: `${Math.round(OPTIMIZER_CONFIG.MEMORY_CLEANUP_THRESHOLD / 1024 / 1024)}MB`,
        });
        this.cleanupCache();
      }

      // Логування попереджень при високому використанні
      if (memUsage.heapUsed > OPTIMIZER_CONFIG.MEMORY_THRESHOLD) {
        this.metrics.warnings++;
        logger.warn('Високе використання пам\'яті', {
          heapUsed: `${Math.round(memUsage.heapUsed / 1024 / 1024)}MB`,
          heapTotal: `${Math.round(memUsage.heapTotal / 1024 / 1024)}MB`,
          external: `${Math.round(memUsage.external / 1024 / 1024)}MB`,
        });
      }
    } catch (error) {
      logger.error('Помилка моніторингу пам\'яті:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Вимірювання часу виконання функції з детальним логуванням
   */
  async measurePerformance<T>(
    fn: () => Promise<T>, 
    context = 'unknown',
    perfContext?: PerformanceContext
  ): Promise<T> {
    const startTime = performance.now();
    const startMemory = process.memoryUsage();

    try {
      const result = await fn();
      const endTime = performance.now();
      const endMemory = process.memoryUsage();

      const executionTime = endTime - startTime;
      const memoryDelta = endMemory.heapUsed - startMemory.heapUsed;

      this.metrics.responseTimes.push({
        context,
        executionTime,
        memoryDelta,
        timestamp: Date.now(),
        userId: perfContext?.userId,
        operation: perfContext?.operation,
      });

      // Зберігаємо тільки останні записи
      if (this.metrics.responseTimes.length > OPTIMIZER_CONFIG.MAX_RESPONSE_TIMES) {
        this.metrics.responseTimes.shift();
      }

      // Логування повільних операцій
      if (executionTime > OPTIMIZER_CONFIG.PERFORMANCE_THRESHOLD) {
        logger.warn('Повільна операція виявлена', {
          context,
          executionTime: `${executionTime.toFixed(2)}ms`,
          threshold: `${OPTIMIZER_CONFIG.PERFORMANCE_THRESHOLD}ms`,
          userId: perfContext?.userId,
          operation: perfContext?.operation,
        });
      }

      return result;
    } catch (error: any) {
      const endTime = performance.now();
      this.metrics.errors++;
      
      this.metrics.responseTimes.push({
        context,
        executionTime: endTime - startTime,
        error: error.message,
        timestamp: Date.now(),
        userId: perfContext?.userId,
        operation: perfContext?.operation,
      });

      logger.error('Помилка виконання операції', {
        context,
        executionTime: `${(endTime - startTime).toFixed(2)}ms`,
        error: error.message,
        userId: perfContext?.userId,
        operation: perfContext?.operation,
      });

      throw error;
    }
  }

  /**
   * Кешування результатів з TTL та детальним логуванням
   */
  async getCachedResult<T>(key: string, ttl = 300000): Promise<T | null> {
    try {
      const cached = this.cache.get(key);

      if (cached && Date.now() - cached.timestamp < ttl) {
        this.metrics.cacheHits++;
        cached.accessCount++;
        cached.lastAccess = Date.now();
        
        logger.debug('Кеш-хіт', {
          key: key.substring(0, 50) + '...',
          accessCount: cached.accessCount,
          age: `${Math.round((Date.now() - cached.timestamp) / 1000)}s`,
        });
        
        return cached.data as T;
      }

      this.metrics.cacheMisses++;
      logger.debug('Кеш-міс', { key: key.substring(0, 50) + '...' });
      return null;
    } catch (error) {
      logger.error('Помилка отримання кешованого результату:', error);
      this.metrics.errors++;
      return null;
    }
  }

  /**
   * Збереження результату в кеш з детальним логуванням
   */
  setCachedResult<T>(key: string, data: T, ttl = 300000): void {
    try {
      const dataSize = JSON.stringify(data).length;
      
      this.cache.set(key, {
        data,
        timestamp: Date.now(),
        ttl,
        accessCount: 0,
        lastAccess: Date.now(),
        size: dataSize,
      });

      // Обмеження розміру кешу
      if (this.cache.size > OPTIMIZER_CONFIG.MAX_CACHE_SIZE) {
        this.evictLeastUsed();
      }

      logger.debug('Результат кешовано', {
        key: key.substring(0, 50) + '...',
        size: `${Math.round(dataSize / 1024)}KB`,
        ttl: `${Math.round(ttl / 1000)}s`,
      });
    } catch (error) {
      logger.error('Помилка кешування результату:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Видалення найменш використовуваних записів
   */
  private evictLeastUsed(): void {
    try {
      const entries = Array.from(this.cache.entries());
      entries.sort((a, b) => {
        // Сортуємо за кількістю доступів, потім за часом останнього доступу
        if (a[1].accessCount !== b[1].accessCount) {
          return a[1].accessCount - b[1].accessCount;
        }
        return a[1].lastAccess - b[1].lastAccess;
      });

      const toRemove = Math.floor(entries.length * 0.2); // Видаляємо 20%
      for (let i = 0; i < toRemove; i++) {
        this.cache.delete(entries[i][0]);
      }

      logger.info('Видалено найменш використовувані записи кешу', {
        removed: toRemove,
        remaining: this.cache.size,
      });
    } catch (error) {
      logger.error('Помилка видалення записів кешу:', error);
    }
  }

  /**
   * Очищення застарілого кешу
   */
  private cleanupExpiredCache(): void {
    try {
      const now = Date.now();
      let cleanedMain = 0;
      let cleanedQuery = 0;

      // Очищення основного кешу
      for (const [key, value] of this.cache.entries()) {
        if (now - value.timestamp > value.ttl) {
          this.cache.delete(key);
          cleanedMain++;
        }
      }

      // Очищення кешу запитів
      for (const [key, value] of this.queryCache.entries()) {
        if (now - value.timestamp > value.ttl) {
          this.queryCache.delete(key);
          cleanedQuery++;
        }
      }

      if (cleanedMain > 0 || cleanedQuery > 0) {
        logger.debug('Очищено застарілий кеш', {
          mainCache: cleanedMain,
          queryCache: cleanedQuery,
        });
      }
    } catch (error) {
      logger.error('Помилка очищення застарілого кешу:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Очищення кешу при високому використанні пам'яті
   */
  private cleanupCache(): void {
    try {
      const cacheSize = this.cache.size;
      const queryCacheSize = this.queryCache.size;

      // Видаляємо 50% найстаріших записів
      const keysToDelete = Array.from(this.cache.keys()).slice(0, Math.floor(cacheSize / 2));
      keysToDelete.forEach(key => this.cache.delete(key));

      const queryKeysToDelete = Array.from(this.queryCache.keys()).slice(
        0,
        Math.floor(queryCacheSize / 2)
      );
      queryKeysToDelete.forEach(key => this.queryCache.delete(key));

      logger.info('Очищено кеш через високе використання пам\'яті', {
        mainCacheRemoved: keysToDelete.length,
        queryCacheRemoved: queryKeysToDelete.length,
        mainCacheRemaining: this.cache.size,
        queryCacheRemaining: this.queryCache.size,
      });
    } catch (error) {
      logger.error('Помилка очищення кешу:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Оптимізація запитів Google Sheets з детальним логуванням
   */
  optimizeSheetsQuery(query: any, data: any[]): any[] {
    try {
      const rules = this.optimizationRules.get('sheets_query');
      if (!rules) {
        logger.warn('Правила оптимізації sheets_query не знайдено');
        return data;
      }

      // Кешування запиту
      const queryKey = this.generateQueryKey(query);
      const cached = this.queryCache.get(queryKey);

      if (cached && Date.now() - cached.timestamp < rules.cacheTTL) {
        this.metrics.cacheHits++;
        cached.accessCount++;
        cached.lastAccess = Date.now();
        
        logger.debug('Використано кешований результат запиту', {
          queryKey: queryKey.substring(0, 50) + '...',
          accessCount: cached.accessCount,
        });
        
        return cached.data as any[];
      }

      // Оптимізація даних
      let optimizedData = data;

      // Обмеження розміру відповіді
      if (Array.isArray(data) && data.length > rules.maxBatchSize) {
        optimizedData = data.slice(0, rules.maxBatchSize);
        logger.warn('Обмежено розмір відповіді запиту', {
          originalSize: data.length,
          limitedSize: rules.maxBatchSize,
        });
      }

      // Кешування результату
      this.queryCache.set(queryKey, {
        data: optimizedData,
        timestamp: Date.now(),
        ttl: rules.cacheTTL,
        accessCount: 1,
        lastAccess: Date.now(),
        size: JSON.stringify(optimizedData).length,
      });

      // Обмеження розміру кешу запитів
      if (this.queryCache.size > OPTIMIZER_CONFIG.MAX_QUERY_CACHE_SIZE) {
        const oldestKey = this.queryCache.keys().next().value;
        this.queryCache.delete(oldestKey);
      }

      this.metrics.queryOptimizations++;
      
      logger.debug('Запит оптимізовано', {
        originalSize: data.length,
        optimizedSize: optimizedData.length,
        cacheSize: this.queryCache.size,
      });
      
      return optimizedData;
    } catch (error) {
      logger.error('Помилка оптимізації запиту:', error);
      this.metrics.errors++;
      return data;
    }
  }

  /**
   * Генерація ключа для кешування запиту
   */
  private generateQueryKey(query: any): string {
    try {
      if (typeof query === 'string') {
        return `query_${Buffer.from(query).toString('base64')}`;
      }
      return `query_${Buffer.from(JSON.stringify(query)).toString('base64')}`;
    } catch (error) {
      logger.error('Помилка генерації ключа запиту:', error);
      return `query_${Date.now()}_${Math.random()}`;
    }
  }

  /**
   * Оптимізація AI запитів з детальним логуванням
   */
  optimizeAIQuery(query: string, context = ''): string | null {
    try {
      const rules = this.optimizationRules.get('ai_query');
      if (!rules) {
        logger.warn('Правила оптимізації ai_query не знайдено');
        return null;
      }

      // Обмеження довжини запиту
      const maxLength = 2000;
      if (query.length > maxLength) {
        const truncatedQuery = query.substring(0, maxLength) + '...';
        logger.warn('Обрізано довгий AI запит', {
          originalLength: query.length,
          truncatedLength: maxLength,
        });
        query = truncatedQuery;
      }

      // Кешування схожих запитів
      const queryKey = this.generateQueryKey(query + context);
      const cached = this.cache.get(queryKey);

      if (cached && Date.now() - cached.timestamp < rules.cacheTTL) {
        this.metrics.cacheHits++;
        cached.accessCount++;
        cached.lastAccess = Date.now();
        
        logger.debug('Використано кешований AI запит', {
          queryKey: queryKey.substring(0, 50) + '...',
          accessCount: cached.accessCount,
        });
        
        return cached.data as string;
      }

      return null; // Повертаємо null, щоб виконати запит
    } catch (error) {
      logger.error('Помилка оптимізації AI запиту:', error);
      this.metrics.errors++;
      return null;
    }
  }

  /**
   * Оптимізація файлових операцій з детальним логуванням
   */
  optimizeFileOperation(filePath: string, operation: string): any | null {
    try {
      const rules = this.optimizationRules.get('file_operation');
      if (!rules) {
        logger.warn('Правила оптимізації file_operation не знайдено');
        return null;
      }

      // Перевірка розміру файлу
      try {
        const stats = fs.statSync(filePath);
        if (stats.size > (rules.maxFileSize || 10 * 1024 * 1024)) {
          logger.warn('Файл занадто великий для операції', {
            filePath,
            fileSize: `${Math.round(stats.size / 1024 / 1024)}MB`,
            maxSize: `${Math.round((rules.maxFileSize || 10 * 1024 * 1024) / 1024 / 1024)}MB`,
            operation,
          });
          throw new Error(`Файл занадто великий: ${stats.size} байт`);
        }
      } catch (error) {
        if (error instanceof Error && error.message.includes('занадто великий')) {
          throw error;
        }
        logger.warn(`Помилка перевірки файлу: ${(error as Error).message}`, { filePath });
      }

      // Кешування метаданих файлу
      const fileKey = `file_${Buffer.from(filePath).toString('base64')}`;
      const cached = this.cache.get(fileKey);

      if (cached && Date.now() - cached.timestamp < (rules.cacheTTL || 1800000)) {
        this.metrics.cacheHits++;
        cached.accessCount++;
        cached.lastAccess = Date.now();
        
        logger.debug('Використано кешовані метадані файлу', {
          filePath,
          accessCount: cached.accessCount,
        });
        
        return cached.data;
      }

      return null;
    } catch (error) {
      logger.error('Помилка оптимізації файлової операції:', error);
      this.metrics.errors++;
      return null;
    }
  }

  /**
   * Отримання статистики продуктивності з детальним аналізом
   */
  getPerformanceStats(): PerformanceStats {
    try {
      const responseTimes = this.metrics.responseTimes;
      const memoryUsage = this.metrics.memoryUsage;

      if (responseTimes.length === 0) {
        return {
          averageResponseTime: 0,
          maxResponseTime: 0,
          minResponseTime: 0,
          cacheHitRate: 0,
          memoryUsage: {
            current: process.memoryUsage(),
            average: 0,
            max: 0,
            trend: 'stable',
          },
          optimizations: this.metrics.queryOptimizations,
          cacheSize: this.cache.size,
          queryCacheSize: this.queryCache.size,
          errors: this.metrics.errors,
          warnings: this.metrics.warnings,
          autoOptimizations: this.metrics.autoOptimizations,
        };
      }

      const times = responseTimes.map(r => r.executionTime);
      const averageResponseTime = times.reduce((a, b) => a + b, 0) / times.length;
      const maxResponseTime = Math.max(...times);
      const minResponseTime = Math.min(...times);

      const totalCacheAccess = this.metrics.cacheHits + this.metrics.cacheMisses;
      const cacheHitRate =
        totalCacheAccess > 0 ? (this.metrics.cacheHits / totalCacheAccess) * 100 : 0;

      const memoryValues = memoryUsage.map(m => m.heapUsed);
      const averageMemory =
        memoryValues.length > 0 ? memoryValues.reduce((a, b) => a + b, 0) / memoryValues.length : 0;
      const maxMemory = memoryValues.length > 0 ? Math.max(...memoryValues) : 0;

      // Аналіз тренду пам'яті
      let memoryTrend: 'increasing' | 'decreasing' | 'stable' = 'stable';
      if (memoryValues.length >= 3) {
        const recent = memoryValues.slice(-3);
        if (recent[2] > recent[1] && recent[1] > recent[0]) {
          memoryTrend = 'increasing';
        } else if (recent[2] < recent[1] && recent[1] < recent[0]) {
          memoryTrend = 'decreasing';
        }
      }

      return {
        averageResponseTime: Math.round(averageResponseTime * 100) / 100,
        maxResponseTime: Math.round(maxResponseTime * 100) / 100,
        minResponseTime: Math.round(minResponseTime * 100) / 100,
        cacheHitRate: Math.round(cacheHitRate * 100) / 100,
        memoryUsage: {
          current: process.memoryUsage(),
          average: Math.round((averageMemory / 1024 / 1024) * 100) / 100, // MB
          max: Math.round((maxMemory / 1024 / 1024) * 100) / 100, // MB
          trend: memoryTrend,
        },
        optimizations: this.metrics.queryOptimizations,
        cacheSize: this.cache.size,
        queryCacheSize: this.queryCache.size,
        errors: this.metrics.errors,
        warnings: this.metrics.warnings,
        autoOptimizations: this.metrics.autoOptimizations,
      };
    } catch (error) {
      logger.error('Помилка отримання статистики продуктивності:', error);
      return {
        averageResponseTime: 0,
        maxResponseTime: 0,
        minResponseTime: 0,
        cacheHitRate: 0,
        memoryUsage: {
          current: process.memoryUsage(),
          average: 0,
          max: 0,
          trend: 'stable',
        },
        optimizations: 0,
        cacheSize: 0,
        queryCacheSize: 0,
        errors: this.metrics.errors,
        warnings: this.metrics.warnings,
        autoOptimizations: this.metrics.autoOptimizations,
      };
    }
  }

  /**
   * Збереження метрик в файл з детальним логуванням
   */
  private async saveMetrics(): Promise<void> {
    try {
      const metricsDir = path.join(process.cwd(), 'data', 'metrics');
      await fs.ensureDir(metricsDir);

      const metricsFile = path.join(metricsDir, `performance_${Date.now()}.json`);
      const stats = this.getPerformanceStats();

      const metricsData = {
        timestamp: Date.now(),
        stats,
        rawMetrics: this.metrics,
        cacheInfo: {
          mainCacheSize: this.cache.size,
          queryCacheSize: this.queryCache.size,
          totalCacheEntries: this.cache.size + this.queryCache.size,
        },
        optimizationRules: Array.from(this.optimizationRules.entries()),
      };

      await fs.writeJson(metricsFile, metricsData, { spaces: 2 });

      logger.info('Метрики продуктивності збережено', {
        file: metricsFile,
        stats: {
          averageResponseTime: stats.averageResponseTime,
          cacheHitRate: stats.cacheHitRate,
          memoryUsage: `${stats.memoryUsage.average}MB`,
          errors: stats.errors,
        },
      });
    } catch (error) {
      logger.error('Помилка збереження метрик:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Отримання рекомендацій по оптимізації з детальним аналізом
   */
  getOptimizationRecommendations(): OptimizationRecommendation[] {
    try {
      const stats = this.getPerformanceStats();
      const recommendations: OptimizationRecommendation[] = [];

      // Рекомендації по часу відповіді
      if (stats.averageResponseTime > OPTIMIZER_CONFIG.PERFORMANCE_THRESHOLD) {
        recommendations.push({
          type: 'performance',
          priority: 'high',
          message: `Середній час відповіді занадто високий (${stats.averageResponseTime}ms > ${OPTIMIZER_CONFIG.PERFORMANCE_THRESHOLD}ms). Рекомендується оптимізувати запити до API.`,
          action: 'Оптимізувати запити, додати кешування, зменшити розмір відповідей',
          impact: 'high',
          estimatedImprovement: 'Зменшення часу відповіді на 30-50%',
        });
      }

      // Рекомендації по кешу
      if (stats.cacheHitRate < OPTIMIZER_CONFIG.CACHE_HIT_THRESHOLD) {
        recommendations.push({
          type: 'cache',
          priority: 'medium',
          message: `Низький відсоток попадань в кеш (${stats.cacheHitRate}% < ${OPTIMIZER_CONFIG.CACHE_HIT_THRESHOLD}%). Рекомендується збільшити TTL кешу.`,
          action: 'Збільшити TTL кешу, додати більше кешованих запитів, оптимізувати ключі кешу',
          impact: 'medium',
          estimatedImprovement: 'Збільшення cache hit rate на 20-30%',
        });
      }

      // Рекомендації по пам'яті
      if (stats.memoryUsage.average > OPTIMIZER_CONFIG.MEMORY_THRESHOLD / 1024 / 1024) {
        recommendations.push({
          type: 'memory',
          priority: 'high',
          message: `Високе використання пам'яті (${stats.memoryUsage.average}MB > ${OPTIMIZER_CONFIG.MEMORY_THRESHOLD / 1024 / 1024}MB). Рекомендується оптимізувати кеш.`,
          action: 'Зменшити розмір кешу, додати автоматичну очистку, оптимізувати структури даних',
          impact: 'high',
          estimatedImprovement: 'Зменшення використання пам\'яті на 25-40%',
        });
      }

      // Рекомендації по помилках
      if (stats.errors > 10) {
        recommendations.push({
          type: 'error',
          priority: 'high',
          message: `Високий рівень помилок (${stats.errors}). Рекомендується додати обробку помилок та retry логіку.`,
          action: 'Додати retry логіку, покращити обробку помилок, додати fallback механізми',
          impact: 'high',
          estimatedImprovement: 'Зменшення кількості помилок на 50-80%',
        });
      }

      // Рекомендації по оптимізації
      if (stats.optimizations < 10) {
        recommendations.push({
          type: 'optimization',
          priority: 'low',
          message: 'Мало оптимізацій запитів. Рекомендується додати більше правил оптимізації.',
          action: 'Додати правила оптимізації для різних типів запитів, налаштувати автоматичну оптимізацію',
          impact: 'low',
          estimatedImprovement: 'Покращення загальної продуктивності на 10-15%',
        });
      }

      logger.debug('Згенеровано рекомендації оптимізації', {
        count: recommendations.length,
        highPriority: recommendations.filter(r => r.priority === 'high').length,
        mediumPriority: recommendations.filter(r => r.priority === 'medium').length,
        lowPriority: recommendations.filter(r => r.priority === 'low').length,
      });

      return recommendations;
    } catch (error) {
      logger.error('Помилка генерації рекомендацій оптимізації:', error);
      return [];
    }
  }

  /**
   * Автоматична оптимізація на основі метрик з детальним логуванням
   */
  autoOptimize(): void {
    try {
      const recommendations = this.getOptimizationRecommendations();
      let optimizationsApplied = 0;

      for (const rec of recommendations) {
        if (rec.priority === 'high') {
          logger.info(`🚀 Автоматична оптимізація: ${rec.message}`);

          switch (rec.type) {
            case 'memory':
              this.cleanupCache();
              optimizationsApplied++;
              break;
            case 'cache':
              // Збільшуємо TTL для популярних запитів
              const sheetsRules = this.optimizationRules.get('sheets_query');
              if (sheetsRules) {
                sheetsRules.cacheTTL = Math.floor(sheetsRules.cacheTTL * 1.5);
                logger.info('Збільшено TTL кешу для sheets_query', {
                  newTTL: `${Math.round(sheetsRules.cacheTTL / 1000)}s`,
                });
              }
              optimizationsApplied++;
              break;
            case 'performance':
              // Зменшуємо розмір batch для запитів
              const sheetsRules2 = this.optimizationRules.get('sheets_query');
              if (sheetsRules2) {
                sheetsRules2.maxBatchSize = Math.floor(sheetsRules2.maxBatchSize * 0.8);
                logger.info('Зменшено розмір batch для sheets_query', {
                  newBatchSize: sheetsRules2.maxBatchSize,
                });
              }
              optimizationsApplied++;
              break;
            case 'error':
              // Збільшуємо кількість спроб
              const aiRules = this.optimizationRules.get('ai_query');
              if (aiRules) {
                aiRules.retryAttempts = Math.min(aiRules.retryAttempts + 1, 5);
                logger.info('Збільшено кількість спроб для ai_query', {
                  newRetryAttempts: aiRules.retryAttempts,
                });
              }
              optimizationsApplied++;
              break;
          }
        }
      }

      this.metrics.autoOptimizations += optimizationsApplied;
      
      if (optimizationsApplied > 0) {
        logger.info('Автоматична оптимізація завершена', {
          optimizationsApplied,
          totalAutoOptimizations: this.metrics.autoOptimizations,
        });
      }
    } catch (error) {
      logger.error('Помилка автоматичної оптимізації:', error);
      this.metrics.errors++;
    }
  }

  /**
   * Скидання метрик з детальним логуванням
   */
  resetMetrics(): void {
    try {
      this.metrics = {
        responseTimes: [],
        cacheHits: 0,
        cacheMisses: 0,
        queryOptimizations: 0,
        memoryUsage: [],
        errors: 0,
        warnings: 0,
        autoOptimizations: 0,
      };

      this.cache.clear();
      this.queryCache.clear();

      logger.info('Метрики продуктивності скинуто', {
        cacheCleared: true,
        queryCacheCleared: true,
      });
    } catch (error) {
      logger.error('Помилка скидання метрик:', error);
    }
  }

  /**
   * Зупинка моніторингу
   */
  stopMonitoring(): void {
    try {
      if (this.memoryInterval) {
        clearInterval(this.memoryInterval);
        this.memoryInterval = null;
      }
      if (this.metricsInterval) {
        clearInterval(this.metricsInterval);
        this.metricsInterval = null;
      }
      if (this.cleanupInterval) {
        clearInterval(this.cleanupInterval);
        this.cleanupInterval = null;
      }

      this.isMonitoring = false;
      logger.info('Моніторинг продуктивності зупинено');
    } catch (error) {
      logger.error('Помилка зупинки моніторингу:', error);
    }
  }

  /**
   * Отримання детальної інформації про стан оптимізатора
   */
  getOptimizerStatus(): any {
    return {
      isMonitoring: this.isMonitoring,
      rulesCount: this.optimizationRules.size,
      cacheSize: this.cache.size,
      queryCacheSize: this.queryCache.size,
      metrics: {
        totalResponseTimes: this.metrics.responseTimes.length,
        totalMemoryRecords: this.metrics.memoryUsage.length,
        errors: this.metrics.errors,
        warnings: this.metrics.warnings,
      },
      intervals: {
        memory: !!this.memoryInterval,
        metrics: !!this.metricsInterval,
        cleanup: !!this.cleanupInterval,
      },
    };
  }
}

// Експорт синглтона
export default new PerformanceOptimizer(); 