/**
 * 🚀 Performance Optimizer Module
 * Оптимізація продуктивності Discord AI Assistant Bot
 *
 * Функції:
 * - Кешування результатів
 * - Оптимізація запитів
 * - Моніторинг продуктивності
 * - Автоматична оптимізація
 */

const { performance } = require('perf_hooks');
const fs = require('fs-extra');
const path = require('path');

class PerformanceOptimizer {
  constructor() {
    this.metrics = {
      responseTimes: [],
      cacheHits: 0,
      cacheMisses: 0,
      queryOptimizations: 0,
      memoryUsage: [],
    };

    this.optimizationRules = new Map();
    this.cache = new Map();
    this.queryCache = new Map();

    this.loadOptimizationRules();
    this.startMonitoring();
  }

  /**
   * Завантаження правил оптимізації
   */
  loadOptimizationRules() {
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
  }

  /**
   * Запуск моніторингу продуктивності
   */
  startMonitoring() {
    // Моніторинг пам'яті кожні 30 секунд
    setInterval(() => {
      const memUsage = process.memoryUsage();
      this.metrics.memoryUsage.push({
        timestamp: Date.now(),
        rss: memUsage.rss,
        heapUsed: memUsage.heapUsed,
        heapTotal: memUsage.heapTotal,
        external: memUsage.external,
      });

      // Зберігаємо тільки останні 100 записів
      if (this.metrics.memoryUsage.length > 100) {
        this.metrics.memoryUsage.shift();
      }

      // Автоматична очистка кешу при високому використанні пам'яті
      if (memUsage.heapUsed > 500 * 1024 * 1024) {
        // 500MB
        this.cleanupCache();
      }
    }, 30000);

    // Збереження метрик кожні 5 хвилин
    setInterval(() => {
      this.saveMetrics();
    }, 300000);
  }

  /**
   * Вимірювання часу виконання функції
   */
  async measurePerformance(fn, context = 'unknown') {
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
      });

      // Зберігаємо тільки останні 1000 записів
      if (this.metrics.responseTimes.length > 1000) {
        this.metrics.responseTimes.shift();
      }

      return result;
    } catch (error) {
      const endTime = performance.now();
      this.metrics.responseTimes.push({
        context,
        executionTime: endTime - startTime,
        error: error.message,
        timestamp: Date.now(),
      });
      throw error;
    }
  }

  /**
   * Кешування результатів з TTL
   */
  async getCachedResult(key, ttl = 300000) {
    const cached = this.cache.get(key);

    if (cached && Date.now() - cached.timestamp < ttl) {
      this.metrics.cacheHits++;
      return cached.data;
    }

    this.metrics.cacheMisses++;
    return null;
  }

  /**
   * Збереження результату в кеш
   */
  setCachedResult(key, data, ttl = 300000) {
    this.cache.set(key, {
      data,
      timestamp: Date.now(),
      ttl,
    });

    // Автоматична очистка застарілих записів
    this.cleanupExpiredCache();
  }

  /**
   * Очищення застарілого кешу
   */
  cleanupExpiredCache() {
    const now = Date.now();
    for (const [key, value] of this.cache.entries()) {
      if (now - value.timestamp > value.ttl) {
        this.cache.delete(key);
      }
    }
  }

  /**
   * Очищення кешу при високому використанні пам'яті
   */
  cleanupCache() {
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

    console.log(
      `🧹 Очищено кеш: ${keysToDelete.length} записів основного кешу, ${queryKeysToDelete.length} записів кешу запитів`
    );
  }

  /**
   * Оптимізація запитів Google Sheets
   */
  optimizeSheetsQuery(query, data) {
    const rules = this.optimizationRules.get('sheets_query');

    // Кешування запиту
    const queryKey = this.generateQueryKey(query);
    const cached = this.queryCache.get(queryKey);

    if (cached && Date.now() - cached.timestamp < rules.cacheTTL) {
      this.metrics.cacheHits++;
      return cached.data;
    }

    // Оптимізація даних
    let optimizedData = data;

    // Обмеження розміру відповіді
    if (Array.isArray(data) && data.length > rules.maxBatchSize) {
      optimizedData = data.slice(0, rules.maxBatchSize);
    }

    // Кешування результату
    this.queryCache.set(queryKey, {
      data: optimizedData,
      timestamp: Date.now(),
    });

    this.metrics.queryOptimizations++;
    return optimizedData;
  }

  /**
   * Генерація ключа для кешування запиту
   */
  generateQueryKey(query) {
    if (typeof query === 'string') {
      return `query_${Buffer.from(query).toString('base64')}`;
    }
    return `query_${Buffer.from(JSON.stringify(query)).toString('base64')}`;
  }

  /**
   * Оптимізація AI запитів
   */
  optimizeAIQuery(query, context = '') {
    const rules = this.optimizationRules.get('ai_query');

    // Обмеження довжини запиту
    const maxLength = 2000;
    if (query.length > maxLength) {
      query = query.substring(0, maxLength) + '...';
    }

    // Кешування схожих запитів
    const queryKey = this.generateQueryKey(query + context);
    const cached = this.cache.get(queryKey);

    if (cached && Date.now() - cached.timestamp < rules.cacheTTL) {
      this.metrics.cacheHits++;
      return cached.data;
    }

    return null; // Повертаємо null, щоб виконати запит
  }

  /**
   * Оптимізація файлових операцій
   */
  optimizeFileOperation(filePath, operation) {
    const rules = this.optimizationRules.get('file_operation');

    // Перевірка розміру файлу
    try {
      const stats = fs.statSync(filePath);
      if (stats.size > rules.maxFileSize) {
        throw new Error(`Файл занадто великий: ${stats.size} байт`);
      }
    } catch (error) {
      console.warn(`⚠️ Помилка перевірки файлу: ${error.message}`);
    }

    // Кешування метаданих файлу
    const fileKey = `file_${Buffer.from(filePath).toString('base64')}`;
    const cached = this.cache.get(fileKey);

    if (cached && Date.now() - cached.timestamp < rules.cacheTTL) {
      this.metrics.cacheHits++;
      return cached.data;
    }

    return null;
  }

  /**
   * Отримання статистики продуктивності
   */
  getPerformanceStats() {
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
        },
        optimizations: this.metrics.queryOptimizations,
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

    return {
      averageResponseTime: Math.round(averageResponseTime * 100) / 100,
      maxResponseTime: Math.round(maxResponseTime * 100) / 100,
      minResponseTime: Math.round(minResponseTime * 100) / 100,
      cacheHitRate: Math.round(cacheHitRate * 100) / 100,
      memoryUsage: {
        current: process.memoryUsage(),
        average: Math.round((averageMemory / 1024 / 1024) * 100) / 100, // MB
        max: Math.round((maxMemory / 1024 / 1024) * 100) / 100, // MB
      },
      optimizations: this.metrics.queryOptimizations,
      cacheSize: this.cache.size,
      queryCacheSize: this.queryCache.size,
    };
  }

  /**
   * Збереження метрик в файл
   */
  async saveMetrics() {
    try {
      const metricsDir = path.join(process.cwd(), 'metrics');
      await fs.ensureDir(metricsDir);

      const metricsFile = path.join(metricsDir, 'performance.json');
      const stats = this.getPerformanceStats();

      await fs.writeJson(
        metricsFile,
        {
          timestamp: Date.now(),
          stats,
          rawMetrics: this.metrics,
        },
        { spaces: 2 }
      );

      console.log(`📊 Метрики продуктивності збережено: ${metricsFile}`);
    } catch (error) {
      console.error(`❌ Помилка збереження метрик: ${error.message}`);
    }
  }

  /**
   * Отримання рекомендацій по оптимізації
   */
  getOptimizationRecommendations() {
    const stats = this.getPerformanceStats();
    const recommendations = [];

    // Рекомендації по часу відповіді
    if (stats.averageResponseTime > 5000) {
      recommendations.push({
        type: 'performance',
        priority: 'high',
        message:
          'Середній час відповіді занадто високий (>5с). Рекомендується оптимізувати запити до API.',
        action: 'Оптимізувати запити, додати кешування',
      });
    }

    // Рекомендації по кешу
    if (stats.cacheHitRate < 50) {
      recommendations.push({
        type: 'cache',
        priority: 'medium',
        message: 'Низький відсоток попадань в кеш (<50%). Рекомендується збільшити TTL кешу.',
        action: 'Збільшити TTL кешу, додати більше кешованих запитів',
      });
    }

    // Рекомендації по пам'яті
    if (stats.memoryUsage.average > 200) {
      recommendations.push({
        type: 'memory',
        priority: 'high',
        message: "Високе використання пам'яті (>200MB). Рекомендується оптимізувати кеш.",
        action: 'Зменшити розмір кешу, додати автоматичну очистку',
      });
    }

    // Рекомендації по оптимізації
    if (stats.optimizations < 10) {
      recommendations.push({
        type: 'optimization',
        priority: 'low',
        message: 'Мало оптимізацій запитів. Рекомендується додати більше правил оптимізації.',
        action: 'Додати правила оптимізації для різних типів запитів',
      });
    }

    return recommendations;
  }

  /**
   * Автоматична оптимізація на основі метрик
   */
  autoOptimize() {
    const recommendations = this.getOptimizationRecommendations();

    for (const rec of recommendations) {
      if (rec.priority === 'high') {
        console.log(`🚀 Автоматична оптимізація: ${rec.message}`);

        switch (rec.type) {
          case 'memory':
            this.cleanupCache();
            break;
          case 'cache':
            // Збільшуємо TTL для популярних запитів
            this.optimizationRules.get('sheets_query').cacheTTL *= 1.5;
            break;
          case 'performance':
            // Зменшуємо розмір batch для запитів
            this.optimizationRules.get('sheets_query').maxBatchSize = Math.floor(
              this.optimizationRules.get('sheets_query').maxBatchSize * 0.8
            );
            break;
        }
      }
    }
  }

  /**
   * Скидання метрик
   */
  resetMetrics() {
    this.metrics = {
      responseTimes: [],
      cacheHits: 0,
      cacheMisses: 0,
      queryOptimizations: 0,
      memoryUsage: [],
    };

    this.cache.clear();
    this.queryCache.clear();

    console.log('🔄 Метрики продуктивності скинуто');
  }
}

// Експорт синглтона
module.exports = new PerformanceOptimizer();
