"use strict";
/**
 * Утиліти для оптимізації продуктивності Discord AI Assistant Bot
 * Моніторинг та оптимізація продуктивності
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.cleanupPerformanceOptimizer = exports.getSystemMetrics = exports.getPerformanceStats = exports.getFromCache = exports.cache = exports.measureOperation = exports.performanceOptimizer = exports.PerformanceOptimizer = void 0;
const logger_1 = __importDefault(require("./logger"));
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
};
/**
 * Клас для оптимізації продуктивності
 */
class PerformanceOptimizer {
    constructor() {
        this.caches = new Map();
        this.metricsHistory = [];
        this.gcInterval = null;
        this.metricsInterval = null;
        this.cacheCleanupInterval = null;
        this.isInitialized = false;
        this.isShuttingDown = false;
        if (PerformanceOptimizer.instance) {
            logger_1.default.debug('🔄 Повернення існуючого екземпляру PerformanceOptimizer');
            return PerformanceOptimizer.instance;
        }
        logger_1.default.info('🔧 Ініціалізація PerformanceOptimizer...');
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
        logger_1.default.info('✅ PerformanceOptimizer успішно ініціалізовано');
    }
    /**
     * Запуск моніторингу продуктивності
     */
    startMonitoring() {
        try {
            logger_1.default.info('📊 Запуск моніторингу продуктивності...');
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
            logger_1.default.info('✅ Моніторинг продуктивності запущено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка запуску моніторингу продуктивності:', error);
            throw new Error(`Помилка запуску моніторингу: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Збір метрик продуктивності
     */
    collectMetrics() {
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
            logger_1.default.debug('📊 Метрики продуктивності зібрано', {
                memory: {
                    rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`,
                    heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
                    heapTotal: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
                },
                cpu: {
                    usage: `${cpuUsage.usage.toFixed(2)}%`,
                    load: cpuUsage.load.toFixed(2),
                },
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка збору метрик продуктивності:', error);
        }
    }
    /**
     * Отримання використання CPU
     */
    getCPUUsage() {
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
                }
                catch (cpuError) {
                    logger_1.default.error('❌ Помилка вимірювання CPU:', cpuError);
                    this.performanceData.cpu.usage = 0;
                }
            }, PERFORMANCE_CONSTANTS.CPU_MEASUREMENT_DELAY);
            return {
                usage: this.performanceData.cpu.usage,
                load: require('os').loadavg()[0] || 0,
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка отримання використання CPU:', error);
            return { usage: 0, load: 0 };
        }
    }
    /**
     * Перевірка порогів продуктивності
     */
    checkThresholds() {
        try {
            const memoryUsage = this.performanceData.memory;
            const cpuUsage = this.performanceData.cpu;
            const heapUsedPercent = (memoryUsage.heapUsed / memoryUsage.heapTotal) * 100;
            // Перевірка пам'яті
            if (heapUsedPercent > PERFORMANCE_CONSTANTS.MEMORY_THRESHOLD) {
                logger_1.default.warn('⚠️ Високе використання пам\'яті', {
                    heapUsed: `${Math.round(heapUsedPercent)}%`,
                    threshold: `${PERFORMANCE_CONSTANTS.MEMORY_THRESHOLD}%`,
                });
                this.optimizeMemory();
            }
            // Перевірка CPU
            if (cpuUsage.usage > PERFORMANCE_CONSTANTS.CPU_THRESHOLD) {
                logger_1.default.warn('⚠️ Високе використання CPU', {
                    usage: `${cpuUsage.usage.toFixed(2)}%`,
                    threshold: `${PERFORMANCE_CONSTANTS.CPU_THRESHOLD}%`,
                });
                this.optimizeCPU();
            }
        }
        catch (error) {
            logger_1.default.error('❌ Помилка перевірки порогів продуктивності:', error);
        }
    }
    /**
     * Оптимізація пам'яті
     */
    optimizeMemory() {
        try {
            logger_1.default.info('🧹 Оптимізація пам\'яті...');
            // Примусова garbage collection
            if (global.gc) {
                global.gc();
                logger_1.default.debug('✅ Garbage collection виконано');
            }
            // Очищення кешів
            this.cleanupCaches();
            // Очищення метрик
            if (this.metricsHistory.length > PERFORMANCE_CONSTANTS.MAX_METRICS_HISTORY / 2) {
                this.metricsHistory = this.metricsHistory.slice(-PERFORMANCE_CONSTANTS.MAX_METRICS_HISTORY / 2);
                logger_1.default.debug('✅ Історія метрик очищено');
            }
            logger_1.default.info('✅ Оптимізація пам\'яті завершено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка оптимізації пам\'яті:', error);
        }
    }
    /**
     * Оптимізація CPU
     */
    optimizeCPU() {
        try {
            logger_1.default.info('⚡ Оптимізація CPU...');
            // Зменшення інтервалів моніторингу
            if (this.metricsInterval) {
                clearInterval(this.metricsInterval);
                this.metricsInterval = setInterval(() => {
                    this.collectMetrics();
                }, PERFORMANCE_CONSTANTS.METRICS_INTERVAL * 2);
            }
            // Очищення повільних операцій
            this.performanceData.slowOperations = [];
            logger_1.default.info('✅ Оптимізація CPU завершено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка оптимізації CPU:', error);
        }
    }
    /**
     * Виконання garbage collection
     */
    performGarbageCollection() {
        try {
            if (global.gc) {
                const beforeMemory = process.memoryUsage();
                global.gc();
                const afterMemory = process.memoryUsage();
                const freedMemory = beforeMemory.heapUsed - afterMemory.heapUsed;
                logger_1.default.debug('🧹 Garbage collection виконано', {
                    freedMemory: `${Math.round(freedMemory / 1024 / 1024)}MB`,
                    beforeHeap: `${Math.round(beforeMemory.heapUsed / 1024 / 1024)}MB`,
                    afterHeap: `${Math.round(afterMemory.heapUsed / 1024 / 1024)}MB`,
                });
            }
        }
        catch (error) {
            logger_1.default.error('❌ Помилка garbage collection:', error);
        }
    }
    /**
     * Очищення кешів
     */
    cleanupCaches() {
        try {
            let totalCleaned = 0;
            const now = Date.now();
            const maxAge = 30 * 60 * 1000; // 30 хвилин
            for (const [cacheName, cache] of this.caches.entries()) {
                let cleanedCount = 0;
                const entriesToDelete = [];
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
                logger_1.default.debug(`🧹 Очищено ${totalCleaned} записів з кешів`);
            }
        }
        catch (error) {
            logger_1.default.error('❌ Помилка очищення кешів:', error);
        }
    }
    /**
     * Вимірювання часу виконання операції
     */
    async measureOperation(operation, operationName, category = 'general') {
        const startTime = performance.now();
        try {
            const result = await operation();
            const duration = performance.now() - startTime;
            this.recordOperation(operationName, duration, category);
            return result;
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.recordOperation(operationName, duration, category, error);
            throw error;
        }
    }
    /**
     * Запис операції
     */
    recordOperation(operationName, duration, category, error) {
        try {
            const metric = {
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
                logger_1.default.warn(`🐌 Повільна операція: ${operationName}`, {
                    duration: `${duration.toFixed(2)}ms`,
                    category,
                    threshold: `${PERFORMANCE_CONSTANTS.SLOW_OPERATION_THRESHOLD}ms`,
                });
            }
            // Логування продуктивності
            logger_1.default.performance(operationName, duration, {
                category,
                error: error ? true : false,
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка запису операції:', error);
        }
    }
    /**
     * Кешування з автоматичним очищенням
     */
    cache(cacheName, key, value, ttl = 300000 // 5 хвилин за замовчуванням
    ) {
        try {
            if (!this.caches.has(cacheName)) {
                this.caches.set(cacheName, new Map());
            }
            const cache = this.caches.get(cacheName);
            const now = Date.now();
            cache.set(key, {
                value,
                timestamp: now,
                accessCount: 0,
                lastAccess: now,
            });
            logger_1.default.debug(`💾 Значення кешовано: ${cacheName}:${key}`);
        }
        catch (error) {
            logger_1.default.error('❌ Помилка кешування:', error);
        }
    }
    /**
     * Отримання з кешу
     */
    getFromCache(cacheName, key) {
        try {
            const cache = this.caches.get(cacheName);
            if (!cache)
                return null;
            const entry = cache.get(key);
            if (!entry)
                return null;
            // Оновлення статистики доступу
            entry.accessCount++;
            entry.lastAccess = Date.now();
            logger_1.default.debug(`📖 Значення отримано з кешу: ${cacheName}:${key}`);
            return entry.value;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка отримання з кешу:', error);
            return null;
        }
    }
    /**
     * Отримання статистики продуктивності
     */
    getPerformanceStats() {
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
        }
        catch (error) {
            logger_1.default.error('❌ Помилка отримання статистики продуктивності:', error);
            return {};
        }
    }
    /**
     * Отримання системних метрик
     */
    getSystemMetrics() {
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
        }
        catch (error) {
            logger_1.default.error('❌ Помилка отримання системних метрик:', error);
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
    cleanup() {
        if (this.isShuttingDown) {
            logger_1.default.warn('⚠️ PerformanceOptimizer вже зупиняється');
            return;
        }
        this.isShuttingDown = true;
        try {
            logger_1.default.info('🧹 Очищення ресурсів PerformanceOptimizer...');
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
            logger_1.default.info('✅ Ресурси PerformanceOptimizer очищено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка очищення PerformanceOptimizer:', error);
        }
        finally {
            this.isShuttingDown = false;
        }
    }
    /**
     * Перевірка стану ініціалізації
     */
    getInitializedState() {
        return this.isInitialized;
    }
    /**
     * Перевірка стану зупинки
     */
    getShuttingDownState() {
        return this.isShuttingDown;
    }
}
exports.PerformanceOptimizer = PerformanceOptimizer;
PerformanceOptimizer.instance = null;
// Експорт єдиного екземпляра
exports.performanceOptimizer = new PerformanceOptimizer();
// Експорт функцій для зручності
const measureOperation = (operation, operationName, category) => {
    return exports.performanceOptimizer.measureOperation(operation, operationName, category);
};
exports.measureOperation = measureOperation;
const cache = (cacheName, key, value, ttl) => {
    exports.performanceOptimizer.cache(cacheName, key, value, ttl);
};
exports.cache = cache;
const getFromCache = (cacheName, key) => {
    return exports.performanceOptimizer.getFromCache(cacheName, key);
};
exports.getFromCache = getFromCache;
const getPerformanceStats = () => exports.performanceOptimizer.getPerformanceStats();
exports.getPerformanceStats = getPerformanceStats;
const getSystemMetrics = () => exports.performanceOptimizer.getSystemMetrics();
exports.getSystemMetrics = getSystemMetrics;
const cleanupPerformanceOptimizer = () => exports.performanceOptimizer.cleanup();
exports.cleanupPerformanceOptimizer = cleanupPerformanceOptimizer;
//# sourceMappingURL=performanceOptimizer.js.map