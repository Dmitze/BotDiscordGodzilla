"use strict";
/**
 * Metrics Service для Discord бота
 * Централізоване управління метриками та моніторингом
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.MetricsService = void 0;
const prom_client_1 = require("prom-client");
const BaseService_1 = require("@/core/BaseService");
// TODO: Створити типизовані утиліти
const logger = {
    info: (message, ...args) => console.log(message, ...args),
    error: (message, ...args) => console.error(message, ...args),
    warn: (message, ...args) => console.warn(message, ...args),
    debug: (message, ...args) => console.debug(message, ...args),
};
class MetricsService extends BaseService_1.BaseService {
    constructor(config) {
        super('MetricsService', config);
        this.registry = null;
        this.metrics = null;
        this.server = null;
        this.updateInterval = null;
        this.stats = {
            service: 'MetricsService',
            uptime: 0,
            requests: 0,
            errors: 0,
            startTime: Date.now(),
            metricsCount: 0,
        };
    }
    /**
     * Ініціалізація Metrics сервісу
     */
    async onInitialize() {
        try {
            logger.info('📊 Ініціалізація Metrics сервісу...');
            // Створення Prometheus реєстру
            await this.createRegistry();
            // Створення метрик
            this.createMetrics();
            // Запуск HTTP сервера
            await this.startServer();
            // Запуск періодичних оновлень
            this.startPeriodicUpdates();
            logger.info('✅ Metrics сервіс ініціалізовано');
        }
        catch (error) {
            logger.error('❌ Помилка ініціалізації Metrics сервісу:', error);
            throw error;
        }
    }
    /**
     * Створення Prometheus реєстру
     */
    async createRegistry() {
        try {
            this.registry = new prom_client_1.Registry();
            // Збір стандартних метрик Node.js
            (0, prom_client_1.collectDefaultMetrics)({ register: this.registry });
            logger.debug('✅ Prometheus реєстр створено');
        }
        catch (error) {
            logger.error('Помилка створення Prometheus реєстру:', error);
            throw error;
        }
    }
    /**
     * Створення метрик
     */
    createMetrics() {
        try {
            if (!this.registry) {
                throw new Error('Реєстр не ініціалізовано');
            }
            this.metrics = {
                // Лічильники
                commandsTotal: new prom_client_1.Counter({
                    name: 'discord_bot_commands_total',
                    help: 'Загальна кількість виконаних команд',
                    labelNames: ['command', 'status'],
                    registers: [this.registry],
                }),
                messagesTotal: new prom_client_1.Counter({
                    name: 'discord_bot_messages_total',
                    help: 'Загальна кількість повідомлень',
                    labelNames: ['type'],
                    registers: [this.registry],
                }),
                errorsTotal: new prom_client_1.Counter({
                    name: 'discord_bot_errors_total',
                    help: 'Загальна кількість помилок',
                    labelNames: ['type', 'service'],
                    registers: [this.registry],
                }),
                // Гейджи
                activeUsers: new prom_client_1.Gauge({
                    name: 'discord_bot_active_users',
                    help: 'Кількість активних користувачів',
                    registers: [this.registry],
                }),
                activeGuilds: new prom_client_1.Gauge({
                    name: 'discord_bot_active_guilds',
                    help: 'Кількість активних серверів',
                    registers: [this.registry],
                }),
                memoryUsage: new prom_client_1.Gauge({
                    name: 'discord_bot_memory_usage_bytes',
                    help: 'Використання пам\'яті в байтах',
                    registers: [this.registry],
                }),
                uptime: new prom_client_1.Gauge({
                    name: 'discord_bot_uptime_seconds',
                    help: 'Час роботи бота в секундах',
                    registers: [this.registry],
                }),
                // Гістограми
                commandDuration: new prom_client_1.Histogram({
                    name: 'discord_bot_command_duration_seconds',
                    help: 'Тривалість виконання команд',
                    labelNames: ['command'],
                    buckets: [0.1, 0.5, 1, 2, 5, 10],
                    registers: [this.registry],
                }),
                apiResponseTime: new prom_client_1.Histogram({
                    name: 'discord_bot_api_response_time_seconds',
                    help: 'Час відповіді API',
                    labelNames: ['service'],
                    buckets: [0.1, 0.5, 1, 2, 5, 10],
                    registers: [this.registry],
                }),
                // Кеш метрики
                cacheHitRate: new prom_client_1.Gauge({
                    name: 'discord_bot_cache_hit_rate_percent',
                    help: 'Відсоток попадань в кеш',
                    registers: [this.registry],
                }),
                cacheSize: new prom_client_1.Gauge({
                    name: 'discord_bot_cache_size',
                    help: 'Розмір кешу',
                    registers: [this.registry],
                }),
                // Черги
                queueLength: new prom_client_1.Gauge({
                    name: 'discord_bot_queue_length',
                    help: 'Довжина черги',
                    labelNames: ['priority'],
                    registers: [this.registry],
                }),
                // Connection Pool
                connectionPoolUsage: new prom_client_1.Gauge({
                    name: 'discord_bot_connection_pool_usage_percent',
                    help: 'Використання connection pool',
                    labelNames: ['service'],
                    registers: [this.registry],
                }),
                // AI метрики
                aiRequestsTotal: new prom_client_1.Counter({
                    name: 'discord_bot_ai_requests_total',
                    help: 'Загальна кількість AI запитів',
                    labelNames: ['provider', 'status'],
                    registers: [this.registry],
                }),
                aiResponseTime: new prom_client_1.Histogram({
                    name: 'discord_bot_ai_response_time_seconds',
                    help: 'Час відповіді AI',
                    labelNames: ['provider'],
                    buckets: [0.1, 0.5, 1, 2, 5, 10, 30],
                    registers: [this.registry],
                }),
                // Google API метрики
                googleApiRequestsTotal: new prom_client_1.Counter({
                    name: 'discord_bot_google_api_requests_total',
                    help: 'Загальна кількість Google API запитів',
                    labelNames: ['service', 'endpoint', 'status'],
                    registers: [this.registry],
                }),
                googleApiResponseTime: new prom_client_1.Histogram({
                    name: 'discord_bot_google_api_response_time_seconds',
                    help: 'Час відповіді Google API',
                    labelNames: ['service'],
                    buckets: [0.1, 0.5, 1, 2, 5, 10],
                    registers: [this.registry],
                }),
            };
            this.stats.metricsCount = Object.keys(this.metrics).length;
            logger.debug('✅ Метрики створено');
        }
        catch (error) {
            logger.error('Помилка створення метрик:', error);
            throw error;
        }
    }
    /**
     * Запуск HTTP сервера
     */
    async startServer() {
        try {
            if (!this.config.metrics.enabled) {
                logger.info('Metrics сервер вимкнено');
                return;
            }
            const http = require('http');
            this.server = http.createServer(async (req, res) => {
                try {
                    if (req.url === this.config.metrics.path) {
                        res.writeHead(200, { 'Content-Type': 'text/plain' });
                        if (this.registry) {
                            const metrics = await this.registry.metrics();
                            res.end(metrics);
                        }
                        else {
                            res.end('# Metrics not available');
                        }
                    }
                    else {
                        res.writeHead(404);
                        res.end('Not Found');
                    }
                }
                catch (error) {
                    logger.error('Помилка обробки metrics запиту:', error);
                    res.writeHead(500);
                    res.end('Internal Server Error');
                }
            });
            this.server.listen(this.config.metrics.port, () => {
                logger.info(`📊 Metrics сервер запущено на порту ${this.config.metrics.port}`);
            });
            this.server.on('error', (error) => {
                logger.error('Помилка metrics сервера:', error);
            });
        }
        catch (error) {
            logger.error('Помилка запуску metrics сервера:', error);
            throw error;
        }
    }
    /**
     * Інкремент лічильника команд
     */
    incrementCommand(command, status = 'success') {
        if (this.metrics) {
            this.metrics.commandsTotal.inc({ command, status });
        }
    }
    /**
     * Інкремент лічильника повідомлень
     */
    incrementMessage(type) {
        if (this.metrics) {
            this.metrics.messagesTotal.inc({ type });
        }
    }
    /**
     * Інкремент лічильника помилок
     */
    incrementError(type, service = 'unknown') {
        if (this.metrics) {
            this.metrics.errorsTotal.inc({ type, service });
        }
    }
    /**
     * Встановлення кількості активних користувачів
     */
    setActiveUsers(count) {
        if (this.metrics) {
            this.metrics.activeUsers.set(count);
        }
    }
    /**
     * Встановлення кількості активних серверів
     */
    setActiveGuilds(count) {
        if (this.metrics) {
            this.metrics.activeGuilds.set(count);
        }
    }
    /**
     * Оновлення використання пам'яті
     */
    updateMemoryUsage() {
        if (this.metrics) {
            const memUsage = process.memoryUsage();
            this.metrics.memoryUsage.set(memUsage.heapUsed);
        }
    }
    /**
     * Оновлення часу роботи
     */
    updateUptime() {
        if (this.metrics) {
            const uptime = process.uptime();
            this.metrics.uptime.set(uptime);
        }
    }
    /**
     * Вимірювання тривалості команди
     */
    measureCommandDuration(command, duration) {
        if (this.metrics) {
            this.metrics.commandDuration.observe({ command }, duration / 1000);
        }
    }
    /**
     * Вимірювання часу відповіді API
     */
    measureApiResponseTime(service, endpoint, duration) {
        if (this.metrics) {
            this.metrics.apiResponseTime.observe({ service }, duration / 1000);
        }
    }
    /**
     * Оновлення метрик кешу
     */
    updateCacheMetrics(cacheStats) {
        if (this.metrics) {
            const totalRequests = cacheStats.hits + cacheStats.misses;
            const hitRate = totalRequests > 0 ? (cacheStats.hits / totalRequests) * 100 : 0;
            this.metrics.cacheHitRate.set(hitRate);
            this.metrics.cacheSize.set(cacheStats.hits + cacheStats.misses);
        }
    }
    /**
     * Оновлення метрик черг
     */
    updateQueueMetrics(queueStats) {
        if (this.metrics) {
            this.metrics.queueLength.set({ priority: 'high' }, queueStats.high.length);
            this.metrics.queueLength.set({ priority: 'normal' }, queueStats.normal.length);
            this.metrics.queueLength.set({ priority: 'low' }, queueStats.low.length);
        }
    }
    /**
     * Оновлення метрик connection pool
     */
    updateConnectionPoolMetrics(connectionStats) {
        if (this.metrics) {
            for (const [service, stats] of Object.entries(connectionStats)) {
                const usage = stats.inUse ? 100 : 0;
                this.metrics.connectionPoolUsage.set({ service }, usage);
            }
        }
    }
    /**
     * Оновлення AI метрик
     */
    updateAIMetrics(provider, status, duration) {
        if (this.metrics) {
            this.metrics.aiRequestsTotal.inc({ provider, status });
            this.metrics.aiResponseTime.observe({ provider }, duration / 1000);
        }
    }
    /**
     * Оновлення Google API метрик
     */
    updateGoogleApiMetrics(service, endpoint, status, duration) {
        if (this.metrics) {
            this.metrics.googleApiRequestsTotal.inc({ service, endpoint, status });
            this.metrics.googleApiResponseTime.observe({ service }, duration / 1000);
        }
    }
    /**
     * Оновлення всіх метрик
     */
    updateAllMetrics() {
        try {
            this.updateMemoryUsage();
            this.updateUptime();
            // TODO: Отримати статистику з інших сервісів
            // const cacheStats = this.bot.serviceContainer.get('CacheService').getCacheStats();
            // this.updateCacheMetrics(cacheStats);
            // const queueStats = this.bot.queueManager.getQueueStats();
            // this.updateQueueMetrics(queueStats);
            // const connectionStats = this.bot.serviceContainer.get('GoogleService').getConnectionStats();
            // this.updateConnectionPoolMetrics(connectionStats);
        }
        catch (error) {
            logger.error('Помилка оновлення метрик:', error);
        }
    }
    /**
     * Запуск періодичних оновлень
     */
    startPeriodicUpdates() {
        this.updateInterval = setInterval(() => {
            this.updateAllMetrics();
        }, 30000); // Кожні 30 секунд
    }
    /**
     * Health check
     */
    async onHealthCheck() {
        try {
            if (!this.config.metrics.enabled) {
                return {
                    healthy: true,
                    service: this.name,
                    details: { enabled: false },
                };
            }
            if (!this.registry || !this.metrics) {
                return {
                    healthy: false,
                    service: this.name,
                    error: 'Метрики не ініціалізовано',
                };
            }
            // Тестовий запит до metrics endpoint
            if (this.server) {
                try {
                    const http = require('http');
                    const response = await new Promise((resolve, reject) => {
                        const req = http.get(`http://localhost:${this.config.metrics.port}${this.config.metrics.path}`, (res) => {
                            let data = '';
                            res.on('data', (chunk) => data += chunk);
                            res.on('end', () => resolve({ statusCode: res.statusCode, data }));
                        });
                        req.on('error', reject);
                        req.setTimeout(5000, () => reject(new Error('Timeout')));
                    });
                    if (response.statusCode !== 200) {
                        return {
                            healthy: false,
                            service: this.name,
                            error: `Metrics endpoint returned ${response.statusCode}`,
                        };
                    }
                }
                catch (error) {
                    return {
                        healthy: false,
                        service: this.name,
                        error: `Metrics endpoint test failed: ${error}`,
                    };
                }
            }
            return {
                healthy: true,
                service: this.name,
                details: {
                    metricsCount: this.stats.metricsCount,
                    serverRunning: !!this.server,
                    port: this.config.metrics.port,
                },
            };
        }
        catch (error) {
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
    async onShutdown() {
        try {
            if (this.updateInterval) {
                clearInterval(this.updateInterval);
                this.updateInterval = null;
            }
            if (this.server) {
                this.server.close();
                this.server = null;
            }
            logger.info('✅ Metrics Service зупинено');
        }
        catch (error) {
            logger.error('❌ Помилка зупинки Metrics Service:', error);
            throw error;
        }
    }
    /**
     * Отримання статистики
     */
    onGetStats() {
        return this.stats;
    }
}
exports.MetricsService = MetricsService;
//# sourceMappingURL=MetricsService.js.map