"use strict";
/**
 * Розширений логер для Discord AI Assistant Bot
 * Рефакторована версія з покращеними можливостями
 * TypeScript версія 3.0.0 - Повністю рефакторовано
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.Logger = void 0;
const winston_1 = __importDefault(require("winston"));
const path_1 = __importDefault(require("path"));
const fs_1 = __importDefault(require("fs"));
const perf_hooks_1 = require("perf_hooks");
// Константи для конфігурації логера
const LOGGER_CONFIG = {
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
    MAX_FILES: 5,
    COMMAND_LOG_SIZE: 5 * 1024 * 1024, // 5MB
    COMMAND_LOG_FILES: 3,
    CLEANUP_INTERVAL: 24 * 60 * 60 * 1000, // 24 години
    MAX_LOG_AGE: 30 * 24 * 60 * 60 * 1000, // 30 днів
    BUFFER_SIZE: 1000,
    FLUSH_INTERVAL: 5000, // 5 секунд
};
class Logger {
    constructor() {
        this.logger = null;
        this.logBuffer = [];
        this.cleanupInterval = null;
        this.flushInterval = null;
        this.isInitialized = false;
        this.logsDir = path_1.default.join(process.cwd(), 'data', 'logs');
        this.stats = {
            totalLogs: 0,
            errors: 0,
            commands: 0,
            apiRequests: 0,
            performance: 0,
            security: 0,
            system: 0,
            debug: 0,
            warnings: 0,
            lastLogTime: new Date(),
            averageLogSize: 0,
            logBufferSize: 0,
        };
        this.initialize();
    }
    /**
     * Санітізація метаданих логів: маскує секрети, обрізає великі значення, прибирає цикли
     */
    sanitizeMeta(meta) {
        const SECRET_KEYS = new Set([
            'token', 'apiKey', 'apikey', 'api_key', 'password', 'pass', 'secret', 'clientSecret', 'authorization', 'auth', 'bearer', 'session', 'cookie', 'cookies'
        ]);
        const MAX_STRING_LEN = 2000; // захист від гігантських полів
        const seen = new WeakSet();
        const redact = (key, value) => {
            if (value == null)
                return value;
            if (SECRET_KEYS.has(key.toLowerCase()))
                return '[REDACTED]';
            if (typeof value === 'string') {
                return value.length > MAX_STRING_LEN ? value.slice(0, MAX_STRING_LEN) + '…' : value;
            }
            if (typeof value === 'object') {
                if (seen.has(value))
                    return '[CIRCULAR]';
                seen.add(value);
                if (Array.isArray(value))
                    return value.map((v) => redact(key, v));
                const out = {};
                for (const [k, v] of Object.entries(value)) {
                    out[k] = redact(k, v);
                }
                return out;
            }
            return value;
        };
        // Глибоке копіювання з санітізацією
        const safe = {};
        for (const [k, v] of Object.entries(meta || {})) {
            safe[k] = redact(k, v);
        }
        return safe;
    }
    /**
     * Ініціалізація логера з детальним логуванням
     */
    initialize() {
        try {
            console.log('🔧 Ініціалізація логера...');
            // Створення папки для логів
            this.ensureLogsDirectory();
            // Конфігурація форматів
            const formats = this.createFormats();
            // Створення транспортів
            const transports = this.createTransports();
            // Створення логера
            this.logger = winston_1.default.createLogger({
                level: this.getLogLevel(),
                format: formats.file,
                transports: transports,
                exitOnError: false,
                silent: false,
            });
            // Налаштування обробки необроблених помилок
            this.setupExceptionHandling();
            // Запуск періодичних завдань
            this.startPeriodicTasks();
            this.isInitialized = true;
            console.log('✅ Логер успішно ініціалізовано');
        }
        catch (error) {
            console.error('❌ Помилка ініціалізації логера:', error);
            this.createFallbackLogger();
        }
    }
    /**
     * Створення папки для логів
     */
    ensureLogsDirectory() {
        try {
            if (!fs_1.default.existsSync(this.logsDir)) {
                fs_1.default.mkdirSync(this.logsDir, { recursive: true });
                console.log(`📁 Створено папку для логів: ${this.logsDir}`);
            }
        }
        catch (error) {
            console.error('❌ Помилка створення папки логів:', error);
            throw new Error(`Неможливо створити папку логів: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Створення форматів логування
     */
    createFormats() {
        return {
            console: winston_1.default.format.combine(winston_1.default.format.colorize(), winston_1.default.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss.SSS' }), winston_1.default.format.errors({ stack: true }), winston_1.default.format.printf(({ timestamp, level, message, service, userId, ...meta }) => {
                let log = `${timestamp} [${level}]`;
                if (service)
                    log += ` [${service}]`;
                if (userId)
                    log += ` [User:${userId}]`;
                log += `: ${message}`;
                const remainingMeta = Object.keys(meta).filter(key => !['timestamp', 'level', 'service', 'userId'].includes(key));
                if (remainingMeta.length > 0) {
                    log += ` ${JSON.stringify(meta)}`;
                }
                return log;
            })),
            file: winston_1.default.format.combine(winston_1.default.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss.SSS' }), winston_1.default.format.errors({ stack: true }), winston_1.default.format.json()),
        };
    }
    /**
     * Створення транспортів
     */
    createTransports() {
        const formats = this.createFormats();
        return [
            // Консольний транспорт
            new winston_1.default.transports.Console({
                format: formats.console,
                level: this.getLogLevel(),
                handleExceptions: true,
                handleRejections: true,
            }),
            // Файл для всіх логів
            new winston_1.default.transports.File({
                filename: path_1.default.join(this.logsDir, 'bot.log'),
                format: formats.file,
                maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
                maxFiles: LOGGER_CONFIG.MAX_FILES,
                level: 'info',
                tailable: true,
                handleExceptions: true,
                handleRejections: true,
            }),
            // Файл для помилок
            new winston_1.default.transports.File({
                filename: path_1.default.join(this.logsDir, 'error.log'),
                format: formats.file,
                maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
                maxFiles: LOGGER_CONFIG.MAX_FILES,
                level: 'error',
                tailable: true,
            }),
            // Файл для команд
            new winston_1.default.transports.File({
                filename: path_1.default.join(this.logsDir, 'commands.log'),
                format: formats.file,
                maxsize: LOGGER_CONFIG.COMMAND_LOG_SIZE,
                maxFiles: LOGGER_CONFIG.COMMAND_LOG_FILES,
                level: 'info',
                tailable: true,
            }),
            // Файл для безпеки
            new winston_1.default.transports.File({
                filename: path_1.default.join(this.logsDir, 'security.log'),
                format: formats.file,
                maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
                maxFiles: LOGGER_CONFIG.MAX_FILES,
                level: 'warn',
                tailable: true,
            }),
            // Файл для продуктивності
            new winston_1.default.transports.File({
                filename: path_1.default.join(this.logsDir, 'performance.log'),
                format: formats.file,
                maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
                maxFiles: LOGGER_CONFIG.MAX_FILES,
                level: 'info',
                tailable: true,
            }),
        ];
    }
    /**
     * Отримання рівня логування
     */
    getLogLevel() {
        const level = process.env['LOG_LEVEL']?.toLowerCase();
        const validLevels = ['error', 'warn', 'info', 'debug'];
        if (level && validLevels.includes(level)) {
            return level;
        }
        return process.env['NODE_ENV'] === 'production' ? 'info' : 'debug';
    }
    /**
     * Налаштування обробки необроблених помилок
     */
    setupExceptionHandling() {
        if (!this.logger)
            return;
        const formats = this.createFormats();
        this.logger.exceptions.handle(new winston_1.default.transports.File({
            filename: path_1.default.join(this.logsDir, 'exceptions.log'),
            format: formats.file,
            maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
            maxFiles: LOGGER_CONFIG.MAX_FILES,
        }));
        this.logger.rejections.handle(new winston_1.default.transports.File({
            filename: path_1.default.join(this.logsDir, 'rejections.log'),
            format: formats.file,
            maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
            maxFiles: LOGGER_CONFIG.MAX_FILES,
        }));
    }
    /**
     * Запуск періодичних завдань
     */
    startPeriodicTasks() {
        // Очищення старих логів
        this.cleanupInterval = setInterval(() => {
            this.cleanupOldLogs();
        }, LOGGER_CONFIG.CLEANUP_INTERVAL);
        // Скидання буфера логів
        this.flushInterval = setInterval(() => {
            this.flushLogBuffer();
        }, LOGGER_CONFIG.FLUSH_INTERVAL);
    }
    /**
     * Створення резервного логера
     */
    createFallbackLogger() {
        console.warn('⚠️ Використання резервного логера');
        this.logger = winston_1.default.createLogger({
            level: 'info',
            format: winston_1.default.format.simple(),
            transports: [
                new winston_1.default.transports.Console(),
            ],
        });
    }
    /**
     * Логування з детальною інформацією
     */
    log(level, message, meta = {}) {
        if (!this.isInitialized || !this.logger) {
            console.log(`[${level.toUpperCase()}]: ${message}`, meta);
            return;
        }
        try {
            const startTime = perf_hooks_1.performance.now();
            // Додавання додаткової інформації
            const enhancedMeta = {
                ...meta,
                timestamp: new Date().toISOString(),
                service: meta.service || 'logger',
                logLevel: level,
                processId: process.pid,
                memory: process.memoryUsage(),
            };
            // Санитизация секретов
            const safeMeta = this.sanitizeMeta(enhancedMeta);
            // Оновлення статистики
            this.updateStats(level, message, safeMeta);
            // Додавання до буфера
            this.addToBuffer(level, message, safeMeta);
            // Логування через winston
            this.logger.log(level, message, safeMeta);
            const duration = perf_hooks_1.performance.now() - startTime;
            if (duration > 100) {
                console.warn(`⚠️ Повільне логування: ${duration.toFixed(2)}ms`);
            }
        }
        catch (error) {
            console.error('❌ Помилка логування:', error);
            console.log(`[${level.toUpperCase()}]: ${message}`, meta);
        }
    }
    /**
     * Оновлення статистики
     */
    updateStats(level, message, meta) {
        this.stats.totalLogs++;
        this.stats.lastLogTime = new Date();
        this.stats.logBufferSize = this.logBuffer.length;
        const logSize = JSON.stringify({ level, message, meta }).length;
        this.stats.averageLogSize = (this.stats.averageLogSize + logSize) / 2;
        switch (level) {
            case 'error':
                this.stats.errors++;
                break;
            case 'warn':
                this.stats.warnings++;
                break;
            case 'debug':
                this.stats.debug++;
                break;
        }
        if (meta.type === 'command')
            this.stats.commands++;
        if (meta.type === 'api_request')
            this.stats.apiRequests++;
        if (meta.type === 'performance')
            this.stats.performance++;
        if (meta.type === 'security')
            this.stats.security++;
        if (meta.type === 'system')
            this.stats.system++;
    }
    /**
     * Додавання до буфера логів
     */
    addToBuffer(level, message, meta) {
        const entry = {
            timestamp: new Date(),
            level,
            message,
            meta,
            size: JSON.stringify({ level, message, meta }).length,
        };
        this.logBuffer.push(entry);
        // Обмеження розміру буфера
        if (this.logBuffer.length > LOGGER_CONFIG.BUFFER_SIZE) {
            this.logBuffer.shift();
        }
    }
    /**
     * Скидання буфера логів
     */
    flushLogBuffer() {
        if (this.logBuffer.length === 0)
            return;
        try {
            const bufferSize = this.logBuffer.length;
            const totalSize = this.logBuffer.reduce((sum, entry) => sum + entry.size, 0);
            this.debug(`Скидання буфера логів: ${bufferSize} записів, ${totalSize} байт`);
            this.logBuffer = [];
            this.stats.logBufferSize = 0;
        }
        catch (error) {
            console.error('❌ Помилка скидання буфера логів:', error);
        }
    }
    /**
     * Очищення старих логів
     */
    cleanupOldLogs() {
        try {
            const files = fs_1.default.readdirSync(this.logsDir);
            const now = Date.now();
            let cleanedCount = 0;
            for (const file of files) {
                const filePath = path_1.default.join(this.logsDir, file);
                const stats = fs_1.default.statSync(filePath);
                if (now - stats.mtime.getTime() > LOGGER_CONFIG.MAX_LOG_AGE) {
                    fs_1.default.unlinkSync(filePath);
                    cleanedCount++;
                }
            }
            if (cleanedCount > 0) {
                this.info(`Очищено ${cleanedCount} старих лог-файлів`);
            }
        }
        catch (error) {
            console.error('❌ Помилка очищення старих логів:', error);
        }
    }
    /**
     * Логування інформації
     */
    info(message, meta = {}) {
        this.log('info', message, meta);
    }
    /**
     * Логування помилок
     */
    error(message, meta = {}) {
        this.log('error', message, meta);
    }
    /**
     * Логування попереджень
     */
    warn(message, meta = {}) {
        this.log('warn', message, meta);
    }
    /**
     * Логування дебагу
     */
    debug(message, meta = {}) {
        this.log('debug', message, meta);
    }
    /**
     * Логування команд з детальною інформацією
     */
    command(command, user, duration, success = true, meta = {}) {
        this.log('info', `Команда виконана: ${command}`, {
            ...meta,
            command,
            user,
            duration: `${duration}ms`,
            success,
            type: 'command',
            performance: duration > 1000 ? 'slow' : duration > 500 ? 'medium' : 'fast',
        });
    }
    /**
     * Логування помилок команд
     */
    commandError(command, user, error, duration, meta = {}) {
        this.log('error', `Помилка команди: ${command}`, {
            ...meta,
            command,
            user,
            error: error.message,
            stack: error.stack,
            duration: `${duration}ms`,
            type: 'command_error',
            errorType: error.constructor.name,
        });
    }
    /**
     * Логування API запитів
     */
    apiRequest(service, endpoint, duration, success = true, meta = {}) {
        this.log('info', `API запит: ${service} - ${endpoint}`, {
            ...meta,
            service,
            endpoint,
            duration: `${duration}ms`,
            success,
            type: 'api_request',
            performance: duration > 5000 ? 'slow' : duration > 1000 ? 'medium' : 'fast',
        });
    }
    /**
     * Логування помилок API
     */
    apiError(service, endpoint, error, duration, meta = {}) {
        this.log('error', `Помилка API: ${service} - ${endpoint}`, {
            ...meta,
            service,
            endpoint,
            error: error.message,
            stack: error.stack,
            duration: `${duration}ms`,
            type: 'api_error',
            errorType: error.constructor.name,
        });
    }
    /**
     * Логування подій безпеки
     */
    security(event, user, details = {}) {
        this.log('warn', `Подія безпеки: ${event}`, {
            ...details,
            event,
            user,
            type: 'security',
            severity: details.severity || 'medium',
        });
    }
    /**
     * Логування продуктивності
     */
    performance(operation, duration, details = {}) {
        this.log('info', `Метрика продуктивності: ${operation}`, {
            ...details,
            operation,
            duration: `${duration}ms`,
            type: 'performance',
            category: details.category || 'general',
        });
    }
    /**
     * Логування системних подій
     */
    system(event, details = {}) {
        this.log('info', `Системна подія: ${event}`, {
            ...details,
            event,
            type: 'system',
            component: details.component || 'unknown',
        });
    }
    /**
     * Отримання детальної статистики логера
     */
    getStats() {
        return {
            ...this.stats,
            logBufferSize: this.logBuffer.length,
        };
    }
    /**
     * Отримання буфера логів
     */
    getLogBuffer() {
        return [...this.logBuffer];
    }
    /**
     * Очищення ресурсів
     */
    async cleanup() {
        try {
            this.info('Очищення ресурсів логера...');
            // Зупинка періодичних завдань
            if (this.cleanupInterval) {
                clearInterval(this.cleanupInterval);
                this.cleanupInterval = null;
            }
            if (this.flushInterval) {
                clearInterval(this.flushInterval);
                this.flushInterval = null;
            }
            // Скидання буфера
            this.flushLogBuffer();
            // Закриття логера
            if (this.logger) {
                await new Promise((resolve) => {
                    this.logger.on('finish', () => resolve());
                    this.logger.end();
                });
            }
            this.info('Ресурси логера очищено');
        }
        catch (error) {
            console.error('❌ Помилка очищення ресурсів логера:', error);
        }
    }
    /**
     * Перевірка стану логера
     */
    isHealthy() {
        return this.isInitialized && this.logger !== null;
    }
}
exports.Logger = Logger;
// Експорт єдиного екземпляра
const logger = new Logger();
exports.default = logger;
//# sourceMappingURL=logger.js.map