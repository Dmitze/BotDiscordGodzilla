"use strict";
/**
 * Розширена система безпеки для Discord AI Assistant Bot
 * Валідація, санітизація та захист від атак
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateCommandOptions = exports.sanitizeInput = exports.cleanupSecurityManager = exports.getSuspiciousActivities = exports.getSecurityStats = exports.validateUrl = exports.checkRateLimit = exports.validateInput = exports.securityManager = exports.SecurityManager = void 0;
const errorHandler_1 = require("./errorHandler");
const logger_1 = __importDefault(require("./logger"));
// Константи для безпеки
const SECURITY_CONSTANTS = {
    MAX_INPUT_LENGTH: 2000,
    MAX_COMMAND_LENGTH: 100,
    MAX_URL_LENGTH: 500,
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
    RATE_LIMIT_WINDOW: 60000, // 1 хвилина
    RATE_LIMIT_MAX: 10, // 10 запитів за хвилину
    SUSPICIOUS_PATTERNS: [
        /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
        /javascript:/gi,
        /on\w+\s*=/gi,
        /data:text\/html/gi,
        /vbscript:/gi,
        /<iframe/gi,
        /<object/gi,
        /<embed/gi,
        /<applet/gi,
        /<meta/gi,
        /<link/gi,
        /<base/gi,
        /<form/gi,
        /<input/gi,
        /<textarea/gi,
        /<select/gi,
        /<button/gi,
        /<label/gi,
        /<fieldset/gi,
        /<legend/gi,
        /<optgroup/gi,
        /<option/gi,
    ],
    ALLOWED_CHARS: /^[a-zA-Z0-9\s\-_.,!?@#$%^&*()+=<>{}[\]|\\/:;"'`~]+$/,
    ALLOWED_URLS: /^https?:\/\/(www\.)?(discord\.com|discordapp\.com|google\.com|docs\.google\.com|drive\.google\.com)/i,
    BLACKLISTED_WORDS: [
        'admin', 'root', 'sudo', 'system', 'exec', 'eval', 'require', 'import',
        'delete', 'drop', 'insert', 'update', 'select', 'union', 'where',
        'script', 'javascript', 'vbscript', 'onload', 'onerror', 'onclick',
    ],
};
class SecurityManager {
    constructor() {
        this.rateLimitMap = new Map();
        this.blacklistCache = new Set();
        this.suspiciousActivities = [];
        this._isInitialized = false;
        if (SecurityManager.instance) {
            return SecurityManager.instance;
        }
        SecurityManager.instance = this;
        this.stats = {
            totalValidations: 0,
            successfulValidations: 0,
            failedValidations: 0,
            suspiciousActivities: 0,
            rateLimitHits: 0,
            blacklistHits: 0,
            xssAttempts: 0,
            sqlInjectionAttempts: 0,
            averageValidationTime: 0,
            totalValidationTime: 0,
        };
        this.initialize();
    }
    /**
     * Ініціалізація системи безпеки
     */
    initialize() {
        try {
            logger_1.default.info('🔒 Ініціалізація системи безпеки...');
            // Завантаження чорного списку
            this.loadBlacklist();
            // Запуск періодичних завдань
            this.startPeriodicTasks();
            this._isInitialized = true;
            logger_1.default.info('✅ Система безпеки успішно ініціалізована');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'initialize' },
            });
            throw new Error('Помилка ініціалізації системи безпеки');
        }
    }
    /**
     * Завантаження чорного списку
     */
    loadBlacklist() {
        try {
            // Тут можна завантажити чорний список з файлу або бази даних
            SECURITY_CONSTANTS.BLACKLISTED_WORDS.forEach(word => {
                this.blacklistCache.add(word.toLowerCase());
            });
            logger_1.default.info(`📋 Завантажено ${this.blacklistCache.size} слів у чорний список`);
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'loadBlacklist' },
            });
        }
    }
    /**
     * Запуск періодичних завдань
     */
    startPeriodicTasks() {
        // Очищення rate limit кешу кожні 5 хвилин
        setInterval(() => {
            this.cleanupRateLimitCache();
        }, 5 * 60 * 1000);
        // Очищення підозрілої активності кожні 10 хвилин
        setInterval(() => {
            this.cleanupSuspiciousActivities();
        }, 10 * 60 * 1000);
        logger_1.default.info('⏰ Періодичні завдання безпеки запущено');
    }
    /**
     * Валідація та санітизація введення
     */
    validateInput(input, context = {}) {
        const startTime = performance.now();
        try {
            logger_1.default.debug('🔍 Валідація введення...', {
                inputLength: input.length,
                inputType: context.inputType,
                userId: context.userId,
                commandName: context.commandName,
            });
            const errors = [];
            const warnings = [];
            let sanitizedValue = input;
            // Перевірка довжини
            if (input.length > SECURITY_CONSTANTS.MAX_INPUT_LENGTH) {
                errors.push(`Введення занадто довге (${input.length} символів, максимум ${SECURITY_CONSTANTS.MAX_INPUT_LENGTH})`);
                sanitizedValue = input.substring(0, SECURITY_CONSTANTS.MAX_INPUT_LENGTH);
            }
            // Перевірка на XSS атаки
            const xssResult = this.checkForXSS(input);
            if (xssResult.found) {
                errors.push('Виявлено потенційну XSS атаку');
                this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
                    subtype: 'xss_attempt',
                    pattern: xssResult.pattern,
                    input: input.substring(0, 100),
                });
                this.stats.xssAttempts++;
            }
            // Перевірка на SQL ін'єкції
            const sqlResult = this.checkForSQLInjection(input);
            if (sqlResult.found) {
                errors.push('Виявлено потенційну SQL ін\'єкцію');
                this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
                    subtype: 'sql_injection_attempt',
                    pattern: sqlResult.pattern,
                    input: input.substring(0, 100),
                });
                this.stats.sqlInjectionAttempts++;
            }
            // Перевірка чорного списку
            const blacklistResult = this.checkBlacklist(input);
            if (blacklistResult.found) {
                warnings.push('Виявлено слова з чорного списку');
                this.stats.blacklistHits++;
            }
            // Перевірка дозволених символів
            if (!SECURITY_CONSTANTS.ALLOWED_CHARS.test(input)) {
                warnings.push('Введення містить недозволені символи');
            }
            // Санітизація
            sanitizedValue = this.sanitizeInput(input);
            const duration = performance.now() - startTime;
            this.updateStats(true, duration);
            const result = {
                isValid: errors.length === 0,
                sanitizedValue,
                errors,
                warnings,
            };
            if (errors.length > 0) {
                this.stats.failedValidations++;
                logger_1.default.warn('❌ Валідація введення невдала', {
                    errors,
                    warnings,
                    inputLength: input.length,
                    userId: context.userId,
                    commandName: context.commandName,
                });
            }
            else {
                this.stats.successfulValidations++;
                logger_1.default.debug('✅ Валідація введення успішна', {
                    inputLength: input.length,
                    warnings,
                    userId: context.userId,
                    commandName: context.commandName,
                });
            }
            return result;
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.updateStats(false, duration);
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                ...(context.userId ? { userId: context.userId } : {}),
                additionalContext: { operation: 'validateInput', input: input.substring(0, 100) },
            });
            return {
                isValid: false,
                sanitizedValue: '',
                errors: ['Помилка валідації введення'],
                warnings: [],
            };
        }
    }
    /**
     * Перевірка на XSS атаки
     */
    checkForXSS(input) {
        for (const pattern of SECURITY_CONSTANTS.SUSPICIOUS_PATTERNS) {
            if (pattern.test(input)) {
                return { found: true, pattern: pattern.source };
            }
        }
        return { found: false };
    }
    /**
     * Перевірка на SQL ін'єкції
     */
    checkForSQLInjection(input) {
        const sqlPatterns = [
            /(\b(union|select|insert|update|delete|drop|create|alter)\b)/i,
            /(\b(where|from|into|values|set)\b)/i,
            /(--|#|\/\*|\*\/)/,
            /(\b(and|or)\b\s+\d+\s*=\s*\d+)/i,
            /(\b(and|or)\b\s+['"]\w+['"]\s*=\s*['"]\w+['"])/i,
        ];
        for (const pattern of sqlPatterns) {
            if (pattern.test(input)) {
                return { found: true, pattern: pattern.source };
            }
        }
        return { found: false };
    }
    /**
     * Перевірка чорного списку
     */
    checkBlacklist(input) {
        const foundWords = [];
        const words = input.toLowerCase().split(/\s+/);
        for (const word of words) {
            if (this.blacklistCache.has(word)) {
                foundWords.push(word);
            }
        }
        return {
            found: foundWords.length > 0,
            words: foundWords,
        };
    }
    /**
     * Санітизація введення
     */
    sanitizeInput(input) {
        let sanitized = input;
        // Видалення HTML тегів
        sanitized = sanitized.replace(/<[^>]*>/g, '');
        // Екранування спеціальних символів
        sanitized = sanitized
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#x27;')
            .replace(/\//g, '&#x2F;');
        // Видалення зайвих пробілів
        sanitized = sanitized.trim().replace(/\s+/g, ' ');
        return sanitized;
    }
    /**
     * Перевірка rate limit
     */
    checkRateLimit(userId) {
        try {
            const now = Date.now();
            const userLimit = this.rateLimitMap.get(userId);
            if (!userLimit || now > userLimit.resetTime) {
                // Створення нового ліміту
                this.rateLimitMap.set(userId, {
                    count: 1,
                    resetTime: now + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
                    lastRequest: now,
                });
                return {
                    allowed: true,
                    remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - 1,
                    resetTime: now + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
                };
            }
            if (userLimit.count >= SECURITY_CONSTANTS.RATE_LIMIT_MAX) {
                this.stats.rateLimitHits++;
                this.recordSecurityEvent('rate_limit', userId, {
                    count: userLimit.count,
                    resetTime: userLimit.resetTime,
                });
                logger_1.default.warn('⏰ Rate limit перевищено', {
                    userId,
                    count: userLimit.count,
                    resetTime: userLimit.resetTime,
                });
                return {
                    allowed: false,
                    remaining: 0,
                    resetTime: userLimit.resetTime,
                };
            }
            // Збільшення лічильника
            userLimit.count++;
            userLimit.lastRequest = now;
            return {
                allowed: true,
                remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - userLimit.count,
                resetTime: userLimit.resetTime,
            };
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                userId,
                additionalContext: { operation: 'checkRateLimit' },
            });
            // У випадку помилки дозволяємо запит
            return {
                allowed: true,
                remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX,
                resetTime: Date.now() + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
            };
        }
    }
    /**
     * Валідація URL
     */
    validateUrl(url) {
        try {
            const errors = [];
            const warnings = [];
            // Перевірка довжини
            if (url.length > SECURITY_CONSTANTS.MAX_URL_LENGTH) {
                errors.push(`URL занадто довгий (${url.length} символів, максимум ${SECURITY_CONSTANTS.MAX_URL_LENGTH})`);
            }
            // Перевірка протоколу
            if (!url.startsWith('http://') && !url.startsWith('https://')) {
                errors.push('URL повинен починатися з http:// або https://');
            }
            // Перевірка дозволених доменів
            if (!SECURITY_CONSTANTS.ALLOWED_URLS.test(url)) {
                warnings.push('URL не з дозволеного домену');
            }
            // Перевірка на підозрілі патерни
            if (url.includes('javascript:') || url.includes('data:text/html')) {
                errors.push('URL містить підозрілі патерни');
            }
            return {
                isValid: errors.length === 0,
                sanitizedValue: url,
                errors,
                warnings,
            };
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'validateUrl', url },
            });
            return {
                isValid: false,
                sanitizedValue: '',
                errors: ['Помилка валідації URL'],
                warnings: [],
            };
        }
    }
    /**
     * Запис події безпеки
     */
    recordSecurityEvent(type, userId, details = {}) {
        try {
            const event = {
                type,
                userId,
                details,
                timestamp: new Date(),
                severity: this.determineEventSeverity(type),
            };
            this.suspiciousActivities.push(event);
            this.stats.suspiciousActivities++;
            this.stats.lastSecurityEvent = event;
            // Обмеження розміру масиву
            if (this.suspiciousActivities.length > 1000) {
                this.suspiciousActivities = this.suspiciousActivities.slice(-500);
            }
            logger_1.default.security(type, userId, {
                details,
                severity: event.severity,
                timestamp: event.timestamp.toISOString(),
            });
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                userId,
                additionalContext: { operation: 'recordSecurityEvent', type },
            });
        }
    }
    /**
     * Визначення серйозності події
     */
    determineEventSeverity(type) {
        switch (type) {
            case 'unauthorized_access':
                return 'high';
            case 'rate_limit':
                return 'medium';
            case 'invalid_input':
            case 'suspicious_activity':
            default:
                return 'low';
        }
    }
    /**
     * Очищення rate limit кешу
     */
    cleanupRateLimitCache() {
        try {
            const now = Date.now();
            let cleanedCount = 0;
            for (const [userId, limit] of this.rateLimitMap.entries()) {
                if (now > limit.resetTime) {
                    this.rateLimitMap.delete(userId);
                    cleanedCount++;
                }
            }
            if (cleanedCount > 0) {
                logger_1.default.debug(`🧹 Очищено ${cleanedCount} застарілих rate limit записів`);
            }
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'cleanupRateLimitCache' },
            });
        }
    }
    /**
     * Очищення підозрілої активності
     */
    cleanupSuspiciousActivities() {
        try {
            const now = new Date();
            const maxAge = 24 * 60 * 60 * 1000; // 24 години
            const initialCount = this.suspiciousActivities.length;
            this.suspiciousActivities = this.suspiciousActivities.filter(activity => now.getTime() - activity.timestamp.getTime() < maxAge);
            const cleanedCount = initialCount - this.suspiciousActivities.length;
            if (cleanedCount > 0) {
                logger_1.default.debug(`🧹 Очищено ${cleanedCount} застарілих подій безпеки`);
            }
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'cleanupSuspiciousActivities' },
            });
        }
    }
    /**
     * Оновлення статистики
     */
    updateStats(success, duration) {
        try {
            this.stats.totalValidations++;
            this.stats.totalValidationTime += duration;
            this.stats.averageValidationTime = this.stats.totalValidationTime / this.stats.totalValidations;
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'updateStats' },
            });
        }
    }
    /**
     * Отримання статистики безпеки
     */
    getStats() {
        return { ...this.stats };
    }
    /**
     * Отримання підозрілої активності
     */
    getSuspiciousActivities() {
        return [...this.suspiciousActivities];
    }
    /**
     * Очищення ресурсів
     */
    cleanup() {
        try {
            this.rateLimitMap.clear();
            this.suspiciousActivities = [];
            this.blacklistCache.clear();
            logger_1.default.info('🧹 Ресурси SecurityManager очищено');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'SecurityManager',
                additionalContext: { operation: 'cleanup' },
            });
        }
    }
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized() {
        return this._isInitialized;
    }
}
exports.SecurityManager = SecurityManager;
SecurityManager.instance = null;
// Експорт єдиного екземпляра
exports.securityManager = new SecurityManager();
// Експорт функцій для зручності
const validateInput = (input, context) => exports.securityManager.validateInput(input, context);
exports.validateInput = validateInput;
const checkRateLimit = (userId) => exports.securityManager.checkRateLimit(userId);
exports.checkRateLimit = checkRateLimit;
const validateUrl = (url) => exports.securityManager.validateUrl(url);
exports.validateUrl = validateUrl;
const getSecurityStats = () => exports.securityManager.getStats();
exports.getSecurityStats = getSecurityStats;
const getSuspiciousActivities = () => exports.securityManager.getSuspiciousActivities();
exports.getSuspiciousActivities = getSuspiciousActivities;
const cleanupSecurityManager = () => exports.securityManager.cleanup();
exports.cleanupSecurityManager = cleanupSecurityManager;
// Функції для зворотної сумісності
const sanitizeInput = (input) => {
    const result = (0, exports.validateInput)(input);
    return result.sanitizedValue;
};
exports.sanitizeInput = sanitizeInput;
const validateCommandOptions = (options) => {
    const input = JSON.stringify(options);
    return (0, exports.validateInput)(input, { inputType: 'command' });
};
exports.validateCommandOptions = validateCommandOptions;
//# sourceMappingURL=security.js.map