"use strict";
/**
 * Клас для управління конфігурацією додатку
 * Завантажує та валідує налаштування з змінних середовища
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.Config = void 0;
const fs_1 = require("fs");
const path_1 = require("path");
const logger_1 = __importDefault(require("@/utils/logger"));
// Константи для конфігурації
const CONFIG_CONSTANTS = {
    DEFAULT_PREFIX: '!',
    DEFAULT_INTENTS: ['Guilds', 'GuildMessages', 'MessageContent', 'GuildMembers'],
    DEFAULT_AI_PROVIDER: 'openai',
    DEFAULT_OPENAI_MODEL: 'gpt-3.5-turbo',
    DEFAULT_OPENAI_MAX_TOKENS: 1000,
    DEFAULT_OPENAI_TEMPERATURE: 0.7,
    DEFAULT_OLLAMA_HOST: 'http://localhost:11434',
    DEFAULT_OLLAMA_MODEL: 'llama2',
    DEFAULT_REDIS_HOST: 'localhost',
    DEFAULT_REDIS_PORT: 6379,
    DEFAULT_REDIS_DATABASE: 0,
    DEFAULT_METRICS_PORT: 9090,
    DEFAULT_METRICS_PATH: '/metrics',
    MAX_OPENAI_TOKENS: 4000,
    MAX_TEMPERATURE: 2.0,
    MIN_TEMPERATURE: 0.0,
};
class Config {
    /**
     * Завантаження конфігурації з змінних середовища (Singleton pattern)
     */
    static load() {
        if (this.instance) {
            logger_1.default.debug('🔄 Повернення кешованої конфігурації');
            return this.instance;
        }
        try {
            logger_1.default.info('🔧 Завантаження конфігурації...');
            const config = {
                discord: this.loadDiscordConfig(),
                google: this.loadGoogleConfig(),
                ai: this.loadAIConfig(),
                redis: this.loadRedisConfig(),
                metrics: this.loadMetricsConfig(),
                security: this.loadSecurityConfig(),
                performance: this.loadPerformanceConfig(),
                logging: this.loadLoggingConfig(),
            };
            this.validate(config);
            this.instance = config;
            logger_1.default.info('✅ Конфігурація успішно завантажена та валідована');
            this.logConfigurationSummary(config);
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження конфігурації:', error);
            throw new Error(`Помилка конфігурації: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Завантаження Discord конфігурації
     */
    static loadDiscordConfig() {
        try {
            logger_1.default.debug('📡 Завантаження Discord конфігурації...');
            const token = this.getRequiredEnv('DISCORD_TOKEN');
            const clientId = this.getRequiredEnv('DISCORD_CLIENT_ID');
            const guildId = this.getRequiredEnv('DISCORD_GUILD_ID');
            // Валідація токена
            if (!token.startsWith('MTA') && !token.startsWith('OTk')) {
                logger_1.default.warn('⚠️ Discord токен може бути некоректним');
            }
            const config = {
                token,
                clientId,
                guildId,
                prefix: this.getEnv('DISCORD_PREFIX', CONFIG_CONSTANTS.DEFAULT_PREFIX),
                intents: this.parseIntents(this.getEnv('DISCORD_INTENTS', CONFIG_CONSTANTS.DEFAULT_INTENTS.join(','))),
            };
            logger_1.default.debug('✅ Discord конфігурація завантажена');
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Discord конфігурації:', error);
            throw error;
        }
    }
    /**
     * Парсинг Discord intents
     */
    static parseIntents(intentsString) {
        try {
            const intents = intentsString.split(',').map(intent => intent.trim());
            const validIntents = intents.filter(intent => CONFIG_CONSTANTS.DEFAULT_INTENTS.includes(intent) ||
                ['DirectMessages', 'GuildPresences', 'GuildVoiceStates'].includes(intent));
            if (validIntents.length !== intents.length) {
                logger_1.default.warn('⚠️ Деякі Discord intents некоректні:', intents.filter(intent => !validIntents.includes(intent)));
            }
            return validIntents.length > 0 ? validIntents : CONFIG_CONSTANTS.DEFAULT_INTENTS;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка парсингу Discord intents:', error);
            return CONFIG_CONSTANTS.DEFAULT_INTENTS;
        }
    }
    /**
     * Завантаження Google конфігурації
     */
    static loadGoogleConfig() {
        try {
            logger_1.default.debug('🌐 Завантаження Google конфігурації...');
            const config = {
                spreadsheetId: this.getRequiredEnv('GOOGLE_SPREADSHEET_ID'),
                driveFolderId: this.getRequiredEnv('GOOGLE_DRIVE_FOLDER_ID'),
                apiKey: this.getRequiredEnv('GOOGLE_API_KEY'),
                applicationCredentials: this.getRequiredEnv('GOOGLE_APPLICATION_CREDENTIALS'),
                appScriptUrl: this.getRequiredEnv('GOOGLE_APP_SCRIPT_URL'),
                sheetName: this.getEnv('GOOGLE_SHEET_NAME', 'Sheet1'),
            };
            // Валідація Google API ключа
            if (!config.apiKey.startsWith('AIza')) {
                logger_1.default.warn('⚠️ Google API ключ може бути некоректним');
            }
            // Завантаження credentials
            const credentials = this.loadGoogleCredentials();
            if (credentials) {
                config.credentials = credentials;
                logger_1.default.debug('✅ Google credentials завантажено');
            }
            else {
                logger_1.default.warn('⚠️ Google credentials не знайдено');
            }
            logger_1.default.debug('✅ Google конфігурація завантажена');
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Google конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Google credentials з файлу або змінних середовища
     */
    static loadGoogleCredentials() {
        try {
            const clientEmail = process.env['GOOGLE_CLIENT_EMAIL'];
            const privateKey = process.env['GOOGLE_PRIVATE_KEY'];
            const projectId = process.env['GOOGLE_PROJECT_ID'];
            // Спроба завантаження з файлу
            const credentialsPath = process.env['GOOGLE_APPLICATION_CREDENTIALS'];
            if (credentialsPath && (0, fs_1.existsSync)(credentialsPath)) {
                try {
                    const credentialsFile = (0, fs_1.readFileSync)(credentialsPath, 'utf8');
                    const credentials = JSON.parse(credentialsFile);
                    if (credentials.client_email && credentials.private_key && credentials.project_id) {
                        logger_1.default.debug('✅ Google credentials завантажено з файлу');
                        return credentials;
                    }
                }
                catch (fileError) {
                    logger_1.default.warn('⚠️ Помилка читання Google credentials файлу:', fileError);
                }
            }
            // Завантаження з змінних середовища
            if (clientEmail && privateKey && projectId) {
                const credentials = {
                    client_email: clientEmail,
                    private_key: privateKey.replace(/\\n/g, '\n'),
                    project_id: projectId,
                };
                logger_1.default.debug('✅ Google credentials завантажено з змінних середовища');
                return credentials;
            }
            return undefined;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Google credentials:', error);
            return undefined;
        }
    }
    /**
     * Завантаження AI конфігурації
     */
    static loadAIConfig() {
        try {
            logger_1.default.debug('🤖 Завантаження AI конфігурації...');
            const provider = this.getEnv('AI_PROVIDER', CONFIG_CONSTANTS.DEFAULT_AI_PROVIDER);
            const config = {
                provider,
                openai: {
                    apiKey: this.getRequiredEnv('OPENAI_API_KEY'),
                    model: this.getEnv('OPENAI_MODEL', CONFIG_CONSTANTS.DEFAULT_OPENAI_MODEL),
                    maxTokens: this.validateNumber(this.getEnv('OPENAI_MAX_TOKENS', CONFIG_CONSTANTS.DEFAULT_OPENAI_MAX_TOKENS.toString()), CONFIG_CONSTANTS.DEFAULT_OPENAI_MAX_TOKENS, 1, CONFIG_CONSTANTS.MAX_OPENAI_TOKENS),
                    temperature: this.validateNumber(this.getEnv('OPENAI_TEMPERATURE', CONFIG_CONSTANTS.DEFAULT_OPENAI_TEMPERATURE.toString()), CONFIG_CONSTANTS.DEFAULT_OPENAI_TEMPERATURE, CONFIG_CONSTANTS.MIN_TEMPERATURE, CONFIG_CONSTANTS.MAX_TEMPERATURE),
                },
                ollama: {
                    host: this.getEnv('OLLAMA_HOST', CONFIG_CONSTANTS.DEFAULT_OLLAMA_HOST),
                    model: this.getEnv('OLLAMA_MODEL', CONFIG_CONSTANTS.DEFAULT_OLLAMA_MODEL),
                },
            };
            // Валідація OpenAI API ключа
            if (!config.openai.apiKey.startsWith('sk-')) {
                logger_1.default.warn('⚠️ OpenAI API ключ може бути некоректним');
            }
            logger_1.default.debug('✅ AI конфігурація завантажена');
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження AI конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Redis конфігурації
     */
    static loadRedisConfig() {
        try {
            logger_1.default.debug('💾 Завантаження Redis конфігурації...');
            const config = {
                host: this.getEnv('REDIS_HOST', CONFIG_CONSTANTS.DEFAULT_REDIS_HOST),
                port: this.validateNumber(this.getEnv('REDIS_PORT', CONFIG_CONSTANTS.DEFAULT_REDIS_PORT.toString()), CONFIG_CONSTANTS.DEFAULT_REDIS_PORT, 1, 65535),
                password: this.getEnv('REDIS_PASSWORD'),
                database: this.validateNumber(this.getEnv('REDIS_DATABASE', CONFIG_CONSTANTS.DEFAULT_REDIS_DATABASE.toString()), CONFIG_CONSTANTS.DEFAULT_REDIS_DATABASE, 0, 15),
                enabled: this.getEnv('REDIS_ENABLED', 'true').toLowerCase() === 'true',
                url: this.getEnv('REDIS_URL'),
            };
            logger_1.default.debug('✅ Redis конфігурація завантажена');
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Redis конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Metrics конфігурації
     */
    static loadMetricsConfig() {
        try {
            logger_1.default.debug('📊 Завантаження Metrics конфігурації...');
            const config = {
                enabled: this.getEnv('METRICS_ENABLED', 'true').toLowerCase() === 'true',
                port: this.validateNumber(this.getEnv('METRICS_PORT', CONFIG_CONSTANTS.DEFAULT_METRICS_PORT.toString()), CONFIG_CONSTANTS.DEFAULT_METRICS_PORT, 1024, 65535),
                path: this.getEnv('METRICS_PATH', CONFIG_CONSTANTS.DEFAULT_METRICS_PATH),
            };
            logger_1.default.debug('✅ Metrics конфігурація завантажена');
            return config;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Metrics конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Security конфігурації
     */
    static loadSecurityConfig() {
        try {
            logger_1.default.debug('🔒 Завантаження Security конфігурації...');
            return {
                rateLimitWindow: this.validateNumber(this.getEnv('RATE_LIMIT_WINDOW', '60000'), 60000, 1000, 300000),
                rateLimitMax: this.validateNumber(this.getEnv('RATE_LIMIT_MAX', '100'), 100, 1, 1000),
                adminRole: this.getEnv('ADMIN_ROLE', 'Admin'),
                botUserRole: this.getEnv('BOT_USER_ROLE', 'Bot User'),
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Security конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Performance конфігурації
     */
    static loadPerformanceConfig() {
        try {
            logger_1.default.debug('⚡ Завантаження Performance конфігурації...');
            return {
                cacheTTL: this.validateNumber(this.getEnv('CACHE_TTL', '300000'), 300000, 1000, 3600000),
                maxSearchResults: this.validateNumber(this.getEnv('MAX_SEARCH_RESULTS', '100'), 100, 1, 1000),
                maxAnalysisRows: this.validateNumber(this.getEnv('MAX_ANALYSIS_ROWS', '1000'), 1000, 1, 10000),
                requestTimeout: this.validateNumber(this.getEnv('REQUEST_TIMEOUT', '30000'), 30000, 1000, 300000),
                maxRetries: this.validateNumber(this.getEnv('MAX_RETRIES', '3'), 3, 0, 10),
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Performance конфігурації:', error);
            throw error;
        }
    }
    /**
     * Завантаження Logging конфігурації
     */
    static loadLoggingConfig() {
        try {
            logger_1.default.debug('📝 Завантаження Logging конфігурації...');
            return {
                level: this.getEnv('LOG_LEVEL', 'info'),
                maxFiles: this.validateNumber(this.getEnv('LOG_MAX_FILES', '5'), 5, 1, 50),
                maxSize: this.getEnv('LOG_MAX_SIZE', '10m'),
                directory: this.getEnv('LOG_DIRECTORY', (0, path_1.join)(process.cwd(), 'logs')),
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завантаження Logging конфігурації:', error);
            throw error;
        }
    }
    /**
     * Валідація числових значень
     */
    static validateNumber(value, defaultValue, min, max) {
        try {
            const num = parseInt(value, 10);
            if (isNaN(num) || num < min || num > max) {
                logger_1.default.warn(`⚠️ Некоректне значення ${value}, використовую ${defaultValue}`);
                return defaultValue;
            }
            return num;
        }
        catch (error) {
            logger_1.default.warn(`⚠️ Помилка парсингу числа ${value}, використовую ${defaultValue}`);
            return defaultValue;
        }
    }
    /**
     * Валідація конфігурації
     */
    static validate(config) {
        logger_1.default.info('🔍 Валідація конфігурації...');
        const errors = [];
        // Валідація Discord
        if (!config.discord.token)
            errors.push('DISCORD_TOKEN is required');
        if (!config.discord.clientId)
            errors.push('DISCORD_CLIENT_ID is required');
        if (!config.discord.guildId)
            errors.push('DISCORD_GUILD_ID is required');
        // Валідація Google
        if (!config.google.spreadsheetId)
            errors.push('GOOGLE_SPREADSHEET_ID is required');
        if (!config.google.apiKey)
            errors.push('GOOGLE_API_KEY is required');
        if (!config.google.appScriptUrl)
            errors.push('GOOGLE_APP_SCRIPT_URL is required');
        // Валідація AI
        if (config.ai.provider === 'openai' && !config.ai.openai.apiKey) {
            errors.push('OPENAI_API_KEY is required when AI_PROVIDER is openai');
        }
        if (errors.length > 0) {
            const errorMessage = `Configuration validation failed:\n${errors.join('\n')}`;
            logger_1.default.error('❌ Помилки валідації конфігурації:', errors);
            throw new Error(errorMessage);
        }
        logger_1.default.info('✅ Конфігурація валідна');
    }
    /**
     * Логування підсумку конфігурації
     */
    static logConfigurationSummary(config) {
        try {
            logger_1.default.info('📋 Підсумок конфігурації:', {
                discord: {
                    clientId: config.discord.clientId,
                    guildId: config.discord.guildId,
                    prefix: config.discord.prefix,
                    intents: config.discord.intents.length,
                },
                google: {
                    spreadsheetId: config.google.spreadsheetId,
                    sheetName: config.google.sheetName,
                    hasCredentials: !!config.google.credentials,
                },
                ai: {
                    provider: config.ai.provider,
                    model: config.ai.provider === 'openai' ? config.ai.openai.model : config.ai.ollama.model,
                },
                redis: {
                    enabled: config.redis.enabled,
                    host: config.redis.host,
                    port: config.redis.port,
                },
                metrics: {
                    enabled: config.metrics.enabled,
                    port: config.metrics.port,
                },
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка логування підсумку конфігурації:', error);
        }
    }
    /**
     * Отримання обов'язкової змінної середовища
     */
    static getRequiredEnv(key) {
        const value = process.env[key];
        if (!value) {
            const error = `Required environment variable ${key} is not set`;
            logger_1.default.error(`❌ ${error}`);
            throw new Error(error);
        }
        return value;
    }
    /**
     * Отримання змінної середовища з значенням за замовчуванням
     */
    static getEnv(key, defaultValue) {
        const value = process.env[key];
        if (!value) {
            logger_1.default.debug(`🔧 Використовую значення за замовчуванням для ${key}: ${defaultValue}`);
        }
        return value || defaultValue;
    }
    /**
     * Очищення кешу конфігурації
     */
    static clearCache() {
        this.instance = null;
        this.configCache.clear();
        logger_1.default.debug('🧹 Кеш конфігурації очищено');
    }
    /**
     * Перезавантаження конфігурації
     */
    static reload() {
        logger_1.default.info('🔄 Перезавантаження конфігурації...');
        this.clearCache();
        return this.load();
    }
}
exports.Config = Config;
Config.instance = null;
Config.configCache = new Map();
//# sourceMappingURL=Config.js.map