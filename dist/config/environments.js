"use strict";
/**
 * 🌍 Конфігурація середовищ розгортання
 * Discord AI Assistant Bot v2.3.0
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.production = exports.staging = exports.testing = exports.development = void 0;
exports.getConfig = getConfig;
exports.validateConfig = validateConfig;
exports.getValidatedConfig = getValidatedConfig;
const path_1 = __importDefault(require("path"));
// Базові налаштування для всіх середовищ
const baseConfig = {
    // Налаштування логування
    logging: {
        level: process.env.LOG_LEVEL || 'info',
        maxFiles: parseInt(process.env.LOG_MAX_FILES || '5'),
        maxSize: process.env.LOG_MAX_SIZE || '10m',
        directory: path_1.default.join(process.cwd(), 'logs'),
    },
    // Налаштування метрик
    metrics: {
        enabled: process.env.METRICS_ENABLED === 'true',
        port: parseInt(process.env.METRICS_PORT || '9090'),
        path: process.env.METRICS_PATH || '/metrics',
    },
    // Налаштування безпеки
    security: {
        rateLimitWindow: parseInt(process.env.RATE_LIMIT_WINDOW || '60000'),
        rateLimitMax: parseInt(process.env.RATE_LIMIT_MAX || '100'),
        adminRole: process.env.ADMIN_ROLE || 'Admin',
        botUserRole: process.env.BOT_USER_ROLE || 'Bot User',
    },
    // Налаштування продуктивності
    performance: {
        cacheTTL: parseInt(process.env.CACHE_TTL || '300000'),
        maxSearchResults: parseInt(process.env.MAX_SEARCH_RESULTS || '100'),
        maxAnalysisRows: parseInt(process.env.MAX_ANALYSIS_ROWS || '1000'),
        requestTimeout: parseInt(process.env.REQUEST_TIMEOUT || '30000'),
        maxRetries: parseInt(process.env.MAX_RETRIES || '3'),
    },
};
// Конфігурація для розробки
const development = {
    ...baseConfig,
    name: 'development',
    nodeEnv: 'development',
    // Discord налаштування
    discord: {
        token: process.env.DISCORD_TOKEN || '',
        clientId: process.env.CLIENT_ID || '',
        guildId: process.env.GUILD_ID || '',
        intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    // Google Services
    google: {
        apiKey: process.env.GOOGLE_API_KEY || '',
        appScriptUrl: process.env.APP_SCRIPT_URL || '',
        sheetName: process.env.SHEET_NAME || 'Data',
    },
    // AI налаштування
    ai: {
        openai: {
            apiKey: process.env.OPENAI_API_KEY || '',
            model: process.env.OPENAI_MODEL || 'gpt-3.5-turbo',
            maxTokens: parseInt(process.env.OPENAI_MAX_TOKENS || '1000'),
            temperature: parseFloat(process.env.OPENAI_TEMPERATURE || '0.7'),
        },
        ollama: {
            enabled: process.env.OLLAMA_ENABLED === 'true',
            url: process.env.OLLAMA_URL || 'http://localhost:11434',
            model: process.env.OLLAMA_MODEL || 'llama2',
        },
    },
    // Redis налаштування
    redis: {
        enabled: process.env.REDIS_ENABLED === 'true',
        host: process.env.REDIS_HOST || 'localhost',
        port: parseInt(process.env.REDIS_PORT || '6379'),
        password: process.env.REDIS_PASSWORD || null,
        db: parseInt(process.env.REDIS_DB || '0'),
    },
    // Налаштування для розробки
    development: {
        debug: true,
        verbose: true,
        hotReload: true,
        testMode: false,
    },
};
exports.development = development;
// Конфігурація для тестування
const testing = {
    ...baseConfig,
    name: 'testing',
    nodeEnv: 'testing',
    // Discord налаштування (тестовий сервер)
    discord: {
        token: process.env.TEST_DISCORD_TOKEN || process.env.DISCORD_TOKEN || '',
        clientId: process.env.TEST_CLIENT_ID || process.env.CLIENT_ID || '',
        guildId: process.env.TEST_GUILD_ID || process.env.GUILD_ID || '',
        intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    // Google Services (тестові)
    google: {
        apiKey: process.env.TEST_GOOGLE_API_KEY || process.env.GOOGLE_API_KEY || '',
        appScriptUrl: process.env.TEST_APP_SCRIPT_URL || process.env.APP_SCRIPT_URL || '',
        sheetName: process.env.TEST_SHEET_NAME || 'TestData',
    },
    // AI налаштування (тестові)
    ai: {
        openai: {
            apiKey: process.env.TEST_OPENAI_API_KEY || process.env.OPENAI_API_KEY || '',
            model: 'gpt-3.5-turbo',
            maxTokens: 500,
            temperature: 0.5,
        },
        ollama: {
            enabled: true,
            url: 'http://localhost:11434',
            model: 'llama2',
        },
    },
    // Redis налаштування (тестовий)
    redis: {
        enabled: true,
        host: 'localhost',
        port: 6379,
        password: null,
        db: 1,
    },
    // Налаштування для тестування
    testing: {
        debug: true,
        verbose: true,
        hotReload: false,
        testMode: true,
        mockExternalServices: true,
    },
};
exports.testing = testing;
// Конфігурація для staging
const staging = {
    ...baseConfig,
    name: 'staging',
    nodeEnv: 'staging',
    // Discord налаштування
    discord: {
        token: process.env.STAGING_DISCORD_TOKEN || process.env.DISCORD_TOKEN || '',
        clientId: process.env.STAGING_CLIENT_ID || process.env.CLIENT_ID || '',
        guildId: process.env.STAGING_GUILD_ID || process.env.GUILD_ID || '',
        intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    // Google Services
    google: {
        apiKey: process.env.STAGING_GOOGLE_API_KEY || process.env.GOOGLE_API_KEY || '',
        appScriptUrl: process.env.STAGING_APP_SCRIPT_URL || process.env.APP_SCRIPT_URL || '',
        sheetName: process.env.STAGING_SHEET_NAME || 'StagingData',
    },
    // AI налаштування
    ai: {
        openai: {
            apiKey: process.env.STAGING_OPENAI_API_KEY || process.env.OPENAI_API_KEY || '',
            model: process.env.OPENAI_MODEL || 'gpt-3.5-turbo',
            maxTokens: parseInt(process.env.OPENAI_MAX_TOKENS || '1000'),
            temperature: parseFloat(process.env.OPENAI_TEMPERATURE || '0.7'),
        },
        ollama: {
            enabled: process.env.OLLAMA_ENABLED === 'true',
            url: process.env.OLLAMA_URL || 'http://localhost:11434',
            model: process.env.OLLAMA_MODEL || 'llama2',
        },
    },
    // Redis налаштування
    redis: {
        enabled: process.env.REDIS_ENABLED === 'true',
        host: process.env.REDIS_HOST || 'localhost',
        port: parseInt(process.env.REDIS_PORT || '6379'),
        password: process.env.REDIS_PASSWORD || null,
        db: parseInt(process.env.REDIS_DB || '0'),
    },
    // Налаштування для staging
    staging: {
        debug: false,
        verbose: true,
        hotReload: false,
        testMode: false,
        monitoring: true,
    },
};
exports.staging = staging;
// Конфігурація для продакшену
const production = {
    ...baseConfig,
    name: 'production',
    nodeEnv: 'production',
    // Discord налаштування
    discord: {
        token: process.env.DISCORD_TOKEN || '',
        clientId: process.env.CLIENT_ID || '',
        guildId: process.env.GUILD_ID || '',
        intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    // Google Services
    google: {
        apiKey: process.env.GOOGLE_API_KEY || '',
        appScriptUrl: process.env.APP_SCRIPT_URL || '',
        sheetName: process.env.SHEET_NAME || 'Data',
    },
    // AI налаштування
    ai: {
        openai: {
            apiKey: process.env.OPENAI_API_KEY || '',
            model: process.env.OPENAI_MODEL || 'gpt-4',
            maxTokens: parseInt(process.env.OPENAI_MAX_TOKENS || '2000'),
            temperature: parseFloat(process.env.OPENAI_TEMPERATURE || '0.7'),
        },
        ollama: {
            enabled: process.env.OLLAMA_ENABLED === 'true',
            url: process.env.OLLAMA_URL || 'http://localhost:11434',
            model: process.env.OLLAMA_MODEL || 'llama2',
        },
    },
    // Redis налаштування
    redis: {
        enabled: process.env.REDIS_ENABLED === 'true',
        host: process.env.REDIS_HOST || 'localhost',
        port: parseInt(process.env.REDIS_PORT || '6379'),
        password: process.env.REDIS_PASSWORD || null,
        db: parseInt(process.env.REDIS_DB || '0'),
    },
    // Налаштування для продакшену
    production: {
        debug: false,
        verbose: false,
        hotReload: false,
        testMode: false,
        monitoring: true,
        clustering: true,
        loadBalancing: true,
    },
};
exports.production = production;
// Функція для отримання конфігурації за середовищем
function getConfig(environment = null) {
    const env = environment || process.env.NODE_ENV || 'development';
    switch (env.toLowerCase()) {
        case 'development':
        case 'dev':
            return development;
        case 'testing':
        case 'test':
            return testing;
        case 'staging':
        case 'stage':
            return staging;
        case 'production':
        case 'prod':
            return production;
        default:
            console.warn(`Невідоме середовище: ${env}, використовую development`);
            return development;
    }
}
// Функція для валідації конфігурації
function validateConfig(config) {
    const errors = [];
    // Перевірка обов'язкових змінних
    if (!config.discord.token) {
        errors.push('DISCORD_TOKEN не встановлено');
    }
    if (!config.discord.clientId) {
        errors.push('CLIENT_ID не встановлено');
    }
    if (!config.discord.guildId) {
        errors.push('GUILD_ID не встановлено');
    }
    if (!config.google.apiKey) {
        errors.push('GOOGLE_API_KEY не встановлено');
    }
    if (!config.ai.openai.apiKey) {
        errors.push('OPENAI_API_KEY не встановлено');
    }
    if (errors.length > 0) {
        throw new Error(`Помилки конфігурації:\n${errors.join('\n')}`);
    }
    return true;
}
// Функція для отримання конфігурації з валідацією
function getValidatedConfig(environment = null) {
    const config = getConfig(environment);
    validateConfig(config);
    return config;
}
//# sourceMappingURL=environments.js.map