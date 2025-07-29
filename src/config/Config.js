/**
 * Централізована конфігурація бота
 */

class Config {
  constructor() {
    this.validateEnvironment();
    this.loadConfig();
  }

  /**
   * Валідація обов'язкових змінних середовища
   */
  validateEnvironment() {
    const required = [
      'DISCORD_TOKEN',
      'GOOGLE_SPREADSHEET_ID'
    ];

    const missing = required.filter(key => !process.env[key]);
    
    if (missing.length > 0) {
      throw new Error(`❌ Відсутні обов'язкові змінні середовища: ${missing.join(', ')}`);
    }
  }

  /**
   * Валідація конфігурації
   */
  async validate() {
    try {
      // Валідація змінних середовища
      this.validateEnvironment();

      // Валідація Discord конфігурації
      if (!this.discord.token) {
        throw new Error('Discord токен не налаштовано');
      }

      // Валідація Google конфігурації
      if (!this.google.spreadsheetId) {
        throw new Error('Google Spreadsheet ID не налаштовано');
      }

      // Валідація AI конфігурації
      if (this.ai.provider === 'openai' && !this.ai.openai.apiKey) {
        logger.warn('OpenAI API ключ не налаштовано - AI функції будуть недоступні');
      }

      // Валідація Redis конфігурації
      if (this.redis.enabled && !this.redis.host) {
        throw new Error('Redis увімкнено, але хост не налаштовано');
      }

      logger.info('✅ Конфігурація валідована');
      return true;
    } catch (error) {
      logger.error('❌ Помилка валідації конфігурації:', error);
      throw error;
    }
  }

  /**
   * Завантаження конфігурації
   */
  loadConfig() {
    // Discord Configuration
    this.discord = {
      token: process.env.DISCORD_TOKEN,
      clientId: process.env.DISCORD_CLIENT_ID,
      guildId: process.env.DISCORD_GUILD_ID,
      prefix: process.env.BOT_PREFIX || '!',
      intents: [
        'Guilds',
        'GuildMessages', 
        'MessageContent',
        'GuildMessageReactions'
      ]
    };

    // Google Services Configuration
    this.google = {
      spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
      driveFolderId: process.env.GOOGLE_DRIVE_FOLDER_ID,
      apiKey: process.env.GOOGLE_API_KEY,
      applicationCredentials: process.env.GOOGLE_APPLICATION_CREDENTIALS,
      appScriptUrl: process.env.APP_SCRIPT_URL,
      sheetName: process.env.SHEET_NAME || 'Аркуш1'
    };

    // AI Configuration
    this.ai = {
      provider: process.env.AI_PROVIDER || 'openai',
      openai: {
        apiKey: process.env.OPENAI_API_KEY,
        model: process.env.OPENAI_MODEL || 'gpt-3.5-turbo',
        maxTokens: parseInt(process.env.OPENAI_MAX_TOKENS) || 2000,
        temperature: parseFloat(process.env.OPENAI_TEMPERATURE) || 0.7
      },
      ollama: {
        host: process.env.OLLAMA_HOST || 'http://localhost:11434',
        model: process.env.OLLAMA_MODEL || 'llama2'
      }
    };

    // Redis Configuration
    this.redis = {
      host: process.env.REDIS_HOST || 'localhost',
      port: parseInt(process.env.REDIS_PORT) || 6379,
      password: process.env.REDIS_PASSWORD,
      database: parseInt(process.env.REDIS_DB) || 0,
      enabled: process.env.REDIS_ENABLED === 'true'
    };

    // Security Configuration
    this.security = {
      rateLimit: {
        enabled: process.env.RATE_LIMIT_ENABLED !== 'false',
        windowMs: parseInt(process.env.RATE_LIMIT_WINDOW) || 900000, // 15 хвилин
        maxRequests: parseInt(process.env.RATE_LIMIT_MAX) || 100
      },
      roles: {
        admin: process.env.ADMIN_ROLE || 'Адміністратор',
        botUser: process.env.BOT_USER_ROLE || 'Бот-Користувач',
        sheetsAccess: process.env.SHEETS_ACCESS_ROLE || 'Sheets-Доступ',
        aiAccess: process.env.AI_ACCESS_ROLE || 'AI-Доступ',
        exportAccess: process.env.EXPORT_ACCESS_ROLE || 'Експорт-Доступ'
      },
      logLevel: process.env.SECURITY_LOG_LEVEL || 'info'
    };

    // File Processing Configuration
    this.files = {
      maxFileSize: parseInt(process.env.MAX_FILE_SIZE) || 10 * 1024 * 1024, // 10MB
      tempDir: process.env.TEMP_DIR || './data/tmp',
      cleanupInterval: parseInt(process.env.FILE_CLEANUP_INTERVAL) || 3600000, // 1 година
      supportedFormats: ['pdf', 'docx', 'doc', 'txt', 'gdoc'],
      downloadTimeout: parseInt(process.env.DOWNLOAD_TIMEOUT) || 30000 // 30 секунд
    };

    // Performance Configuration
    this.performance = {
      cacheTtl: parseInt(process.env.CACHE_TTL) || 300000, // 5 хвилин
      maxSearchResults: parseInt(process.env.MAX_SEARCH_RESULTS) || 20,
      maxAnalysisRows: parseInt(process.env.MAX_ANALYSIS_ROWS) || 50,
      requestTimeout: parseInt(process.env.REQUEST_TIMEOUT) || 30000, // 30 секунд
      maxRetries: parseInt(process.env.MAX_RETRIES) || 3
    };

    // Logging Configuration
    this.logging = {
      level: process.env.LOG_LEVEL || 'info',
      maxFiles: parseInt(process.env.LOG_MAX_FILES) || 5,
      maxSize: process.env.LOG_MAX_SIZE || '10m'
    };

    // Metrics Configuration
    this.metrics = {
      enabled: process.env.METRICS_ENABLED === 'true',
      port: parseInt(process.env.PROMETHEUS_PORT) || 9090,
      path: process.env.METRICS_PATH || '/metrics'
    };

    // Export Configuration
    this.export = {
      maxFileSize: parseInt(process.env.EXPORT_MAX_FILE_SIZE) || 25 * 1024 * 1024, // 25MB
      tempFileTtl: parseInt(process.env.TEMP_FILE_TTL) || 60000, // 1 хвилина
      formats: ['xlsx', 'csv', 'pdf', 'docx'],
      includeMetadata: process.env.INCLUDE_METADATA !== 'false'
    };
  }

  /**
   * Отримання URL для Google Sheets API
   */
  getGoogleSheetsUrl(range = null) {
    const sheetRange = range || this.google.sheetName;
    return `https://sheets.googleapis.com/v4/spreadsheets/${this.google.spreadsheetId}/values/${sheetRange}?key=${this.google.apiKey}`;
  }

  /**
   * Отримання URL для Google Sheets Cells API
   */
  getGoogleSheetsCellsUrl(range) {
    return `https://sheets.googleapis.com/v4/spreadsheets/${this.google.spreadsheetId}/values/${range}?key=${this.google.apiKey}`;
  }

  /**
   * Перевірка чи увімкнено AI
   */
  isAIEnabled() {
    return !!(this.ai.openai.apiKey || this.ai.ollama.host);
  }

  /**
   * Перевірка чи увімкнено метрики
   */
  isMetricsEnabled() {
    return this.metrics.enabled;
  }

  /**
   * Отримання середовища
   */
  getEnvironment() {
    return process.env.NODE_ENV || 'development';
  }

  /**
   * Перевірка чи це продакшен
   */
  isProduction() {
    return this.getEnvironment() === 'production';
  }

  /**
   * Отримання всієї конфігурації
   */
  getAll() {
    return {
      discord: this.discord,
      google: this.google,
      ai: this.ai,
      redis: this.redis,
      security: this.security,
      files: this.files,
      performance: this.performance,
      logging: this.logging,
      metrics: this.metrics,
      export: this.export,
      environment: this.getEnvironment()
    };
  }

  /**
   * Валідація конфігурації
   */
  validate() {
    const errors = [];

    // Перевірка Discord
    if (!this.discord.token) {
      errors.push('DISCORD_TOKEN is required');
    }

    // Перевірка Google
    if (!this.google.spreadsheetId) {
      errors.push('GOOGLE_SPREADSHEET_ID is required');
    }

    // Перевірка AI
    if (this.isAIEnabled()) {
      if (this.ai.provider === 'openai' && !this.ai.openai.apiKey) {
        errors.push('OPENAI_API_KEY is required when AI provider is openai');
      }
    }

    // Перевірка Redis
    if (this.redis.enabled && !this.redis.host) {
      errors.push('REDIS_HOST is required when Redis is enabled');
    }

    if (errors.length > 0) {
      throw new Error(`Configuration validation failed: ${errors.join(', ')}`);
    }

    return true;
  }

  /**
   * Отримання конфігурації для конкретного модуля
   */
  getModuleConfig(moduleName) {
    switch (moduleName) {
      case 'discord':
        return this.discord;
      case 'google':
        return this.google;
      case 'ai':
        return this.ai;
      case 'redis':
        return this.redis;
      case 'security':
        return this.security;
      case 'files':
        return this.files;
      case 'performance':
        return this.performance;
      case 'logging':
        return this.logging;
      case 'metrics':
        return this.metrics;
      case 'export':
        return this.export;
      default:
        throw new Error(`Unknown module: ${moduleName}`);
    }
  }

  /**
   * Отримання значення з конфігурації
   */
  get(key, defaultValue = null) {
    const keys = key.split('.');
    let value = this;

    for (const k of keys) {
      if (value && typeof value === 'object' && k in value) {
        value = value[k];
      } else {
        return defaultValue;
      }
    }

    return value !== undefined ? value : defaultValue;
  }

  /**
   * Встановлення значення в конфігурацію
   */
  set(key, value) {
    const keys = key.split('.');
    const lastKey = keys.pop();
    let obj = this;

    for (const k of keys) {
      if (!(k in obj) || typeof obj[k] !== 'object') {
        obj[k] = {};
      }
      obj = obj[k];
    }

    obj[lastKey] = value;
  }
}

// Експорт екземпляру
const config = new Config();

module.exports = config; 