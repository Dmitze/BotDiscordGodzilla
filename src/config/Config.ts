/**
 * Клас для управління конфігурацією додатку
 * Завантажує та валідує налаштування з змінних середовища
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import { readFileSync, existsSync } from 'fs';
import { join } from 'path';
import type {
  BotConfig,
  DiscordConfig,
  GoogleConfig,
  AIConfig,
  RedisConfig,
  MetricsConfig,
} from '@/types';
import type { DriveConfig } from '@/types/drive';
import logger from '@/utils/logger';

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
  DEFAULT_METRICS_PORT: 9091,
  DEFAULT_METRICS_PATH: '/metrics',
  MAX_OPENAI_TOKENS: 4000,
  MAX_TEMPERATURE: 2.0,
  MIN_TEMPERATURE: 0.0,
} as const;

export class Config {
  private static instance: BotConfig | null = null;
  private static readonly configCache = new Map<string, any>();
  
  /**
   * Повертає поточну конфігурацію з кешу або виконує завантаження
   */
  public static get(): BotConfig {
    if (this.instance) {
      return this.instance;
    }
    return this.load();
  }
  
  /**
   * Завантаження конфігурації (Singleton)
   */
  public static load(): BotConfig {
    try {
      logger.info('🔧 Завантаження конфігурації...');

      const config: BotConfig = {
        discord: this.loadDiscordConfig(),
        google: this.loadGoogleConfig(),
        ai: this.loadAIConfig(),
        redis: this.loadRedisConfig(),
        metrics: this.loadMetricsConfig(),
        security: this.loadSecurityConfig(),
        performance: this.loadPerformanceConfig(),
        logging: this.loadLoggingConfig(),
        drive: this.loadDriveConfig(),
      };

      this.validate(config);
      this.instance = config;

      logger.info('✅ Конфігурація успішно завантажена та валідована');
      this.logConfigurationSummary(config);

      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження конфігурації:', error as any);
      throw new Error(
        `Помилка конфігурації: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Завантаження Drive конфігурації
   */
  private static loadDriveConfig(): DriveConfig {
    try {
      logger.debug('🗂️ Завантаження Drive конфігурації...');

      const csv = (v: string) =>
        v.trim() === ''
          ? []
          : v
              .split(',')
              .map(s => s.trim())
              .filter(Boolean);

      const allowedMimeRaw = this.getEnv('DRIVE_ALLOWED_MIME', '*');
      const allowedMime = allowedMimeRaw === '*' ? ['*'] : csv(allowedMimeRaw);

      const ownerAllowlist = csv(this.getEnv('DRIVE_OWNER_ALLOWLIST', ''));

      const config: DriveConfig = {
        folderId: this.getRequiredEnv('GOOGLE_DRIVE_FOLDER_ID'),
        pageSize: this.validateNumber(this.getEnv('DRIVE_PAGE_SIZE', '25'), 25, 5, 100),
        allowedMime,
        fileMaxSizeMb: this.validateNumber(this.getEnv('FILE_MAX_SIZE_MB', '8'), 8, 1, 24),
        enableTextIndex: this.getEnv('DRIVE_ENABLE_TEXT_INDEX', 'false').toLowerCase() === 'true',
        indexCron: this.getEnv('DRIVE_INDEX_CRON', '*/30 * * * *'),
        maxConcurrency: this.validateNumber(this.getEnv('DRIVE_MAX_CONCURRENCY', '3'), 3, 1, 10),
        ttlListSec: this.validateNumber(this.getEnv('DRIVE_LIST_TTL_SEC', '60'), 60, 10, 3600),
        ttlTextSec: this.validateNumber(this.getEnv('DRIVE_TEXT_TTL_SEC', '21600'), 21600, 60, 604800),
        ownerAllowlist,
        hideWebLink: this.getEnv('DRIVE_HIDE_WEBLINK', 'true').toLowerCase() === 'true',
      };

      logger.debug('✅ Drive конфігурація завантажена');
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження Drive конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Discord конфігурації
   */
  private static loadDiscordConfig(): DiscordConfig {
    try {
      logger.debug('📡 Завантаження Discord конфігурації...');

      const token = this.getRequiredEnv('DISCORD_TOKEN');
      const clientId = this.getRequiredEnv('DISCORD_CLIENT_ID');
      const guildId = this.getEnv('DISCORD_GUILD_ID', '');

      // Валідація токена
      if (!token.startsWith('MTA') && !token.startsWith('OTk')) {
        logger.warn('⚠️ Discord токен може бути некоректним');
      }

      // Флаги режимов
      const enableChat = this.getEnv('ENABLE_CHAT', 'true').toLowerCase() === 'true';
      const enableSlash = this.getEnv('ENABLE_SLASH', 'false').toLowerCase() === 'true';
      const enableMessageContentIntent =
        this.getEnv('ENABLE_MESSAGE_CONTENT_INTENT', 'true').toLowerCase() === 'true';

      let parsedIntents = this.parseIntents(
        this.getEnv('DISCORD_INTENTS', CONFIG_CONSTANTS.DEFAULT_INTENTS.join(','))
      );

      // Принудительно управляем MessageContent через флаг
      const hasMessageContent = parsedIntents.includes('MessageContent');
      if (enableMessageContentIntent && !hasMessageContent) {
        parsedIntents = [...parsedIntents, 'MessageContent'];
      }
      if (!enableMessageContentIntent && hasMessageContent) {
        parsedIntents = parsedIntents.filter(i => i !== 'MessageContent');
      }

      const config: DiscordConfig = {
        token,
        clientId,
        guildId,
        prefix: this.getEnv('DISCORD_PREFIX', CONFIG_CONSTANTS.DEFAULT_PREFIX),
        intents: parsedIntents,
        enableChat,
        enableSlash,
        enableMessageContentIntent,
      };

      // Попередження: чат включений, але MessageContent вимкнено
      if (config.enableChat && !config.enableMessageContentIntent) {
        logger.warn('⚠️ ENABLE_CHAT=true, але ENABLE_MESSAGE_CONTENT_INTENT=false — чат-режим не працюватиме.');
      }

      logger.debug('✅ Discord конфігурація завантажена');
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження Discord конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Парсинг Discord intents
   */
  private static parseIntents(intentsString: string): string[] {
    try {
      const intents = intentsString.split(',').map(intent => intent.trim());
      const allowed = new Set<string>([
        ...(CONFIG_CONSTANTS.DEFAULT_INTENTS as unknown as string[]),
        'DirectMessages',
        'GuildPresences',
        'GuildVoiceStates',
      ]);
      const validIntents = intents.filter(intent => allowed.has(intent));

      if (validIntents.length !== intents.length) {
        logger.warn(
          '⚠️ Деякі Discord intents некоректні:',
          intents.filter(intent => !validIntents.includes(intent))
        );
      }

      return validIntents.length > 0
        ? validIntents
        : ([...CONFIG_CONSTANTS.DEFAULT_INTENTS] as unknown as string[]);
    } catch (error) {
      logger.error('❌ Помилка парсингу Discord intents:', error as any);
      return [...(CONFIG_CONSTANTS.DEFAULT_INTENTS as unknown as string[])];
    }
  }

  /**
   * Завантаження Google конфігурації
   */
  private static loadGoogleConfig(): GoogleConfig {
    try {
      logger.debug('🌐 Завантаження Google конфігурації...');

      const config: GoogleConfig = {
        spreadsheetId: this.getRequiredEnv('GOOGLE_SPREADSHEET_ID'),
        driveFolderId: this.getRequiredEnv('GOOGLE_DRIVE_FOLDER_ID'),
        apiKey: this.getRequiredEnv('GOOGLE_API_KEY'),
        applicationCredentials: this.getRequiredEnv('GOOGLE_APPLICATION_CREDENTIALS'),
        appScriptUrl: this.getRequiredEnv('GOOGLE_APP_SCRIPT_URL'),
        sheetName: this.getEnv('GOOGLE_SHEET_NAME', 'Sheet1'),
        // OCR settings with safe offline defaults
        ocrProvider: (this.getEnv('OCR_PROVIDER', 'off') as 'vision' | 'tesseract' | 'off'),
        ocrCacheTTL: this.validateNumber(this.getEnv('OCR_CACHE_TTL', '3600'), 3600, 60, 604800),
        tesseractLangs: this.getEnv('TESSERACT_LANGS', 'eng'),
        tesseractLangPath: this.getEnv('TESSERACT_LANG_PATH', ''),
        // Analytics cache TTL
        analyticsCacheTTL: this.validateNumber(
          this.getEnv('ANALYTICS_CACHE_TTL', '900'),
          900,
          60,
          604800
        ),
      };

      // Валідація Google API ключа
      if (!config.apiKey.startsWith('AIza')) {
        logger.warn('⚠️ Google API ключ може бути некоректним');
      }

      // Завантаження credentials
      const credentials = this.loadGoogleCredentials();
      if (credentials) {
        config.credentials = credentials;
        logger.debug('✅ Google credentials завантажено');
      } else {
        logger.warn('⚠️ Google credentials не знайдено');
      }

      logger.debug('✅ Google конфігурація завантажена');
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження Google конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Google credentials з файлу або змінних середовища
   */
  private static loadGoogleCredentials(): GoogleConfig['credentials'] {
    try {
      const clientEmail = process.env['GOOGLE_CLIENT_EMAIL'];
      const privateKey = process.env['GOOGLE_PRIVATE_KEY'];
      const projectId = process.env['GOOGLE_PROJECT_ID'];

      // Спроба завантаження з файлу
      const credentialsPath = process.env['GOOGLE_APPLICATION_CREDENTIALS'];
      if (credentialsPath && existsSync(credentialsPath)) {
        try {
          const credentialsFile = readFileSync(credentialsPath, 'utf8');
          const credentials = JSON.parse(credentialsFile);

          if (credentials.client_email && credentials.private_key && credentials.project_id) {
            logger.debug('✅ Google credentials завантажено з файлу');
            return credentials;
          }
        } catch (fileError) {
          logger.warn('⚠️ Помилка читання Google credentials файлу:', fileError as any);
        }
      }

      // Завантаження з змінних середовища
      if (clientEmail && privateKey && projectId) {
        const credentials = {
          client_email: clientEmail,
          private_key: privateKey.replace(/\\n/g, '\n'),
          project_id: projectId,
        };

        logger.debug('✅ Google credentials завантажено з змінних середовища');
        return credentials;
      }

      return undefined;
    } catch (error) {
      logger.error('❌ Помилка завантаження Google credentials:', error as any);
      return undefined;
    }
  }

  /**
   * Завантаження AI конфігурації
   */
  private static loadAIConfig(): AIConfig {
    try {
      logger.debug('🤖 Завантаження AI конфігурації...');

      const provider = this.getEnv('AI_PROVIDER', CONFIG_CONSTANTS.DEFAULT_AI_PROVIDER) as
        | 'openai'
        | 'ollama';

      const config: AIConfig = {
        provider,
        openai: {
          apiKey:
            provider === 'openai'
              ? this.getRequiredEnv('OPENAI_API_KEY')
              : this.getEnv('OPENAI_API_KEY', ''),
          model: this.getEnv('OPENAI_MODEL', CONFIG_CONSTANTS.DEFAULT_OPENAI_MODEL),
          maxTokens: this.validateNumber(
            this.getEnv('OPENAI_MAX_TOKENS', CONFIG_CONSTANTS.DEFAULT_OPENAI_MAX_TOKENS.toString()),
            CONFIG_CONSTANTS.DEFAULT_OPENAI_MAX_TOKENS,
            1,
            CONFIG_CONSTANTS.MAX_OPENAI_TOKENS
          ),
          temperature: this.validateNumber(
            this.getEnv(
              'OPENAI_TEMPERATURE',
              CONFIG_CONSTANTS.DEFAULT_OPENAI_TEMPERATURE.toString()
            ),
            CONFIG_CONSTANTS.DEFAULT_OPENAI_TEMPERATURE,
            CONFIG_CONSTANTS.MIN_TEMPERATURE,
            CONFIG_CONSTANTS.MAX_TEMPERATURE
          ),
        },
        ollama: {
          host: this.getEnv('OLLAMA_HOST', CONFIG_CONSTANTS.DEFAULT_OLLAMA_HOST),
          model: this.getEnv('OLLAMA_MODEL', CONFIG_CONSTANTS.DEFAULT_OLLAMA_MODEL),
        },
      };

      // Валідація OpenAI API ключа
      if (config.provider === 'openai' && !config.openai.apiKey.startsWith('sk-')) {
        logger.warn('⚠️ OpenAI API ключ може бути некоректним');
      }

      logger.debug('✅ AI конфігурація завантажена');
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження AI конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Redis конфігурації
   */
  private static loadRedisConfig(): RedisConfig {
    try {
      logger.debug('💾 Завантаження Redis конфігурації...');

      const config: RedisConfig = {
        host: this.getEnv('REDIS_HOST', CONFIG_CONSTANTS.DEFAULT_REDIS_HOST),
        port: this.validateNumber(
          this.getEnv('REDIS_PORT', CONFIG_CONSTANTS.DEFAULT_REDIS_PORT.toString()),
          CONFIG_CONSTANTS.DEFAULT_REDIS_PORT,
          1,
          65535
        ),
        password: this.getEnv('REDIS_PASSWORD', ''),
        database: this.validateNumber(
          this.getEnv('REDIS_DATABASE', CONFIG_CONSTANTS.DEFAULT_REDIS_DATABASE.toString()),
          CONFIG_CONSTANTS.DEFAULT_REDIS_DATABASE,
          0,
          15
        ),
        enabled: this.getEnv('REDIS_ENABLED', 'true').toLowerCase() === 'true',
        url: this.getEnv('REDIS_URL', ''),
      };

      logger.debug('✅ Redis конфігурація завантажена');
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження Redis конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Metrics конфігурації
   */
  private static loadMetricsConfig(): MetricsConfig {
    try {
      logger.debug('📊 Завантаження Metrics конфігурації...', {
        component: 'Config',
        type: 'config',
        section: 'metrics',
        event: 'load_start',
      });

      const enabled = this.getEnv('METRICS_ENABLED', 'true').toLowerCase() === 'true';
      const port = this.validateNumber(
        this.getEnv('METRICS_PORT', CONFIG_CONSTANTS.DEFAULT_METRICS_PORT.toString()),
        CONFIG_CONSTANTS.DEFAULT_METRICS_PORT,
        1024,
        65535
      );
      const rawPath = this.getEnv('METRICS_PATH', CONFIG_CONSTANTS.DEFAULT_METRICS_PATH);
      const path = rawPath.startsWith('/') ? rawPath : `/${rawPath}`;
      if (rawPath !== path) {
        logger.warn('⚠️ METRICS_PATH не починається зі "/", виконую нормалізацію', {
          component: 'Config',
          type: 'config',
          section: 'metrics',
          event: 'path_normalized',
          rawPath,
          normalized: path,
        });
      }

      const config: MetricsConfig = { enabled, port, path };
      logger.debug('✅ Metrics конфігурація завантажена', {
        component: 'Config',
        type: 'config',
        section: 'metrics',
        event: 'load_success',
        port: config.port,
        path: config.path,
        enabled: config.enabled,
      });
      return config;
    } catch (error) {
      logger.error('❌ Помилка завантаження Metrics конфігурації:', {
        component: 'Config',
        type: 'config',
        section: 'metrics',
        event: 'load_failed',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
      });
      throw error;
    }
  }

  /**
   * Завантаження Security конфігурації
   */
  private static loadSecurityConfig() {
    try {
      logger.debug('🔒 Завантаження Security конфігурації...');

      return {
        rateLimitWindow: this.validateNumber(
          this.getEnv('RATE_LIMIT_WINDOW', '60000'),
          60000,
          1000,
          300000
        ),
        rateLimitMax: this.validateNumber(this.getEnv('RATE_LIMIT_MAX', '100'), 100, 1, 1000),
        adminRole: this.getEnv('ADMIN_ROLE', 'Admin'),
        botUserRole: this.getEnv('BOT_USER_ROLE', 'Bot User'),
      };
    } catch (error) {
      logger.error('❌ Помилка завантаження Security конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Performance конфігурації
   */
  private static loadPerformanceConfig() {
    try {
      logger.debug('⚡ Завантаження Performance конфігурації...');

      return {
        cacheTTL: this.validateNumber(this.getEnv('CACHE_TTL', '300000'), 300000, 1000, 3600000),
        maxSearchResults: this.validateNumber(
          this.getEnv('MAX_SEARCH_RESULTS', '100'),
          100,
          1,
          1000
        ),
        maxAnalysisRows: this.validateNumber(
          this.getEnv('MAX_ANALYSIS_ROWS', '1000'),
          1000,
          1,
          10000
        ),
        requestTimeout: this.validateNumber(
          this.getEnv('REQUEST_TIMEOUT', '30000'),
          30000,
          1000,
          300000
        ),
        maxRetries: this.validateNumber(this.getEnv('MAX_RETRIES', '3'), 3, 0, 10),
      };
    } catch (error) {
      logger.error('❌ Помилка завантаження Performance конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Завантаження Logging конфігурації
   */
  private static loadLoggingConfig() {
    try {
      logger.debug('📝 Завантаження Logging конфігурації...');

      return {
        level: this.getEnv('LOG_LEVEL', 'info'),
        maxFiles: this.validateNumber(this.getEnv('LOG_MAX_FILES', '5'), 5, 1, 50),
        maxSize: this.getEnv('LOG_MAX_SIZE', '10m'),
        directory: this.getEnv('LOG_DIRECTORY', join(process.cwd(), 'logs')),
      };
    } catch (error) {
      logger.error('❌ Помилка завантаження Logging конфігурації:', error as any);
      throw error;
    }
  }

  /**
   * Валідація числових значень
   */
  private static validateNumber(
    value: string,
    defaultValue: number,
    min: number,
    max: number
  ): number {
    try {
      const num = parseInt(value, 10);
      if (isNaN(num) || num < min || num > max) {
        logger.warn(`⚠️ Некоректне значення ${value}, використовую ${defaultValue}`);
        return defaultValue;
      }
      return num;
    } catch (error) {
      logger.warn(`⚠️ Помилка парсингу числа ${value}, використовую ${defaultValue}`);
      return defaultValue;
    }
  }

  /**
   * Валідація конфігурації
   */
  private static validate(config: BotConfig): void {
    logger.info('🔍 Валідація конфігурації...', {
      component: 'Config',
      type: 'config',
      event: 'validate_start',
    });

    const errors: string[] = [];

    // Валідація Discord
    if (!config.discord.token) errors.push('DISCORD_TOKEN is required');
    if (!config.discord.clientId) errors.push('DISCORD_CLIENT_ID is required');
    // guildId обов'язковий лише коли ENABLE_SLASH=true
    if (config.discord.enableSlash && !config.discord.guildId) {
      errors.push('DISCORD_GUILD_ID is required when ENABLE_SLASH=true');
    }

    // Валідація Google
    if (!config.google.spreadsheetId) errors.push('GOOGLE_SPREADSHEET_ID is required');
    if (!config.google.apiKey) errors.push('GOOGLE_API_KEY is required');
    if (!config.google.appScriptUrl) errors.push('GOOGLE_APP_SCRIPT_URL is required');

    // Валідація AI
    if (config.ai.provider === 'openai' && !config.ai.openai.apiKey) {
      errors.push('OPENAI_API_KEY is required when AI_PROVIDER is openai');
    }

    if (errors.length > 0) {
      const errorMessage = `Configuration validation failed:\n${errors.join('\n')}`;
      logger.error('❌ Помилки валідації конфігурації:', errors);
      throw new Error(errorMessage);
    }

    // Додаткова порада щодо Metrics path
    if (!config.metrics.path.startsWith('/')) {
      logger.warn('⚠️ Metrics path не починається зі "/". Рекомендується формат "/metrics"', {
        component: 'Config',
        type: 'config',
        section: 'metrics',
        event: 'path_warning',
        path: config.metrics.path,
      });
    }

    logger.info('✅ Конфігурація валідна', {
      component: 'Config',
      type: 'config',
      event: 'validate_success',
    });
  }

  /**
   * Завантаження конфігурації
   */
  public static loadLegacy(): BotConfig { return this.load(); }

  /**
   * Логування підсумку конфігурації
   */
  private static logConfigurationSummary(config: BotConfig): void {
    try {
      logger.info('📋 Підсумок конфігурації:', {
        discord: {
          clientId: config.discord.clientId ? '***' : '',
          guildId: config.discord.guildId ? '***' : '',
          prefix: config.discord.prefix,
          intents: config.discord.intents.length,
        },
        google: {
          spreadsheetId: config.google.spreadsheetId,
          sheetName: config.google.sheetName,
          hasCredentials: !!config.google.credentials,
        },
        drive: {
          folderId: config.drive.folderId,
          pageSize: config.drive.pageSize,
          allowedMime: config.drive.allowedMime[0] === '*' ? '*' : config.drive.allowedMime.length,
          hideWebLink: config.drive.hideWebLink,
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
    } catch (error) {
      logger.error('❌ Помилка логування підсумку конфігурації:', error as any);
    }
  }
  /**
   * Отримання обов'язкової змінної середовища
   */
  private static getRequiredEnv(key: string): string {
    const value = process.env[key];
    if (!value) {
      // М'який режим для Google-ключів у локальному smoke-запуску
      const allowGoogleStubs = ((process.env['ALLOW_GOOGLE_STUBS'] as string) || '').toLowerCase() === 'true';
      const softGoogleKeys = new Set<string>([
        'GOOGLE_API_KEY',
        'GOOGLE_APPLICATION_CREDENTIALS',
        'GOOGLE_SPREADSHEET_ID',
        'GOOGLE_DRIVE_FOLDER_ID',
        'GOOGLE_APP_SCRIPT_URL',
      ]);

      if (allowGoogleStubs && softGoogleKeys.has(key)) {
        const stubMap: Record<string, string> = {
          GOOGLE_API_KEY: 'AIzaStub',
          GOOGLE_APPLICATION_CREDENTIALS: '',
          GOOGLE_SPREADSHEET_ID: 'stub-spreadsheet-id',
          GOOGLE_DRIVE_FOLDER_ID: 'stub-drive-folder-id',
          GOOGLE_APP_SCRIPT_URL: 'http://localhost/stub-app-script',
        };
        const stub = stubMap[key] ?? 'stub';
        logger.warn(`⚠️ [SOFT] ${key} не задан, використовую заглушку для локального smoke-тесту`);
        return stub;
      }

      const error = `Required environment variable ${key} is not set`;
      logger.error(`❌ ${error}`);
      throw new Error(error);
    }
    return value;
  }

  /**
   * Отримання змінної середовища з значенням за замовчуванням
   */
  private static getEnv(key: string, defaultValue: string): string {
    const value = process.env[key];
    if (!value) {
      logger.debug(`🔧 Використовую значення за замовчуванням для ${key}: ${defaultValue}`);
    }
    return value || defaultValue;
  }

  /**
   * Очищення кешу конфігурації
   */
  public static clearCache(): void {
    this.instance = null;
    this.configCache.clear();
    logger.debug('🧹 Кеш конфігурації очищено');
  }

  /**
   * Перезавантаження конфігурації
   */
  public static reload(): BotConfig {
    logger.info('🔄 Перезавантаження конфігурації...');
    this.clearCache();
    return this.load();
  }
}
