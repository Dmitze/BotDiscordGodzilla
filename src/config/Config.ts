/**
 * Клас для управління конфігурацією додатку
 * Завантажує та валідує налаштування з змінних середовища
 */

import type { BotConfig, DiscordConfig, GoogleConfig, AIConfig, RedisConfig, MetricsConfig } from '@/types';

export class Config {
  /**
   * Завантаження конфігурації з змінних середовища
   */
  public static load(): BotConfig {
    const config: BotConfig = {
      discord: Config.loadDiscordConfig(),
      google: Config.loadGoogleConfig(),
      ai: Config.loadAIConfig(),
      redis: Config.loadRedisConfig(),
      metrics: Config.loadMetricsConfig(),
    };

    Config.validate(config);
    return config;
  }

  /**
   * Завантаження Discord конфігурації
   */
  private static loadDiscordConfig(): DiscordConfig {
    return {
      token: Config.getRequiredEnv('DISCORD_TOKEN'),
      clientId: Config.getRequiredEnv('DISCORD_CLIENT_ID'),
      guildId: Config.getRequiredEnv('DISCORD_GUILD_ID'),
      prefix: Config.getEnv('DISCORD_PREFIX', '!'),
      intents: Config.getEnv('DISCORD_INTENTS', 'Guilds,GuildMessages,MessageContent,GuildMembers').split(','),
    };
  }

  /**
   * Завантаження Google конфігурації
   */
  private static loadGoogleConfig(): GoogleConfig {
    const config: GoogleConfig = {
      spreadsheetId: Config.getRequiredEnv('GOOGLE_SPREADSHEET_ID'),
      driveFolderId: Config.getRequiredEnv('GOOGLE_DRIVE_FOLDER_ID'),
      apiKey: Config.getRequiredEnv('GOOGLE_API_KEY'),
      applicationCredentials: Config.getRequiredEnv('GOOGLE_APPLICATION_CREDENTIALS'),
      appScriptUrl: Config.getRequiredEnv('GOOGLE_APP_SCRIPT_URL'),
      sheetName: Config.getEnv('GOOGLE_SHEET_NAME', 'Sheet1'),
    };

    const credentials = Config.loadGoogleCredentials();
    if (credentials) {
      config.credentials = credentials;
    }

    return config;
  }

  /**
   * Завантаження Google credentials
   */
  private static loadGoogleCredentials(): GoogleConfig['credentials'] {
    const clientEmail = process.env['GOOGLE_CLIENT_EMAIL'];
    const privateKey = process.env['GOOGLE_PRIVATE_KEY'];
    const projectId = process.env['GOOGLE_PROJECT_ID'];

    if (clientEmail && privateKey && projectId) {
      return {
        client_email: clientEmail,
        private_key: privateKey.replace(/\\n/g, '\n'),
        project_id: projectId,
      };
    }

    return undefined;
  }

  /**
   * Завантаження AI конфігурації
   */
  private static loadAIConfig(): AIConfig {
    return {
      provider: (Config.getEnv('AI_PROVIDER', 'openai') as 'openai' | 'ollama'),
      openai: {
        apiKey: Config.getRequiredEnv('OPENAI_API_KEY'),
        model: Config.getEnv('OPENAI_MODEL', 'gpt-3.5-turbo'),
        maxTokens: parseInt(Config.getEnv('OPENAI_MAX_TOKENS', '1000'), 10),
        temperature: parseFloat(Config.getEnv('OPENAI_TEMPERATURE', '0.7')),
      },
      ollama: {
        host: Config.getEnv('OLLAMA_HOST', 'http://localhost:11434'),
        model: Config.getEnv('OLLAMA_MODEL', 'llama2'),
      },
    };
  }

  /**
   * Завантаження Redis конфігурації
   */
  private static loadRedisConfig(): RedisConfig {
    return {
      host: Config.getEnv('REDIS_HOST', 'localhost'),
      port: parseInt(Config.getEnv('REDIS_PORT', '6379'), 10),
      password: Config.getEnv('REDIS_PASSWORD'),
      database: parseInt(Config.getEnv('REDIS_DATABASE', '0'), 10),
      enabled: Config.getEnv('REDIS_ENABLED', 'true').toLowerCase() === 'true',
      url: Config.getEnv('REDIS_URL'),
    };
  }

  /**
   * Завантаження Metrics конфігурації
   */
  private static loadMetricsConfig(): MetricsConfig {
    return {
      enabled: Config.getEnv('METRICS_ENABLED', 'true').toLowerCase() === 'true',
      port: parseInt(Config.getEnv('METRICS_PORT', '9090'), 10),
      path: Config.getEnv('METRICS_PATH', '/metrics'),
    };
  }

  /**
   * Валідація конфігурації
   */
  private static validate(config: BotConfig): void {
    const errors: string[] = [];

    // Валідація Discord
    if (!config.discord.token) errors.push('DISCORD_TOKEN is required');
    if (!config.discord.clientId) errors.push('DISCORD_CLIENT_ID is required');
    if (!config.discord.guildId) errors.push('DISCORD_GUILD_ID is required');

    // Валідація Google
    if (!config.google.spreadsheetId) errors.push('GOOGLE_SPREADSHEET_ID is required');
    if (!config.google.apiKey) errors.push('GOOGLE_API_KEY is required');

    // Валідація AI
    if (config.ai.provider === 'openai' && !config.ai.openai.apiKey) {
      errors.push('OPENAI_API_KEY is required when AI_PROVIDER is openai');
    }

    if (errors.length > 0) {
      throw new Error(`Configuration validation failed:\n${errors.join('\n')}`);
    }
  }

  /**
   * Отримання обов'язкової змінної середовища
   */
  private static getRequiredEnv(key: string): string {
    const value = process.env[key];
    if (!value) {
      throw new Error(`Required environment variable ${key} is not set`);
    }
    return value;
  }

  /**
   * Отримання змінної середовища з значенням за замовчуванням
   */
  private static getEnv(key: string, defaultValue: string): string {
    return process.env[key] || defaultValue;
  }
} 