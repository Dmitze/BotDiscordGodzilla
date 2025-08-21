/**
 * Основний клас Discord бота
 * Управляє всіма компонентами та сервісами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import { Client, GatewayIntentBits, Collection, Events, Interaction } from 'discord.js';
import type { BotConfig, BaseCommand } from '@/types';
import { ServiceContainer } from './ServiceContainer';
import { BaseService as BaseServiceClass } from './BaseService';
import { CommandManager } from './CommandManager';
import { ErrorHandler } from './ErrorHandler';
import EventManager from './EventManager';
import ServiceManager from './ServiceManager';
import logger from '@/utils/logger';
import { ChatRouter } from '@/chat/ChatRouter';
import { IntentDetector } from '@/chat/IntentDetector';
import { MemoryService } from '@/chat/MemoryService';

// Константи для конфігурації бота
const BOT_CONSTANTS = {
  READY_TIMEOUT: 120000, // 120 секунд
  COMMAND_TIMEOUT: 15000, // 15 секунд
  MAX_RECONNECT_ATTEMPTS: 5,
  RECONNECT_DELAY: 5000, // 5 секунд
  HEALTH_CHECK_INTERVAL: 60000, // 1 хвилина
  MAX_MEMORY_USAGE: 512 * 1024 * 1024, // 512MB
  COMMAND_RATE_LIMIT: 10, // команд за хвилину
  INTERACTION_RATE_LIMIT: 50, // interactions за хвилину
} as const;

interface BotStats {
  uptime: number;
  commands: number;
  interactions: number;
  errors: number;
  reconnects: number;
  lastActivity: Date;
  memory: NodeJS.MemoryUsage;
  rateLimitHits: number;
  slowCommands: number;
}

interface RateLimitInfo {
  count: number;
  resetTime: number;
}

export class Bot extends BaseServiceClass {
  public readonly client: Client;
  public readonly serviceContainer: ServiceContainer;
  public readonly commandManager: CommandManager;
  public readonly errorHandler: ErrorHandler;
  public readonly eventManager: EventManager;
  public readonly serviceManager: ServiceManager;
  private readonly chatMemory: MemoryService;
  private readonly intentDetector: IntentDetector;
  private readonly chatRouter: ChatRouter;

  private commands = new Collection<string, BaseCommand>();
  private isReady = false;
  private isConnecting = false;
  private reconnectAttempts = 0;
  private stats: BotStats;
  private healthCheckInterval: NodeJS.Timeout | null = null;
  private lastInteractionTime: Date = new Date();
  private rateLimitMap = new Map<string, RateLimitInfo>();

  constructor(config: BotConfig) {
    super('DiscordBot', config);

    // Ініціалізація статистики
    this.stats = {
      uptime: 0,
      commands: 0,
      interactions: 0,
      errors: 0,
      reconnects: 0,
      lastActivity: new Date(),
      memory: process.memoryUsage(),
      rateLimitHits: 0,
      slowCommands: 0,
    };

    // Вираховуємо intents з конфігурації
    const intentMap: Record<string, number> = {
      Guilds: GatewayIntentBits.Guilds,
      GuildMessages: GatewayIntentBits.GuildMessages,
      DirectMessages: GatewayIntentBits.DirectMessages,
      GuildMembers: GatewayIntentBits.GuildMembers,
      GuildPresences: GatewayIntentBits.GuildPresences,
      MessageContent: GatewayIntentBits.MessageContent,
      GuildMessageReactions: GatewayIntentBits.GuildMessageReactions,
      GuildMessageTyping: GatewayIntentBits.GuildMessageTyping,
      DirectMessageReactions: GatewayIntentBits.DirectMessageReactions,
      DirectMessageTyping: GatewayIntentBits.DirectMessageTyping,
    };

    const requestedIntentNames = Array.isArray(config.discord.intents)
      ? config.discord.intents
      : [];
    const intentsResolved = requestedIntentNames
      .map(name => intentMap[name])
      .filter((v): v is number => typeof v === 'number');

    // Fail-fast: чат включен, но нет MessageContent
    if (config.discord.enableChat && !requestedIntentNames.includes('MessageContent')) {
      const meta = {
        type: 'bot',
        event: 'missing_intent',
        enableChat: config.discord.enableChat,
        enableMessageContentIntent: config.discord.enableMessageContentIntent,
        intents: requestedIntentNames,
      } as const;
      logger.error(
        'ENABLE_CHAT=true, но MessageContent не включен в intents. Включите ENABLE_MESSAGE_CONTENT_INTENT=true и добавьте MessageContent в DISCORD_INTENTS.',
        meta
      );
      throw new Error('Чат-режим требует разрешения Message Content Intent');
    }

    this.client = new Client({
      intents: intentsResolved,
      failIfNotExists: false,
    });

    // Ініціалізація менеджерів та сервісів
    this.serviceContainer = new ServiceContainer(config);
    this.commandManager = new CommandManager(this, config);
    this.errorHandler = new ErrorHandler(this.serviceContainer);
    this.eventManager = new EventManager(this);
    this.serviceManager = new ServiceManager(this);

    // Чат-режим: пам'ять, детектор намірів і роутер повідомлень
    this.chatMemory = new MemoryService({ maxTokens: 2000, summaryAfter: 1500 });
    this.intentDetector = new IntentDetector();
    this.chatRouter = new ChatRouter(this.client, this.chatMemory, this.intentDetector, this.getService.bind(this));

    // Налаштування обробників подій
    this.setupEventHandlers();

    logger.info('🤖 Екземпляр Discord бота створено');
  }

  /**
   * Логування стартової статистики бота з безпечним метаданими
   */
  private logStartupStats(): void {
    try {
      const s = this.getStats();
      const guilds = this.client.guilds.cache.size;
      const channels = this.client.channels.cache.size;
      const meta = {
        type: 'bot',
        event: 'startup_stats',
        guilds,
        channels,
      } as const;
      logger.info(
        `📊 Стартові метрики: uptime=${s.uptime}ms, errors=${(s as any).errors ?? 0}, reconnects=${(s as any).reconnects ?? 0}`,
        meta
      );
    } catch (e) {
      const meta =
        e instanceof Error
          ? {
              type: 'bot',
              event: 'startup_stats_failed',
              errorName: e.name,
              errorMessage: e.message,
              stack: e.stack,
            }
          : { type: 'bot', event: 'startup_stats_failed', errorMessage: String(e) };
      logger.warn('⚠️ Не вдалося залогувати стартові метрики', meta);
    }
  }

  /**
   * Отримання сервісу за назвою (проксі метод для сумісності з ServiceManager)
   */
  public getService(name: string): unknown {
    // Спочатку пробуємо через ServiceManager, якщо вже ініціалізований
    if (this.serviceManager && typeof this.serviceManager.getService === 'function') {
      return this.serviceManager.getService(name);
    }
    // В іншому випадку – напряму з контейнера сервісів
    try {
      return this.serviceContainer.get(name as unknown as string);
    } catch {
      return undefined;
    }
  }

  /**
   * Ініціалізація бота з детальним логуванням
   */
  protected async onInitialize(): Promise<void> {
    const startTime = Date.now();

    try {
      logger.info('🚀 Початок ініціалізації Discord бота...');

      // Перевірка системних ресурсів
      await this.checkSystemResources();

      // Ініціалізація обробника помилок
      logger.info('🛡️ Ініціалізація обробника помилок...');
      await this.errorHandler.initialize();

      // Ініціалізація менеджера подій
      logger.info('📡 Ініціалізація менеджера подій...');
      await this.eventManager.initialize();

      // Ініціалізація сервісів
      logger.info('🔧 Ініціалізація сервісів...');
      await this.serviceContainer.initialize();

      // Ініціалізація менеджера сервісів
      logger.info('⚙️ Ініціалізація менеджера сервісів...');
      await this.serviceManager.initialize();

      // Ініціалізація менеджера команд
      logger.info('📝 Ініціалізація менеджера команд...');
      await this.commandManager.initialize();

      // Підключення до Discord
      logger.info('🔌 Підключення до Discord...');
      await this.connectToDiscord();

      // Очікування готовності клієнта
      logger.info('⏳ Очікування готовності клієнта...');
      await this.waitForReady();

      // Підключаємо чат-роутер після готовності клієнта (тільки якщо чат увімкнено)
      if (this.config.discord.enableChat) {
        try {
          this.chatRouter.bind();
        } catch (e) {
          logger.error('❌ Неможливо підключити ChatRouter', {
            type: 'bot',
            event: 'chat_bind_failed',
            error: e instanceof Error ? e.message : String(e),
          });
          // Не зупиняємо весь бот, але логгуємо критично
        }
      } else {
        logger.info('💤 Chat mode disabled — ChatRouter не підключено');
      }
      // Запуск health check (пропускаємо у тестовому середовищі)
      if (process.env['NODE_ENV'] === 'test' || process.env['DISABLE_BOT_HEALTHCHECK'] === 'true') {
        logger.info('🧪 Режим тесту/відключено: health check бота не запускається', { type: 'bot', event: 'health_check_skipped' } as const);
      } else {
        this.startHealthCheck();
      }

      const initDuration = Date.now() - startTime;
      const meta = { type: 'bot', event: 'initialized', durationMs: initDuration };
      logger.info(`✅ Discord бот успішно ініціалізовано за ${initDuration}ms`, meta);

      // Логування статистики запуску
      this.logStartupStats();
    } catch (error) {
      const initDuration = Date.now() - startTime;
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'initialize_failed',
              durationMs: initDuration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'bot',
              event: 'initialize_failed',
              durationMs: initDuration,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка ініціалізації бота після ${initDuration}ms`, meta);

      // Спроба очищення ресурсів
      await this.cleanupOnError();

      throw new Error(
        `Помилка ініціалізації бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Завершення роботи бота з детальним логуванням
   */
  protected async onShutdown(): Promise<void> {
    const shutdownStartTime = Date.now();

    try {
      logger.info('🛑 Початок завершення роботи Discord бота...');

      // Зупинка health check
      this.stopHealthCheck();

      // Завершення менеджера подій
      logger.info('📡 Завершення менеджера подій...');
      await this.eventManager.shutdown();

      // Завершення менеджера команд (команди не потребують спеціального shutdown у поточній реалізації)
      logger.info('📝 Завершення менеджера команд...');

      // Завершення менеджера сервісів
      logger.info('⚙️ Завершення менеджера сервісів...');
      await this.serviceManager.shutdown();

      // Завершення сервісів
      logger.info('🔧 Завершення сервісів...');
      await this.serviceContainer.shutdown();

      // Завершення обробника помилок
      logger.info('🛡️ Завершення обробника помилок...');
      await this.errorHandler.shutdown();

      // Відключення від Discord
      logger.info('🔌 Відключення від Discord...');
      this.client.destroy();

      const shutdownDuration = Date.now() - shutdownStartTime;
      const meta = { type: 'bot', event: 'shutdown', durationMs: shutdownDuration };
      logger.info(`✅ Discord бот успішно завершено за ${shutdownDuration}ms`, meta);
    } catch (error) {
      const shutdownDuration = Date.now() - shutdownStartTime;
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'shutdown_failed',
              durationMs: shutdownDuration,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : {
              type: 'bot',
              event: 'shutdown_failed',
              durationMs: shutdownDuration,
              errorMessage: String(error),
            };
      logger.error(`❌ Помилка завершення бота після ${shutdownDuration}ms`, meta);

      // Примусова зупинка при помилці
      logger.warn('🔄 Примусова зупинка Discord клієнта...');
      this.client.destroy();

      throw new Error(
        `Помилка завершення бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Health check бота з розширеною інформацією
   */
  protected async onHealthCheck(): Promise<{
    healthy: boolean;
    service: string;
    error?: string;
    details?: Record<string, unknown>;
  }> {
    try {
      const isConnected = this.client.isReady();
      const servicesHealth = await this.serviceContainer.getHealthStatus();

      const allServicesHealthy = Object.values(servicesHealth).every(health => health.healthy);
      const healthy = isConnected && allServicesHealthy;

      return {
        healthy,
        service: this.name,
        details: {
          connected: isConnected,
          ready: this.isReady,
          services: servicesHealth,
          // Команди не надають окремого health, додаємо базову статистику команд
          commands: { total: this.stats.commands, slow: this.stats.slowCommands },
          stats: this.getStats(),
          uptime: this.getStats().uptime,
          memory: process.memoryUsage(),
          lastActivity: this.lastInteractionTime,
          rateLimitHits: this.stats.rateLimitHits,
          slowCommands: this.stats.slowCommands,
        },
      };
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'healthcheck_failed',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'healthcheck_failed', errorMessage: String(error) };
      logger.error('❌ Помилка health check бота', meta);
      return {
        healthy: false,
        service: this.name,
        error: `Health check failed: ${error instanceof Error ? error.message : 'Невідома помилка'}`,
      };
    }
  }

  /**
   * Отримання детальної статистики бота
   */
  protected onGetStats(): Partial<import('@/types').ServiceStats> {
    this.stats.uptime = Date.now() - this.startTime;
    this.stats.memory = process.memoryUsage();
    // Повертаємо як розширену статистику сервісу, BaseService додасть базові поля
    return { ...this.stats } as Record<string, unknown>;
  }

  /**
   * Перевірка системних ресурсів
   */
  private async checkSystemResources(): Promise<void> {
    try {
      logger.info('🔍 Перевірка системних ресурсів бота...');

      const memoryUsage = process.memoryUsage();
      const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;

      if (heapUsedMB > 200) {
        logger.warn(`⚠️ Високе використання пам'яті бота: ${Math.round(heapUsedMB)}MB`);
      }
      // Політика: офлайн/ізольоване середовище — пропускаємо зовнішні мережеві перевірки

      logger.info('✅ Системні ресурси бота перевірено');
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'resources_check_failed',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'resources_check_failed', errorMessage: String(error) };
      logger.error('❌ Помилка перевірки системних ресурсів бота', meta);
      throw error;
    }
  }

  /**
   * Підключення до Discord з обробкою помилок
   */
  private async connectToDiscord(): Promise<void> {
    if (this.isConnecting) {
      logger.warn('⚠️ Вже виконується підключення до Discord');
      return;
    }

    this.isConnecting = true;

    try {
      logger.info('🔌 Спроба підключення до Discord...');
      await this.client.login(this.config.discord.token);
      logger.info('✅ Успішно підключено до Discord');

      // Скидання лічильника спроб перепідключення
      this.reconnectAttempts = 0;
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'login_failed',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'login_failed', errorMessage: String(error) };
      logger.error('❌ Помилка підключення до Discord', meta);

      if (this.reconnectAttempts < BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
        this.reconnectAttempts++;
        logger.info(
          `🔄 Спроба перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS}...`
        );

        setTimeout(() => {
          this.connectToDiscord();
        }, BOT_CONSTANTS.RECONNECT_DELAY);
      } else {
        throw new Error(
          `Не вдалося підключитися до Discord після ${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} спроб`
        );
      }
    } finally {
      this.isConnecting = false;
    }
  }

  /**
   * Налаштування обробників подій з детальним логуванням
   */
  private setupEventHandlers(): void {
    // Ready event
    this.client.on(Events.ClientReady, () => {
      this.isReady = true;
      this.stats.lastActivity = new Date();

      logger.info(`🤖 Бот ${this.client.user?.tag} готовий до роботи`);
      logger.info(
        `📊 Статистика: ${this.client.guilds.cache.size} серверів, ${this.client.channels.cache.size} каналів`
      );

      // Встановлення статусу бота
      this.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
    });

    // Interaction event
    this.client.on(Events.InteractionCreate, async (interaction: Interaction) => {
      this.stats.interactions++;
      this.lastInteractionTime = new Date();

      try {
        // ВАЖЛИВО: Обробку ChatInput-команд виконує CommandManager через власний обробник InteractionCreate
        // Щоб уникнути подвійної відповіді та помилки Unknown interaction — пропускаємо їх тут
        if (interaction.isChatInputCommand && interaction.isChatInputCommand()) {
          return; // CommandManager обробить самостійно
        }

        // Перевірка rate limit для інших видів interactions (кнопки, select menu)
        if (this.isRateLimited(interaction.user?.id || 'unknown')) {
          logger.warn(`⚠️ Rate limit для користувача ${interaction.user?.id}`);
          await this.handleRateLimit(interaction);
          return;
        }

        if (interaction.isButton()) {
          await this.handleButtonInteraction(interaction);
        } else if (interaction.isSelectMenu()) {
          await this.handleSelectMenuInteraction(interaction);
        }
      } catch (error) {
        this.stats.errors++;
        const meta =
          error instanceof Error
            ? {
                type: 'bot',
                event: 'interaction_failed',
                errorName: error.name,
                errorMessage: error.message,
                stack: error.stack,
              }
            : { type: 'bot', event: 'interaction_failed', errorMessage: String(error) };
        logger.error('❌ Помилка обробки interaction', meta);
        await this.handleInteractionError(interaction, error);
      }
    });

    // Error event
    this.client.on(Events.Error, error => {
      this.stats.errors++;
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'client_error',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'client_error', errorMessage: String(error) };
      logger.error('❌ Помилка Discord клієнта', meta);

      // Спроба перепідключення при критичних помилках
      if (this.shouldReconnect(error)) {
        this.scheduleReconnect();
      }
    });

    // Shard disconnect event (v14)
    this.client.on(Events.ShardDisconnect, (event, shardId) => {
      this.isReady = false;
      const meta = {
        type: 'bot',
        event: 'shard_disconnect',
        shardId,
        code: (event as unknown as { code?: unknown })?.code,
      };
      logger.warn('🔌 Discord shard відключено', meta);

      // Автоматичне перепідключення
      this.scheduleReconnect();
    });

    // Shard reconnecting event (v14)
    this.client.on(Events.ShardReconnecting, shardId => {
      this.stats.reconnects++;
      const meta = { type: 'bot', event: 'shard_reconnecting', shardId };
      logger.info('🔄 Discord shard перепідключається...', meta);
    });

    // Guild events
    this.client.on(Events.GuildCreate, guild => {
      logger.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`);
    });

    this.client.on(Events.GuildDelete, guild => {
      logger.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`);
    });

    logger.info('✅ Обробники подій Discord налаштовано');
  }

  

  /**
   * Обробка кнопкових interactions
   */
  private async handleButtonInteraction(interaction: any): Promise<void> {
    logger.debug(`🔘 Обробка кнопкового interaction: ${interaction.customId}`);
    // Тут можна додати логіку обробки кнопок
  }

  /**
   * Обробка select menu interactions
   */
  private async handleSelectMenuInteraction(interaction: any): Promise<void> {
    logger.debug(`📋 Обробка select menu interaction: ${interaction.customId}`);
    // Тут можна додати логіку обробки select menu
  }

  /**
   * Обробка помилок interactions
   */
  private async handleInteractionError(interaction: Interaction, error: unknown): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

    try {
      if (interaction.isRepliable()) {
        if (interaction.replied || interaction.deferred) {
          await interaction.editReply({
            content: `❌ Помилка виконання: ${errorMessage}`,
          });
        } else {
          await interaction.reply({
            content: `❌ Помилка виконання: ${errorMessage}`,
            ephemeral: true,
          });
        }
      }
    } catch (replyError) {
      const meta =
        replyError instanceof Error
          ? {
              type: 'bot',
              event: 'interaction_reply_failed',
              errorName: replyError.name,
              errorMessage: replyError.message,
              stack: replyError.stack,
            }
          : { type: 'bot', event: 'interaction_reply_failed', errorMessage: String(replyError) };
      logger.error('❌ Помилка відповіді на помилку interaction', meta);
    }
  }

  /**
   * Перевірка rate limit
   */
  private isRateLimited(userId: string): boolean {
    const now = Date.now();
    const userLimit = this.rateLimitMap.get(userId);

    if (!userLimit) {
      this.rateLimitMap.set(userId, { count: 1, resetTime: now + 60000 });
      return false;
    }

    if (now > userLimit.resetTime) {
      this.rateLimitMap.set(userId, { count: 1, resetTime: now + 60000 });
      return false;
    }

    if (userLimit.count >= BOT_CONSTANTS.COMMAND_RATE_LIMIT) {
      this.stats.rateLimitHits++;
      return true;
    }

    userLimit.count++;
    return false;
  }

  /**
   * Обробка rate limit
   */
  private async handleRateLimit(interaction: Interaction): Promise<void> {
    try {
      if (interaction.isRepliable()) {
        await interaction.reply({
          content: '⚠️ Забагато запитів. Спробуйте пізніше.',
          ephemeral: true,
        });
      }
    } catch (error) {
      const meta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'rate_limit_reply_failed',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'rate_limit_reply_failed', errorMessage: String(error) };
      logger.error('❌ Помилка обробки rate limit', meta);
    }
  }

  /**
   * Очікування готовності клієнта з таймаутом
   */
  private waitForReady(): Promise<void> {
    return new Promise((resolve) => {
      if (this.isReady) {
        resolve();
        return;
      }

      const timeout = setTimeout(() => {
        // М'який таймаут: не кидаємо помилку, а продовжуємо у деградованому режимі
        logger.warn(
          `⏰ Таймаут очікування готовності клієнта (${BOT_CONSTANTS.READY_TIMEOUT}ms). Продовжуємо у режимі очікування Ready...`,
          { type: 'bot', event: 'ready_timeout_soft' } as const
        );
        resolve();
      }, BOT_CONSTANTS.READY_TIMEOUT);

      this.client.once(Events.ClientReady, () => {
        clearTimeout(timeout);
        resolve();
      });
    });
  }

  /**
   * Перевірка чи потрібно перепідключення
   */
  private shouldReconnect(error: Error): boolean {
    const reconnectErrors = [
      'ECONNRESET',
      'ENOTFOUND',
      'ETIMEDOUT',
      'ECONNREFUSED',
      'WebSocket connection was closed',
    ];

    return reconnectErrors.some(
      errorType => error.message.includes(errorType) || error.name.includes(errorType)
    );
  }

  /**
   * Планування перепідключення
   */
  private scheduleReconnect(): void {
    if (this.reconnectAttempts >= BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
      logger.error(
        `❌ Досягнуто максимальну кількість спроб перепідключення (${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS})`
      );
      return;
    }

    this.reconnectAttempts++;
    const meta = {
      type: 'bot',
      event: 'reconnect_scheduled',
      reconnectAttempts: this.reconnectAttempts,
      delayMs: BOT_CONSTANTS.RECONNECT_DELAY,
    } as const;
    logger.info(
      `🔄 Планування перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} через ${BOT_CONSTANTS.RECONNECT_DELAY}ms`,
      meta
    );

    setTimeout(() => {
      this.connectToDiscord();
    }, BOT_CONSTANTS.RECONNECT_DELAY);
  }

  /**
   * Запуск health check
   */
  private startHealthCheck(): void {
    // Захист від повторного запуску
    if (this.healthCheckInterval) return;
    this.healthCheckInterval = setInterval(async () => {
      try {
        const health = await this.onHealthCheck();
        if (!health.healthy) {
          const meta = {
            type: 'bot',
            event: 'health_check_warning',
            healthy: health.healthy,
            details: health.details,
          } as const;
          logger.warn('⚠️ Health check виявив проблеми', meta);
        }
      } catch (error) {
        const meta =
          error instanceof Error
            ? {
                type: 'bot',
                event: 'health_check_error',
                errorName: error.name,
                errorMessage: error.message,
                stack: error.stack,
              }
            : { type: 'bot', event: 'health_check_error', errorMessage: String(error) };
        logger.error('❌ Помилка health check', meta);
      }
    }, BOT_CONSTANTS.HEALTH_CHECK_INTERVAL);

    const meta = { type: 'bot', event: 'health_check_started' } as const;
    logger.info('🏥 Health check запущено', meta);
  }

  /**
   * Зупинка health check
   */
  private stopHealthCheck(): void {
    if (this.healthCheckInterval) {
      clearInterval(this.healthCheckInterval);
      this.healthCheckInterval = null;
      const meta = { type: 'bot', event: 'health_check_stopped' } as const;
      logger.info('🏥 Health check зупинено', meta);
    }
  }

  /**
   * Очищення ресурсів при помилці
   */
  private async cleanupOnError(): Promise<void> {
    try {
      const meta = { type: 'bot', event: 'cleanup_start' } as const;
      logger.info('🧹 Очищення ресурсів при помилці...', meta);

      this.stopHealthCheck();

      if (this.client) {
        this.client.destroy();
      }

      logger.info('✅ Ресурси очищено', { type: 'bot', event: 'cleanup_done' } as const);
    } catch (error) {
      const errMeta =
        error instanceof Error
          ? {
              type: 'bot',
              event: 'cleanup_error',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            }
          : { type: 'bot', event: 'cleanup_error', errorMessage: String(error) };
      logger.error('❌ Помилка очищення ресурсів', errMeta);
    }
  }

  /**
   * Отримання всіх команд
   */
  public getCommands(): Collection<string, BaseCommand> {
    return this.commands;
  }

  /**
   * Перевірка чи бот готовий
   */
  public isBotReady(): boolean {
    return this.isReady && this.client.isReady();
  }

  /**
   * Отримання детальної статистики
   */
  public getDetailedStats(): BotStats & {
    isReady: boolean;
    isConnecting: boolean;
    reconnectAttempts: number;
  } {
    const base = this.getStats();
    return {
      ...(base as unknown as BotStats),
      isReady: this.isReady,
      isConnecting: this.isConnecting,
      reconnectAttempts: this.reconnectAttempts,
    };
  }
}
