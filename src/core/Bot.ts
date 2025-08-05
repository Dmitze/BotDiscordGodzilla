/**
 * Основний клас Discord бота
 * Управляє всіма компонентами та сервісами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import { Client, GatewayIntentBits, Collection, Events, Interaction, CommandInteraction, ClientEvents } from 'discord.js';
import type { BotConfig, BaseService, CommandInteraction as CommandInteractionType, BaseCommand } from '@/types';
import { ServiceContainer } from './ServiceContainer';
import { BaseService as BaseServiceClass } from './BaseService';
import { CommandManager } from './CommandManager';
import { ErrorHandler } from './ErrorHandler';
import { EventManager } from './EventManager';
import { ServiceManager } from './ServiceManager';
import logger from '@/utils/logger';

// Константи для конфігурації бота
const BOT_CONSTANTS = {
  READY_TIMEOUT: 30000, // 30 секунд
  COMMAND_TIMEOUT: 15000, // 15 секунд
  MAX_RECONNECT_ATTEMPTS: 5,
  RECONNECT_DELAY: 5000, // 5 секунд
  HEALTH_CHECK_INTERVAL: 60000, // 1 хвилина
} as const;

interface BotStats {
  uptime: number;
  commands: number;
  interactions: number;
  errors: number;
  reconnects: number;
  lastActivity: Date;
  memory: NodeJS.MemoryUsage;
}

export class Bot extends BaseServiceClass {
  public readonly client: Client;
  public readonly serviceContainer: ServiceContainer;
  public readonly commandManager: CommandManager;
  public readonly errorHandler: ErrorHandler;
  public readonly eventManager: EventManager;
  public readonly serviceManager: ServiceManager;
  
  private commands = new Collection<string, BaseCommand>();
  private isReady = false;
  private isConnecting = false;
  private reconnectAttempts = 0;
  private stats: BotStats;
  private healthCheckInterval: NodeJS.Timeout | null = null;
  private lastInteractionTime: Date = new Date();

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
    };

    // Створення Discord клієнта з розширеними intents
    this.client = new Client({
      intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
        GatewayIntentBits.GuildMembers,
        GatewayIntentBits.DirectMessages,
        GatewayIntentBits.GuildPresences,
      ],
      failIfNotExists: false,
      retryLimit: 3,
    });

    // Ініціалізація менеджерів та сервісів
    this.serviceContainer = new ServiceContainer(config);
    this.commandManager = new CommandManager(this.client, config);
    this.errorHandler = new ErrorHandler(this.serviceContainer);
    this.eventManager = new EventManager(this);
    this.serviceManager = new ServiceManager(this);

    // Налаштування обробників подій
    this.setupEventHandlers();
    
    logger.info('🤖 Екземпляр Discord бота створено');
  }

  /**
   * Ініціалізація бота з детальним логуванням
   */
  protected async onInitialize(): Promise<void> {
    const startTime = Date.now();
    
    try {
      logger.info('🚀 Початок ініціалізації Discord бота...');
      
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
      
      // Запуск health check
      this.startHealthCheck();
      
      const initDuration = Date.now() - startTime;
      logger.info(`✅ Discord бот успішно ініціалізовано за ${initDuration}ms`);
      
      // Логування статистики запуску
      this.logStartupStats();
      
    } catch (error) {
      const initDuration = Date.now() - startTime;
      logger.error(`❌ Помилка ініціалізації бота після ${initDuration}ms:`, error);
      
      // Спроба очищення ресурсів
      await this.cleanupOnError();
      
      throw new Error(`Помилка ініціалізації бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
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
      
      // Завершення менеджера команд
      logger.info('📝 Завершення менеджера команд...');
      await this.commandManager.shutdown();
      
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
      logger.info(`✅ Discord бот успішно завершено за ${shutdownDuration}ms`);
      
    } catch (error) {
      const shutdownDuration = Date.now() - shutdownStartTime;
      logger.error(`❌ Помилка завершення бота після ${shutdownDuration}ms:`, error);
      
      // Примусова зупинка при помилці
      logger.warn('🔄 Примусова зупинка Discord клієнта...');
      this.client.destroy();
      
      throw new Error(`Помилка завершення бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
    }
  }

  /**
   * Health check бота з розширеною інформацією
   */
  protected async onHealthCheck(): Promise<{ healthy: boolean; service: string; error?: string; details?: Record<string, unknown> }> {
    try {
      const isConnected = this.client.isReady();
      const servicesHealth = await this.serviceContainer.getHealthStatus();
      const commandsHealth = await this.commandManager.getHealthStatus();
      
      const allServicesHealthy = Object.values(servicesHealth).every(health => health.healthy);
      const allCommandsHealthy = Object.values(commandsHealth).every(health => health.healthy);
      
      const healthy = isConnected && allServicesHealthy && allCommandsHealthy;
      
      return {
        healthy,
        service: this.name,
        details: {
          connected: isConnected,
          ready: this.isReady,
          services: servicesHealth,
          commands: commandsHealth,
          stats: this.getStats(),
          uptime: this.getStats().uptime,
          memory: process.memoryUsage(),
          lastActivity: this.lastInteractionTime,
        },
      };
    } catch (error) {
      logger.error('❌ Помилка health check бота:', error);
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
  protected onGetStats(): BotStats {
    this.stats.uptime = Date.now() - this.startTime;
    this.stats.memory = process.memoryUsage();
    return { ...this.stats };
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
      logger.error('❌ Помилка підключення до Discord:', error);
      
      if (this.reconnectAttempts < BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
        this.reconnectAttempts++;
        logger.info(`🔄 Спроба перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS}...`);
        
        setTimeout(() => {
          this.connectToDiscord();
        }, BOT_CONSTANTS.RECONNECT_DELAY);
      } else {
        throw new Error(`Не вдалося підключитися до Discord після ${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} спроб`);
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
      logger.info(`📊 Статистика: ${this.client.guilds.cache.size} серверів, ${this.client.channels.cache.size} каналів`);
      
      // Встановлення статусу бота
      this.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
    });

    // Interaction event
    this.client.on(Events.InteractionCreate, async (interaction: Interaction) => {
      this.stats.interactions++;
      this.lastInteractionTime = new Date();
      
      try {
        if (interaction.isCommand()) {
          await this.handleCommand(interaction as CommandInteraction);
        } else if (interaction.isButton()) {
          await this.handleButtonInteraction(interaction);
        } else if (interaction.isSelectMenu()) {
          await this.handleSelectMenuInteraction(interaction);
        }
      } catch (error) {
        this.stats.errors++;
        logger.error('❌ Помилка обробки interaction:', error);
        await this.handleInteractionError(interaction, error);
      }
    });

    // Error event
    this.client.on(Events.Error, (error) => {
      this.stats.errors++;
      logger.error('❌ Помилка Discord клієнта:', error);
      
      // Спроба перепідключення при критичних помилках
      if (this.shouldReconnect(error)) {
        this.scheduleReconnect();
      }
    });

    // Disconnect event
    this.client.on(Events.Disconnect, (event) => {
      this.isReady = false;
      logger.warn('🔌 Discord клієнт відключено:', event);
      
      // Автоматичне перепідключення
      this.scheduleReconnect();
    });

    // Reconnecting event
    this.client.on(Events.Reconnecting, () => {
      this.stats.reconnects++;
      logger.info('🔄 Discord клієнт перепідключається...');
    });

    // Guild events
    this.client.on(Events.GuildCreate, (guild) => {
      logger.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`);
    });

    this.client.on(Events.GuildDelete, (guild) => {
      logger.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`);
    });

    logger.info('✅ Обробники подій Discord налаштовано');
  }

  /**
   * Обробка команд з детальним логуванням
   */
  private async handleCommand(interaction: CommandInteraction): Promise<void> {
    const startTime = Date.now();
    const commandName = interaction.commandName;
    const userId = interaction.user.id;
    const guildId = interaction.guildId;
    
    logger.info(`📝 Обробка команди: ${commandName} від користувача ${userId} в сервері ${guildId}`);
    
    try {
      const command = this.commands.get(commandName);
      if (!command) {
        logger.warn(`⚠️ Команда не знайдена: ${commandName}`);
        await interaction.reply({ 
          content: '❌ Команда не знайдена або не зареєстрована', 
          ephemeral: true 
        });
        return;
      }

      // Встановлення таймауту для команди
      const commandTimeout = setTimeout(() => {
        logger.warn(`⏰ Таймаут команди: ${commandName}`);
      }, BOT_CONSTANTS.COMMAND_TIMEOUT);

      await command.execute(interaction);
      
      clearTimeout(commandTimeout);
      this.stats.commands++;
      
      const duration = Date.now() - startTime;
      logger.info(`✅ Команда ${commandName} виконана за ${duration}ms`);
      
    } catch (error) {
      const duration = Date.now() - startTime;
      logger.error(`❌ Помилка виконання команди ${commandName} після ${duration}ms:`, error);
      throw error;
    }
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
            content: `❌ Помилка виконання: ${errorMessage}` 
          });
        } else {
          await interaction.reply({ 
            content: `❌ Помилка виконання: ${errorMessage}`, 
            ephemeral: true 
          });
        }
      }
    } catch (replyError) {
      logger.error('❌ Помилка відповіді на помилку interaction:', replyError);
    }
  }

  /**
   * Очікування готовності клієнта з таймаутом
   */
  private waitForReady(): Promise<void> {
    return new Promise((resolve, reject) => {
      if (this.isReady) {
        resolve();
        return;
      }

      const timeout = setTimeout(() => {
        reject(new Error(`Таймаут очікування готовності клієнта (${BOT_CONSTANTS.READY_TIMEOUT}ms)`));
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
    
    return reconnectErrors.some(errorType => 
      error.message.includes(errorType) || error.name.includes(errorType)
    );
  }

  /**
   * Планування перепідключення
   */
  private scheduleReconnect(): void {
    if (this.reconnectAttempts >= BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
      logger.error(`❌ Досягнуто максимальну кількість спроб перепідключення (${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS})`);
      return;
    }

    this.reconnectAttempts++;
    logger.info(`🔄 Планування перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} через ${BOT_CONSTANTS.RECONNECT_DELAY}ms`);
    
    setTimeout(() => {
      this.connectToDiscord();
    }, BOT_CONSTANTS.RECONNECT_DELAY);
  }

  /**
   * Запуск health check
   */
  private startHealthCheck(): void {
    this.healthCheckInterval = setInterval(async () => {
      try {
        const health = await this.onHealthCheck();
        if (!health.healthy) {
          logger.warn('⚠️ Health check виявив проблеми:', health);
        }
      } catch (error) {
        logger.error('❌ Помилка health check:', error);
      }
    }, BOT_CONSTANTS.HEALTH_CHECK_INTERVAL);
    
    logger.info('🏥 Health check запущено');
  }

  /**
   * Зупинка health check
   */
  private stopHealthCheck(): void {
    if (this.healthCheckInterval) {
      clearInterval(this.healthCheckInterval);
      this.healthCheckInterval = null;
      logger.info('🏥 Health check зупинено');
    }
  }

  /**
   * Очищення ресурсів при помилці
   */
  private async cleanupOnError(): Promise<void> {
    try {
      logger.info('🧹 Очищення ресурсів при помилці...');
      
      this.stopHealthCheck();
      
      if (this.client) {
        this.client.destroy();
      }
      
      logger.info('✅ Ресурси очищено');
    } catch (error) {
      logger.error('❌ Помилка очищення ресурсів:', error);
    }
  }

  /**
   * Логування статистики запуску
   */
  private logStartupStats(): void {
    try {
      const stats = this.getStats();
      logger.info('📊 Статистика запуску бота:', {
        uptime: `${Math.round(stats.uptime / 1000)}s`,
        commands: stats.commands,
        interactions: stats.interactions,
        errors: stats.errors,
        reconnects: stats.reconnects,
        memory: {
          rss: `${Math.round(stats.memory.rss / 1024 / 1024)}MB`,
          heapUsed: `${Math.round(stats.memory.heapUsed / 1024 / 1024)}MB`,
        },
      });
    } catch (error) {
      logger.error('❌ Помилка логування статистики запуску:', error);
    }
  }

  /**
   * Реєстрація команди
   */
  public registerCommand(command: BaseCommand): void {
    try {
      const commandName = command.getName();
      this.commands.set(commandName, command);
      logger.debug(`✅ Команда зареєстрована: ${commandName}`);
    } catch (error) {
      logger.error('❌ Помилка реєстрації команди:', error);
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
    return {
      ...this.getStats(),
      isReady: this.isReady,
      isConnecting: this.isConnecting,
      reconnectAttempts: this.reconnectAttempts,
    };
  }
} 