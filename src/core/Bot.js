/**
 * Основний клас Discord бота
 * Рефакторована архітектура з Dependency Injection
 */

const { Client, GatewayIntentBits, Collection } = require('discord.js');
const logger = require('../utils/logger');
const CommandManager = require('./CommandManager');
const EventManager = require('./EventManager');

class Bot {
  constructor(serviceContainer, errorHandler) {
    this.serviceContainer = serviceContainer;
    this.errorHandler = errorHandler;
    this.config = serviceContainer.getConfig();
    
    this.client = null;
    this.commands = new Collection();
    this.isReady = false;

    // Менеджери
    this.commandManager = null;
    this.eventManager = null;
  }

  /**
   * Ініціалізація бота
   */
  async initialize() {
    try {
      logger.info('🤖 Ініціалізація Discord бота...');

      // Створення Discord клієнта
      await this.createClient();

      // Ініціалізація менеджерів
      await this.initializeManagers();

      // Підключення до Discord
      await this.connect();

      // Реєстрація бота в Service Container
      this.serviceContainer.register('bot', () => this);

      this.isReady = true;
      logger.info('✅ Discord бот успішно ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації бота:', error);
      await this.errorHandler.handle(error, { context: 'Bot.initialize' });
      throw error;
    }
  }

  /**
   * Створення Discord клієнта
   */
  async createClient() {
    try {
      this.client = new Client({
        intents: [
          GatewayIntentBits.Guilds,
          GatewayIntentBits.GuildMessages,
          GatewayIntentBits.MessageContent,
          GatewayIntentBits.GuildMessageReactions,
        ],
      });

      // Налаштування обробників подій клієнта
      this.setupClientEventHandlers();

      logger.info('✅ Discord клієнт створено');
    } catch (error) {
      logger.error('❌ Помилка створення Discord клієнта:', error);
      throw error;
    }
  }

  /**
   * Налаштування обробників подій клієнта
   */
  setupClientEventHandlers() {
    // Ready event
    this.client.once('ready', () => {
      logger.info(`🤖 Бот ${this.client.user.tag} підключено до Discord`);
      logger.info(`📊 Бот працює на ${this.client.guilds.cache.size} серверах`);
    });

    // Error event
    this.client.on('error', async (error) => {
      logger.error('❌ Discord клієнт помилка:', error);
      await this.errorHandler.handle(error, { context: 'DiscordClient.error' });
    });

    // Warn event
    this.client.on('warn', (warning) => {
      logger.warn('⚠️ Discord клієнт попередження:', warning);
    });

    // Debug event (тільки в development)
    if (process.env.NODE_ENV === 'development') {
      this.client.on('debug', (info) => {
        logger.debug('🔍 Discord debug:', info);
      });
    }
  }

  /**
   * Ініціалізація менеджерів
   */
  async initializeManagers() {
    try {
      // Command Manager
      this.commandManager = new CommandManager(this);
      await this.commandManager.initialize();

      // Event Manager
      this.eventManager = new EventManager(this);
      await this.eventManager.initialize();

      logger.info('✅ Менеджери ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджерів:', error);
      throw error;
    }
  }

  /**
   * Підключення до Discord
   */
  async connect() {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error('Timeout connecting to Discord'));
      }, 30000); // 30 секунд timeout

      this.client.once('ready', () => {
        clearTimeout(timeout);
        resolve();
      });

      this.client.on('error', (error) => {
        clearTimeout(timeout);
        reject(error);
      });

      this.client.login(this.config.discord.token).catch(reject);
    });
  }

  /**
   * Отримання команди за назвою
   */
  getCommand(name) {
    return this.commands.get(name);
  }

  /**
   * Отримання сервісу
   */
  getService(name) {
    return this.serviceContainer.get(name);
  }

  /**
   * Обробка помилки
   */
  async handleError(error, context = {}) {
    return await this.errorHandler.handle(error, {
      ...context,
      bot: true,
    });
  }

  /**
   * Перевірка готовності бота
   */
  isReady() {
    return this.isReady && this.client && this.client.isReady();
  }

  /**
   * Отримання статистики бота
   */
  getStats() {
    if (!this.client) {
      return { status: 'not_initialized' };
    }

    return {
      status: this.isReady ? 'ready' : 'initializing',
      user: this.client.user ? {
        id: this.client.user.id,
        tag: this.client.user.tag,
        username: this.client.user.username,
      } : null,
      guilds: this.client.guilds.cache.size,
      channels: this.client.channels.cache.size,
      users: this.client.users.cache.size,
      commands: this.commands.size,
      uptime: this.client.uptime,
      ping: this.client.ws.ping,
      memory: process.memoryUsage(),
    };
  }

  /**
   * Health check
   */
  isHealthy() {
    return this.isReady && 
           this.client && 
           this.client.isReady() && 
           this.client.ws.ping < 1000; // Ping менше 1 секунди
  }

  /**
   * Завершення роботи бота
   */
  async shutdown() {
    logger.info('🛑 Завершення роботи Discord бота...');

    try {
      // Завершення менеджерів
      if (this.eventManager) {
        await this.eventManager.shutdown();
      }

      if (this.commandManager) {
        await this.commandManager.shutdown();
      }

      // Відключення від Discord
      if (this.client) {
        this.client.destroy();
      }

      this.isReady = false;
      logger.info('✅ Discord бот завершено');
    } catch (error) {
      logger.error('❌ Помилка завершення бота:', error);
      throw error;
    }
  }

  /**
   * Перезапуск бота
   */
  async restart() {
    logger.info('🔄 Перезапуск Discord бота...');

    try {
      await this.shutdown();
      await this.initialize();
      logger.info('✅ Discord бот перезапущено');
    } catch (error) {
      logger.error('❌ Помилка перезапуску бота:', error);
      throw error;
    }
  }

  /**
   * Отримання конфігурації
   */
  getConfig() {
    return this.config;
  }

  /**
   * Отримання Service Container
   */
  getServiceContainer() {
    return this.serviceContainer;
  }

  /**
   * Отримання Error Handler
   */
  getErrorHandler() {
    return this.errorHandler;
  }
}

module.exports = Bot;

