/**
 * Головний файл Discord AI Assistant Bot
 * Рефакторована архітектура з Dependency Injection та Service Layer
 * Версія: 3.0.0
 */

const Bot = require('./core/Bot');
const logger = require('./utils/logger');
const { Config } = require('./config/Config');
const { ServiceContainer } = require('./core/ServiceContainer');
const { ErrorHandler } = require('./core/ErrorHandler');

/**
 * Головний клас додатку
 */
class Application {
  constructor() {
    this.bot = null;
    this.serviceContainer = null;
    this.errorHandler = null;
    this.config = null;
    this.isInitialized = false;
  }

  /**
   * Ініціалізація додатку
   */
  async initialize() {
    try {
      logger.info('🚀 Ініціалізація Discord AI Assistant Bot v3.0.0...');

      // 1. Завантаження конфігурації
      await this.loadConfiguration();

      // 2. Ініціалізація Service Container
      await this.initializeServiceContainer();

      // 3. Ініціалізація Error Handler
      await this.initializeErrorHandler();

      // 4. Ініціалізація бота
      await this.initializeBot();

      // 5. Налаштування graceful shutdown
      this.setupGracefulShutdown();

      this.isInitialized = true;
      logger.info('✅ Додаток успішно ініціалізовано');
    } catch (error) {
      logger.error('❌ Критична помилка при ініціалізації:', error);
      throw error;
    }
  }

  /**
   * Завантаження конфігурації
   */
  async loadConfiguration() {
    try {
      this.config = new Config();
      await this.config.validate();
      logger.info('✅ Конфігурація завантажена');
    } catch (error) {
      logger.error('❌ Помилка завантаження конфігурації:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація Service Container
   */
  async initializeServiceContainer() {
    try {
      this.serviceContainer = new ServiceContainer(this.config);
      await this.serviceContainer.initialize();
      logger.info('✅ Service Container ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Service Container:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація Error Handler
   */
  async initializeErrorHandler() {
    try {
      this.errorHandler = new ErrorHandler(this.serviceContainer);
      await this.errorHandler.initialize();
      logger.info('✅ Error Handler ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Error Handler:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація бота
   */
  async initializeBot() {
    try {
      this.bot = new Bot(this.serviceContainer, this.errorHandler);
      await this.bot.initialize();
      logger.info('✅ Discord бот ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації бота:', error);
      throw error;
    }
  }

  /**
   * Налаштування graceful shutdown
   */
  setupGracefulShutdown() {
    const shutdown = async (signal) => {
      logger.info(`📡 Отримано сигнал ${signal}, завершення роботи...`);

      try {
        if (this.bot) {
          await this.bot.shutdown();
        }

        if (this.serviceContainer) {
          await this.serviceContainer.shutdown();
        }

        logger.info('✅ Graceful shutdown завершено');
        process.exit(0);
      } catch (error) {
        logger.error('❌ Помилка при shutdown:', error);
        process.exit(1);
      }
    };

    // Обробка сигналів
    process.on('SIGINT', () => shutdown('SIGINT'));
    process.on('SIGTERM', () => shutdown('SIGTERM'));
    process.on('SIGQUIT', () => shutdown('SIGQUIT'));

    // Обробка необроблених помилок
    process.on('uncaughtException', (error) => {
      logger.error('❌ Необроблена помилка:', error);
      if (this.errorHandler) {
        this.errorHandler.handleUncaughtException(error);
      }
    });

    process.on('unhandledRejection', (reason, promise) => {
      logger.error('❌ Необроблений rejection:', reason);
      if (this.errorHandler) {
        this.errorHandler.handleUnhandledRejection(reason, promise);
      }
    });
  }

  /**
   * Отримання статистики
   */
  getStats() {
    if (!this.isInitialized) return null;

    return {
      bot: this.bot?.getStats(),
      services: this.serviceContainer?.getStats(),
      errors: this.errorHandler?.getStats(),
      uptime: process.uptime(),
      memory: process.memoryUsage(),
    };
  }

  /**
   * Перезапуск додатку
   */
  async restart() {
    logger.info('🔄 Перезапуск додатку...');

    try {
      if (this.bot) {
        await this.bot.shutdown();
      }

      if (this.serviceContainer) {
        await this.serviceContainer.shutdown();
      }

      await this.initialize();
      logger.info('✅ Додаток успішно перезапущено');
    } catch (error) {
      logger.error('❌ Помилка при перезапуску:', error);
      throw error;
    }
  }

  /**
   * Завершення роботи
   */
  async shutdown() {
    logger.info('🛑 Завершення роботи додатку...');

    try {
      if (this.bot) {
        await this.bot.shutdown();
      }

      if (this.serviceContainer) {
        await this.serviceContainer.shutdown();
      }

      logger.info('✅ Додаток успішно завершено');
    } catch (error) {
      logger.error('❌ Помилка при завершенні:', error);
      throw error;
    }
  }
}

// Глобальний екземпляр додатку
let app = null;

/**
 * Головна функція запуску
 */
async function main() {
  try {
    app = new Application();
    await app.initialize();
  } catch (error) {
    logger.error('❌ Критична помилка при запуску:', error);
    process.exit(1);
  }
}

/**
 * Функції для зовнішнього використання
 */
module.exports = {
  main,
  getStats: () => app?.getStats(),
  restart: () => app?.restart(),
  shutdown: () => app?.shutdown(),
  getApp: () => app,
};

// Запуск додатку, якщо файл виконано напряму
if (require.main === module) {
  main().catch((error) => {
    logger.error('❌ Помилка в головній функції:', error);
    process.exit(1);
  });
}
