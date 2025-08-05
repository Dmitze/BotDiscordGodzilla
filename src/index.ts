/**
 * Основний файл Discord AI Assistant Bot
 * Точка входу в додаток
 * Версія 3.0.0 - Повністю рефакторовано з TypeScript
 */

import { config } from 'dotenv';
import { join } from 'path';
import { existsSync } from 'fs';
import type { BotConfig } from '@/types';
import { Bot } from '@/core/Bot';
import { Config } from '@/config/Config';
import logger from '@/utils/logger';

// Константи для конфігурації
const APP_CONFIG = {
  VERSION: '3.0.0',
  NAME: 'Discord AI Assistant Bot',
  STARTUP_TIMEOUT: 30000, // 30 секунд
  SHUTDOWN_TIMEOUT: 10000, // 10 секунд
  RESTART_DELAY: 5000, // 5 секунд
} as const;

// Завантаження змінних середовища
try {
  const envPath = join(process.cwd(), '.env');
  if (existsSync(envPath)) {
    config({ path: envPath });
    logger.info('✅ Змінні середовища завантажено з .env файлу');
  } else {
    config();
    logger.warn('⚠️ .env файл не знайдено, використовую системні змінні');
  }
} catch (error) {
  logger.error('❌ Помилка завантаження змінних середовища:', error);
  throw new Error('Неможливо завантажити змінні середовища');
}

class Application {
  private bot: Bot | null = null;
  private config: BotConfig;
  private isStarting: boolean = false;
  private isShuttingDown: boolean = false;
  private startupTime: number = 0;
  private restartCount: number = 0;
  private readonly maxRestarts: number = 5;

  constructor() {
    try {
      logger.info(`🚀 Ініціалізація ${APP_CONFIG.NAME} v${APP_CONFIG.VERSION}`);
      this.config = Config.load();
      logger.info('✅ Конфігурація завантажена успішно');
    } catch (error) {
      logger.error('❌ Критична помилка ініціалізації додатку:', error);
      throw new Error(`Помилка ініціалізації: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
    }
  }

  /**
   * Запуск додатку з детальним логуванням
   */
  public async start(): Promise<void> {
    if (this.isStarting) {
      logger.warn('⚠️ Додаток вже запускається');
      return;
    }

    if (this.isShuttingDown) {
      logger.warn('⚠️ Неможливо запустити додаток під час зупинки');
      return;
    }

    this.isStarting = true;
    this.startupTime = Date.now();

    try {
      logger.info('🔄 Початок запуску додатку...');
      
      // Валідація конфігурації
      await this.validateConfiguration();
      
      // Створення та ініціалізація бота
      logger.info('🤖 Створення екземпляру бота...');
      this.bot = new Bot(this.config);
      
      logger.info('⚙️ Ініціалізація бота...');
      await this.bot.initialize();
      
      const startupDuration = Date.now() - this.startupTime;
      logger.info(`✅ Додаток успішно запущено за ${startupDuration}ms`);
      
      // Скидання лічильника перезапусків при успішному запуску
      this.restartCount = 0;
      
      // Обробка сигналів завершення
      this.setupGracefulShutdown();
      
      // Логування статистики запуску
      this.logStartupStats();
      
    } catch (error) {
      const startupDuration = Date.now() - this.startupTime;
      logger.error(`❌ Помилка запуску додатку після ${startupDuration}ms:`, error);
      
      // Спроба очищення ресурсів
      await this.cleanupOnError();
      
      throw new Error(`Помилка запуску: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
    } finally {
      this.isStarting = false;
    }
  }

  /**
   * Зупинка додатку з детальним логуванням
   */
  public async stop(): Promise<void> {
    if (this.isShuttingDown) {
      logger.warn('⚠️ Додаток вже зупиняється');
      return;
    }

    this.isShuttingDown = true;
    const shutdownStartTime = Date.now();

    try {
      logger.info('🛑 Початок зупинки додатку...');
      
      if (this.bot) {
        logger.info('🤖 Зупинка бота...');
        await this.bot.shutdown();
        this.bot = null;
      }
      
      const shutdownDuration = Date.now() - shutdownStartTime;
      logger.info(`✅ Додаток успішно зупинено за ${shutdownDuration}ms`);
      
    } catch (error) {
      const shutdownDuration = Date.now() - shutdownStartTime;
      logger.error(`❌ Помилка зупинки додатку після ${shutdownDuration}ms:`, error);
      
      // Примусова зупинка при помилці
      logger.warn('🔄 Примусова зупинка процесу...');
      process.exit(1);
    } finally {
      this.isShuttingDown = false;
    }
  }

  /**
   * Отримання детальної статистики
   */
  public getStats(): any {
    try {
      if (!this.bot) {
        return {
          status: 'not_initialized',
          uptime: process.uptime(),
          memory: process.memoryUsage(),
          version: APP_CONFIG.VERSION,
        };
      }

      const botStats = this.bot.getStats();
      const memoryUsage = process.memoryUsage();
      
      return {
        status: 'running',
        bot: botStats,
        uptime: process.uptime(),
        memory: {
          rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`,
          heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
          heapTotal: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
          external: `${Math.round(memoryUsage.external / 1024 / 1024)}MB`,
        },
        version: APP_CONFIG.VERSION,
        restartCount: this.restartCount,
        startupTime: this.startupTime,
      };
    } catch (error) {
      logger.error('❌ Помилка отримання статистики:', error);
      return {
        status: 'error',
        error: error instanceof Error ? error.message : 'Невідома помилка',
        uptime: process.uptime(),
        version: APP_CONFIG.VERSION,
      };
    }
  }

  /**
   * Перезапуск додатку з обмеженнями
   */
  public async restart(): Promise<void> {
    if (this.restartCount >= this.maxRestarts) {
      const error = `Досягнуто максимальну кількість перезапусків (${this.maxRestarts})`;
      logger.error(`❌ ${error}`);
      throw new Error(error);
    }

    this.restartCount++;
    logger.info(`🔄 Перезапуск додатку (спроба ${this.restartCount}/${this.maxRestarts})...`);

    try {
      // Зупинка поточного екземпляру
      if (this.bot) {
        logger.info('🛑 Зупинка поточного екземпляру...');
        await this.bot.shutdown();
        this.bot = null;
      }

      // Затримка перед перезапуском
      logger.info(`⏳ Затримка ${APP_CONFIG.RESTART_DELAY}ms перед перезапуском...`);
      await new Promise(resolve => setTimeout(resolve, APP_CONFIG.RESTART_DELAY));

      // Запуск нового екземпляру
      logger.info('🚀 Запуск нового екземпляру...');
      await this.start();
      
      logger.info('✅ Додаток успішно перезапущено');
    } catch (error) {
      logger.error('❌ Помилка при перезапуску:', error);
      throw error;
    }
  }

  /**
   * Валідація конфігурації
   */
  private async validateConfiguration(): Promise<void> {
    try {
      logger.info('🔍 Валідація конфігурації...');
      
      // Перевірка обов'язкових полів
      const requiredFields = [
        'discord.token',
        'discord.clientId',
        'discord.guildId',
        'google.apiKey',
        'google.appScriptUrl',
        'ai.openai.apiKey'
      ];

      for (const field of requiredFields) {
        const value = this.getNestedValue(this.config, field);
        if (!value) {
          throw new Error(`Відсутнє обов'язкове поле конфігурації: ${field}`);
        }
      }

      logger.info('✅ Конфігурація валідна');
    } catch (error) {
      logger.error('❌ Помилка валідації конфігурації:', error);
      throw error;
    }
  }

  /**
   * Отримання вкладених значень об'єкта
   */
  private getNestedValue(obj: any, path: string): any {
    return path.split('.').reduce((current, key) => current?.[key], obj);
  }

  /**
   * Очищення ресурсів при помилці
   */
  private async cleanupOnError(): Promise<void> {
    try {
      logger.info('🧹 Очищення ресурсів при помилці...');
      
      if (this.bot) {
        try {
          await this.bot.shutdown();
        } catch (shutdownError) {
          logger.error('❌ Помилка при очищенні бота:', shutdownError);
        }
        this.bot = null;
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
      logger.info('📊 Статистика запуску:', {
        version: stats.version,
        uptime: `${Math.round(stats.uptime)}s`,
        memory: stats.memory,
        restartCount: stats.restartCount,
      });
    } catch (error) {
      logger.error('❌ Помилка логування статистики запуску:', error);
    }
  }

  /**
   * Налаштування graceful shutdown з покращеною обробкою
   */
  private setupGracefulShutdown(): void {
    const shutdown = async (signal: string) => {
      logger.info(`📡 Отримано сигнал ${signal}, початок graceful shutdown...`);
      
      try {
        // Встановлення таймауту для shutdown
        const shutdownTimeout = setTimeout(() => {
          logger.error('⏰ Таймаут graceful shutdown, примусова зупинка');
          process.exit(1);
        }, APP_CONFIG.SHUTDOWN_TIMEOUT);

        await this.stop();
        clearTimeout(shutdownTimeout);
        
        logger.info('✅ Graceful shutdown завершено успішно');
        process.exit(0);
      } catch (error) {
        logger.error('❌ Помилка graceful shutdown:', error);
        process.exit(1);
      }
    };

    // Обробка сигналів завершення
    process.on('SIGINT', () => shutdown('SIGINT'));
    process.on('SIGTERM', () => shutdown('SIGTERM'));
    process.on('SIGQUIT', () => shutdown('SIGQUIT'));

    // Обробка необроблених помилок
    process.on('uncaughtException', (error) => {
      logger.error('💥 Необроблена помилка:', {
        name: error.name,
        message: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString(),
      });
      
      this.handleCriticalError(error);
    });

    process.on('unhandledRejection', (reason, promise) => {
      logger.error('💥 Необроблений rejection:', {
        reason: reason instanceof Error ? reason.message : String(reason),
        promise: promise.toString(),
        timestamp: new Date().toISOString(),
      });
      
      this.handleCriticalError(reason instanceof Error ? reason : new Error(String(reason)));
    });

    logger.info('🛡️ Graceful shutdown налаштовано');
  }

  /**
   * Обробка критичних помилок
   */
  private async handleCriticalError(error: Error): Promise<void> {
    try {
      logger.error('🚨 Обробка критичної помилки...');
      
      // Спроба graceful shutdown
      await this.stop();
    } catch (shutdownError) {
      logger.error('❌ Помилка при обробці критичної помилки:', shutdownError);
    } finally {
      // Примусова зупинка через 5 секунд
      setTimeout(() => {
        logger.error('⏰ Примусова зупинка через критичну помилку');
        process.exit(1);
      }, 5000);
    }
  }
}

// Глобальний екземпляр додатку
let app: Application | null = null;

/**
 * Головна функція запуску з покращеною обробкою помилок
 */
async function main(): Promise<void> {
  const startTime = Date.now();
  
  try {
    logger.info(`🎯 Запуск ${APP_CONFIG.NAME} v${APP_CONFIG.VERSION}`);
    
    app = new Application();
    await app.start();
    
    const totalStartupTime = Date.now() - startTime;
    logger.info(`🎉 Додаток повністю запущено за ${totalStartupTime}ms`);
    
  } catch (error) {
    const totalStartupTime = Date.now() - startTime;
    logger.error(`💥 Критична помилка при запуску після ${totalStartupTime}ms:`, error);
    
    // Детальне логування помилки
    if (error instanceof Error) {
      logger.error('Деталі помилки:', {
        name: error.name,
        message: error.message,
        stack: error.stack,
      });
    }
    
    process.exit(1);
  }
}

/**
 * Функції для зовнішнього використання з покращеною обробкою помилок
 */
export {
  main,
  getStats: () => {
    try {
      return app?.getStats() || { status: 'not_initialized' };
    } catch (error) {
      logger.error('❌ Помилка отримання статистики:', error);
      return { status: 'error', error: error instanceof Error ? error.message : 'Невідома помилка' };
    }
  },
  restart: async () => {
    try {
      if (!app) {
        throw new Error('Додаток не ініціалізовано');
      }
      return await app.restart();
    } catch (error) {
      logger.error('❌ Помилка перезапуску:', error);
      throw error;
    }
  },
  shutdown: async () => {
    try {
      if (!app) {
        logger.warn('⚠️ Додаток не ініціалізовано для зупинки');
        return;
      }
      return await app.stop();
    } catch (error) {
      logger.error('❌ Помилка зупинки:', error);
      throw error;
    }
  },
  getApp: () => app,
  APP_CONFIG,
};

// Запуск додатку, якщо файл виконано напряму
if (require.main === module) {
  main().catch((error) => {
    logger.error('💥 Фатальна помилка в головній функції:', error);
    process.exit(1);
  });
} 