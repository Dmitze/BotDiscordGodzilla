/**
 * Основний файл Discord AI Assistant Bot
 * Точка входу в додаток
 */

import { config } from 'dotenv';
import type { BotConfig } from '@/types';
import { Bot } from '@/core/Bot';
import { Config } from '@/config/Config';

// Завантаження змінних середовища
config();

class Application {
  private bot: Bot | null = null;
  private config: BotConfig;

  constructor() {
    this.config = Config.load();
  }

  /**
   * Запуск додатку
   */
  public async start(): Promise<void> {
    try {
      console.log('🚀 Запуск Discord AI Assistant Bot...');
      
      // Створення та ініціалізація бота
      this.bot = new Bot(this.config);
      await this.bot.initialize();
      
      console.log('✅ Додаток успішно запущено');
      
      // Обробка сигналів завершення
      this.setupGracefulShutdown();
      
    } catch (error) {
      console.error('❌ Помилка запуску додатку:', error);
      process.exit(1);
    }
  }

  /**
   * Зупинка додатку
   */
  public async stop(): Promise<void> {
    try {
      console.log('🛑 Зупинка додатку...');
      
      if (this.bot) {
        await this.bot.shutdown();
      }
      
      console.log('✅ Додаток успішно зупинено');
      process.exit(0);
      
    } catch (error) {
      console.error('❌ Помилка зупинки додатку:', error);
      process.exit(1);
    }
  }

  /**
   * Налаштування graceful shutdown
   */
  private setupGracefulShutdown(): void {
    const shutdown = async (signal: string) => {
      console.log(`\n📡 Отримано сигнал ${signal}, початок graceful shutdown...`);
      await this.stop();
    };

    process.on('SIGINT', () => shutdown('SIGINT'));
    process.on('SIGTERM', () => shutdown('SIGTERM'));
    process.on('SIGQUIT', () => shutdown('SIGQUIT'));

    // Обробка необроблених помилок
    process.on('uncaughtException', (error) => {
      console.error('💥 Необроблена помилка:', error);
      this.stop();
    });

    process.on('unhandledRejection', (reason, promise) => {
      console.error('💥 Необроблений rejection:', reason, 'в', promise);
      this.stop();
    });
  }
}

// Запуск додатку
const app = new Application();
app.start().catch((error) => {
  console.error('❌ Критична помилка:', error);
  process.exit(1);
}); 