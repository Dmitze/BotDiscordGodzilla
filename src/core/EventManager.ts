/**
 * Event Manager для Discord бота
 * Централізована обробка Discord подій
 * TypeScript версія
 */

import logger from '../utils/logger';
import { Client, Guild, Message } from 'discord.js';

interface Bot {
  client: Client;
}

type EventHandler = (...args: any[]) => Promise<void> | void;

class EventManager {
  private bot: Bot;
  private events: Map<string, EventHandler>;
  private isActive: boolean;

  constructor(bot: Bot) {
    this.bot = bot;
    this.events = new Map();
    this.isActive = false;
  }

  /**
   * Ініціалізація менеджера подій
   */
  async initialize(): Promise<void> {
    try {
      logger.info('📡 Ініціалізація менеджера подій...');

      // Реєстрація стандартних подій
      this.registerDefaultEvents();

      this.isActive = true;
      logger.info('✅ Менеджер подій ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера подій:', error);
      throw error;
    }
  }

  /**
   * Реєстрація стандартних подій
   */
  private registerDefaultEvents(): void {
    // Ready event
    this.registerEvent('ready', () => {
      logger.info(`🤖 Бот ${this.bot.client.user?.tag} готовий до роботи!`);
      this.bot.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
    });

    // Error event
    this.registerEvent('error', (error: Error) => {
      logger.error('Discord клієнт помилка:', error);
    });

    // Warn event
    this.registerEvent('warn', (warning: string) => {
      logger.warn('Discord клієнт попередження:', warning);
    });

    // Disconnect event
    this.registerEvent('disconnect', () => {
      logger.warn('Discord клієнт відключено');
    });

    // Reconnecting event
    this.registerEvent('reconnecting', () => {
      logger.info('Discord клієнт перепідключається...');
    });

    // Guild Create event
    this.registerEvent('guildCreate', (guild: Guild) => {
      logger.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`);
    });

    // Guild Delete event
    this.registerEvent('guildDelete', (guild: Guild) => {
      logger.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`);
    });

    // Message Create event (для логування)
    this.registerEvent('messageCreate', (message: Message) => {
      if (message.author.bot) return;

      logger.debug(
        `💬 Повідомлення від ${message.author.tag}: ${message.content.substring(0, 50)}...`
      );
    });

    logger.debug('✅ Стандартні події зареєстровано');
  }

  /**
   * Реєстрація події
   */
  registerEvent(eventName: string, handler: EventHandler): void {
    try {
      if (this.events.has(eventName)) {
        logger.warn(`Подія "${eventName}" вже зареєстрована, перезаписуємо`);
      }

      // Обгортка для обробки помилок
      const wrappedHandler: EventHandler = async (...args: any[]) => {
        try {
          await handler(...args);
        } catch (error) {
          logger.error(`Помилка обробки події "${eventName}":`, error);
        }
      };

      this.events.set(eventName, wrappedHandler);
      this.bot.client.on(eventName, wrappedHandler);

      logger.debug(`✅ Подія "${eventName}" зареєстрована`);
    } catch (error) {
      logger.error(`Помилка реєстрації події "${eventName}":`, error);
    }
  }

  /**
   * Видалення події
   */
  removeEvent(eventName: string): void {
    try {
      const handler = this.events.get(eventName);
      if (handler) {
        this.bot.client.off(eventName, handler);
        this.events.delete(eventName);
        logger.debug(`✅ Подія "${eventName}" видалена`);
      }
    } catch (error) {
      logger.error(`Помилка видалення події "${eventName}":`, error);
    }
  }

  /**
   * Отримання списку зареєстрованих подій
   */
  getRegisteredEvents(): string[] {
    return Array.from(this.events.keys());
  }

  /**
   * Перевірка активності
   */
  isActive(): boolean {
    return this.isActive;
  }

  /**
   * Завершення роботи
   */
  async shutdown(): Promise<void> {
    logger.info('🛑 Завершення роботи менеджера подій...');

    try {
      // Видалення всіх подій
      for (const eventName of this.events.keys()) {
        this.removeEvent(eventName);
      }

      this.isActive = false;
      logger.info('✅ Менеджер подій завершено');
    } catch (error) {
      logger.error('❌ Помилка завершення менеджера подій:', error);
    }
  }
}

export default EventManager; 