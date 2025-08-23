/**
 * Event Manager для Discord бота
 * Централізована обробка Discord подій
 * TypeScript версія
 */

import logger from '@/utils/logger';
import type { Client, Guild, Message} from 'discord.js';
import { Events } from 'discord.js';

interface Bot {
  client: Client;
}

type EventHandler = (...args: any[]) => Promise<void> | void;

class EventManager {
  private bot: Bot;
  private events: Map<string, EventHandler>;
  private active: boolean;

  constructor(bot: Bot) {
    this.bot = bot;
    this.events = new Map();
    this.active = false;
  }

  /**
   * Ініціалізація менеджера подій
   */
  async initialize(): Promise<void> {
    try {
      logger.info('📡 Ініціалізація менеджера подій...', {
        type: 'system',
        event: 'event_manager_init',
      });

      // Реєстрація стандартних подій
      this.registerDefaultEvents();

      this.active = true;

      logger.info('✅ Менеджер подій ініціалізовано', {
        type: 'system',
        event: 'event_manager_init_success',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації менеджера подій', {
          type: 'system',
          event: 'event_manager_init_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка ініціалізації менеджера подій', {
          type: 'system',
          event: 'event_manager_init_failed',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Реєстрація стандартних подій
   */
  private registerDefaultEvents(): void {
    // Ready event (v14)
    this.registerEvent(Events.ClientReady, () => {
      logger.info(`🤖 Бот ${this.bot.client.user?.tag} готовий до роботи!`, {
        type: 'event',
        event: 'ready',
      });
      this.bot.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
    });

    // Error event (v14)
    this.registerEvent(Events.Error, (error: Error) => {
      logger.error('Discord клієнт помилка', {
        type: 'system',
        event: 'client_error',
        errorName: error.name,
        errorMessage: error.message,
        stack: error.stack,
      });
    });

    // Warn event (v14)
    this.registerEvent(Events.Warn, (warning: string) => {
      logger.warn('Discord клієнт попередження', {
        type: 'system',
        event: 'client_warn',
        warning,
      });
    });

    // Shard disconnect event (v14)
    this.registerEvent(Events.ShardDisconnect, () => {
      logger.warn('Discord клієнт відключено', {
        type: 'event',
        event: 'shard_disconnect',
      });
    });

    // Shard reconnecting event (v14)
    this.registerEvent(Events.ShardReconnecting, () => {
      logger.info('Discord клієнт перепідключається...', {
        type: 'event',
        event: 'shard_reconnecting',
      });
    });

    // Guild Create event (v14)
    this.registerEvent(Events.GuildCreate, (guild: Guild) => {
      logger.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`, {
        type: 'event',
        event: 'guildCreate',
        guildId: guild.id,
      });
    });

    // Guild Delete event (v14)
    this.registerEvent(Events.GuildDelete, (guild: Guild) => {
      logger.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`, {
        type: 'event',
        event: 'guildDelete',
        guildId: guild.id,
      });
    });

    // Message Create event (для логування, v14)
    this.registerEvent(Events.MessageCreate, (message: Message) => {
      if (message.author.bot) return;

      logger.debug(
        `💬 Повідомлення від ${message.author.tag}: ${message.content.substring(0, 50)}...`,
        {
          type: 'event',
          event: 'messageCreate',
          userId: message.author.id,
          ...(message.guild?.id ? { guildId: message.guild.id } : {}),
          ...(message.channel?.id ? { channelId: message.channel.id } : {}),
        }
      );
    });

    logger.debug('✅ Стандартні події зареєстровано', {
      type: 'system',
      event: 'default_events_registered',
    });
  }

  /**
   * Реєстрація події
   */
  registerEvent(eventName: string, handler: EventHandler): void {
    try {
      if (this.events.has(eventName)) {
        logger.warn(`Подія "${eventName}" вже зареєстрована, перезаписуємо`, {
          type: 'system',
          event: 'event_register_warning',
          eventName,
        });
      }

      // Обгортка для обробки помилок
      const wrappedHandler: EventHandler = async (...args: any[]) => {
        try {
          await handler(...args);
        } catch (error) {
          if (error instanceof Error) {
            logger.error(`Помилка обробки події "${eventName}"`, {
              type: 'event',
              event: eventName,
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
            });
          } else {
            logger.error(`Помилка обробки події "${eventName}"`, {
              type: 'event',
              event: eventName,
              errorMessage: String(error),
            });
          }
        }
      };

      this.events.set(eventName, wrappedHandler);
      this.bot.client.on(eventName, wrappedHandler);

      logger.debug(`✅ Подія "${eventName}" зареєстрована`, {
        type: 'system',
        event: 'event_registered',
        eventName,
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error(`Помилка реєстрації події "${eventName}"`, {
          type: 'system',
          event: 'event_register_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error(`Помилка реєстрації події "${eventName}"`, {
          type: 'system',
          event: 'event_register_failed',
          errorMessage: String(error),
        });
      }
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
        logger.debug(`✅ Подія "${eventName}" видалена`, {
          type: 'system',
          event: 'event_removed',
          eventName,
        });
      }
    } catch (error) {
      if (error instanceof Error) {
        logger.error(`Помилка видалення події "${eventName}"`, {
          type: 'system',
          event: 'event_remove_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error(`Помилка видалення події "${eventName}"`, {
          type: 'system',
          event: 'event_remove_failed',
          errorMessage: String(error),
        });
      }
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
    return this.active;
  }

  /**
   * Завершення роботи
   */
  async shutdown(): Promise<void> {
    logger.info('🛑 Завершення роботи менеджера подій...', {
      type: 'system',
      event: 'event_manager_shutdown',
    });

    try {
      // Видалення всіх подій
      for (const eventName of this.events.keys()) {
        this.removeEvent(eventName);
      }

      this.active = false;

      logger.info('✅ Менеджер подій завершено', {
        type: 'system',
        event: 'event_manager_shutdown_success',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка завершення менеджера подій', {
          type: 'system',
          event: 'event_manager_shutdown_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка завершення менеджера подій', {
          type: 'system',
          event: 'event_manager_shutdown_failed',
          errorMessage: String(error),
        });
      }
    }
  }
}

export default EventManager;
