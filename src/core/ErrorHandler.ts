/**
 * Error Handler для Discord бота
 * Централізована обробка помилок з покращеною архітектурою
 * TypeScript версія
 */

import logger from '@/utils/logger';

interface ErrorType {
  severity: 'critical' | 'error' | 'warn' | 'info';
  category: string;
  retryable: boolean;
  maxRetries: number;
  notificationThreshold: number;
}

interface ErrorInfo extends ErrorType {
  type: string;
}

interface ErrorContext {
  type?: string;
  promise?: Promise<any>;
  [key: string]: any;
}

interface Notification {
  error: Error;
  errorInfo: ErrorInfo;
  context: ErrorContext;
  timestamp: Date;
  id: string;
}

interface ErrorStats {
  totalErrors: number;
  errorCounts: Record<string, number>;
  notificationQueueSize: number;
  isActive: boolean;
}

interface ServiceContainer {
  get(service: string): any;
  shutdown(): Promise<void>;
}

class ErrorHandler {
  private serviceContainer: ServiceContainer;
  private errorTypes: Map<string, ErrorType>;
  private errorCounts: Map<string, number>;
  private active: boolean;
  private notificationQueue: Notification[];
  private maxQueueSize: number;

  constructor(serviceContainer: ServiceContainer) {
    this.serviceContainer = serviceContainer;
    this.errorTypes = new Map();
    this.errorCounts = new Map();
    this.active = false;

    this.notificationQueue = [];
    this.maxQueueSize = 100;
  }

  /**
   * Ініціалізація обробника помилок
   */
  async initialize(): Promise<void> {
    try {
      logger.info('🛡️ Ініціалізація обробника помилок...', {
        type: 'system',
        event: 'error_handler_init',
      });

      // Реєстрація типів помилок
      this.registerErrorTypes();

      // Запуск обробки черги сповіщень
      this.startNotificationProcessor();

      this.active = true;

      logger.info('✅ Обробник помилок ініціалізовано', {
        type: 'system',
        event: 'error_handler_init_success',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації обробника помилок', {
          type: 'system',
          event: 'error_handler_init_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          severity: 'critical',
        });
      } else {
        logger.error('❌ Помилка ініціалізації обробника помилок', {
          type: 'system',
          event: 'error_handler_init_failed',
          severity: 'critical',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Реєстрація типів помилок
   */
  private registerErrorTypes(): void {
    // Discord API помилки
    this.errorTypes.set('DiscordAPIError', {
      severity: 'error',
      category: 'discord',
      retryable: true,
      maxRetries: 3,
      notificationThreshold: 5,
    });

    // Google API помилки
    this.errorTypes.set('GoogleAPIError', {
      severity: 'error',
      category: 'google',
      retryable: true,
      maxRetries: 3,
      notificationThreshold: 3,
    });

    // AI API помилки
    this.errorTypes.set('AIAPIError', {
      severity: 'error',
      category: 'ai',
      retryable: true,
      maxRetries: 2,
      notificationThreshold: 2,
    });

    // Валідація помилки
    this.errorTypes.set('ValidationError', {
      severity: 'warn',
      category: 'validation',
      retryable: false,
      maxRetries: 0,
      notificationThreshold: 10,
    });

    // Помилки кешу
    this.errorTypes.set('CacheError', {
      severity: 'warn',
      category: 'cache',
      retryable: true,
      maxRetries: 2,
      notificationThreshold: 5,
    });

    // Помилки файлів
    this.errorTypes.set('FileError', {
      severity: 'error',
      category: 'file',
      retryable: true,
      maxRetries: 2,
      notificationThreshold: 3,
    });

    // Помилки мережі
    this.errorTypes.set('NetworkError', {
      severity: 'error',
      category: 'network',
      retryable: true,
      maxRetries: 5,
      notificationThreshold: 2,
    });

    // Помилки бази даних
    this.errorTypes.set('DatabaseError', {
      severity: 'error',
      category: 'database',
      retryable: true,
      maxRetries: 3,
      notificationThreshold: 1,
    });

    logger.debug('✅ Типи помилок зареєстровано', {
      type: 'system',
      event: 'error_types_registered',
    });
  }

  /**
   * Обробка помилки
   */
  async handle(
    error: Error,
    context: ErrorContext = {}
  ): Promise<{
    handled: boolean;
    errorInfo?: ErrorInfo;
    retryable?: boolean;
    maxRetries?: number;
    error?: Error;
  }> {
    try {
      // Класифікація помилки
      const errorInfo = this.classifyError(error);

      // Логування помилки
      this.logError(error, errorInfo, context);

      // Підрахунок помилок
      this.incrementErrorCount(errorInfo.type);

      // Перевірка чи потрібно відправити сповіщення
      if (this.shouldSendNotification(errorInfo.type)) {
        await this.queueNotification(error, errorInfo, context);
      }

      // Повернення інформації про помилку
      return {
        handled: true,
        errorInfo,
        retryable: errorInfo.retryable,
        maxRetries: errorInfo.maxRetries,
      };
    } catch (handleError: any) {
      if (handleError instanceof Error) {
        logger.error('❌ Помилка в обробнику помилок', {
          type: 'system',
          event: 'error_handler_runtime_error',
          errorName: handleError.name,
          errorMessage: handleError.message,
          stack: handleError.stack,
          severity: 'high',
        });
      } else {
        logger.error('❌ Помилка в обробнику помилок', {
          type: 'system',
          event: 'error_handler_runtime_error',
          severity: 'high',
          errorMessage: String(handleError),
        });
      }
      return {
        handled: false,
        error: handleError,
      };
    }
  }

  /**
   * Обробка необроблених помилок
   */
  handleUncaughtException(error: Error): void {
    logger.error('🚨 КРИТИЧНА ПОМИЛКА - Uncaught Exception', {
      type: 'system',
      event: 'uncaught_exception',
      errorName: error.name,
      errorMessage: error.message,
      stack: error.stack,
      severity: 'critical',
    });

    const errorInfo: ErrorInfo = {
      type: 'UncaughtException',
      severity: 'critical',
      category: 'system',
      retryable: false,
      maxRetries: 0,
      notificationThreshold: 0,
    };

    this.logError(error, errorInfo, { type: 'uncaughtException' });
    this.incrementErrorCount('UncaughtException');

    // Критичні помилки завжди потребують сповіщення
    this.queueNotification(error, errorInfo, { type: 'uncaughtException' });

    // Спроба graceful shutdown
    this.attemptGracefulShutdown();
  }

  /**
   * Обробка необроблених rejections
   */
  handleUnhandledRejection(reason: any, promise: Promise<any>): void {
    const meta =
      reason instanceof Error
        ? {
            errorName: reason.name,
            errorMessage: reason.message,
            stack: reason.stack,
          }
        : { errorMessage: String(reason) };
    logger.error('🚨 КРИТИЧНА ПОМИЛКА - Unhandled Rejection', {
      type: 'system',
      event: 'unhandled_rejection',
      severity: 'critical',
      ...meta,
    });

    const errorInfo: ErrorInfo = {
      type: 'UnhandledRejection',
      severity: 'critical',
      category: 'system',
      retryable: false,
      maxRetries: 0,
      notificationThreshold: 0,
    };

    const errorObj: Error = reason instanceof Error ? reason : new Error(String(reason));
    this.logError(errorObj, errorInfo, { type: 'unhandledRejection', promise });
    this.incrementErrorCount('UnhandledRejection');

    // Критичні помилки завжди потребують сповіщення
    this.queueNotification(errorObj, errorInfo, { type: 'unhandledRejection', promise });

    // Спроба graceful shutdown
    this.attemptGracefulShutdown();
  }

  /**
   * Спроба graceful shutdown
   */
  private async attemptGracefulShutdown(): Promise<void> {
    try {
      logger.warn('🛑 Спроба graceful shutdown через критичну помилку...');

      if (this.serviceContainer) {
        await this.serviceContainer.shutdown();
      }

      logger.info('✅ Graceful shutdown завершено');
      process.exit(1);
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка при graceful shutdown', {
          type: 'system',
          event: 'graceful_shutdown_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          severity: 'high',
        });
      } else {
        logger.error('❌ Помилка при graceful shutdown', {
          type: 'system',
          event: 'graceful_shutdown_failed',
          severity: 'high',
          errorMessage: String(error),
        });
      }
      process.exit(1);
    }
  }

  /**
   * Додавання сповіщення до черги
   */
  private async queueNotification(
    error: Error,
    errorInfo: ErrorInfo,
    context: ErrorContext
  ): Promise<void> {
    const notification: Notification = {
      error,
      errorInfo,
      context,
      timestamp: new Date(),
      id: this.generateNotificationId(),
    };

    this.notificationQueue.push(notification);

    // Обмеження розміру черги
    if (this.notificationQueue.length > this.maxQueueSize) {
      this.notificationQueue.shift();
    }

    logger.debug('📧 Сповіщення додано до черги', {
      type: 'system',
      event: 'notification_queued',
      notificationId: notification.id,
    });
  }

  /**
   * Запуск обробника сповіщень
   */
  private startNotificationProcessor(): void {
    setInterval(async () => {
      if (this.notificationQueue.length > 0) {
        const notification = this.notificationQueue.shift()!;
        await this.sendNotification(
          notification.error,
          notification.errorInfo,
          notification.context
        );
      }
    }, 5000); // Обробка кожні 5 секунд
  }

  /**
   * Генерація ID для сповіщення
   */
  private generateNotificationId(): string {
    return `notif_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Класифікація помилки
   */
  private classifyError(error: Error): ErrorInfo {
    const errorName = error.name || 'UnknownError';
    const defaultInfo: ErrorInfo = {
      type: errorName,
      severity: 'error',
      category: 'unknown',
      retryable: false,
      maxRetries: 0,
      notificationThreshold: 5,
    };

    const errorType = this.errorTypes.get(errorName);
    return errorType ? { ...errorType, type: errorName } : defaultInfo;
  }

  /**
   * Логування помилки
   */
  private logError(error: Error, errorInfo: ErrorInfo, context: ErrorContext): void {
    const logData = {
      error: {
        name: error.name,
        message: error.message,
        stack: error.stack,
      },
      errorInfo,
      context,
      timestamp: new Date().toISOString(),
    };

    switch (errorInfo.severity) {
      case 'critical':
        logger.error('🚨 КРИТИЧНА ПОМИЛКА:', logData);
        break;
      case 'error':
        logger.error('❌ ПОМИЛКА:', logData);
        break;
      case 'warn':
        logger.warn('⚠️ ПОПЕРЕДЖЕННЯ:', logData);
        break;
      default:
        logger.info('ℹ️ ІНФОРМАЦІЯ:', logData);
    }
  }

  /**
   * Підрахунок помилок
   */
  private incrementErrorCount(type: string): void {
    const currentCount = this.errorCounts.get(type) || 0;
    this.errorCounts.set(type, currentCount + 1);
  }

  /**
   * Відправка сповіщення
   */
  private async sendNotification(
    error: Error,
    errorInfo: ErrorInfo,
    context: ErrorContext
  ): Promise<void> {
    try {
      // Спроба відправити Discord сповіщення
      await this.sendDiscordNotification(error, errorInfo, context);

      // Можна додати інші канали сповіщень (email, Slack, etc.)
    } catch (notificationError) {
      if (notificationError instanceof Error) {
        logger.error('❌ Помилка відправки сповіщення', {
          type: 'system',
          event: 'notification_send_failed',
          errorName: notificationError.name,
          errorMessage: notificationError.message,
          stack: notificationError.stack,
          severity: 'medium',
        });
      } else {
        logger.error('❌ Помилка відправки сповіщення', {
          type: 'system',
          event: 'notification_send_failed',
          severity: 'medium',
          errorMessage: String(notificationError),
        });
      }
    }
  }

  /**
   * Перевірка чи потрібно відправити сповіщення
   */
  private shouldSendNotification(errorType: string): boolean {
    const errorInfo = this.errorTypes.get(errorType);
    if (!errorInfo) return false;

    const currentCount = this.errorCounts.get(errorType) || 0;
    return currentCount >= errorInfo.notificationThreshold;
  }

  /**
   * Відправка Discord сповіщення
   */
  private async sendDiscordNotification(
    error: Error,
    errorInfo: ErrorInfo,
    context: ErrorContext
  ): Promise<void> {
    try {
      // Отримання Discord клієнта з Service Container
      const bot = this.serviceContainer?.get('bot');
      if (!bot || !bot.client) {
        logger.warn('Discord клієнт недоступний для сповіщення', {
          type: 'system',
          event: 'notification_client_unavailable',
        });
        return;
      }

      const channel = this.findNotificationChannel(bot.client);
      if (!channel) {
        logger.warn('Канал сповіщень не знайдено', {
          type: 'system',
          event: 'notification_channel_not_found',
        });
        return;
      }

      const embed = {
        color: this.getErrorColor(errorInfo.severity),
        title: `🚨 Помилка: ${errorInfo.type}`,
        description: this.getUserFriendlyMessage(error, errorInfo),
        fields: [
          {
            name: 'Категорія',
            value: errorInfo.category,
            inline: true,
          },
          {
            name: 'Серйозність',
            value: errorInfo.severity,
            inline: true,
          },
          {
            name: 'Контекст',
            value: JSON.stringify(context, null, 2),
            inline: false,
          },
        ],
        timestamp: new Date(),
        footer: {
          text: 'Discord AI Bot - Error Handler',
        },
      };

      await channel.send({ embeds: [embed] });
      logger.info('✅ Discord сповіщення відправлено', {
        type: 'system',
        event: 'notification_sent',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка відправки Discord сповіщення', {
          type: 'system',
          event: 'discord_notification_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка відправки Discord сповіщення', {
          type: 'system',
          event: 'discord_notification_failed',
          errorMessage: String(error),
        });
      }
    }
  }

  /**
   * Отримання кольору для помилки
   */
  private getErrorColor(severity: string): number {
    switch (severity) {
      case 'critical':
        return 0xff0000; // Червоний
      case 'error':
        return 0xff6b6b; // Світло-червоний
      case 'warn':
        return 0xffa500; // Помаранчевий
      default:
        return 0x808080; // Сірий
    }
  }

  /**
   * Пошук каналу для сповіщень
   */
  private findNotificationChannel(client: any): any {
    try {
      // Пошук каналу з назвою "errors" або "logs"
      for (const guild of client.guilds.cache.values()) {
        const errorChannel = guild.channels.cache.find(
          (channel: any) =>
            channel.type === 0 && // Text channel
            (channel.name.includes('error') ||
              channel.name.includes('log') ||
              channel.name.includes('admin'))
        );

        if (errorChannel) {
          return errorChannel;
        }
      }

      return null;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('Помилка пошуку каналу сповіщень', {
          type: 'system',
          event: 'find_notification_channel_failed',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('Помилка пошуку каналу сповіщень', {
          type: 'system',
          event: 'find_notification_channel_failed',
          errorMessage: String(error),
        });
      }
      return null;
    }
  }

  /**
   * Отримання зрозумілого повідомлення про помилку
   */
  private getUserFriendlyMessage(error: Error, errorInfo: ErrorInfo): string {
    const baseMessage = error.message || 'Невідома помилка';

    switch (errorInfo.type) {
      case 'DiscordAPIError':
        return `Помилка Discord API: ${baseMessage}`;
      case 'GoogleAPIError':
        return `Помилка Google API: ${baseMessage}`;
      case 'AIAPIError':
        return `Помилка AI сервісу: ${baseMessage}`;
      case 'ValidationError':
        return `Помилка валідації: ${baseMessage}`;
      case 'CacheError':
        return `Помилка кешу: ${baseMessage}`;
      case 'FileError':
        return `Помилка файлу: ${baseMessage}`;
      case 'NetworkError':
        return `Мережева помилка: ${baseMessage}`;
      case 'DatabaseError':
        return `Помилка бази даних: ${baseMessage}`;
      default:
        return baseMessage;
    }
  }

  /**
   * Отримання статистики помилок
   */
  getStats(): ErrorStats {
    return {
      totalErrors: Array.from(this.errorCounts.values()).reduce((a, b) => a + b, 0),
      errorCounts: Object.fromEntries(this.errorCounts),
      notificationQueueSize: this.notificationQueue.length,
      isActive: this.active,
    };
  }

  /**
   * Очищення статистики помилок
   */
  clearErrorStats(): void {
    this.errorCounts.clear();
    this.notificationQueue.length = 0;
    logger.info('✅ Статистика помилок очищена', {
      type: 'system',
      event: 'error_stats_cleared',
    });
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
    logger.info('🛑 Завершення роботи Error Handler...', {
      type: 'system',
      event: 'error_handler_shutdown',
    });

    this.active = false;

    // Обробка залишкових сповіщень
    while (this.notificationQueue.length > 0) {
      const notification = this.notificationQueue.shift()!;
      await this.sendNotification(notification.error, notification.errorInfo, notification.context);
    }

    logger.info('✅ Error Handler завершено', {
      type: 'system',
      event: 'error_handler_shutdown_success',
    });
  }
}

export { ErrorHandler };
