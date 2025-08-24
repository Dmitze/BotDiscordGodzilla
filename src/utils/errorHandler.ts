/**
 * Розширений обробник помилок для Discord AI Assistant Bot
 * Централізована обробка та логування помилок
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import logger, { type LogMeta } from './logger';

// Константи для обробки помилок
const ERROR_HANDLER_CONSTANTS = {
  MAX_ERROR_DETAILS: 1000,
  MAX_STACK_TRACE_LINES: 20,
  ERROR_CATEGORIES: {
    VALIDATION: 'validation',
    NETWORK: 'network',
    DATABASE: 'database',
    AUTHENTICATION: 'authentication',
    AUTHORIZATION: 'authorization',
    RATE_LIMIT: 'rate_limit',
    TIMEOUT: 'timeout',
    RESOURCE: 'resource',
    SYSTEM: 'system',
    UNKNOWN: 'unknown',
  } as const,
  SEVERITY_LEVELS: {
    LOW: 'low',
    MEDIUM: 'medium',
    HIGH: 'high',
    CRITICAL: 'critical',
  } as const,
} as const;

export interface ErrorDetails {
  name: string;
  message: string;
  stack?: string;
  code?: string;
  cause?: Error;
  timestamp: Date;
  category: string;
  severity: string;
  context?: Record<string, unknown>;
  userId?: string;
  guildId?: string;
  channelId?: string;
  commandName?: string;
  serviceName?: string;
  requestId?: string;
  correlationId?: string;
}

export interface ErrorHandlerStats {
  totalErrors: number;
  errorsByCategory: Record<string, number>;
  errorsBySeverity: Record<string, number>;
  errorsByService: Record<string, number>;
  recentErrors: ErrorDetails[];
  lastError?: ErrorDetails;
  averageErrorRate: number;
  criticalErrors: number;
}

export class ErrorHandler {
  private static instance: ErrorHandler | null = null;
  private errorStats!: ErrorHandlerStats;
  private errorHistory: ErrorDetails[] = [];
  private readonly maxErrorHistory = 1000;
  private _isInitialized = false;

  constructor() {
    if (ErrorHandler.instance) {
      return ErrorHandler.instance;
    }
    ErrorHandler.instance = this;

    this.errorStats = {
      totalErrors: 0,
      errorsByCategory: {},
      errorsBySeverity: {},
      errorsByService: {},
      recentErrors: [],
      averageErrorRate: 0,
      criticalErrors: 0,
    };

    this.initialize();
  }

  /**
   * Ініціалізація обробника помилок
   */
  private initialize(): void {
    try {
      logger.info('🔧 Ініціалізація ErrorHandler...');

      // Налаштування глобальних обробників помилок
      this.setupGlobalErrorHandlers();

      this._isInitialized = true;
      logger.info('✅ ErrorHandler успішно ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації ErrorHandler:', {
        type: 'system',
        event: 'error_handler_init_failed',
        error,
      });
      this.createFallbackErrorHandler();
    }
  }

  /**
   * Налаштування глобальних обробників помилок
   */
  private setupGlobalErrorHandlers(): void {
    // Обробка необроблених помилок
    process.on('uncaughtException', error => {
      this.handleUncaughtException(error);
    });

    // Обробка необроблених rejections
    process.on('unhandledRejection', (reason, promise) => {
      this.handleUnhandledRejection(reason, promise);
    });

    // Обробка попереджень
    process.on('warning', warning => {
      this.handleWarning(warning);
    });

    logger.info('🛡️ Глобальні обробники помилок налаштовано');
  }

  /**
   * Обробка необробленої помилки
   */
  private handleUncaughtException(error: Error): void {
    const errorDetails: ErrorDetails = {
      name: error.name,
      message: error.message,
      ...(error.stack ? { stack: error.stack } : {}),
      timestamp: new Date(),
      category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
      severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL,
      context: {
        type: 'uncaught_exception',
        processId: process.pid,
        uptime: process.uptime(),
        memory: process.memoryUsage(),
      },
    };

    this.logError(errorDetails);

    // Логування критичної помилки
    logger.error('💥 Критична необроблена помилка', {
      type: 'system',
      event: 'uncaught_exception',
      severity: 'critical',
      errorName: error.name,
      errorMessage: error.message,
      stack: this.truncateStackTrace(error.stack),
      processId: process.pid,
    });

    // Зупинка процесу при критичній помилці
    logger.error('🛑 Зупинка процесу через критичну помилку');
    process.exit(1);
  }

  /**
   * Обробка необробленого rejection
   */
  private handleUnhandledRejection(reason: unknown, promise: Promise<unknown>): void {
    const errorDetails: ErrorDetails = {
      name: 'UnhandledRejection',
      message: reason instanceof Error ? reason.message : String(reason),
      ...(reason instanceof Error && reason.stack ? { stack: reason.stack } : {}),
      timestamp: new Date(),
      category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
      severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH,
      context: {
        type: 'unhandled_rejection',
        promise: promise.toString(),
        processId: process.pid,
        uptime: process.uptime(),
      },
    };

    this.logError(errorDetails);

    logger.error('💥 Необроблений rejection', {
      type: 'system',
      event: 'unhandled_rejection',
      severity: 'high',
      reason: reason instanceof Error ? reason.message : String(reason),
      errorName: reason instanceof Error ? reason.name : undefined,
      errorType: reason instanceof Error ? reason.constructor.name : typeof reason,
      promise: promise.toString(),
      processId: process.pid,
    });
  }

  /**
   * Обробка попередження
   */
  private handleWarning(warning: Error): void {
    const errorDetails: ErrorDetails = {
      name: warning.name,
      message: warning.message,
      ...(warning.stack ? { stack: warning.stack } : {}),
      timestamp: new Date(),
      category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.SYSTEM,
      severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW,
      context: {
        type: 'warning',
        processId: process.pid,
        uptime: process.uptime(),
      },
    };

    this.logError(errorDetails);

    logger.warn('⚠️ Попередження системи', {
      type: 'system',
      event: 'warning',
      severity: 'low',
      errorName: warning.name,
      errorMessage: warning.message,
      processId: process.pid,
    });
  }

  /**
   * Основний метод обробки помилок
   */
  public handleError(
    error: unknown,
    context: {
      userId?: string;
      guildId?: string;
      channelId?: string;
      commandName?: string;
      serviceName?: string;
      requestId?: string;
      correlationId?: string;
      additionalContext?: Record<string, unknown>;
    } = {}
  ): ErrorDetails {
    try {
      const errorDetails = this.createErrorDetails(error, context);
      this.logError(errorDetails);
      this.updateStats(errorDetails);

      return errorDetails;
    } catch (handlerError) {
      logger.error('❌ Помилка в ErrorHandler:', {
        type: 'system',
        event: 'error_handler_runtime_error',
        error: handlerError,
      });
      return this.createFallbackErrorDetails(error);
    }
  }

  /**
   * Створення деталей помилки
   */
  private createErrorDetails(
    error: unknown,
    context: {
      userId?: string;
      guildId?: string;
      channelId?: string;
      commandName?: string;
      serviceName?: string;
      requestId?: string;
      correlationId?: string;
      additionalContext?: Record<string, unknown>;
    }
  ): ErrorDetails {
    const errorObj = error instanceof Error ? error : new Error(String(error));

    return {
      name: errorObj.name,
      message: errorObj.message,
      ...(errorObj.stack ? { stack: errorObj.stack } : {}),
      code: (error as any)?.code,
      ...('cause' in (errorObj as any) && (errorObj as any).cause !== undefined
        ? { cause: (errorObj as any).cause as Error }
        : {}),
      timestamp: new Date(),
      category: this.categorizeError(errorObj),
      severity: this.determineSeverity(errorObj),
      context: {
        ...context.additionalContext,
        errorType: errorObj.constructor.name,
        hasStack: !!errorObj.stack,
      },
      ...(context.userId ? { userId: context.userId } : {}),
      ...(context.guildId ? { guildId: context.guildId } : {}),
      ...(context.channelId ? { channelId: context.channelId } : {}),
      ...(context.commandName ? { commandName: context.commandName } : {}),
      ...(context.serviceName ? { serviceName: context.serviceName } : {}),
      ...(context.requestId ? { requestId: context.requestId } : {}),
      ...(context.correlationId ? { correlationId: context.correlationId } : {}),
    };
  }

  /**
   * Категоризація помилки
   */
  private categorizeError(error: Error): string {
    const message = error.message.toLowerCase();
    const name = error.name.toLowerCase();

    if (message.includes('validation') || name.includes('validation')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.VALIDATION;
    }
    if (
      message.includes('network') ||
      message.includes('connection') ||
      message.includes('timeout')
    ) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.NETWORK;
    }
    if (message.includes('database') || message.includes('sql') || message.includes('query')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.DATABASE;
    }
    if (message.includes('auth') || message.includes('token') || message.includes('permission')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.AUTHENTICATION;
    }
    if (message.includes('rate limit') || message.includes('too many requests')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.RATE_LIMIT;
    }
    if (message.includes('timeout') || message.includes('timed out')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.TIMEOUT;
    }
    if (message.includes('resource') || message.includes('memory') || message.includes('disk')) {
      return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.RESOURCE;
    }

    return ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.UNKNOWN;
  }

  /**
   * Визначення серйозності помилки
   */
  private determineSeverity(error: Error): string {
    const message = error.message.toLowerCase();
    const name = error.name.toLowerCase();

    if (name.includes('critical') || message.includes('critical')) {
      return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL;
    }
    if (name.includes('fatal') || message.includes('fatal')) {
      return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL;
    }
    if (message.includes('timeout') || message.includes('connection failed')) {
      return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH;
    }
    if (message.includes('validation') || message.includes('invalid')) {
      return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM;
    }
    if (message.includes('warning') || name.includes('warning')) {
      return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW;
    }

    return ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM;
  }

  /**
   * Логування помилки
   */
  private logError(errorDetails: ErrorDetails): void {
    try {
      const logMeta: LogMeta = {
        errorName: errorDetails.name,
        errorMessage: errorDetails.message,
        errorCategory: errorDetails.category,
        errorSeverity: errorDetails.severity,
        errorCode: errorDetails.code,
        ...(errorDetails.userId ? { userId: errorDetails.userId } : {}),
        ...(errorDetails.guildId ? { guildId: errorDetails.guildId } : {}),
        ...(errorDetails.channelId ? { channelId: errorDetails.channelId } : {}),
        ...(errorDetails.commandName ? { commandName: errorDetails.commandName } : {}),
        ...(errorDetails.serviceName ? { serviceName: errorDetails.serviceName } : {}),
        ...(errorDetails.requestId ? { requestId: errorDetails.requestId } : {}),
        ...(errorDetails.correlationId ? { correlationId: errorDetails.correlationId } : {}),
        timestamp: errorDetails.timestamp.toISOString(),
        type: 'system',
        event: 'error',
        severity: errorDetails.severity,
      };

      // Логування в залежності від серйозності
      switch (errorDetails.severity) {
        case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL:
          logger.error('💥 Критична помилка', logMeta);
          break;
        case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.HIGH:
          logger.error('❌ Серйозна помилка', logMeta);
          break;
        case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM:
          logger.warn('⚠️ Помилка', logMeta);
          break;
        case ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW:
          logger.debug('ℹ️ Попередження', logMeta);
          break;
        default:
          logger.error('❌ Помилка', logMeta);
      }

      // Логування stack trace для серйозних помилок
      if (
        errorDetails.severity !== ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.LOW &&
        errorDetails.stack
      ) {
        logger.debug('📋 Stack trace', {
          type: 'system',
          event: 'stack_trace',
          errorName: errorDetails.name,
          stack: this.truncateStackTrace(errorDetails.stack),
        });
      }
    } catch (logError) {
      console.error('❌ Помилка логування помилки:', logError);
    }
  }

  /**
   * Оновлення статистики помилок
   */
  private updateStats(errorDetails: ErrorDetails): void {
    try {
      this.errorStats.totalErrors++;
      this.errorStats.lastError = errorDetails;

      // Оновлення статистики по категоріях
      this.errorStats.errorsByCategory[errorDetails.category] =
        (this.errorStats.errorsByCategory[errorDetails.category] || 0) + 1;

      // Оновлення статистики по серйозності
      this.errorStats.errorsBySeverity[errorDetails.severity] =
        (this.errorStats.errorsBySeverity[errorDetails.severity] || 0) + 1;

      // Оновлення статистики по сервісах
      if (errorDetails.serviceName) {
        this.errorStats.errorsByService[errorDetails.serviceName] =
          (this.errorStats.errorsByService[errorDetails.serviceName] || 0) + 1;
      }

      // Оновлення критичних помилок
      if (errorDetails.severity === ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.CRITICAL) {
        this.errorStats.criticalErrors++;
      }

      // Додавання до історії
      this.errorHistory.push(errorDetails);
      if (this.errorHistory.length > this.maxErrorHistory) {
        this.errorHistory.shift();
      }

      // Оновлення середньої частоти помилок
      const uptime = process.uptime();
      this.errorStats.averageErrorRate = uptime > 0 ? this.errorStats.totalErrors / uptime : 0;
    } catch (statsError) {
      console.error('❌ Помилка оновлення статистики помилок:', statsError);
    }
  }

  /**
   * Обрізання stack trace
   */
  private truncateStackTrace(stack?: string): string {
    if (!stack) return '';

    const lines = stack.split('\n');
    const truncatedLines = lines.slice(0, ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES);

    if (lines.length > ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES) {
      truncatedLines.push(
        `... (${lines.length - ERROR_HANDLER_CONSTANTS.MAX_STACK_TRACE_LINES} more lines)`
      );
    }

    return truncatedLines.join('\n');
  }

  /**
   * Створення fallback обробника помилок
   */
  private createFallbackErrorHandler(): void {
    console.error('🔧 Створення fallback обробника помилок...');

    process.on('uncaughtException', error => {
      console.error('💥 Критична помилка (fallback):', error);
      process.exit(1);
    });

    process.on('unhandledRejection', reason => {
      console.error('💥 Необроблений rejection (fallback):', reason);
    });
  }

  /**
   * Створення fallback деталей помилки
   */
  private createFallbackErrorDetails(error: unknown): ErrorDetails {
    return {
      name: 'UnknownError',
      message: error instanceof Error ? error.message : String(error),
      timestamp: new Date(),
      category: ERROR_HANDLER_CONSTANTS.ERROR_CATEGORIES.UNKNOWN,
      severity: ERROR_HANDLER_CONSTANTS.SEVERITY_LEVELS.MEDIUM,
    };
  }

  /**
   * Отримання статистики помилок
   */
  public getStats(): ErrorHandlerStats {
    return { ...this.errorStats };
  }

  /**
   * Отримання історії помилок
   */
  public getErrorHistory(): ErrorDetails[] {
    return [...this.errorHistory];
  }

  /**
   * Очищення історії помилок
   */
  public clearErrorHistory(): void {
    this.errorHistory = [];
    logger.info('🧹 Історія помилок очищено');
  }

  /**
   * Перевірка стану ініціалізації
   */
  public isInitialized(): boolean {
    return this._isInitialized;
  }
}

// Експорт єдиного екземпляра
export const errorHandler = new ErrorHandler();

// Експорт функцій для зручності
export const handleError = (
  error: unknown,
  context?: {
    userId?: string;
    guildId?: string;
    channelId?: string;
    commandName?: string;
    serviceName?: string;
    requestId?: string;
    correlationId?: string;
    additionalContext?: Record<string, unknown>;
  }
) => errorHandler.handleError(error, context);

export const getErrorStats = () => errorHandler.getStats();
export const getErrorHistory = () => errorHandler.getErrorHistory();
export const clearErrorHistory = () => errorHandler.clearErrorHistory();
