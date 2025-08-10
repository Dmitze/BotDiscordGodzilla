/**
 * Утиліта для повторних спроб операцій
 * TypeScript версія
 */

import logger from './logger';

interface RetryOptions {
  maxAttempts?: number;
  delay?: number;
  backoff?: 'fixed' | 'exponential' | 'linear';
  factor?: number;
  maxDelay?: number;
  timeout?: number;
  onRetry?: (attempt: number, error: Error) => void;
  shouldRetry?: (error: Error) => boolean;
}

interface RetryResult<T> {
  success: boolean;
  data?: T;
  error?: Error;
  attempts: number;
  totalTime: number;
}

class RetryManager {
  private static defaultOptions: Required<RetryOptions> = {
    maxAttempts: 3,
    delay: 1000,
    backoff: 'exponential',
    factor: 2,
    maxDelay: 30000,
    timeout: 30000,
    onRetry: () => {},
    shouldRetry: () => true,
  };

  /**
   * Виконання операції з повторними спробами
   */
  static async execute<T>(
    operation: () => Promise<T>,
    options: RetryOptions = {}
  ): Promise<RetryResult<T>> {
    const config = { ...this.defaultOptions, ...options };
    const startTime = Date.now();
    let lastError: Error;

    for (let attempt = 1; attempt <= config.maxAttempts; attempt++) {
      try {
        // Створюємо timeout promise
        const timeoutPromise = new Promise<never>((_, reject) => {
          setTimeout(() => reject(new Error('Operation timeout')), config.timeout);
        });

        // Виконуємо операцію з timeout
        const result = await Promise.race([operation(), timeoutPromise]);

        return {
          success: true,
          data: result,
          attempts: attempt,
          totalTime: Date.now() - startTime,
        };
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        // Перевіряємо чи потрібно повторювати
        if (!config.shouldRetry(lastError)) {
          return {
            success: false,
            error: lastError,
            attempts: attempt,
            totalTime: Date.now() - startTime,
          };
        }

        // Остання спроба
        if (attempt === config.maxAttempts) {
          logger.error('❌ Операція невдала', {
            type: 'retry',
            event: 'final_failure',
            attempt,
            maxAttempts: config.maxAttempts,
            totalTimeMs: Date.now() - startTime,
            error: lastError.message,
            errorType: lastError.constructor.name,
            stack: lastError.stack,
            backoff: config.backoff,
          });
          return {
            success: false,
            error: lastError,
            attempts: attempt,
            totalTime: Date.now() - startTime,
          };
        }

        // Викликаємо callback
        config.onRetry(attempt, lastError);

        // Розраховуємо затримку
        const delay = this.calculateDelay(attempt, config);
        
        logger.warn('⚠️ Планування повторної спроби', {
          type: 'retry',
          event: 'retry_scheduled',
          attempt,
          nextDelayMs: delay,
          maxAttempts: config.maxAttempts,
          error: lastError.message,
          errorType: lastError.constructor.name,
          backoff: config.backoff,
        });

        // Чекаємо перед наступною спробою
        await this.sleep(delay);
      }
    }

    return {
      success: false,
      error: lastError!,
      attempts: config.maxAttempts,
      totalTime: Date.now() - startTime,
    };
  }

  /**
   * Розрахунок затримки між спробами
   */
  private static calculateDelay(attempt: number, config: Required<RetryOptions>): number {
    let delay: number;

    switch (config.backoff) {
      case 'fixed':
        delay = config.delay;
        break;
      case 'linear':
        delay = config.delay * attempt;
        break;
      case 'exponential':
        delay = config.delay * Math.pow(config.factor, attempt - 1);
        break;
      default:
        delay = config.delay;
    }

    return Math.min(delay, config.maxDelay);
  }

  /**
   * Затримка виконання
   */
  private static sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Створення функції з повторними спробами
   */
  static createRetryFunction<T extends (...args: any[]) => Promise<any>>(
    fn: T,
    options: RetryOptions = {}
  ): (...args: Parameters<T>) => Promise<RetryResult<Awaited<ReturnType<T>>>> {
    return async (...args: Parameters<T>) => {
      return this.execute(() => fn(...args), options);
    };
  }

  /**
   * Retry для HTTP запитів
   */
  static async httpRequest<T>(
    requestFn: () => Promise<T>,
    options: RetryOptions = {}
  ): Promise<RetryResult<T>> {
    const httpOptions: RetryOptions = {
      shouldRetry: (error: Error) => {
        // Повторюємо для 5xx помилок та мережевих помилок
        const anyErr = error as any;
        const status = typeof anyErr?.status === 'number' ? anyErr.status : undefined;
        const code = typeof anyErr?.code === 'string' ? anyErr.code : undefined;
        return (typeof status === 'number' && status >= 500) || code === 'ECONNRESET' || code === 'ETIMEDOUT';
      },
      ...options,
    };

    return this.execute(requestFn, httpOptions);
  }

  /**
   * Retry для операцій з базою даних
   */
  static async databaseOperation<T>(
    operation: () => Promise<T>,
    options: RetryOptions = {}
  ): Promise<RetryResult<T>> {
    const dbOptions: RetryOptions = {
      shouldRetry: (error: Error) => {
        // Повторюємо для тимчасових помилок БД
        const errorMessage = (error?.message ?? '').toLowerCase();
        return errorMessage.includes('connection') || 
               errorMessage.includes('timeout') ||
               errorMessage.includes('deadlock') ||
               errorMessage.includes('temporary');
      },
      maxAttempts: 5,
      delay: 2000,
      ...options,
    };

    return this.execute(operation, dbOptions);
  }

  /**
   * Retry для файлових операцій
   */
  static async fileOperation<T>(
    operation: () => Promise<T>,
    options: RetryOptions = {}
  ): Promise<RetryResult<T>> {
    const fileOptions: RetryOptions = {
      shouldRetry: (error: Error) => {
        // Повторюємо для тимчасових помилок файлової системи
        const errorCode = (error as any)?.code;
        return typeof errorCode === 'string' && (
               errorCode === 'EBUSY' || 
               errorCode === 'EACCES' || 
               errorCode === 'ENOENT' ||
               errorCode === 'EAGAIN');
      },
      maxAttempts: 3,
      delay: 1000,
      ...options,
    };

    return this.execute(operation, fileOptions);
  }

  /**
   * Retry для Discord API операцій
   */
  static async discordOperation<T>(
    operation: () => Promise<T>,
    options: RetryOptions = {}
  ): Promise<RetryResult<T>> {
    const discordOptions: RetryOptions = {
      shouldRetry: (error: Error) => {
        // Повторюємо для rate limits та тимчасових помилок Discord
        const statusCode = (error as any)?.status;
        return (typeof statusCode === 'number' && (statusCode === 429 || statusCode >= 500));
      },
      maxAttempts: 3,
      delay: 1000,
      backoff: 'exponential',
      ...options,
    };

    return this.execute(operation, discordOptions);
  }
}

export default RetryManager;
export { RetryManager };