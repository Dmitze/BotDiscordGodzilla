/**
 * Розширений логер для Discord AI Assistant Bot
 * Рефакторована версія з покращеними можливостями
 * TypeScript версія 3.0.0 - Повністю рефакторовано
 */

import winston from 'winston';
import path from 'path';
import fs from 'fs';
import { performance } from 'perf_hooks';

// Константи для конфігурації логера
const LOGGER_CONFIG = {
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  MAX_FILES: 5,
  COMMAND_LOG_SIZE: 5 * 1024 * 1024, // 5MB
  COMMAND_LOG_FILES: 3,
  CLEANUP_INTERVAL: 24 * 60 * 60 * 1000, // 24 години
  MAX_LOG_AGE: 30 * 24 * 60 * 60 * 1000, // 30 днів
  BUFFER_SIZE: 1000,
  FLUSH_INTERVAL: 5000, // 5 секунд
} as const;

interface LogMeta {
  [key: string]: any;
  timestamp?: string;
  level?: string;
  service?: string;
  userId?: string;
  guildId?: string;
  channelId?: string;
  requestId?: string;
  correlationId?: string;
  // Явно оголошені поля, що часто використовуються
  type?: string;
  severity?: string;
  category?: string;
  component?: string;
  logLevel?: string;
  processId?: number;
  memory?: NodeJS.MemoryUsage;
}

interface LoggerStats {
  totalLogs: number;
  errors: number;
  commands: number;
  apiRequests: number;
  performance: number;
  security: number;
  system: number;
  debug: number;
  warnings: number;
  lastLogTime: Date;
  averageLogSize: number;
  logBufferSize: number;
}

interface LogEntry {
  timestamp: Date;
  level: string;
  message: string;
  meta: LogMeta;
  size: number;
}

class Logger {
  private logger: winston.Logger | null = null;
  private stats: LoggerStats;
  private logBuffer: LogEntry[] = [];
  private cleanupInterval: NodeJS.Timeout | null = null;
  private flushInterval: NodeJS.Timeout | null = null;
  private isInitialized = false;
  private readonly logsDir: string;

  constructor() {
    this.logsDir = path.join(process.cwd(), 'data', 'logs');
    this.stats = {
      totalLogs: 0,
      errors: 0,
      commands: 0,
      apiRequests: 0,
      performance: 0,
      security: 0,
      system: 0,
      debug: 0,
      warnings: 0,
      lastLogTime: new Date(),
      averageLogSize: 0,
      logBufferSize: 0,
    };

    this.initialize();
  }

  /**
   * Санітізація метаданих логів: маскує секрети, обрізає великі значення, прибирає цикли
   */
  private sanitizeMeta(meta: LogMeta): LogMeta {
    const SECRET_KEYS = new Set([
      'token',
      'apiKey',
      'apikey',
      'api_key',
      'password',
      'pass',
      'secret',
      'clientSecret',
      'authorization',
      'auth',
      'bearer',
      'session',
      'cookie',
      'cookies',
    ]);

    const MAX_STRING_LEN = 2000; // захист від гігантських полів

    const seen = new WeakSet();

    const redact = (key: string, value: unknown): unknown => {
      if (value == null) return value;
      if (SECRET_KEYS.has(key.toLowerCase())) return '[REDACTED]';
      if (typeof value === 'string') {
        return value.length > MAX_STRING_LEN ? value.slice(0, MAX_STRING_LEN) + '…' : value;
      }
      if (typeof value === 'object') {
        if (seen.has(value as object)) return '[CIRCULAR]';
        seen.add(value as object);
        if (Array.isArray(value)) return value.map(v => redact(key, v));
        const out: Record<string, unknown> = {};
        for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
          out[k] = redact(k, v);
        }
        return out;
      }
      return value;
    };

    // Глибоке копіювання з санітізацією
    const safe: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(meta || {})) {
      safe[k] = redact(k, v);
    }
    return safe as LogMeta;
  }

  /**
   * Ініціалізація логера з детальним логуванням
   */
  private initialize(): void {
    try {
      console.log('🔧 Ініціалізація логера...');

      // Створення папки для логів
      this.ensureLogsDirectory();

      // Конфігурація форматів
      const formats = this.createFormats();

      // Створення транспортів
      const transports = this.createTransports();

      // Створення логера
      this.logger = winston.createLogger({
        level: this.getLogLevel(),
        format: formats.file,
        transports: transports,
        exitOnError: false,
        silent: false,
      });

      // Налаштування обробки необроблених помилок
      this.setupExceptionHandling();

      // Запуск періодичних завдань (пропускаємо у тестах)
      if (process.env['NODE_ENV'] !== 'test' && !process.env['JEST_WORKER_ID']) {
        this.startPeriodicTasks();
      } else {
        console.log('⏭️ Пропуск періодичних завдань логера у тестовому середовищі');
      }

      this.isInitialized = true;
      console.log('✅ Логер успішно ініціалізовано');
    } catch (error) {
      console.error('❌ Помилка ініціалізації логера:', error);
      this.createFallbackLogger();
    }
  }

  /**
   * Створення папки для логів
   */
  private ensureLogsDirectory(): void {
    try {
      if (!fs.existsSync(this.logsDir)) {
        fs.mkdirSync(this.logsDir, { recursive: true });
        console.log(`📁 Створено папку для логів: ${this.logsDir}`);
      }
    } catch (error) {
      console.error('❌ Помилка створення папки логів:', error);
      throw new Error(
        `Неможливо створити папку логів: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
    }
  }

  /**
   * Створення форматів логування
   */
  private createFormats() {
    return {
      console: winston.format.combine(
        winston.format.colorize(),
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss.SSS' }),
        winston.format.errors({ stack: true }),
        winston.format.printf(({ timestamp, level, message, service, userId, ...meta }) => {
          let log = `${timestamp} [${level}]`;
          if (service) log += ` [${service}]`;
          if (userId) log += ` [User:${userId}]`;
          log += `: ${message}`;

          const remainingMeta = Object.keys(meta).filter(
            key => !['timestamp', 'level', 'service', 'userId'].includes(key)
          );

          if (remainingMeta.length > 0) {
            log += ` ${JSON.stringify(meta)}`;
          }

          return log;
        })
      ),
      file: winston.format.combine(
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss.SSS' }),
        winston.format.errors({ stack: true }),
        winston.format.json()
      ),
    };
  }

  /**
   * Створення транспортів
   */
  private createTransports(): winston.transport[] {
    const formats = this.createFormats();
    // У тестовому середовищі використовуємо лише консольний транспорт,
    // щоб уникнути помилки winston "write after end" при завершенні Jest
    if (process.env['NODE_ENV'] === 'test' || process.env['JEST_WORKER_ID']) {
      return [
        new winston.transports.Console({
          format: formats.console,
          level: this.getLogLevel(),
          handleExceptions: true,
          handleRejections: true,
        }),
      ];
    }

    return [
      // Консольний транспорт
      new winston.transports.Console({
        format: formats.console,
        level: this.getLogLevel(),
        handleExceptions: true,
        handleRejections: true,
      }),

      // Файл для всіх логів
      new winston.transports.File({
        filename: path.join(this.logsDir, 'bot.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
        level: 'info',
        tailable: true,
        handleExceptions: true,
        handleRejections: true,
      }),

      // Файл для помилок
      new winston.transports.File({
        filename: path.join(this.logsDir, 'error.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
        level: 'error',
        tailable: true,
      }),

      // Файл для команд
      new winston.transports.File({
        filename: path.join(this.logsDir, 'commands.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.COMMAND_LOG_SIZE,
        maxFiles: LOGGER_CONFIG.COMMAND_LOG_FILES,
        level: 'info',
        tailable: true,
      }),

      // Файл для безпеки
      new winston.transports.File({
        filename: path.join(this.logsDir, 'security.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
        level: 'warn',
        tailable: true,
      }),

      // Файл для продуктивності
      new winston.transports.File({
        filename: path.join(this.logsDir, 'performance.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
        level: 'info',
        tailable: true,
      }),
    ];
  }

  /**
   * Отримання рівня логування
   */
  private getLogLevel(): string {
    // У тестах за замовчуванням знижуємо рівень логів
    if (process.env['NODE_ENV'] === 'test' || process.env['JEST_WORKER_ID']) {
      return (process.env['LOG_LEVEL']?.toLowerCase()) || 'error';
    }
    const level = process.env['LOG_LEVEL']?.toLowerCase();
    const validLevels = ['error', 'warn', 'info', 'debug'];

    if (level && validLevels.includes(level)) {
      return level;
    }

    return process.env['NODE_ENV'] === 'production' ? 'info' : 'debug';
  }

  /**
   * Налаштування обробки необроблених помилок
   */
  private setupExceptionHandling(): void {
    if (!this.logger) return;

    const formats = this.createFormats();

    this.logger.exceptions.handle(
      new winston.transports.File({
        filename: path.join(this.logsDir, 'exceptions.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
      })
    );

    this.logger.rejections.handle(
      new winston.transports.File({
        filename: path.join(this.logsDir, 'rejections.log'),
        format: formats.file,
        maxsize: LOGGER_CONFIG.MAX_FILE_SIZE,
        maxFiles: LOGGER_CONFIG.MAX_FILES,
      })
    );
  }

  /**
   * Запуск періодичних завдань
   */
  private startPeriodicTasks(): void {
    if (process.env['NODE_ENV'] === 'test' || process.env['JEST_WORKER_ID']) {
      return;
    }
    // Очищення старих логів
    this.cleanupInterval = setInterval(() => {
      this.cleanupOldLogs();
    }, LOGGER_CONFIG.CLEANUP_INTERVAL);

    // Скидання буфера логів
    this.flushInterval = setInterval(() => {
      this.flushLogBuffer();
    }, LOGGER_CONFIG.FLUSH_INTERVAL);
  }

  /**
   * Створення резервного логера
   */
  private createFallbackLogger(): void {
    console.warn('⚠️ Використання резервного логера');

    this.logger = winston.createLogger({
      level: 'info',
      format: winston.format.simple(),
      transports: [new winston.transports.Console()],
    });
  }

  /**
   * Логування з детальною інформацією
   */
  private log(level: string, message: string, meta: LogMeta = {}): void {
    if (!this.isInitialized || !this.logger) {
      console.log(`[${level.toUpperCase()}]: ${message}`, meta);
      return;
    }
    try {
      const startTime = performance.now();

      // Додавання додаткової інформації
      const enhancedMeta: LogMeta = {
        ...meta,
        timestamp: new Date().toISOString(),
        service: meta.service || 'logger',
        logLevel: level,
        processId: process.pid,
        memory: process.memoryUsage(),
      };

      // Санитизация секретів і циклів
      const safeMeta = this.sanitizeMeta(enhancedMeta);

      // Оновлення статистики
      this.updateStats(level, message, safeMeta);

      // Додавання до буфера
      this.addToBuffer(level, message, safeMeta);

      // Логування через winston
      this.logger.log(level, message, safeMeta);

      const duration = performance.now() - startTime;
      if (duration > 100) {
        console.warn(`⚠️ Повільне логування: ${duration.toFixed(2)}ms`);
      }
    } catch (error) {
      console.error('❌ Помилка логування:', error);
      console.log(`[${level.toUpperCase()}]: ${message}`, meta);
    }
  }

  /**
   * Оновлення статистики
   */
  private updateStats(level: string, message: string, meta: LogMeta): void {
    this.stats.totalLogs++;
    this.stats.lastLogTime = new Date();
    this.stats.logBufferSize = this.logBuffer.length;

    const logSize = JSON.stringify({ level, message, meta }).length;
    this.stats.averageLogSize = (this.stats.averageLogSize + logSize) / 2;

    switch (level) {
      case 'error':
        this.stats.errors++;
        break;
      case 'warn':
        this.stats.warnings++;
        break;
      case 'debug':
        this.stats.debug++;
        break;
    }

    if ((meta as any)['type'] === 'command') this.stats.commands++;
    if ((meta as any)['type'] === 'api_request') this.stats.apiRequests++;
    if ((meta as any)['type'] === 'performance') this.stats.performance++;
    if ((meta as any)['type'] === 'security') this.stats.security++;
    if ((meta as any)['type'] === 'system') this.stats.system++;
  }

  /**
   * Додавання до буфера логів
   */
  private addToBuffer(level: string, message: string, meta: LogMeta): void {
    const entry: LogEntry = {
      timestamp: new Date(),
      level,
      message,
      meta,
      size: JSON.stringify({ level, message, meta }).length,
    };

    this.logBuffer.push(entry);

    // Обмеження розміру буфера
    if (this.logBuffer.length > LOGGER_CONFIG.BUFFER_SIZE) {
      this.logBuffer.shift();
    }
  }

  /**
   * Скидання буфера логів
   */
  private flushLogBuffer(verbose: boolean = true): void {
    if (this.logBuffer.length === 0) return;

    try {
      const bufferSize = this.logBuffer.length;
      const totalSize = this.logBuffer.reduce((sum, entry) => sum + entry.size, 0);

      if (verbose) {
        this.debug(`Скидання буфера логів: ${bufferSize} записів, ${totalSize} байт`);
      }

      this.logBuffer = [];
      this.stats.logBufferSize = 0;
    } catch (error) {
      console.error('❌ Помилка скидання буфера логів:', error);
    }
  }

  /**
   * Очищення старих логів
   */
  private cleanupOldLogs(): void {
    try {
      const files = fs.readdirSync(this.logsDir);
      const now = Date.now();
      let cleanedCount = 0;

      for (const file of files) {
        const filePath = path.join(this.logsDir, file);
        const stats = fs.statSync(filePath);

        if (now - stats.mtime.getTime() > LOGGER_CONFIG.MAX_LOG_AGE) {
          fs.unlinkSync(filePath);
          cleanedCount++;
        }
      }

      if (cleanedCount > 0) {
        this.info(`Очищено ${cleanedCount} старих лог-файлів`);
      }
    } catch (error) {
      console.error('❌ Помилка очищення старих логів:', error);
    }
  }

  /**
   * Логування інформації
   */
  public info(message: string, meta: LogMeta = {}): void {
    this.log('info', message, meta);
  }

  /**
   * Логування помилок
   */
  public error(message: string, meta: LogMeta = {}): void {
    this.log('error', message, meta);
  }

  /**
   * Логування попереджень
   */
  public warn(message: string, meta: LogMeta = {}): void {
    this.log('warn', message, meta);
  }

  /**
   * Логування дебагу
   */
  public debug(message: string, meta: LogMeta = {}): void {
    this.log('debug', message, meta);
  }

  /**
   * Логування команд з детальною інформацією
   */
  public command(
    command: string,
    user: string,
    duration: number,
    success: boolean = true,
    meta: LogMeta = {}
  ): void {
    this.log('info', `Команда виконана: ${command}`, {
      ...meta,
      command,
      user,
      duration: `${duration}ms`,
      success,
      type: 'command',
      performance: duration > 1000 ? 'slow' : duration > 500 ? 'medium' : 'fast',
    });
  }

  /**
   * Логування помилок команд
   */
  public commandError(
    command: string,
    user: string,
    error: Error,
    duration: number,
    meta: LogMeta = {}
  ): void {
    this.log('error', `Помилка команди: ${command}`, {
      ...meta,
      command,
      user,
      error: error.message,
      stack: error.stack,
      duration: `${duration}ms`,
      type: 'command_error',
      errorType: error.constructor.name,
    });
  }

  /**
   * Логування API запитів
   */
  public apiRequest(
    service: string,
    endpoint: string,
    duration: number,
    success: boolean = true,
    meta: LogMeta = {}
  ): void {
    this.log('info', `API запит: ${service} - ${endpoint}`, {
      ...meta,
      service,
      endpoint,
      duration: `${duration}ms`,
      success,
      type: 'api_request',
      performance: duration > 5000 ? 'slow' : duration > 1000 ? 'medium' : 'fast',
    });
  }

  /**
   * Логування помилок API
   */
  public apiError(
    service: string,
    endpoint: string,
    error: Error,
    duration: number,
    meta: LogMeta = {}
  ): void {
    this.log('error', `Помилка API: ${service} - ${endpoint}`, {
      ...meta,
      service,
      endpoint,
      error: error.message,
      stack: error.stack,
      duration: `${duration}ms`,
      type: 'api_error',
      errorType: error.constructor.name,
    });
  }

  /**
   * Логування подій безпеки
   */
  public security(event: string, user: string, details: LogMeta = {}): void {
    this.log('warn', `Подія безпеки: ${event}`, {
      ...details,
      event,
      user,
      type: 'security',
      severity: (details as any)['severity'] || 'medium',
    });
  }

  /**
   * Логування продуктивності
   */
  public performance(operation: string, duration: number, details: LogMeta = {}): void {
    this.log('info', `Метрика продуктивності: ${operation}`, {
      ...details,
      operation,
      duration: `${duration}ms`,
      type: 'performance',
      category: (details as any)['category'] || 'general',
    });
  }

  /**
   * Логування системних подій
   */
  public system(event: string, details: LogMeta = {}): void {
    this.log('info', `Системна подія: ${event}`, {
      ...details,
      event,
      type: 'system',
      component: (details as any)['component'] || 'unknown',
    });
  }

  /**
   * Отримання детальної статистики логера
   */
  public getStats(): LoggerStats {
    return {
      ...this.stats,
      logBufferSize: this.logBuffer.length,
    };
  }

  /**
   * Отримання буфера логів
   */
  public getLogBuffer(): LogEntry[] {
    return [...this.logBuffer];
  }

  /**
   * Очищення ресурсів
   */
  public async cleanup(): Promise<void> {
    try {
      // Переводимо логер у пасивний режим, щоб уникнути записів у закриті потоки
      this.isInitialized = false;
      console.log('🧹 Очищення ресурсів логера...');

      // Зупинка періодичних завдань
      if (this.cleanupInterval) {
        clearInterval(this.cleanupInterval);
        this.cleanupInterval = null;
      }

      if (this.flushInterval) {
        clearInterval(this.flushInterval);
        this.flushInterval = null;
      }

      // Скидання буфера без додаткового логування
      this.flushLogBuffer(false);

      // Закриття транспортів та звільнення інстансу
      if (this.logger) {
        try {
          this.logger.close();
        } catch (e) {
          console.warn('⚠️ Помилка при закритті логера:', e);
        }
        this.logger = null;
      }

      console.log('✅ Ресурси логера очищено');
    } catch (error) {
      console.error('❌ Помилка очищення ресурсів логера:', error);
    }
  }

  /**
   * Перевірка стану логера
   */
  public isHealthy(): boolean {
    return this.isInitialized && this.logger !== null;
  }
}

// Експорт єдиного екземпляра
const logger = new Logger();

export default logger;
export { Logger, type LogMeta, type LoggerStats, type LogEntry };
