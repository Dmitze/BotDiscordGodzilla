/**
 * Розширений логер для Discord AI Assistant Bot
 * Рефакторована версія з покращеними можливостями
 */

const winston = require('winston');
const path = require('path');
const fs = require('fs');

class Logger {
  constructor() {
    this.logger = null;
    this.initialize();
  }

  /**
   * Ініціалізація логера
   */
  initialize() {
    try {
      // Створення папки для логів
      const logsDir = path.join(process.cwd(), 'data', 'logs');
      if (!fs.existsSync(logsDir)) {
        fs.mkdirSync(logsDir, { recursive: true });
      }

      // Конфігурація форматів
      const formats = {
        console: winston.format.combine(
          winston.format.colorize(),
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          winston.format.printf(({ timestamp, level, message, ...meta }) => {
            let log = `${timestamp} [${level}]: ${message}`;
            if (Object.keys(meta).length > 0) {
              log += ` ${JSON.stringify(meta)}`;
            }
            return log;
          })
        ),
        file: winston.format.combine(
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          winston.format.errors({ stack: true }),
          winston.format.json()
        ),
      };

      // Створення транспортів
      const transports = [
        // Консольний транспорт
        new winston.transports.Console({
          format: formats.console,
          level: process.env.LOG_LEVEL || 'info',
        }),

        // Файл для всіх логів
        new winston.transports.File({
          filename: path.join(logsDir, 'bot.log'),
          format: formats.file,
          maxsize: 10 * 1024 * 1024, // 10MB
          maxFiles: 5,
          level: 'info',
        }),

        // Файл для помилок
        new winston.transports.File({
          filename: path.join(logsDir, 'error.log'),
          format: formats.file,
          maxsize: 10 * 1024 * 1024, // 10MB
          maxFiles: 5,
          level: 'error',
        }),

        // Файл для команд
        new winston.transports.File({
          filename: path.join(logsDir, 'commands.log'),
          format: formats.file,
          maxsize: 5 * 1024 * 1024, // 5MB
          maxFiles: 3,
          level: 'info',
        }),
      ];

      // Створення логера
      this.logger = winston.createLogger({
        level: process.env.LOG_LEVEL || 'info',
        format: formats.file,
        transports: transports,
        exitOnError: false,
      });

      // Обробка необроблених помилок
      this.logger.exceptions.handle(
        new winston.transports.File({
          filename: path.join(logsDir, 'exceptions.log'),
          format: formats.file,
        })
      );

      this.logger.rejections.handle(
        new winston.transports.File({
          filename: path.join(logsDir, 'rejections.log'),
          format: formats.file,
        })
      );
    } catch (error) {
      console.error('Помилка ініціалізації логера:', error);
      // Fallback до простого логера
      this.createFallbackLogger();
    }
  }

  /**
   * Створення fallback логера
   */
  createFallbackLogger() {
    this.logger = {
      info: (message, meta) => console.log(`[INFO] ${message}`, meta || ''),
      error: (message, meta) => console.error(`[ERROR] ${message}`, meta || ''),
      warn: (message, meta) => console.warn(`[WARN] ${message}`, meta || ''),
      debug: (message, meta) => console.debug(`[DEBUG] ${message}`, meta || ''),
      verbose: (message, meta) => console.log(`[VERBOSE] ${message}`, meta || ''),
    };
  }

  /**
   * Логування інформації
   */
  info(message, meta = {}) {
    if (this.logger) {
      this.logger.info(message, meta);
    }
  }

  /**
   * Логування помилок
   */
  error(message, meta = {}) {
    if (this.logger) {
      this.logger.error(message, meta);
    }
  }

  /**
   * Логування попереджень
   */
  warn(message, meta = {}) {
    if (this.logger) {
      this.logger.warn(message, meta);
    }
  }

  /**
   * Логування для дебагу
   */
  debug(message, meta = {}) {
    if (this.logger) {
      this.logger.debug(message, meta);
    }
  }

  /**
   * Логування команд
   */
  command(command, user, duration, success = true) {
    this.info(`Команда виконана`, {
      command,
      user: user.tag,
      userId: user.id,
      duration,
      success,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування помилок команд
   */
  commandError(command, user, error, duration) {
    this.error(`Помилка команди`, {
      command,
      user: user.tag,
      userId: user.id,
      error: error.message,
      stack: error.stack,
      duration,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування API запитів
   */
  apiRequest(service, endpoint, duration, success = true) {
    this.info(`API запит`, {
      service,
      endpoint,
      duration,
      success,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування помилок API
   */
  apiError(service, endpoint, error, duration) {
    this.error(`API помилка`, {
      service,
      endpoint,
      error: error.message,
      duration,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування безпеки
   */
  security(event, user, details = {}) {
    this.warn(`Подія безпеки`, {
      event,
      user: user.tag,
      userId: user.id,
      details,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування продуктивності
   */
  performance(operation, duration, details = {}) {
    this.debug(`Продуктивність`, {
      operation,
      duration,
      details,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Логування системних подій
   */
  system(event, details = {}) {
    this.info(`Системна подія`, {
      event,
      details,
      timestamp: new Date().toISOString(),
    });
  }

  /**
   * Отримання статистики логів
   */
  getStats() {
    return {
      level: this.logger.level,
      transports: this.logger.transports.length,
      timestamp: new Date().toISOString(),
    };
  }

  /**
   * Очищення старих логів
   */
  async cleanup() {
    try {
      const logsDir = path.join(process.cwd(), 'logs');
      const files = fs.readdirSync(logsDir);

      for (const file of files) {
        const filePath = path.join(logsDir, file);
        const stats = fs.statSync(filePath);
        const daysOld = (Date.now() - stats.mtime.getTime()) / (1000 * 60 * 60 * 24);

        // Видаляємо файли старіше 30 днів
        if (daysOld > 30) {
          fs.unlinkSync(filePath);
          this.info(`Видалено старий лог файл: ${file}`);
        }
      }
    } catch (error) {
      this.error('Помилка очищення логів:', error);
    }
  }
}

// Створення глобального екземпляру логера
const logger = new Logger();

// Експорт функцій для зворотної сумісності
module.exports = {
  info: (message, meta) => logger.info(message, meta),
  error: (message, meta) => logger.error(message, meta),
  warn: (message, meta) => logger.warn(message, meta),
  debug: (message, meta) => logger.debug(message, meta),
  command: (command, user, duration, success) => logger.command(command, user, duration, success),
  commandError: (command, user, error, duration) =>
    logger.commandError(command, user, error, duration),
  apiRequest: (service, endpoint, duration, success) =>
    logger.apiRequest(service, endpoint, duration, success),
  apiError: (service, endpoint, error, duration) =>
    logger.apiError(service, endpoint, error, duration),
  security: (event, user, details) => logger.security(event, user, details),
  performance: (operation, duration, details) => logger.performance(operation, duration, details),
  system: (event, details) => logger.system(event, details),
  getStats: () => logger.getStats(),
  cleanup: () => logger.cleanup(),
};
