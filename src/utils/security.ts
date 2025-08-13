/**
 * Розширена система безпеки для Discord AI Assistant Bot
 * Валідація, санітизація та захист від атак
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import type { LogMeta, SecurityEvent, SecurityValidationResult } from '@/types';
import { handleError } from './errorHandler';
import logger from './logger';

// Константи для безпеки
const SECURITY_CONSTANTS = {
  MAX_INPUT_LENGTH: 2000,
  MAX_COMMAND_LENGTH: 100,
  MAX_URL_LENGTH: 500,
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  RATE_LIMIT_WINDOW: 60000, // 1 хвилина
  RATE_LIMIT_MAX: 10, // 10 запитів за хвилину
  SUSPICIOUS_PATTERNS: [
    /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
    /javascript:/gi,
    /on\w+\s*=/gi,
    /data:text\/html/gi,
    /vbscript:/gi,
    /<iframe/gi,
    /<object/gi,
    /<embed/gi,
    /<applet/gi,
    /<meta/gi,
    /<link/gi,
    /<base/gi,
    /<form/gi,
    /<input/gi,
    /<textarea/gi,
    /<select/gi,
    /<button/gi,
    /<label/gi,
    /<fieldset/gi,
    /<legend/gi,
    /<optgroup/gi,
    /<option/gi,
  ],
  ALLOWED_CHARS: /^[a-zA-Z0-9\s\-_.,!?@#$%^&*()+=<>{}[\]|\\/:;"'`~]+$/,
  ALLOWED_URLS:
    /^https?:\/\/(www\.)?(discord\.com|discordapp\.com|google\.com|docs\.google\.com|drive\.google\.com)/i,
  BLACKLISTED_WORDS: [
    'admin',
    'root',
    'sudo',
    'system',
    'exec',
    'eval',
    'require',
    'import',
    'delete',
    'drop',
    'insert',
    'update',
    'select',
    'union',
    'where',
    'script',
    'javascript',
    'vbscript',
    'onload',
    'onerror',
    'onclick',
  ],
} as const;

export interface SecurityStats {
  totalValidations: number;
  successfulValidations: number;
  failedValidations: number;
  suspiciousActivities: number;
  rateLimitHits: number;
  blacklistHits: number;
  xssAttempts: number;
  sqlInjectionAttempts: number;
  lastSecurityEvent?: SecurityEvent;
  averageValidationTime: number;
  totalValidationTime: number;
}

export interface RateLimitInfo {
  count: number;
  resetTime: number;
  lastRequest: number;
}

export class SecurityManager {
  private static instance: SecurityManager | null = null;
  private stats!: SecurityStats;
  private rateLimitMap = new Map<string, RateLimitInfo>();
  private blacklistCache = new Set<string>();
  private suspiciousActivities: SecurityEvent[] = [];
  private _isInitialized = false;

  constructor() {
    if (SecurityManager.instance) {
      return SecurityManager.instance;
    }
    SecurityManager.instance = this;

    this.stats = {
      totalValidations: 0,
      successfulValidations: 0,
      failedValidations: 0,
      suspiciousActivities: 0,
      rateLimitHits: 0,
      blacklistHits: 0,
      xssAttempts: 0,
      sqlInjectionAttempts: 0,
      averageValidationTime: 0,
      totalValidationTime: 0,
    };

    this.initialize();
  }

  /**
   * Ініціалізація системи безпеки
   */
  private initialize(): void {
    try {
      logger.info('🔒 Ініціалізація системи безпеки...');

      // Завантаження чорного списку
      this.loadBlacklist();

      // Запуск періодичних завдань
      this.startPeriodicTasks();

      this._isInitialized = true;
      logger.info('✅ Система безпеки успішно ініціалізована');
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'initialize' },
      });
      throw new Error('Помилка ініціалізації системи безпеки');
    }
  }

  /**
   * Завантаження чорного списку
   */
  private loadBlacklist(): void {
    try {
      // Тут можна завантажити чорний список з файлу або бази даних
      SECURITY_CONSTANTS.BLACKLISTED_WORDS.forEach(word => {
        this.blacklistCache.add(word.toLowerCase());
      });

      logger.info(`📋 Завантажено ${this.blacklistCache.size} слів у чорний список`);
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'loadBlacklist' },
      });
    }
  }

  /**
   * Запуск періодичних завдань
   */
  private startPeriodicTasks(): void {
    // Очищення rate limit кешу кожні 5 хвилин
    setInterval(
      () => {
        this.cleanupRateLimitCache();
      },
      5 * 60 * 1000
    );

    // Очищення підозрілої активності кожні 10 хвилин
    setInterval(
      () => {
        this.cleanupSuspiciousActivities();
      },
      10 * 60 * 1000
    );

    logger.info('⏰ Періодичні завдання безпеки запущено');
  }

  /**
   * Валідація та санітизація введення
   */
  public validateInput(
    input: string,
    context: {
      userId?: string;
      guildId?: string;
      channelId?: string;
      commandName?: string;
      inputType?: 'command' | 'message' | 'url' | 'file';
    } = {}
  ): SecurityValidationResult {
    const startTime = performance.now();

    try {
      logger.debug('🔍 Валідація введення...', {
        inputLength: input.length,
        inputType: context.inputType,
        userId: context.userId,
        commandName: context.commandName,
      } as LogMeta);

      const errors: string[] = [];
      const warnings: string[] = [];
      let sanitizedValue = input;

      // Перевірка довжини
      if (input.length > SECURITY_CONSTANTS.MAX_INPUT_LENGTH) {
        errors.push(
          `Введення занадто довге (${input.length} символів, максимум ${SECURITY_CONSTANTS.MAX_INPUT_LENGTH})`
        );
        sanitizedValue = input.substring(0, SECURITY_CONSTANTS.MAX_INPUT_LENGTH);
      }

      // Перевірка на XSS атаки
      const xssResult = this.checkForXSS(input);
      if (xssResult.found) {
        errors.push('Виявлено потенційну XSS атаку');
        this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
          subtype: 'xss_attempt',
          pattern: xssResult.pattern,
          input: input.substring(0, 100),
        });
        this.stats.xssAttempts++;
      }

      // Перевірка на SQL ін'єкції
      const sqlResult = this.checkForSQLInjection(input);
      if (sqlResult.found) {
        errors.push("Виявлено потенційну SQL ін'єкцію");
        this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
          subtype: 'sql_injection_attempt',
          pattern: sqlResult.pattern,
          input: input.substring(0, 100),
        });
        this.stats.sqlInjectionAttempts++;
      }

      // Перевірка чорного списку
      const blacklistResult = this.checkBlacklist(input);
      if (blacklistResult.found) {
        warnings.push('Виявлено слова з чорного списку');
        this.stats.blacklistHits++;
      }

      // Перевірка дозволених символів
      if (!SECURITY_CONSTANTS.ALLOWED_CHARS.test(input)) {
        warnings.push('Введення містить недозволені символи');
      }

      // Санітизація
      sanitizedValue = this.sanitizeInput(input);

      const duration = performance.now() - startTime;
      this.updateStats(true, duration);

      const result: SecurityValidationResult = {
        isValid: errors.length === 0,
        sanitizedValue,
        errors,
        warnings,
      };

      if (errors.length > 0) {
        this.stats.failedValidations++;
        logger.warn('❌ Валідація введення невдала', {
          errors,
          warnings,
          inputLength: input.length,
          userId: context.userId,
          commandName: context.commandName,
        } as LogMeta);
      } else {
        this.stats.successfulValidations++;
        logger.debug('✅ Валідація введення успішна', {
          inputLength: input.length,
          warnings,
          userId: context.userId,
          commandName: context.commandName,
        } as LogMeta);
      }

      return result;
    } catch (error) {
      const duration = performance.now() - startTime;
      this.updateStats(false, duration);

      handleError(error, {
        serviceName: 'SecurityManager',
        ...(context.userId ? { userId: context.userId } : {}),
        additionalContext: { operation: 'validateInput', input: input.substring(0, 100) },
      });

      return {
        isValid: false,
        sanitizedValue: '',
        errors: ['Помилка валідації введення'],
        warnings: [],
      };
    }
  }

  /**
   * Перевірка на XSS атаки
   */
  private checkForXSS(input: string): { found: boolean; pattern?: string } {
    for (const pattern of SECURITY_CONSTANTS.SUSPICIOUS_PATTERNS) {
      if (pattern.test(input)) {
        return { found: true, pattern: pattern.source };
      }
    }
    return { found: false };
  }

  /**
   * Перевірка на SQL ін'єкції
   */
  private checkForSQLInjection(input: string): { found: boolean; pattern?: string } {
    const sqlPatterns = [
      /(\b(union|select|insert|update|delete|drop|create|alter)\b)/i,
      /(\b(where|from|into|values|set)\b)/i,
      /(--|#|\/\*|\*\/)/,
      /(\b(and|or)\b\s+\d+\s*=\s*\d+)/i,
      /(\b(and|or)\b\s+['"]\w+['"]\s*=\s*['"]\w+['"])/i,
    ];

    for (const pattern of sqlPatterns) {
      if (pattern.test(input)) {
        return { found: true, pattern: pattern.source };
      }
    }
    return { found: false };
  }

  /**
   * Перевірка чорного списку
   */
  private checkBlacklist(input: string): { found: boolean; words: string[] } {
    const foundWords: string[] = [];
    const words = input.toLowerCase().split(/\s+/);

    for (const word of words) {
      if (this.blacklistCache.has(word)) {
        foundWords.push(word);
      }
    }

    return {
      found: foundWords.length > 0,
      words: foundWords,
    };
  }

  /**
   * Санітизація введення
   */
  private sanitizeInput(input: string): string {
    let sanitized = input;

    // Видалення HTML тегів
    sanitized = sanitized.replace(/<[^>]*>/g, '');

    // Екранування спеціальних символів
    sanitized = sanitized
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#x27;')
      .replace(/\//g, '&#x2F;');

    // Видалення зайвих пробілів
    sanitized = sanitized.trim().replace(/\s+/g, ' ');

    return sanitized;
  }

  /**
   * Перевірка rate limit
   */
  public checkRateLimit(userId: string): {
    allowed: boolean;
    remaining: number;
    resetTime: number;
  } {
    try {
      const now = Date.now();
      const userLimit = this.rateLimitMap.get(userId);

      if (!userLimit || now > userLimit.resetTime) {
        // Створення нового ліміту
        this.rateLimitMap.set(userId, {
          count: 1,
          resetTime: now + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
          lastRequest: now,
        });

        return {
          allowed: true,
          remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - 1,
          resetTime: now + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
        };
      }

      if (userLimit.count >= SECURITY_CONSTANTS.RATE_LIMIT_MAX) {
        this.stats.rateLimitHits++;
        this.recordSecurityEvent('rate_limit', userId, {
          count: userLimit.count,
          resetTime: userLimit.resetTime,
        });

        logger.warn('⏰ Rate limit перевищено', {
          userId,
          count: userLimit.count,
          resetTime: userLimit.resetTime,
        } as LogMeta);

        return {
          allowed: false,
          remaining: 0,
          resetTime: userLimit.resetTime,
        };
      }

      // Збільшення лічильника
      userLimit.count++;
      userLimit.lastRequest = now;

      return {
        allowed: true,
        remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - userLimit.count,
        resetTime: userLimit.resetTime,
      };
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        userId,
        additionalContext: { operation: 'checkRateLimit' },
      });

      // У випадку помилки дозволяємо запит
      return {
        allowed: true,
        remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX,
        resetTime: Date.now() + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW,
      };
    }
  }

  /**
   * Валідація URL
   */
  public validateUrl(url: string): SecurityValidationResult {
    try {
      const errors: string[] = [];
      const warnings: string[] = [];

      // Перевірка довжини
      if (url.length > SECURITY_CONSTANTS.MAX_URL_LENGTH) {
        errors.push(
          `URL занадто довгий (${url.length} символів, максимум ${SECURITY_CONSTANTS.MAX_URL_LENGTH})`
        );
      }

      // Перевірка протоколу
      if (!url.startsWith('http://') && !url.startsWith('https://')) {
        errors.push('URL повинен починатися з http:// або https://');
      }

      // Перевірка дозволених доменів
      if (!SECURITY_CONSTANTS.ALLOWED_URLS.test(url)) {
        warnings.push('URL не з дозволеного домену');
      }

      // Перевірка на підозрілі патерни
      if (url.includes('javascript:') || url.includes('data:text/html')) {
        errors.push('URL містить підозрілі патерни');
      }

      return {
        isValid: errors.length === 0,
        sanitizedValue: url,
        errors,
        warnings,
      };
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'validateUrl', url },
      });

      return {
        isValid: false,
        sanitizedValue: '',
        errors: ['Помилка валідації URL'],
        warnings: [],
      };
    }
  }

  /**
   * Запис події безпеки
   */
  private recordSecurityEvent(
    type: SecurityEvent['type'],
    userId: string,
    details: Record<string, unknown> = {}
  ): void {
    try {
      const event: SecurityEvent = {
        type,
        userId,
        details,
        timestamp: new Date(),
        severity: this.determineEventSeverity(type),
      };

      this.suspiciousActivities.push(event);
      this.stats.suspiciousActivities++;
      this.stats.lastSecurityEvent = event;

      // Обмеження розміру масиву
      if (this.suspiciousActivities.length > 1000) {
        this.suspiciousActivities = this.suspiciousActivities.slice(-500);
      }

      logger.security(type, userId, {
        details,
        severity: event.severity,
        timestamp: event.timestamp.toISOString(),
      } as LogMeta);
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        userId,
        additionalContext: { operation: 'recordSecurityEvent', type },
      });
    }
  }

  /**
   * Визначення серйозності події
   */
  private determineEventSeverity(type: SecurityEvent['type']): SecurityEvent['severity'] {
    switch (type) {
      case 'unauthorized_access':
        return 'high';
      case 'rate_limit':
        return 'medium';
      case 'invalid_input':
      case 'suspicious_activity':
      default:
        return 'low';
    }
  }

  /**
   * Очищення rate limit кешу
   */
  private cleanupRateLimitCache(): void {
    try {
      const now = Date.now();
      let cleanedCount = 0;

      for (const [userId, limit] of this.rateLimitMap.entries()) {
        if (now > limit.resetTime) {
          this.rateLimitMap.delete(userId);
          cleanedCount++;
        }
      }

      if (cleanedCount > 0) {
        logger.debug(`🧹 Очищено ${cleanedCount} застарілих rate limit записів`);
      }
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'cleanupRateLimitCache' },
      });
    }
  }

  /**
   * Очищення підозрілої активності
   */
  private cleanupSuspiciousActivities(): void {
    try {
      const now = new Date();
      const maxAge = 24 * 60 * 60 * 1000; // 24 години
      const initialCount = this.suspiciousActivities.length;

      this.suspiciousActivities = this.suspiciousActivities.filter(
        activity => now.getTime() - activity.timestamp.getTime() < maxAge
      );

      const cleanedCount = initialCount - this.suspiciousActivities.length;
      if (cleanedCount > 0) {
        logger.debug(`🧹 Очищено ${cleanedCount} застарілих подій безпеки`);
      }
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'cleanupSuspiciousActivities' },
      });
    }
  }

  /**
   * Оновлення статистики
   */
  private updateStats(_success: boolean, duration: number): void {
    try {
      this.stats.totalValidations++;
      this.stats.totalValidationTime += duration;
      this.stats.averageValidationTime =
        this.stats.totalValidationTime / this.stats.totalValidations;
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'updateStats' },
      });
    }
  }

  /**
   * Отримання статистики безпеки
   */
  public getStats(): SecurityStats {
    return { ...this.stats };
  }

  /**
   * Отримання підозрілої активності
   */
  public getSuspiciousActivities(): SecurityEvent[] {
    return [...this.suspiciousActivities];
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    try {
      this.rateLimitMap.clear();
      this.suspiciousActivities = [];
      this.blacklistCache.clear();

      logger.info('🧹 Ресурси SecurityManager очищено');
    } catch (error) {
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'cleanup' },
      });
    }
  }

  /**
   * Перевірка стану ініціалізації
   */
  public isInitialized(): boolean {
    return this._isInitialized;
  }
}

// Експорт єдиного екземпляра
export const securityManager = new SecurityManager();

// Експорт функцій для зручності
export const validateInput = (
  input: string,
  context?: {
    userId?: string;
    guildId?: string;
    channelId?: string;
    commandName?: string;
    inputType?: 'command' | 'message' | 'url' | 'file';
  }
) => securityManager.validateInput(input, context);

export const checkRateLimit = (userId: string) => securityManager.checkRateLimit(userId);
export const validateUrl = (url: string) => securityManager.validateUrl(url);
export const getSecurityStats = () => securityManager.getStats();
export const getSuspiciousActivities = () => securityManager.getSuspiciousActivities();
export const cleanupSecurityManager = () => securityManager.cleanup();

// Функції для зворотної сумісності
// Overloads to support legacy and extended usage
export function sanitizeInput(input: string): string;
export function sanitizeInput(
  input: string,
  inputType: 'command' | 'message' | 'url' | 'file'
): SecurityValidationResult;
export function sanitizeInput(
  input: string,
  inputType?: 'command' | 'message' | 'url' | 'file'
): string | SecurityValidationResult {
  if (inputType) {
    return validateInput(input, { inputType });
  }
  const result = validateInput(input);
  return result.sanitizedValue;
}

export const validateCommandOptions = (
  options: any,
  _schema?: Record<string, any>
): SecurityValidationResult => {
  // Currently ignoring custom schema; the SecurityManager performs intrinsic checks.
  const input = JSON.stringify(options);
  return validateInput(input, { inputType: 'command' });
};
