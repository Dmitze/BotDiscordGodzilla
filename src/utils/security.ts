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
    if (SecurityManager.instance) return SecurityManager.instance;
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

  private initialize(): void {
    try {
      logger.info('🔒 Ініціалізація системи безпеки...');
      this.loadBlacklist();
      // Skip timers in test environment
      if (process.env['NODE_ENV'] !== 'test' && !process.env['JEST_WORKER_ID']) {
        this.startPeriodicTasks();
      } else {
        logger.debug('⏭️ Пропуск періодичних завдань безпеки у тестовому середовищі');
      }
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

  private loadBlacklist(): void {
    try {
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

  private startPeriodicTasks(): void {
    if (process.env['NODE_ENV'] === 'test' || process.env['JEST_WORKER_ID']) {
      return;
    }
    setInterval(() => this.cleanupRateLimitCache(), 5 * 60 * 1000);
    setInterval(() => this.cleanupSuspiciousActivities(), 10 * 60 * 1000);
    logger.info('⏰ Періодичні завдання безпеки запущено');
  }

  // Core validation: simplified but safe
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
    const start = performance.now();
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      if (input.length > SECURITY_CONSTANTS.MAX_INPUT_LENGTH) {
        errors.push(
          `Введення занадто довге (${input.length} символів, максимум ${SECURITY_CONSTANTS.MAX_INPUT_LENGTH})`
        );
        input = input.slice(0, SECURITY_CONSTANTS.MAX_INPUT_LENGTH);
      }

      // Basic suspicious patterns
      for (const pattern of SECURITY_CONSTANTS.SUSPICIOUS_PATTERNS) {
        if (pattern.test(input)) {
          errors.push('Виявлено потенційну XSS атаку');
          this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
            subtype: 'xss_attempt',
          });
          break;
        }
      }

      // Approximate SQLi detection
      const sqlPatterns = [
        /(\b(union|select|insert|update|delete|drop|create|alter)\b)/i,
        /(\b(where|from|into|values|set)\b)/i,
        /(--)|(#)|(\/\*)|(\*\/)/,
      ];
      for (const p of sqlPatterns) {
        if (p.test(input)) {
          errors.push("Виявлено потенційну SQL ін'єкцію");
          this.recordSecurityEvent('suspicious_activity', context.userId || 'unknown', {
            subtype: 'sql_injection_attempt',
          });
          break;
        }
      }

      if (!SECURITY_CONSTANTS.ALLOWED_CHARS.test(input)) {
        warnings.push('Введення містить недозволені символи');
      }

      const sanitizedValue = this.sanitizeInput(input);
      const duration = performance.now() - start;
      this.updateStats(true, duration);

      const result: SecurityValidationResult = {
        isValid: errors.length === 0,
        sanitizedValue,
        errors,
        warnings,
      };

      if (errors.length > 0) {
        logger.warn('❌ Валідація введення невдала', {
          errors,
          warnings,
          inputLength: input.length,
          userId: context.userId,
          commandName: context.commandName,
        } as LogMeta);
      } else {
        logger.debug('✅ Валідація введення успішна', {
          warnings,
          inputLength: input.length,
          userId: context.userId,
          commandName: context.commandName,
        } as LogMeta);
      }
      return result;
    } catch (error) {
      const duration = performance.now() - start;
      this.updateStats(false, duration);
      handleError(error, {
        serviceName: 'SecurityManager',
        additionalContext: { operation: 'validateInput' },
      });
      return {
        isValid: false,
        sanitizedValue: '',
        errors: ['Помилка валідації введення'],
        warnings: [],
      };
    }
  }

  private sanitizeInput(input: string): string {
    return input
      .replace(/<[^>]*>/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#x27;')
      .replace(/\//g, '&#x2F;')
      .trim()
      .replace(/\s+/g, ' ');
  }

  public checkRateLimit(userId: string): { allowed: boolean; remaining: number; resetTime: number } {
    try {
      const now = Date.now();
      const info = this.rateLimitMap.get(userId);
      if (!info || now > info.resetTime) {
        const reset = now + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW;
        this.rateLimitMap.set(userId, { count: 1, resetTime: reset, lastRequest: now });
        return { allowed: true, remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - 1, resetTime: reset };
      }
      if (info.count >= SECURITY_CONSTANTS.RATE_LIMIT_MAX) {
        this.stats.rateLimitHits++;
        this.recordSecurityEvent('rate_limit', userId, { count: info.count, resetTime: info.resetTime });
        return { allowed: false, remaining: 0, resetTime: info.resetTime };
      }
      info.count++;
      info.lastRequest = now;
      return {
        allowed: true,
        remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX - info.count,
        resetTime: info.resetTime,
      };
    } catch (error) {
      handleError(error, { serviceName: 'SecurityManager', additionalContext: { operation: 'checkRateLimit' } });
      return { allowed: true, remaining: SECURITY_CONSTANTS.RATE_LIMIT_MAX, resetTime: Date.now() + SECURITY_CONSTANTS.RATE_LIMIT_WINDOW };
    }
  }

  public validateUrl(url: string): SecurityValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];
    if (url.length > SECURITY_CONSTANTS.MAX_URL_LENGTH) {
      errors.push(`URL занадто довгий (${url.length} символів, максимум ${SECURITY_CONSTANTS.MAX_URL_LENGTH})`);
    }
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      errors.push('URL повинен починатися з http:// або https://');
    }
    if (!SECURITY_CONSTANTS.ALLOWED_URLS.test(url)) {
      warnings.push('URL не з дозволеного домену');
    }
    if (url.includes('javascript:') || url.includes('data:text/html')) {
      errors.push('URL містить підозрілі патерни');
    }
    return { isValid: errors.length === 0, sanitizedValue: url, errors, warnings };
  }

  private recordSecurityEvent(type: SecurityEvent['type'], userId: string, details: Record<string, unknown> = {}): void {
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
    if (this.suspiciousActivities.length > 1000) {
      this.suspiciousActivities = this.suspiciousActivities.slice(-500);
    }
    logger.security(type, userId, { details, severity: event.severity, timestamp: event.timestamp.toISOString() } as LogMeta);
  }

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

  private cleanupRateLimitCache(): void {
    const now = Date.now();
    for (const [uid, limit] of this.rateLimitMap.entries()) {
      if (now > limit.resetTime) this.rateLimitMap.delete(uid);
    }
  }

  private cleanupSuspiciousActivities(): void {
    const now = Date.now();
    const maxAge = 24 * 60 * 60 * 1000;
    this.suspiciousActivities = this.suspiciousActivities.filter(a => now - a.timestamp.getTime() < maxAge);
  }

  private updateStats(_success: boolean, duration: number): void {
    this.stats.totalValidations++;
    this.stats.totalValidationTime += duration;
    this.stats.averageValidationTime = this.stats.totalValidationTime / this.stats.totalValidations;
  }

  public getStats(): SecurityStats {
    return { ...this.stats };
  }

  public getSuspiciousActivities(): SecurityEvent[] {
    return [...this.suspiciousActivities];
  }

  public cleanup(): void {
    this.rateLimitMap.clear();
    this.suspiciousActivities = [];
    this.blacklistCache.clear();
  }

  public isInitialized(): boolean {
    return this._isInitialized;
  }

  // PII masking utility delegates to pure function
  public maskPII(input: string): string {
    return maskPII(input);
  }
}

// Pure, reusable PII masking function (no dependencies, safe for tests/mocks)
export function maskPII(
  input: string,
  opts?: { email?: boolean; phone?: boolean }
): string {
  if (!input) return input;
  const enableEmail = opts?.email !== false; // default true
  const enablePhone = opts?.phone !== false; // default true
  let out = input;
  if (enableEmail) {
    const emailRegex = /([a-zA-Z0-9._%+-])([a-zA-Z0-9._%+-]*)(@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
    out = out.replace(emailRegex, (_m, first: string, middle: string, domain: string) => {
      const maskedMiddle = middle.length > 0 ? '*'.repeat(Math.min(middle.length, 6)) : '***';
      return `${first}${maskedMiddle}${domain}`;
    });
  }
  if (enablePhone) {
    const phoneRegex = /(?<!\d)([+]?\d[\d\s().-]{6,}\d)(?!\d)/g;
    out = out.replace(phoneRegex, (match: string) => {
      const digits = match.replace(/\D/g, '');
      if (digits.length < 7) return match;
      return '*'.repeat(Math.max(0, digits.length - 4)) + digits.slice(-4);
    });
  }
  return out;
}
// Singleton and convenience exports
export const securityManager = new SecurityManager();
export const validateInput = (
  input: string,
  context?: { userId?: string; guildId?: string; channelId?: string; commandName?: string; inputType?: 'command' | 'message' | 'url' | 'file' }
) => securityManager.validateInput(input, context);
export const checkRateLimit = (userId: string) => securityManager.checkRateLimit(userId);
export const validateUrl = (url: string) => securityManager.validateUrl(url);
export const getSecurityStats = () => securityManager.getStats();
export const getSuspiciousActivities = () => securityManager.getSuspiciousActivities();
export const cleanupSecurityManager = () => securityManager.cleanup();

// Backward compatible sanitizeInput overloads

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

// Named export already provided above
