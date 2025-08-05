/**
 * Модуль безпеки для Discord AI Bot
 * Включає управління ролями, rate limiting та валідацію
 * TypeScript версія 3.0.0 - Повністю рефакторовано
 */

import { GuildMember, CommandInteraction, PermissionFlagsBits } from 'discord.js';
import logger from './logger';

// Конфігурація ролей
const ROLES = {
  ADMIN: 'Адміністратор',
  BOT_USER: 'Бот-Користувач',
  SHEETS_ACCESS: 'Sheets-Доступ',
  AI_ACCESS: 'AI-Доступ',
  EXPORT_ACCESS: 'Експорт-Доступ',
  MODERATOR: 'Модератор',
  VIEWER: 'Переглядач',
} as const;

// Конфігурація rate limiting
const RATE_LIMITS = {
  SEARCH: { max: 10, window: 60, penalty: 30 }, // 10 пошуків за хвилину
  AI_ANALYSIS: { max: 5, window: 120, penalty: 60 }, // 5 AI-аналізів за 2 хвилини
  EXPORT: { max: 3, window: 300, penalty: 120 }, // 3 експорти за 5 хвилин
  GENERAL: { max: 20, window: 60, penalty: 30 }, // 20 загальних команд за хвилину
  ADMIN: { max: 100, window: 60, penalty: 0 }, // Адміністратори мають більше прав
} as const;

// Конфігурація безпеки
const SECURITY_CONFIG = {
  MAX_INPUT_LENGTH: 1000,
  MAX_COMMAND_LENGTH: 100,
  MAX_SEARCH_LENGTH: 500,
  CLEANUP_INTERVAL: 5 * 60 * 1000, // 5 хвилин
  CACHE_MAX_SIZE: 10000,
  SUSPICIOUS_PATTERNS: [
    /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
    /javascript:/gi,
    /on\w+\s*=/gi,
    /data:text\/html/gi,
    /vbscript:/gi,
    /onload/gi,
    /onerror/gi,
    /eval\s*\(/gi,
    /document\./gi,
    /window\./gi,
    /alert\s*\(/gi,
    /confirm\s*\(/gi,
    /prompt\s*\(/gi,
  ],
} as const;

interface RateLimitEntry {
  count: number;
  resetTime: number;
  penaltyEndTime: number;
  violations: number;
  lastRequest: number;
}

interface SecurityStats {
  totalChecks: number;
  deniedAccess: number;
  rateLimited: number;
  securityEvents: number;
  suspiciousInputs: number;
  cacheHits: number;
  cacheMisses: number;
  lastCleanup: Date;
}

interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  sanitizedValue?: string;
}

interface PermissionCheckResult {
  allowed: boolean;
  reason?: string;
  requiredRoles?: string[];
  userRoles?: string[];
  rateLimited?: boolean;
  penaltyTime?: number;
}

// In-memory кеш для rate limiting (в продакшені використовуйте Redis)
const rateLimitCache = new Map<string, RateLimitEntry>();

// Кеш для ролей користувачів
const roleCache = new Map<string, { roles: string[]; timestamp: number }>();

// Статистика безпеки
const securityStats: SecurityStats = {
  totalChecks: 0,
  deniedAccess: 0,
  rateLimited: 0,
  securityEvents: 0,
  suspiciousInputs: 0,
  cacheHits: 0,
  cacheMisses: 0,
  lastCleanup: new Date(),
};

/**
 * Перевірка наявності ролі у користувача з кешуванням
 */
function hasRole(member: GuildMember | null, requiredRoles: string | string[]): boolean {
  const startTime = performance.now();
  
  try {
    if (!member || !member.roles) {
      logger.warn('Invalid member object provided to hasRole', { 
        hasMember: !!member, 
        hasRoles: !!member?.roles 
      });
      return false;
    }

    const userId = member.id;
    const now = Date.now();
    const cacheKey = `${userId}:roles`;
    
    // Перевірка кешу
    const cached = roleCache.get(cacheKey);
    if (cached && (now - cached.timestamp) < 300000) { // 5 хвилин кеш
      securityStats.cacheHits++;
      const userRoles = cached.roles;
      
      if (Array.isArray(requiredRoles)) {
        return requiredRoles.some(role => userRoles.includes(role));
      }
      return userRoles.includes(requiredRoles);
    }

    securityStats.cacheMisses++;
    
    // Отримання ролей з Discord
    const userRoles = member.roles.cache.map(role => role.name);
    
    // Кешування ролей
    roleCache.set(cacheKey, { roles: userRoles, timestamp: now });
    
    // Обмеження розміру кешу
    if (roleCache.size > SECURITY_CONFIG.CACHE_MAX_SIZE) {
      const oldestKey = roleCache.keys().next().value;
      roleCache.delete(oldestKey);
    }

    const hasRequiredRole = Array.isArray(requiredRoles) 
      ? requiredRoles.some(role => userRoles.includes(role))
      : userRoles.includes(requiredRoles);

    const duration = performance.now() - startTime;
    logger.debug(`Role check completed in ${duration.toFixed(2)}ms`, {
      userId,
      hasRole: hasRequiredRole,
      userRoles,
      requiredRoles,
    });

    return hasRequiredRole;
    
  } catch (error) {
    logger.error('Error in hasRole function:', error);
    return false;
  }
}

/**
 * Перевірка прав доступу для команди з детальним логуванням
 */
async function checkPermission(
  interaction: CommandInteraction, 
  requiredRoles: string | string[], 
  commandName: string
): Promise<PermissionCheckResult> {
  const startTime = performance.now();
  
  try {
    securityStats.totalChecks++;
    
    const userId = interaction.user.id;
    const userTag = interaction.user.tag;
    const guildId = interaction.guildId;
    
    logger.debug(`Permission check started for ${userTag}`, {
      command: commandName,
      userId,
      guildId,
      requiredRoles,
    });

    // Перевірка чи це серверний канал
    if (!interaction.guild) {
      logger.warn('Command attempted in DM', { userTag, command: commandName });
      return {
        allowed: false,
        reason: 'Ця команда доступна тільки на сервері',
      };
    }

    // Перевірка ролей
    const member = interaction.member as GuildMember;
    if (!hasRole(member, requiredRoles)) {
      securityStats.deniedAccess++;
      
      const userRoles = member.roles.cache.map(role => role.name);
      
      logger.warn('Access denied due to insufficient roles', {
        userTag,
        command: commandName,
        userRoles,
        requiredRoles,
        guildId,
      });
      
      return {
        allowed: false,
        reason: `У вас немає дозволу для використання команди \`${commandName}\``,
        requiredRoles: Array.isArray(requiredRoles) ? requiredRoles : [requiredRoles],
        userRoles,
      };
    }

    // Rate limiting
    const rateLimitResult = await checkRateLimit(userId, commandName);
    if (rateLimitResult.limited) {
      securityStats.rateLimited++;
      
      logger.warn('Access denied due to rate limiting', {
        userTag,
        command: commandName,
        penaltyTime: rateLimitResult.penaltyTime,
        violations: rateLimitResult.violations,
      });
      
      return {
        allowed: false,
        reason: 'Ви надіслали забагато запитів. Будь ласка, зачекайте.',
        rateLimited: true,
        penaltyTime: rateLimitResult.penaltyTime,
      };
    }

    const duration = performance.now() - startTime;
    logger.info(`Access granted for ${userTag} to command: ${commandName}`, {
      duration: `${duration.toFixed(2)}ms`,
      userId,
      guildId,
    });
    
    return { allowed: true };
    
  } catch (error) {
    const duration = performance.now() - startTime;
    logger.error(`Permission check error after ${duration.toFixed(2)}ms:`, error);
    
    return {
      allowed: false,
      reason: 'Помилка перевірки прав доступу',
    };
  }
}

/**
 * Розширена перевірка rate limiting
 */
async function checkRateLimit(userId: string, commandType: string): Promise<{
  limited: boolean;
  penaltyTime?: number;
  violations: number;
  remainingRequests?: number;
}> {
  try {
    const now = Date.now();
    const key = `${userId}:${commandType}`;
    const limit = RATE_LIMITS[commandType as keyof typeof RATE_LIMITS] || RATE_LIMITS.GENERAL;

    const entry = rateLimitCache.get(key);

    // Перевірка штрафного часу
    if (entry && now < entry.penaltyEndTime) {
      return {
        limited: true,
        penaltyTime: entry.penaltyEndTime - now,
        violations: entry.violations,
      };
    }

    if (!entry || now > entry.resetTime) {
      // Новий період або перший запит
      rateLimitCache.set(key, {
        count: 1,
        resetTime: now + (limit.window * 1000),
        penaltyEndTime: 0,
        violations: 0,
        lastRequest: now,
      });
      
      return {
        limited: false,
        violations: 0,
        remainingRequests: limit.max - 1,
      };
    }

    // Перевірка ліміту
    if (entry.count >= limit.max) {
      // Ліміт перевищено - встановлюємо штраф
      const penaltyDuration = limit.penalty * 1000;
      entry.penaltyEndTime = now + penaltyDuration;
      entry.violations++;
      
      logger.warn('Rate limit exceeded', {
        userId,
        commandType,
        violations: entry.violations,
        penaltyDuration,
      });
      
      return {
        limited: true,
        penaltyTime: penaltyDuration,
        violations: entry.violations,
      };
    }

    // Збільшуємо лічильник
    entry.count++;
    entry.lastRequest = now;
    
    return {
      limited: false,
      violations: entry.violations,
      remainingRequests: limit.max - entry.count,
    };
    
  } catch (error) {
    logger.error('Rate limit check error:', error);
    return { limited: false, violations: 0 }; // У випадку помилки дозволяємо доступ
  }
}

/**
 * Розширена санітизація вхідних даних
 */
function sanitizeInput(input: string, type: 'general' | 'search' | 'command' = 'general'): ValidationResult {
  const startTime = performance.now();
  
  try {
    if (!input || typeof input !== 'string') {
      return {
        isValid: false,
        errors: ['Вхідні дані відсутні або некоректні'],
        warnings: [],
      };
    }

    let sanitized = input.trim();
    const warnings: string[] = [];
    const errors: string[] = [];

    // Перевірка довжини
    const maxLength = type === 'search' 
      ? SECURITY_CONFIG.MAX_SEARCH_LENGTH 
      : type === 'command' 
        ? SECURITY_CONFIG.MAX_COMMAND_LENGTH 
        : SECURITY_CONFIG.MAX_INPUT_LENGTH;

    if (sanitized.length > maxLength) {
      errors.push(`Вхідні дані занадто довгі (максимум ${maxLength} символів)`);
      sanitized = sanitized.substring(0, maxLength);
    }

    // Перевірка на підозрілі патерни
    let suspiciousFound = false;
    SECURITY_CONFIG.SUSPICIOUS_PATTERNS.forEach((pattern, index) => {
      if (pattern.test(sanitized)) {
        suspiciousFound = true;
        securityStats.suspiciousInputs++;
        warnings.push(`Виявлено підозрілий патерн #${index + 1}`);
        sanitized = sanitized.replace(pattern, '');
      }
    });

    if (suspiciousFound) {
      logger.warn('Suspicious input detected', {
        originalLength: input.length,
        sanitizedLength: sanitized.length,
        type,
      });
    }

    // Специфічна обробка для різних типів
    switch (type) {
      case 'search':
        // Для пошуку дозволяємо більше символів
        sanitized = sanitized.replace(/[<>]/g, '');
        break;
      case 'command':
        // Для команд більш строга фільтрація
        sanitized = sanitized.replace(/[^a-zA-Z0-9\s\-_.,!?()]/g, '');
        break;
      default:
        // Загальна фільтрація
        sanitized = sanitized.replace(/[<>]/g, '');
    }

    // Додаткові перевірки
    if (sanitized.includes('http') || sanitized.includes('www')) {
      warnings.push('Виявлено потенційні посилання');
    }

    if (sanitized.includes('@') && sanitized.includes('.')) {
      warnings.push('Виявлено потенційну email адресу');
    }

    const duration = performance.now() - startTime;
    logger.debug(`Input sanitization completed in ${duration.toFixed(2)}ms`, {
      originalLength: input.length,
      sanitizedLength: sanitized.length,
      type,
      warnings: warnings.length,
      errors: errors.length,
    });

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue: sanitized,
    };
    
  } catch (error) {
    logger.error('Input sanitization error:', error);
    return {
      isValid: false,
      errors: ['Помилка санітизації вхідних даних'],
      warnings: [],
    };
  }
}

/**
 * Розширена валідація опцій команди
 */
function validateCommandOptions(options: any, schema: Record<string, any>): ValidationResult {
  const startTime = performance.now();
  const errors: string[] = [];
  const warnings: string[] = [];

  try {
    for (const [key, rules] of Object.entries(schema)) {
      const value = options[key];

      // Перевірка обов'язковості
      if (rules.required && (value === undefined || value === null || value === '')) {
        errors.push(`Поле '${key}' є обов'язковим`);
        continue;
      }

      if (value !== undefined && value !== null) {
        // Перевірка типу
        if (rules.type && typeof value !== rules.type) {
          errors.push(`Поле '${key}' має бути типу ${rules.type}, отримано ${typeof value}`);
        }

        // Перевірка довжини для рядків
        if (typeof value === 'string') {
          if (rules.minLength && value.length < rules.minLength) {
            errors.push(`Поле '${key}' має бути не менше ${rules.minLength} символів`);
          }

          if (rules.maxLength && value.length > rules.maxLength) {
            errors.push(`Поле '${key}' має бути не більше ${rules.maxLength} символів`);
          }

          // Санітизація рядків
          const sanitized = sanitizeInput(value, rules.sanitizeType || 'general');
          if (!sanitized.isValid) {
            errors.push(...sanitized.errors);
          }
          if (sanitized.warnings.length > 0) {
            warnings.push(...sanitized.warnings.map(w => `${key}: ${w}`));
          }
        }

        // Перевірка діапазону для чисел
        if (typeof value === 'number') {
          if (rules.min !== undefined && value < rules.min) {
            errors.push(`Поле '${key}' має бути не менше ${rules.min}`);
          }

          if (rules.max !== undefined && value > rules.max) {
            errors.push(`Поле '${key}' має бути не більше ${rules.max}`);
          }
        }

        // Перевірка патерну
        if (rules.pattern && !rules.pattern.test(value)) {
          errors.push(`Поле '${key}' не відповідає необхідному формату`);
        }

        // Перевірка enum
        if (rules.enum && !rules.enum.includes(value)) {
          errors.push(`Поле '${key}' має бути одним з: ${rules.enum.join(', ')}`);
        }
      }
    }
  } catch (error) {
    logger.error('Command options validation error:', error);
    errors.push('Помилка валідації опцій команди');
  }

  const duration = performance.now() - startTime;
  logger.debug(`Command options validation completed in ${duration.toFixed(2)}ms`, {
    fields: Object.keys(schema).length,
    errors: errors.length,
    warnings: warnings.length,
  });

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
  };
}

/**
 * Розширене логування подій безпеки
 */
function logSecurityEvent(event: string, data: Record<string, any>): void {
  try {
    securityStats.securityEvents++;
    
    const enhancedData = {
      ...data,
      timestamp: new Date().toISOString(),
      eventType: event,
      severity: data.severity || 'medium',
    };
    
    logger.security(event, data.user || 'unknown', enhancedData);
    
    // Додаткове логування для критичних подій
    if (data.severity === 'high' || data.severity === 'critical') {
      logger.error('Critical security event detected', enhancedData);
    }
    
  } catch (error) {
    logger.error('Security event logging error:', error);
  }
}

/**
 * Отримання детальної статистики безпеки
 */
function getSecurityStats(): SecurityStats & {
  cacheSize: number;
  rateLimitCacheSize: number;
  uptime: number;
} {
  return { 
    ...securityStats,
    cacheSize: roleCache.size,
    rateLimitCacheSize: rateLimitCache.size,
    uptime: Date.now() - securityStats.lastCleanup.getTime(),
  };
}

/**
 * Очищення застарілих записів rate limiting
 */
function cleanupRateLimitCache(): void {
  try {
    const now = Date.now();
    const keysToDelete: string[] = [];
    let cleanedCount = 0;

    for (const [key, entry] of rateLimitCache.entries()) {
      if (now > entry.resetTime && now > entry.penaltyEndTime) {
        keysToDelete.push(key);
        cleanedCount++;
      }
    }

    keysToDelete.forEach(key => rateLimitCache.delete(key));

    // Очищення кешу ролей
    const roleKeysToDelete: string[] = [];
    for (const [key, entry] of roleCache.entries()) {
      if (now - entry.timestamp > 300000) { // 5 хвилин
        roleKeysToDelete.push(key);
      }
    }
    roleKeysToDelete.forEach(key => roleCache.delete(key));

    if (cleanedCount > 0 || roleKeysToDelete.length > 0) {
      logger.debug(`Security cache cleanup: ${cleanedCount} rate limit entries, ${roleKeysToDelete.length} role entries`);
    }
    
    securityStats.lastCleanup = new Date();
    
  } catch (error) {
    logger.error('Security cache cleanup error:', error);
  }
}

/**
 * Повне очищення ресурсів
 */
function cleanup(): void {
  try {
    rateLimitCache.clear();
    roleCache.clear();
    securityStats.lastCleanup = new Date();
    
    logger.info('Security module cleanup completed', {
      rateLimitCacheSize: 0,
      roleCacheSize: 0,
    });
  } catch (error) {
    logger.error('Security cleanup error:', error);
  }
}

/**
 * Перевірка стану модуля безпеки
 */
function isHealthy(): boolean {
  try {
    const stats = getSecurityStats();
    const memoryUsage = process.memoryUsage();
    
    // Перевірка використання пам'яті
    const memoryLimit = 100 * 1024 * 1024; // 100MB
    if (memoryUsage.heapUsed > memoryLimit) {
      logger.warn('Security module memory usage high', {
        heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
        limit: `${Math.round(memoryLimit / 1024 / 1024)}MB`,
      });
    }
    
    return true;
  } catch (error) {
    logger.error('Security module health check failed:', error);
    return false;
  }
}

// Автоматичне очищення кешу
setInterval(cleanupRateLimitCache, SECURITY_CONFIG.CLEANUP_INTERVAL);

// Періодична перевірка стану
setInterval(() => {
  if (!isHealthy()) {
    logger.warn('Security module health check failed, performing cleanup');
    cleanup();
  }
}, 10 * 60 * 1000); // 10 хвилин

export {
  ROLES,
  RATE_LIMITS,
  SECURITY_CONFIG,
  hasRole,
  checkPermission,
  checkRateLimit,
  sanitizeInput,
  validateCommandOptions,
  logSecurityEvent,
  getSecurityStats,
  cleanup,
  isHealthy,
  type PermissionCheckResult,
  type ValidationResult,
}; 