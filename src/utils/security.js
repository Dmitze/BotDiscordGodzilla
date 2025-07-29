/**
 * Модуль безпеки для Discord AI Bot
 * Включає управління ролями, rate limiting та валідацію
 */

const logger = require('./logger');

// Конфігурація ролей
const ROLES = {
  ADMIN: 'Адміністратор',
  BOT_USER: 'Бот-Користувач',
  SHEETS_ACCESS: 'Sheets-Доступ',
  AI_ACCESS: 'AI-Доступ',
  EXPORT_ACCESS: 'Експорт-Доступ',
};

// Конфігурація rate limiting
const RATE_LIMITS = {
  SEARCH: { max: 10, window: 60 }, // 10 пошуків за хвилину
  AI_ANALYSIS: { max: 5, window: 120 }, // 5 AI-аналізів за 2 хвилини
  EXPORT: { max: 3, window: 300 }, // 3 експорти за 5 хвилин
  GENERAL: { max: 20, window: 60 }, // 20 загальних команд за хвилину
};

// In-memory кеш для rate limiting (в продакшені використовуйте Redis)
const rateLimitCache = new Map();

/**
 * Перевірка наявності ролі у користувача
 * @param {GuildMember} member - Discord member
 * @param {string|Array} requiredRoles - Потрібні ролі
 * @returns {boolean} - Чи має користувач необхідні ролі
 */
function hasRole(member, requiredRoles) {
  try {
    if (!member || !member.roles) {
      logger.warn('Invalid member object provided to hasRole');
      return false;
    }

    const userRoles = member.roles.cache.map(role => role.name);

    if (Array.isArray(requiredRoles)) {
      return requiredRoles.some(role => userRoles.includes(role));
    }

    return userRoles.includes(requiredRoles);
  } catch (error) {
    logger.error('Error in hasRole function:', error);
    return false;
  }
}

/**
 * Перевірка прав доступу для команди
 * @param {CommandInteraction} interaction - Discord interaction
 * @param {string|Array} requiredRoles - Потрібні ролі
 * @param {string} commandName - Назва команди для логування
 * @returns {Promise<boolean>} - Чи має користувач доступ
 */
async function checkPermission(interaction, requiredRoles, commandName) {
  try {
    // Перевірка чи це серверний канал
    if (!interaction.guild) {
      await interaction.reply({
        content: '❌ Ця команда доступна тільки на сервері',
        ephemeral: true,
      });
      return false;
    }

    // Перевірка ролей
    if (!hasRole(interaction.member, requiredRoles)) {
      logger.warn(`Access denied for ${interaction.user.tag} to command: ${commandName}`);
      await interaction.reply({
        content: `❌ У вас немає дозволу для використання команди \`${commandName}\`.\nПотрібні ролі: ${
          Array.isArray(requiredRoles) ? requiredRoles.join(', ') : requiredRoles
        }`,
        ephemeral: true,
      });
      return false;
    }

    // Rate limiting
    const isLimited = await checkRateLimit(interaction.user.id, commandName);
    if (isLimited) {
      await interaction.reply({
        content: '⚠️ Ви надіслали забагато запитів. Будь ласка, зачекайте.',
        ephemeral: true,
      });
      return false;
    }

    logger.info(`Access granted for ${interaction.user.tag} to command: ${commandName}`);
    return true;
  } catch (error) {
    logger.error('Permission check error:', error);
    try {
      await interaction.reply({
        content: '❌ Помилка перевірки прав доступу',
        ephemeral: true,
      });
    } catch (replyError) {
      logger.error('Error sending permission error reply:', replyError);
    }
    return false;
  }
}

/**
 * Перевірка rate limiting
 * @param {string} userId - ID користувача
 * @param {string} commandType - Тип команди
 * @returns {Promise<boolean>} - Чи перевищено ліміт
 */
async function checkRateLimit(userId, commandType) {
  try {
    const now = Date.now();
    const limit = RATE_LIMITS[commandType] || RATE_LIMITS.GENERAL;

    if (!rateLimitCache.has(userId)) {
      rateLimitCache.set(userId, {});
    }

    const userLimits = rateLimitCache.get(userId);

    if (!userLimits[commandType]) {
      userLimits[commandType] = [];
    }

    // Видалення застарілих записів
    userLimits[commandType] = userLimits[commandType].filter(
      timestamp => now - timestamp < limit.window * 1000
    );

    // Перевірка ліміту
    if (userLimits[commandType].length >= limit.max) {
      logger.warn(`Rate limit exceeded for user ${userId} on command ${commandType}`);
      return true;
    }

    // Додавання нового запиту
    userLimits[commandType].push(now);

    return false;
  } catch (error) {
    logger.error('Rate limit check error:', error);
    return false; // У випадку помилки дозволяємо запит
  }
}

/**
 * Очищення застарілих записів rate limiting
 */
function cleanupRateLimitCache() {
  try {
    const now = Date.now();
    const maxWindow = Math.max(...Object.values(RATE_LIMITS).map(limit => limit.window));

    for (const [userId, userLimits] of rateLimitCache.entries()) {
      for (const [commandType, timestamps] of Object.entries(userLimits)) {
        const limit = RATE_LIMITS[commandType] || RATE_LIMITS.GENERAL;
        userLimits[commandType] = timestamps.filter(
          timestamp => now - timestamp < limit.window * 1000
        );
      }

      // Видалення користувачів без активних лімітів
      if (Object.values(userLimits).every(timestamps => timestamps.length === 0)) {
        rateLimitCache.delete(userId);
      }
    }
  } catch (error) {
    logger.error('Rate limit cache cleanup error:', error);
  }
}

// Очищення кешу кожні 5 хвилин
setInterval(cleanupRateLimitCache, 5 * 60 * 1000);

/**
 * Санітизація вхідних даних
 * @param {string} input - Вхідні дані
 * @param {string} type - Тип санітизації
 * @returns {string} - Очищені дані
 */
function sanitizeInput(input, type = 'general') {
  try {
    if (typeof input !== 'string') {
      return '';
    }

    let sanitized = input.trim();

    switch (type) {
      case 'search':
        // Обмеження довжини пошукового запиту
        sanitized = sanitized.slice(0, 100);
        // Видалення небезпечних символів
        sanitized = sanitized.replace(/[<>\"'&]/g, '');
        break;

      case 'filename':
        // Обмеження для назв файлів
        sanitized = sanitized.slice(0, 50);
        sanitized = sanitized.replace(/[<>:"/\\|?*]/g, '');
        break;

      case 'url':
        // Базова валідація URL
        if (!sanitized.startsWith('http://') && !sanitized.startsWith('https://')) {
          sanitized = '';
        }
        break;

      default:
        // Загальна санітизація
        sanitized = sanitized.slice(0, 200);
        sanitized = sanitized.replace(/[<>\"'&]/g, '');
    }

    return sanitized;
  } catch (error) {
    logger.error('Input sanitization error:', error);
    return '';
  }
}

/**
 * Валідація опцій команди
 * @param {Object} options - Опції команди
 * @param {Object} schema - Схема валідації
 * @returns {Object} - Результат валідації
 */
function validateCommandOptions(options, schema) {
  try {
    const result = {
      isValid: true,
      errors: [],
      sanitized: {},
    };

    for (const [key, config] of Object.entries(schema)) {
      const value = options[key];

      // Перевірка обов'язковості
      if (config.required && (value === undefined || value === null || value === '')) {
        result.isValid = false;
        result.errors.push(`Поле '${key}' є обов'язковим`);
        continue;
      }

      // Перевірка типу
      if (value !== undefined && value !== null) {
        if (config.type && typeof value !== config.type) {
          result.isValid = false;
          result.errors.push(`Поле '${key}' має неправильний тип`);
          continue;
        }

        // Санітизація
        if (config.sanitize) {
          result.sanitized[key] = sanitizeInput(value, config.sanitize);
        } else {
          result.sanitized[key] = value;
        }

        // Валідація довжини
        if (config.maxLength && result.sanitized[key].length > config.maxLength) {
          result.isValid = false;
          result.errors.push(`Поле '${key}' занадто довге (макс. ${config.maxLength} символів)`);
        }

        // Валідація мінімальної довжини
        if (config.minLength && result.sanitized[key].length < config.minLength) {
          result.isValid = false;
          result.errors.push(`Поле '${key}' занадто коротке (мін. ${config.minLength} символів)`);
        }
      }
    }

    return result;
  } catch (error) {
    logger.error('Command options validation error:', error);
    return {
      isValid: false,
      errors: ['Помилка валідації опцій'],
      sanitized: {},
    };
  }
}

/**
 * Логування подій безпеки
 * @param {string} event - Тип події
 * @param {Object} data - Дані події
 */
function logSecurityEvent(event, data) {
  try {
    const securityLog = {
      timestamp: new Date().toISOString(),
      event,
      data,
      level: 'security',
    };

    logger.info('Security event:', securityLog);

    // Тут можна додати додаткове логування в файл безпеки
    // або відправку сповіщень адміністратору
  } catch (error) {
    logger.error('Security event logging error:', error);
  }
}

/**
 * Отримання статистики безпеки
 * @returns {Object} - Статистика
 */
function getSecurityStats() {
  try {
    return {
      activeRateLimits: rateLimitCache.size,
      totalRoles: Object.keys(ROLES).length,
      rateLimitConfig: RATE_LIMITS,
    };
  } catch (error) {
    logger.error('Security stats error:', error);
    return {};
  }
}

module.exports = {
  ROLES,
  RATE_LIMITS,
  hasRole,
  checkPermission,
  checkRateLimit,
  cleanupRateLimitCache,
  sanitizeInput,
  validateCommandOptions,
  logSecurityEvent,
  getSecurityStats,
};
