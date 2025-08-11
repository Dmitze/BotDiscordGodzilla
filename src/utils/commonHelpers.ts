/**
 * Загальні допоміжні функції для усунення дублювання коду
 * Централізовані утіліти для використання по всьому проекту
 * Версія 1.0.0 - Створено для рефакторингу
 */

import {
  EmbedBuilder,
  ChatInputCommandInteraction,
  User,
  Guild,
  ButtonBuilder,
  ButtonStyle,
  ActionRowBuilder,
} from 'discord.js';
import logger from './logger';

// Константи для стандартних значень
export const EMBED_COLORS = {
  SUCCESS: 0x00FF00,
  ERROR: 0xFF0000,
  WARNING: 0xFFA500,
  INFO: 0x0099FF,
  PRIMARY: 0x00AE86,
  SECONDARY: 0x808080
} as const;

export const EMBED_LIMITS = {
  TITLE_MAX: 256,
  DESCRIPTION_MAX: 4096,
  FIELD_NAME_MAX: 256,
  FIELD_VALUE_MAX: 1024,
  FIELDS_MAX: 25,
  FOOTER_MAX: 2048,
  AUTHOR_MAX: 256
} as const;

export const TIME_CONSTANTS = {
  SECOND: 1000,
  MINUTE: 60 * 1000,
  HOUR: 60 * 60 * 1000,
  DAY: 24 * 60 * 60 * 1000,
  WEEK: 7 * 24 * 60 * 60 * 1000
} as const;

/**
 * Створення стандартизованих embed повідомлень
 */
export class EmbedFactory {
  private static readonly defaultFooter = 'Discord AI Assistant Bot';
  private static readonly defaultIconURL = 'https://cdn.discordapp.com/embed/avatars/0.png';

  /**
   * Базовий embed з загальними налаштуваннями
   */
  static createBase(title: string, description: string, color: number = EMBED_COLORS.PRIMARY): EmbedBuilder {
    return new EmbedBuilder()
      .setColor(color)
      .setTitle(this.truncateText(title, EMBED_LIMITS.TITLE_MAX))
      .setDescription(this.truncateText(description, EMBED_LIMITS.DESCRIPTION_MAX))
      .setTimestamp()
      .setFooter({ 
        text: this.defaultFooter, 
        iconURL: this.defaultIconURL 
      });
  }

  /**
   * Embed для успішних операцій
   */
  static success(title: string, description: string): EmbedBuilder {
    return this.createBase(`✅ ${title}`, description, EMBED_COLORS.SUCCESS);
  }

  /**
   * Embed для помилок
   */
  static error(title: string, description: string, showSupport: boolean = true): EmbedBuilder {
    const embed = this.createBase(`❌ ${title}`, description, EMBED_COLORS.ERROR);
    
    if (showSupport) {
      embed.addFields({
        name: '📞 Потрібна допомога?',
        value: 'Зверніться до адміністрації сервера або перевірте документацію.',
        inline: false
      });
    }
    
    return embed;
  }

  /**
   * Embed для попереджень
   */
  static warning(title: string, description: string): EmbedBuilder {
    return this.createBase(`⚠️ ${title}`, description, EMBED_COLORS.WARNING);
  }

  /**
   * Embed для інформаційних повідомлень
   */
  static info(title: string, description: string): EmbedBuilder {
    return this.createBase(`ℹ️ ${title}`, description, EMBED_COLORS.INFO);
  }

  /**
   * Embed для завантаження/очікування
   */
  static loading(title: string = 'Завантаження', description: string = 'Зачекайте...'): EmbedBuilder {
    return this.createBase(`⏳ ${title}`, description, EMBED_COLORS.WARNING);
  }

  /**
   * Embed з полями даних
   */
  static dataFields(
    title: string, 
    description: string, 
    fields: Array<{ name: string; value: string; inline?: boolean }>
  ): EmbedBuilder {
    const embed = this.createBase(title, description);
    
    fields.slice(0, EMBED_LIMITS.FIELDS_MAX).forEach(field => {
      embed.addFields({
        name: this.truncateText(field.name, EMBED_LIMITS.FIELD_NAME_MAX),
        value: this.truncateText(field.value, EMBED_LIMITS.FIELD_VALUE_MAX),
        inline: field.inline || false
      });
    });
    
    return embed;
  }

  /**
   * Embed з пагінацією
   */
  static paginated(
    title: string,
    description: string,
    currentPage: number,
    totalPages: number,
    _data?: any
  ): EmbedBuilder {
    const embed = this.createBase(title, description);
    
    embed.setFooter({
      text: `${this.defaultFooter} • Сторінка ${currentPage}/${totalPages}`,
      iconURL: this.defaultIconURL
    });
    
    return embed;
  }

  /**
   * Обрізання тексту до максимальної довжини
   */
  private static truncateText(text: string, maxLength: number): string {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
  }
}

/**
 * Утіліти для роботи з часом
 */
export class TimeUtils {
  /**
   * Форматування мілісекунд у читабельний формат
   */
  static formatDuration(ms: number): string {
    if (ms < TIME_CONSTANTS.SECOND) {
      return `${ms}ms`;
    } else if (ms < TIME_CONSTANTS.MINUTE) {
      return `${Math.round(ms / TIME_CONSTANTS.SECOND)}s`;
    } else if (ms < TIME_CONSTANTS.HOUR) {
      const minutes = Math.floor(ms / TIME_CONSTANTS.MINUTE);
      const seconds = Math.round((ms % TIME_CONSTANTS.MINUTE) / TIME_CONSTANTS.SECOND);
      return seconds > 0 ? `${minutes}m ${seconds}s` : `${minutes}m`;
    } else if (ms < TIME_CONSTANTS.DAY) {
      const hours = Math.floor(ms / TIME_CONSTANTS.HOUR);
      const minutes = Math.round((ms % TIME_CONSTANTS.HOUR) / TIME_CONSTANTS.MINUTE);
      return minutes > 0 ? `${hours}h ${minutes}m` : `${hours}h`;
    } else {
      const days = Math.floor(ms / TIME_CONSTANTS.DAY);
      const hours = Math.round((ms % TIME_CONSTANTS.DAY) / TIME_CONSTANTS.HOUR);
      return hours > 0 ? `${days}d ${hours}h` : `${days}d`;
    }
  }

  /**
   * Форматування timestamp у читабельний формат
   */
  static formatTimestamp(timestamp: number, format: 'relative' | 'absolute' | 'datetime' = 'relative'): string {
    const date = new Date(timestamp);
    const now = new Date();
    const diff = now.getTime() - timestamp;

    switch (format) {
      case 'relative':
        if (diff < TIME_CONSTANTS.MINUTE) {
          return 'щойно';
        } else if (diff < TIME_CONSTANTS.HOUR) {
          const minutes = Math.floor(diff / TIME_CONSTANTS.MINUTE);
          return `${minutes} хв тому`;
        } else if (diff < TIME_CONSTANTS.DAY) {
          const hours = Math.floor(diff / TIME_CONSTANTS.HOUR);
          return `${hours} год тому`;
        } else if (diff < TIME_CONSTANTS.WEEK) {
          const days = Math.floor(diff / TIME_CONSTANTS.DAY);
          return `${days} дн тому`;
        } else {
          return date.toLocaleDateString('uk-UA');
        }

      case 'absolute':
        return date.toLocaleDateString('uk-UA');

      case 'datetime':
        return date.toLocaleString('uk-UA');

      default:
        return date.toISOString();
    }
  }

  /**
   * Перевірка чи час в межах діапазону
   */
  static isWithinRange(timestamp: number, rangeMs: number): boolean {
    return Date.now() - timestamp <= rangeMs;
  }

  /**
   * Отримання часу до наступного інтервалу
   */
  static getTimeUntilNextInterval(intervalMs: number): number {
    const now = Date.now();
    const nextInterval = Math.ceil(now / intervalMs) * intervalMs;
    return nextInterval - now;
  }
}

/**
 * Утіліти для валідації та форматування
 */
export class ValidationUtils {
  /**
   * Перевірка email адреси
   */
  static isValidEmail(email: string): boolean {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  /**
   * Перевірка URL
   */
  static isValidURL(url: string): boolean {
    try {
      new URL(url);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Перевірка Discord ID
   */
  static isValidDiscordId(id: string): boolean {
    const idRegex = /^\d{17,19}$/;
    return idRegex.test(id);
  }

  /**
   * Санітизація тексту для безпеки
   */
  static sanitizeText(text: string): string {
    return text
      .replace(/[<>]/g, '') // Видаляємо потенційно небезпечні символи
      .replace(/@(everyone|here)/gi, '@\u200b$1') // Нейтралізуємо mass mentions
      .trim();
  }

  /**
   * Валідація числового діапазону
   */
  static isInRange(value: number, min: number, max: number): boolean {
    return value >= min && value <= max;
  }

  /**
   * Перевірка довжини рядка
   */
  static isValidLength(text: string, minLength: number, maxLength: number): boolean {
    return text.length >= minLength && text.length <= maxLength;
  }
}

/**
 * Утіліти для роботи з файлами та даними
 */
export class DataUtils {
  /**
   * Форматування розміру файлу
   */
  static formatFileSize(bytes: number): string {
    const units = ['B', 'KB', 'MB', 'GB', 'TB'];
    let size = bytes;
    let unitIndex = 0;

    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }

    return `${size.toFixed(1)} ${units[unitIndex]}`;
  }

  /**
   * Форматування чисел з роздільниками
   */
  static formatNumber(num: number, locale: string = 'uk-UA'): string {
    return new Intl.NumberFormat(locale).format(num);
  }

  /**
   * Форматування відсотків
   */
  static formatPercentage(value: number, total: number, decimals: number = 1): string {
    const percentage = total > 0 ? (value / total) * 100 : 0;
    return `${percentage.toFixed(decimals)}%`;
  }

  /**
   * Глибоке клонування об'єкта
   */
  static deepClone<T>(obj: T): T {
    if (obj === null || typeof obj !== 'object') return obj;
    if (obj instanceof Date) return new Date(obj.getTime()) as any;
    if (obj instanceof Array) return obj.map(item => this.deepClone(item)) as any;
    if (typeof obj === 'object') {
      const copy: any = {};
      Object.keys(obj).forEach(key => {
        copy[key] = this.deepClone((obj as any)[key]);
      });
      return copy;
    }
    return obj;
  }

  /**
   * Безпечне парсення JSON
   */
  static safeJsonParse<T>(json: string, defaultValue: T): T {
    try {
      return JSON.parse(json);
    } catch {
      return defaultValue;
    }
  }

  /**
   * Групування масиву за ключем
   */
  static groupBy<T>(array: T[], keyFn: (item: T) => string): Record<string, T[]> {
    return array.reduce((groups, item) => {
      const key = keyFn(item);
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {} as Record<string, T[]>);
  }

  /**
   * Пагінація масиву
   */
  static paginate<T>(array: T[], page: number, pageSize: number): {
    items: T[];
    totalPages: number;
    currentPage: number;
    hasNext: boolean;
    hasPrev: boolean;
  } {
    const totalPages = Math.ceil(array.length / pageSize);
    const currentPage = Math.max(1, Math.min(page, totalPages));
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const items = array.slice(startIndex, endIndex);

    return {
      items,
      totalPages,
      currentPage,
      hasNext: currentPage < totalPages,
      hasPrev: currentPage > 1
    };
  }
}

/**
 * Утіліти для Discord специфічних операцій
 */
export class DiscordUtils {
  /**
   * Безпечне відправлення відповіді
   */
  static async safeReply(
    interaction: ChatInputCommandInteraction,
    content: { embeds?: EmbedBuilder[]; content?: string; ephemeral?: boolean }
  ): Promise<boolean> {
    try {
      if (interaction.replied || interaction.deferred) {
        await interaction.followUp(content);
      } else {
        await interaction.reply(content);
      }
      return true;
    } catch (error) {
      logger.error(`❌ Помилка відправлення відповіді Discord: ${error instanceof Error ? error.message : String(error)}`);
      return false;
    }
  }

  /**
   * Створення кнопок пагінації
   */
  static createPaginationButtons(
    currentPage: number,
    totalPages: number,
    disabled: boolean = false
  ): ActionRowBuilder<ButtonBuilder> {
    const row = new ActionRowBuilder<ButtonBuilder>();

    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_first')
        .setEmoji('⏪')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(disabled || currentPage <= 1),
      
      new ButtonBuilder()
        .setCustomId('pagination_prev')
        .setEmoji('◀️')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(disabled || currentPage <= 1),
      
      new ButtonBuilder()
        .setCustomId('pagination_info')
        .setLabel(`${currentPage}/${totalPages}`)
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(true),
      
      new ButtonBuilder()
        .setCustomId('pagination_next')
        .setEmoji('▶️')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(disabled || currentPage >= totalPages),
      
      new ButtonBuilder()
        .setCustomId('pagination_last')
        .setEmoji('⏩')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(disabled || currentPage >= totalPages)
    );

    return row;
  }

  /**
   * Отримання відображуваного імені користувача
   */
  static getUserDisplayName(user: User, guild?: Guild): string {
    if (guild) {
      const member = guild.members.cache.get(user.id);
      return member?.displayName || user.displayName || user.username;
    }
    return user.displayName || user.username;
  }

  /**
   * Форматування згадки користувача
   */
  static formatUserMention(userId: string): string {
    return `<@${userId}>`;
  }

  /**
   * Форматування згадки каналу
   */
  static formatChannelMention(channelId: string): string {
    return `<#${channelId}>`;
  }

  /**
   * Форматування згадки ролі
   */
  static formatRoleMention(roleId: string): string {
    return `<@&${roleId}>`;
  }
}

/**
 * Утіліти для обробки помилок
 */
export class ErrorUtils {
  /**
   * Логування помилки з контекстом
   */
  static logError(
    error: unknown,
    context: {
      operation: string;
      userId?: string;
      commandName?: string;
      additionalData?: Record<string, unknown>;
    }
  ): void {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stack = error instanceof Error ? error.stack : undefined;

    logger.error(`❌ Помилка в операції: ${context.operation}`, {
      type: 'utility',
      component: 'ErrorUtils.logError',
      operation: context.operation,
      error: errorMessage,
      stack,
      ...(context.userId ? { userId: context.userId } : {}),
      ...(context.commandName ? { command: context.commandName } : {}),
      ...context.additionalData,
    });
  }

  /**
   * Створення стандартного embed для помилки
   */
  static createErrorEmbed(
    error: unknown,
    showDetails: boolean = false
  ): EmbedBuilder {
    const errorMessage = error instanceof Error ? error.message : String(error);
    
    if (showDetails) {
      return EmbedFactory.error(
        'Помилка виконання',
        `Деталі: ${errorMessage.substring(0, 1000)}`
      );
    } else {
      return EmbedFactory.error(
        'Помилка виконання',
        'Виникла неочікувана помилка. Спробуйте пізніше.'
      );
    }
  }

  /**
   * Перевірка чи помилка критична
   */
  static isCriticalError(error: unknown): boolean {
    if (error instanceof Error) {
      const criticalPatterns = [
        'ECONNREFUSED',
        'ENOTFOUND',
        'Database',
        'Permission denied',
        'Out of memory'
      ];
      
      return criticalPatterns.some(pattern => 
        error.message.includes(pattern) || error.stack?.includes(pattern)
      );
    }
    return false;
  }
}

/**
 * Утіліти для retry логіки
 */
export class RetryUtils {
  /**
   * Виконання функції з повторними спробами
   */
  static async withRetry<T>(
    fn: () => Promise<T>,
    options: {
      maxAttempts?: number;
      delay?: number;
      backoff?: 'linear' | 'exponential';
      shouldRetry?: (error: unknown) => boolean;
    } = {}
  ): Promise<T> {
    const {
      maxAttempts = 3,
      delay = 1000,
      backoff = 'exponential',
      shouldRetry = () => true
    } = options;

    let lastError: unknown;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error;

        if (attempt === maxAttempts || !shouldRetry(error)) {
          throw error;
        }

        const waitTime = backoff === 'exponential' 
          ? delay * Math.pow(2, attempt - 1)
          : delay * attempt;

        logger.warn(`🔄 Повторна спроба ${attempt}/${maxAttempts} через ${waitTime}ms`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    }

    throw lastError;
  }
}