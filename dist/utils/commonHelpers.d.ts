/**
 * Загальні допоміжні функції для усунення дублювання коду
 * Централізовані утіліти для використання по всьому проекту
 * Версія 1.0.0 - Створено для рефакторингу
 */
import { EmbedBuilder, ChatInputCommandInteraction, User, Guild, ButtonBuilder, ActionRowBuilder } from 'discord.js';
export declare const EMBED_COLORS: {
    readonly SUCCESS: 65280;
    readonly ERROR: 16711680;
    readonly WARNING: 16753920;
    readonly INFO: 39423;
    readonly PRIMARY: 44678;
    readonly SECONDARY: 8421504;
};
export declare const EMBED_LIMITS: {
    readonly TITLE_MAX: 256;
    readonly DESCRIPTION_MAX: 4096;
    readonly FIELD_NAME_MAX: 256;
    readonly FIELD_VALUE_MAX: 1024;
    readonly FIELDS_MAX: 25;
    readonly FOOTER_MAX: 2048;
    readonly AUTHOR_MAX: 256;
};
export declare const TIME_CONSTANTS: {
    readonly SECOND: 1000;
    readonly MINUTE: number;
    readonly HOUR: number;
    readonly DAY: number;
    readonly WEEK: number;
};
/**
 * Створення стандартизованих embed повідомлень
 */
export declare class EmbedFactory {
    private static readonly defaultFooter;
    private static readonly defaultIconURL;
    /**
     * Базовий embed з загальними налаштуваннями
     */
    static createBase(title: string, description: string, color?: number): EmbedBuilder;
    /**
     * Embed для успішних операцій
     */
    static success(title: string, description: string): EmbedBuilder;
    /**
     * Embed для помилок
     */
    static error(title: string, description: string, showSupport?: boolean): EmbedBuilder;
    /**
     * Embed для попереджень
     */
    static warning(title: string, description: string): EmbedBuilder;
    /**
     * Embed для інформаційних повідомлень
     */
    static info(title: string, description: string): EmbedBuilder;
    /**
     * Embed для завантаження/очікування
     */
    static loading(title?: string, description?: string): EmbedBuilder;
    /**
     * Embed з полями даних
     */
    static dataFields(title: string, description: string, fields: Array<{
        name: string;
        value: string;
        inline?: boolean;
    }>): EmbedBuilder;
    /**
     * Embed з пагінацією
     */
    static paginated(title: string, description: string, currentPage: number, totalPages: number, _data?: any): EmbedBuilder;
    /**
     * Обрізання тексту до максимальної довжини
     */
    private static truncateText;
}
/**
 * Утіліти для роботи з часом
 */
export declare class TimeUtils {
    /**
     * Форматування мілісекунд у читабельний формат
     */
    static formatDuration(ms: number): string;
    /**
     * Форматування timestamp у читабельний формат
     */
    static formatTimestamp(timestamp: number, format?: 'relative' | 'absolute' | 'datetime'): string;
    /**
     * Перевірка чи час в межах діапазону
     */
    static isWithinRange(timestamp: number, rangeMs: number): boolean;
    /**
     * Отримання часу до наступного інтервалу
     */
    static getTimeUntilNextInterval(intervalMs: number): number;
}
/**
 * Утіліти для валідації та форматування
 */
export declare class ValidationUtils {
    /**
     * Перевірка email адреси
     */
    static isValidEmail(email: string): boolean;
    /**
     * Перевірка URL
     */
    static isValidURL(url: string): boolean;
    /**
     * Перевірка Discord ID
     */
    static isValidDiscordId(id: string): boolean;
    /**
     * Санітизація тексту для безпеки
     */
    static sanitizeText(text: string): string;
    /**
     * Валідація числового діапазону
     */
    static isInRange(value: number, min: number, max: number): boolean;
    /**
     * Перевірка довжини рядка
     */
    static isValidLength(text: string, minLength: number, maxLength: number): boolean;
}
/**
 * Утіліти для роботи з файлами та даними
 */
export declare class DataUtils {
    /**
     * Форматування розміру файлу
     */
    static formatFileSize(bytes: number): string;
    /**
     * Форматування чисел з роздільниками
     */
    static formatNumber(num: number, locale?: string): string;
    /**
     * Форматування відсотків
     */
    static formatPercentage(value: number, total: number, decimals?: number): string;
    /**
     * Глибоке клонування об'єкта
     */
    static deepClone<T>(obj: T): T;
    /**
     * Безпечне парсення JSON
     */
    static safeJsonParse<T>(json: string, defaultValue: T): T;
    /**
     * Групування масиву за ключем
     */
    static groupBy<T>(array: T[], keyFn: (item: T) => string): Record<string, T[]>;
    /**
     * Пагінація масиву
     */
    static paginate<T>(array: T[], page: number, pageSize: number): {
        items: T[];
        totalPages: number;
        currentPage: number;
        hasNext: boolean;
        hasPrev: boolean;
    };
}
/**
 * Утіліти для Discord специфічних операцій
 */
export declare class DiscordUtils {
    /**
     * Безпечне відправлення відповіді
     */
    static safeReply(interaction: ChatInputCommandInteraction, content: {
        embeds?: EmbedBuilder[];
        content?: string;
        ephemeral?: boolean;
    }): Promise<boolean>;
    /**
     * Створення кнопок пагінації
     */
    static createPaginationButtons(currentPage: number, totalPages: number, disabled?: boolean): ActionRowBuilder<ButtonBuilder>;
    /**
     * Отримання відображуваного імені користувача
     */
    static getUserDisplayName(user: User, guild?: Guild): string;
    /**
     * Форматування згадки користувача
     */
    static formatUserMention(userId: string): string;
    /**
     * Форматування згадки каналу
     */
    static formatChannelMention(channelId: string): string;
    /**
     * Форматування згадки ролі
     */
    static formatRoleMention(roleId: string): string;
}
/**
 * Утіліти для обробки помилок
 */
export declare class ErrorUtils {
    /**
     * Логування помилки з контекстом
     */
    static logError(error: unknown, context: {
        operation: string;
        userId?: string;
        commandName?: string;
        additionalData?: Record<string, unknown>;
    }): void;
    /**
     * Створення стандартного embed для помилки
     */
    static createErrorEmbed(error: unknown, showDetails?: boolean): EmbedBuilder;
    /**
     * Перевірка чи помилка критична
     */
    static isCriticalError(error: unknown): boolean;
}
/**
 * Утіліти для retry логіки
 */
export declare class RetryUtils {
    /**
     * Виконання функції з повторними спробами
     */
    static withRetry<T>(fn: () => Promise<T>, options?: {
        maxAttempts?: number;
        delay?: number;
        backoff?: 'linear' | 'exponential';
        shouldRetry?: (error: unknown) => boolean;
    }): Promise<T>;
}
//# sourceMappingURL=commonHelpers.d.ts.map