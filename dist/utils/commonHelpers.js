"use strict";
/**
 * Загальні допоміжні функції для усунення дублювання коду
 * Централізовані утіліти для використання по всьому проекту
 * Версія 1.0.0 - Створено для рефакторингу
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.RetryUtils = exports.ErrorUtils = exports.DiscordUtils = exports.DataUtils = exports.ValidationUtils = exports.TimeUtils = exports.EmbedFactory = exports.TIME_CONSTANTS = exports.EMBED_LIMITS = exports.EMBED_COLORS = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("./logger"));
// Константи для стандартних значень
exports.EMBED_COLORS = {
    SUCCESS: 0x00FF00,
    ERROR: 0xFF0000,
    WARNING: 0xFFA500,
    INFO: 0x0099FF,
    PRIMARY: 0x00AE86,
    SECONDARY: 0x808080
};
exports.EMBED_LIMITS = {
    TITLE_MAX: 256,
    DESCRIPTION_MAX: 4096,
    FIELD_NAME_MAX: 256,
    FIELD_VALUE_MAX: 1024,
    FIELDS_MAX: 25,
    FOOTER_MAX: 2048,
    AUTHOR_MAX: 256
};
exports.TIME_CONSTANTS = {
    SECOND: 1000,
    MINUTE: 60 * 1000,
    HOUR: 60 * 60 * 1000,
    DAY: 24 * 60 * 60 * 1000,
    WEEK: 7 * 24 * 60 * 60 * 1000
};
/**
 * Створення стандартизованих embed повідомлень
 */
class EmbedFactory {
    /**
     * Базовий embed з загальними налаштуваннями
     */
    static createBase(title, description, color = exports.EMBED_COLORS.PRIMARY) {
        return new discord_js_1.EmbedBuilder()
            .setColor(color)
            .setTitle(this.truncateText(title, exports.EMBED_LIMITS.TITLE_MAX))
            .setDescription(this.truncateText(description, exports.EMBED_LIMITS.DESCRIPTION_MAX))
            .setTimestamp()
            .setFooter({
            text: this.defaultFooter,
            iconURL: this.defaultIconURL
        });
    }
    /**
     * Embed для успішних операцій
     */
    static success(title, description) {
        return this.createBase(`✅ ${title}`, description, exports.EMBED_COLORS.SUCCESS);
    }
    /**
     * Embed для помилок
     */
    static error(title, description, showSupport = true) {
        const embed = this.createBase(`❌ ${title}`, description, exports.EMBED_COLORS.ERROR);
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
    static warning(title, description) {
        return this.createBase(`⚠️ ${title}`, description, exports.EMBED_COLORS.WARNING);
    }
    /**
     * Embed для інформаційних повідомлень
     */
    static info(title, description) {
        return this.createBase(`ℹ️ ${title}`, description, exports.EMBED_COLORS.INFO);
    }
    /**
     * Embed для завантаження/очікування
     */
    static loading(title = 'Завантаження', description = 'Зачекайте...') {
        return this.createBase(`⏳ ${title}`, description, exports.EMBED_COLORS.WARNING);
    }
    /**
     * Embed з полями даних
     */
    static dataFields(title, description, fields) {
        const embed = this.createBase(title, description);
        fields.slice(0, exports.EMBED_LIMITS.FIELDS_MAX).forEach(field => {
            embed.addFields({
                name: this.truncateText(field.name, exports.EMBED_LIMITS.FIELD_NAME_MAX),
                value: this.truncateText(field.value, exports.EMBED_LIMITS.FIELD_VALUE_MAX),
                inline: field.inline || false
            });
        });
        return embed;
    }
    /**
     * Embed з пагінацією
     */
    static paginated(title, description, currentPage, totalPages, _data) {
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
    static truncateText(text, maxLength) {
        if (text.length <= maxLength)
            return text;
        return text.substring(0, maxLength - 3) + '...';
    }
}
exports.EmbedFactory = EmbedFactory;
EmbedFactory.defaultFooter = 'Discord AI Assistant Bot';
EmbedFactory.defaultIconURL = 'https://cdn.discordapp.com/embed/avatars/0.png';
/**
 * Утіліти для роботи з часом
 */
class TimeUtils {
    /**
     * Форматування мілісекунд у читабельний формат
     */
    static formatDuration(ms) {
        if (ms < exports.TIME_CONSTANTS.SECOND) {
            return `${ms}ms`;
        }
        else if (ms < exports.TIME_CONSTANTS.MINUTE) {
            return `${Math.round(ms / exports.TIME_CONSTANTS.SECOND)}s`;
        }
        else if (ms < exports.TIME_CONSTANTS.HOUR) {
            const minutes = Math.floor(ms / exports.TIME_CONSTANTS.MINUTE);
            const seconds = Math.round((ms % exports.TIME_CONSTANTS.MINUTE) / exports.TIME_CONSTANTS.SECOND);
            return seconds > 0 ? `${minutes}m ${seconds}s` : `${minutes}m`;
        }
        else if (ms < exports.TIME_CONSTANTS.DAY) {
            const hours = Math.floor(ms / exports.TIME_CONSTANTS.HOUR);
            const minutes = Math.round((ms % exports.TIME_CONSTANTS.HOUR) / exports.TIME_CONSTANTS.MINUTE);
            return minutes > 0 ? `${hours}h ${minutes}m` : `${hours}h`;
        }
        else {
            const days = Math.floor(ms / exports.TIME_CONSTANTS.DAY);
            const hours = Math.round((ms % exports.TIME_CONSTANTS.DAY) / exports.TIME_CONSTANTS.HOUR);
            return hours > 0 ? `${days}d ${hours}h` : `${days}d`;
        }
    }
    /**
     * Форматування timestamp у читабельний формат
     */
    static formatTimestamp(timestamp, format = 'relative') {
        const date = new Date(timestamp);
        const now = new Date();
        const diff = now.getTime() - timestamp;
        switch (format) {
            case 'relative':
                if (diff < exports.TIME_CONSTANTS.MINUTE) {
                    return 'щойно';
                }
                else if (diff < exports.TIME_CONSTANTS.HOUR) {
                    const minutes = Math.floor(diff / exports.TIME_CONSTANTS.MINUTE);
                    return `${minutes} хв тому`;
                }
                else if (diff < exports.TIME_CONSTANTS.DAY) {
                    const hours = Math.floor(diff / exports.TIME_CONSTANTS.HOUR);
                    return `${hours} год тому`;
                }
                else if (diff < exports.TIME_CONSTANTS.WEEK) {
                    const days = Math.floor(diff / exports.TIME_CONSTANTS.DAY);
                    return `${days} дн тому`;
                }
                else {
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
    static isWithinRange(timestamp, rangeMs) {
        return Date.now() - timestamp <= rangeMs;
    }
    /**
     * Отримання часу до наступного інтервалу
     */
    static getTimeUntilNextInterval(intervalMs) {
        const now = Date.now();
        const nextInterval = Math.ceil(now / intervalMs) * intervalMs;
        return nextInterval - now;
    }
}
exports.TimeUtils = TimeUtils;
/**
 * Утіліти для валідації та форматування
 */
class ValidationUtils {
    /**
     * Перевірка email адреси
     */
    static isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    }
    /**
     * Перевірка URL
     */
    static isValidURL(url) {
        try {
            new URL(url);
            return true;
        }
        catch {
            return false;
        }
    }
    /**
     * Перевірка Discord ID
     */
    static isValidDiscordId(id) {
        const idRegex = /^\d{17,19}$/;
        return idRegex.test(id);
    }
    /**
     * Санітизація тексту для безпеки
     */
    static sanitizeText(text) {
        return text
            .replace(/[<>]/g, '') // Видаляємо потенційно небезпечні символи
            .replace(/@(everyone|here)/gi, '@\u200b$1') // Нейтралізуємо mass mentions
            .trim();
    }
    /**
     * Валідація числового діапазону
     */
    static isInRange(value, min, max) {
        return value >= min && value <= max;
    }
    /**
     * Перевірка довжини рядка
     */
    static isValidLength(text, minLength, maxLength) {
        return text.length >= minLength && text.length <= maxLength;
    }
}
exports.ValidationUtils = ValidationUtils;
/**
 * Утіліти для роботи з файлами та даними
 */
class DataUtils {
    /**
     * Форматування розміру файлу
     */
    static formatFileSize(bytes) {
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
    static formatNumber(num, locale = 'uk-UA') {
        return new Intl.NumberFormat(locale).format(num);
    }
    /**
     * Форматування відсотків
     */
    static formatPercentage(value, total, decimals = 1) {
        const percentage = total > 0 ? (value / total) * 100 : 0;
        return `${percentage.toFixed(decimals)}%`;
    }
    /**
     * Глибоке клонування об'єкта
     */
    static deepClone(obj) {
        if (obj === null || typeof obj !== 'object')
            return obj;
        if (obj instanceof Date)
            return new Date(obj.getTime());
        if (obj instanceof Array)
            return obj.map(item => this.deepClone(item));
        if (typeof obj === 'object') {
            const copy = {};
            Object.keys(obj).forEach(key => {
                copy[key] = this.deepClone(obj[key]);
            });
            return copy;
        }
        return obj;
    }
    /**
     * Безпечне парсення JSON
     */
    static safeJsonParse(json, defaultValue) {
        try {
            return JSON.parse(json);
        }
        catch {
            return defaultValue;
        }
    }
    /**
     * Групування масиву за ключем
     */
    static groupBy(array, keyFn) {
        return array.reduce((groups, item) => {
            const key = keyFn(item);
            if (!groups[key]) {
                groups[key] = [];
            }
            groups[key].push(item);
            return groups;
        }, {});
    }
    /**
     * Пагінація масиву
     */
    static paginate(array, page, pageSize) {
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
exports.DataUtils = DataUtils;
/**
 * Утіліти для Discord специфічних операцій
 */
class DiscordUtils {
    /**
     * Безпечне відправлення відповіді
     */
    static async safeReply(interaction, content) {
        try {
            if (interaction.replied || interaction.deferred) {
                await interaction.followUp(content);
            }
            else {
                await interaction.reply(content);
            }
            return true;
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка відправлення відповіді Discord: ${error instanceof Error ? error.message : String(error)}`);
            return false;
        }
    }
    /**
     * Створення кнопок пагінації
     */
    static createPaginationButtons(currentPage, totalPages, disabled = false) {
        const row = new discord_js_1.ActionRowBuilder();
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_first')
            .setEmoji('⏪')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(disabled || currentPage <= 1), new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_prev')
            .setEmoji('◀️')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(disabled || currentPage <= 1), new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_info')
            .setLabel(`${currentPage}/${totalPages}`)
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(true), new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_next')
            .setEmoji('▶️')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(disabled || currentPage >= totalPages), new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_last')
            .setEmoji('⏩')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(disabled || currentPage >= totalPages));
        return row;
    }
    /**
     * Отримання відображуваного імені користувача
     */
    static getUserDisplayName(user, guild) {
        if (guild) {
            const member = guild.members.cache.get(user.id);
            return member?.displayName || user.displayName || user.username;
        }
        return user.displayName || user.username;
    }
    /**
     * Форматування згадки користувача
     */
    static formatUserMention(userId) {
        return `<@${userId}>`;
    }
    /**
     * Форматування згадки каналу
     */
    static formatChannelMention(channelId) {
        return `<#${channelId}>`;
    }
    /**
     * Форматування згадки ролі
     */
    static formatRoleMention(roleId) {
        return `<@&${roleId}>`;
    }
}
exports.DiscordUtils = DiscordUtils;
/**
 * Утіліти для обробки помилок
 */
class ErrorUtils {
    /**
     * Логування помилки з контекстом
     */
    static logError(error, context) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        const stack = error instanceof Error ? error.stack : undefined;
        logger_1.default.error(`❌ Помилка в операції: ${context.operation}`, {
            error: errorMessage,
            stack,
            userId: context.userId,
            command: context.commandName,
            ...context.additionalData
        });
    }
    /**
     * Створення стандартного embed для помилки
     */
    static createErrorEmbed(error, showDetails = false) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        if (showDetails) {
            return EmbedFactory.error('Помилка виконання', `Деталі: ${errorMessage.substring(0, 1000)}`);
        }
        else {
            return EmbedFactory.error('Помилка виконання', 'Виникла неочікувана помилка. Спробуйте пізніше.');
        }
    }
    /**
     * Перевірка чи помилка критична
     */
    static isCriticalError(error) {
        if (error instanceof Error) {
            const criticalPatterns = [
                'ECONNREFUSED',
                'ENOTFOUND',
                'Database',
                'Permission denied',
                'Out of memory'
            ];
            return criticalPatterns.some(pattern => error.message.includes(pattern) || error.stack?.includes(pattern));
        }
        return false;
    }
}
exports.ErrorUtils = ErrorUtils;
/**
 * Утіліти для retry логіки
 */
class RetryUtils {
    /**
     * Виконання функції з повторними спробами
     */
    static async withRetry(fn, options = {}) {
        const { maxAttempts = 3, delay = 1000, backoff = 'exponential', shouldRetry = () => true } = options;
        let lastError;
        for (let attempt = 1; attempt <= maxAttempts; attempt++) {
            try {
                return await fn();
            }
            catch (error) {
                lastError = error;
                if (attempt === maxAttempts || !shouldRetry(error)) {
                    throw error;
                }
                const waitTime = backoff === 'exponential'
                    ? delay * Math.pow(2, attempt - 1)
                    : delay * attempt;
                logger_1.default.warn(`🔄 Повторна спроба ${attempt}/${maxAttempts} через ${waitTime}ms`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
            }
        }
        throw lastError;
    }
}
exports.RetryUtils = RetryUtils;
//# sourceMappingURL=commonHelpers.js.map