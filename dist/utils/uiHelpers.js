"use strict";
/**
 * Модуль для покращеного UI/UX
 * Включає красиві embed повідомлення, інтерактивні кнопки та прогрес-бари
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.EMOJIS = exports.COLORS = exports.UIHelper = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("./logger"));
// Кольори для різних типів повідомлень
const COLORS = {
    SUCCESS: 0x00FF00, // Зелений
    ERROR: 0xFF0000, // Червоний
    WARNING: 0xFFA500, // Помаранчевий
    INFO: 0x0099FF, // Синій
    AI: 0x9B59B6, // Фіолетовий
    SEARCH: 0x3498DB, // Голубий
    FILES: 0xE67E22, // Оранжевий
    EXPORT: 0x27AE60 // Темно-зелений
};
exports.COLORS = COLORS;
// Емодзі для різних дій
const EMOJIS = {
    SUCCESS: '✅',
    ERROR: '❌',
    WARNING: '⚠️',
    INFO: 'ℹ️',
    AI: '🤖',
    SEARCH: '🔍',
    FILES: '📁',
    EXPORT: '📤',
    LOADING: '⏳',
    DONE: '🎉',
    HELP: '❓',
    SETTINGS: '⚙️',
    STATS: '📊',
    SECURITY: '🔒'
};
exports.EMOJIS = EMOJIS;
/**
 * Клас для створення покращених UI елементів
 */
class UIHelper {
    /**
     * Створення базового embed
     */
    static createBaseEmbed(title, description, color = COLORS.INFO) {
        return new discord_js_1.EmbedBuilder()
            .setColor(color)
            .setTitle(title)
            .setDescription(description)
            .setTimestamp()
            .setFooter({
            text: 'Discord AI Assistant Bot',
            iconURL: 'https://cdn.discordapp.com/emojis/1234567890.png'
        });
    }
    /**
     * Створення embed для результатів пошуку
     */
    static createSearchResultsEmbed(results, query, page = 0, totalPages = 1) {
        const embed = this.createBaseEmbed(`${EMOJIS.SEARCH} Результати пошуку`, `**Запит:** \`${query}\`\n**Знайдено:** ${results.length} записів`, COLORS.SEARCH);
        // Додавання результатів
        results.slice(0, 10).forEach((result, index) => {
            const rowNumber = page * 10 + index + 1;
            embed.addFields({
                name: `${EMOJIS.INFO} Запис ${rowNumber}`,
                value: this.formatSearchResult(result),
                inline: false
            });
        });
        // Додавання інформації про сторінки
        if (totalPages > 1) {
            embed.addFields({
                name: `${EMOJIS.INFO} Навігація`,
                value: `Сторінка ${page + 1} з ${totalPages}`,
                inline: true
            });
        }
        return embed;
    }
    /**
     * Форматування результату пошуку
     */
    static formatSearchResult(result) {
        if (Array.isArray(result)) {
            return result.map((item, index) => `${index + 1}. ${item}`).join('\n');
        }
        if (typeof result === 'object') {
            return Object.entries(result)
                .map(([key, value]) => `**${key}:** ${value}`)
                .join('\n');
        }
        return result.toString();
    }
    /**
     * Створення embed для AI відповіді
     */
    static createAIResponseEmbed(query, response, confidence = 1.0) {
        const embed = this.createBaseEmbed(`${EMOJIS.AI} AI-асистент`, `**Ваш запит:** ${query}`, COLORS.AI);
        // Додавання відповіді
        embed.addFields({
            name: `${EMOJIS.INFO} Відповідь`,
            value: response.length > 1024 ? response.substring(0, 1021) + '...' : response,
            inline: false
        });
        // Додавання впевненості
        if (confidence < 0.7) {
            embed.addFields({
                name: `${EMOJIS.WARNING} Впевненість`,
                value: `${Math.round(confidence * 100)}% - низька впевненість`,
                inline: true
            });
        }
        return embed;
    }
    /**
     * Створення embed для роботи з файлами
     */
    static createFileEmbed(action, fileName, content = null, metadata = null) {
        const titles = {
            'пошук': `${EMOJIS.FILES} Пошук файлів`,
            'читати': `${EMOJIS.FILES} Читання файлу`,
            'аналіз': `${EMOJIS.AI} AI-аналіз файлу`,
            'звіт': `${EMOJIS.EXPORT} Створення звіту`
        };
        const embed = this.createBaseEmbed(titles[action] || `${EMOJIS.FILES} Робота з файлами`, `**Файл:** ${fileName}`, COLORS.FILES);
        if (content) {
            embed.addFields({
                name: `${EMOJIS.INFO} Вміст`,
                value: content.length > 1024 ? content.substring(0, 1021) + '...' : content,
                inline: false
            });
        }
        if (metadata) {
            embed.addFields({
                name: `${EMOJIS.INFO} Метадані`,
                value: Object.entries(metadata)
                    .map(([key, value]) => `**${key}:** ${value}`)
                    .join('\n'),
                inline: true
            });
        }
        return embed;
    }
    /**
     * Створення embed для експорту
     */
    static createExportEmbed(format, recordCount, fileName) {
        const embed = this.createBaseEmbed(`${EMOJIS.EXPORT} Експорт завершено`, `**Формат:** ${format.toUpperCase()}\n**Записів:** ${recordCount}`, COLORS.EXPORT);
        embed.addFields({
            name: `${EMOJIS.INFO} Файл`,
            value: fileName,
            inline: true
        });
        return embed;
    }
    /**
     * Створення embed для помилок
     */
    static createErrorEmbed(error, context = '') {
        const embed = this.createBaseEmbed(`${EMOJIS.ERROR} Помилка`, context || 'Сталася помилка при виконанні команди', COLORS.ERROR);
        embed.addFields({
            name: `${EMOJIS.INFO} Деталі`,
            value: typeof error === 'string' ? error : error.message || error.toString(),
            inline: false
        });
        return embed;
    }
    /**
     * Створення embed для успіху
     */
    static createSuccessEmbed(message, details = null) {
        const embed = this.createBaseEmbed(`${EMOJIS.SUCCESS} Успішно`, message, COLORS.SUCCESS);
        if (details) {
            embed.addFields({
                name: `${EMOJIS.INFO} Деталі`,
                value: details,
                inline: false
            });
        }
        return embed;
    }
    /**
     * Створення кнопок для навігації
     */
    static createNavigationButtons(currentPage, totalPages, customIds = {}) {
        const row = new discord_js_1.ActionRowBuilder();
        // Кнопка "Попередня"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(customIds['prev'] ?? 'prev_page')
            .setLabel('◀️ Попередня')
            .setStyle(discord_js_1.ButtonStyle.Primary)
            .setDisabled(currentPage === 0));
        // Кнопка "Наступна"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(customIds['next'] ?? 'next_page')
            .setLabel('Наступна ▶️')
            .setStyle(discord_js_1.ButtonStyle.Primary)
            .setDisabled(currentPage >= totalPages - 1));
        // Кнопка "Закрити"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(customIds['close'] ?? 'close')
            .setLabel('❌ Закрити')
            .setStyle(discord_js_1.ButtonStyle.Danger));
        return row;
    }
    /**
     * Створення кнопок для дій
     */
    static createActionButtons(actions) {
        const row = new discord_js_1.ActionRowBuilder();
        actions.forEach(action => {
            const button = new discord_js_1.ButtonBuilder()
                .setCustomId(action.id)
                .setLabel(action.label)
                .setStyle(action.style || discord_js_1.ButtonStyle.Primary);
            if (action.emoji) {
                button.setEmoji(action.emoji);
            }
            if (action.disabled) {
                button.setDisabled(true);
            }
            row.addComponents(button);
        });
        return row;
    }
    /**
     * Створення прогрес-бару
     */
    static createProgressBar(current, total, width = 20) {
        const progress = Math.round((current / total) * width);
        const bar = '█'.repeat(progress) + '░'.repeat(width - progress);
        const percentage = Math.round((current / total) * 100);
        return `\`[${bar}]\` ${percentage}% (${current}/${total})`;
    }
    /**
     * Створення embed з прогрес-баром
     */
    static createProgressEmbed(title, current, total, status = '') {
        const embed = this.createBaseEmbed(`${EMOJIS.LOADING} ${title}`, this.createProgressBar(current, total), COLORS.INFO);
        if (status) {
            embed.addFields({
                name: `${EMOJIS.INFO} Статус`,
                value: status,
                inline: false
            });
        }
        return embed;
    }
    /**
     * Створення embed для довідки
     */
    static createHelpEmbed(category = 'general') {
        const helpData = {
            general: {
                title: `${EMOJIS.HELP} Довідка по командам`,
                description: 'Виберіть категорію команд для отримання детальної інформації',
                fields: [
                    { name: '🔍 Пошук', value: 'Команди для пошуку даних', inline: true },
                    { name: '🤖 AI', value: 'AI-асистент та аналіз', inline: true },
                    { name: '📁 Файли', value: 'Робота з файлами', inline: true },
                    { name: '📤 Експорт', value: 'Експорт даних', inline: true },
                    { name: '⚙️ Адміністративні', value: 'Управління ботом', inline: true }
                ]
            },
            search: {
                title: `${EMOJIS.SEARCH} Команди пошуку`,
                description: 'Команди для пошуку та фільтрації даних',
                fields: [
                    { name: '/пошук', value: 'Пошук за конкретним полем', inline: false },
                    { name: '/розумний-пошук', value: 'Пошук за кількома критеріями', inline: false },
                    { name: '/залишки', value: 'Показує підсумкові значення', inline: false },
                    { name: '/оновити', value: 'Показує останні записи', inline: false }
                ]
            },
            ai: {
                title: `${EMOJIS.AI} AI-функції`,
                description: 'Команди для роботи з AI-асистентом',
                fields: [
                    { name: '/ai', value: 'Природномовний запит до AI', inline: false },
                    { name: 'Приклади запитів:', value: '• "знайди товари iPhone"\n• "проаналізуй залишки"\n• "створіть звіт по продажах"', inline: false }
                ]
            },
            files: {
                title: `${EMOJIS.FILES} Робота з файлами`,
                description: 'Команди для роботи з Google Drive',
                fields: [
                    { name: '/файли пошук', value: 'Пошук файлів в Google Drive', inline: false },
                    { name: '/файли читати', value: 'Читання вмісту файлу', inline: false },
                    { name: '/файли аналіз', value: 'AI-аналіз файлу', inline: false },
                    { name: '/файли звіт', value: 'Створення звіту з файлу', inline: false }
                ]
            }
        };
        const data = (helpData[category] ?? helpData['general']);
        const embed = this.createBaseEmbed(data.title, data.description, COLORS.INFO);
        data.fields.forEach(field => {
            embed.addFields(field);
        });
        return embed;
    }
    /**
     * Створення embed для статистики
     */
    static createStatsEmbed(stats) {
        const embed = this.createBaseEmbed(`${EMOJIS.STATS} Статистика бота`, 'Детальна статистика використання бота', COLORS.INFO);
        // Загальна статистика
        embed.addFields({
            name: `${EMOJIS.INFO} Загальна статистика`,
            value: `**Команд виконано:** ${stats.totalCommands || 0}\n**Унікальних користувачів:** ${stats.uniqueUsers || 0}\n**Активних розмов:** ${stats.activeConversations || 0}`,
            inline: false
        });
        // Статистика по командах
        if (stats.commandStats) {
            const commandStats = Object.entries(stats.commandStats)
                .map(([cmd, count]) => `**${cmd}:** ${count}`)
                .join('\n');
            embed.addFields({
                name: `${EMOJIS.INFO} Популярні команди`,
                value: commandStats || 'Немає даних',
                inline: true
            });
        }
        // AI статистика
        if (stats.aiStats) {
            embed.addFields({
                name: `${EMOJIS.AI} AI статистика`,
                value: `**Запитів:** ${stats.aiStats.requests || 0}\n**Провайдер:** ${stats.aiStats.provider || 'N/A'}\n**Середній час відповіді:** ${stats.aiStats.avgResponseTime || 0}мс`,
                inline: true
            });
        }
        return embed;
    }
    /**
     * Створення embed для безпеки
     */
    static createSecurityEmbed(event, details) {
        const embed = this.createBaseEmbed(`${EMOJIS.SECURITY} Подія безпеки`, `**Тип події:** ${event}`, COLORS.WARNING);
        if (details) {
            Object.entries(details).forEach(([key, value]) => {
                embed.addFields({
                    name: key,
                    value: value.toString(),
                    inline: true
                });
            });
        }
        return embed;
    }
    /**
     * Створення інтерактивного меню
     */
    static createInteractiveMenu(title, options, description = '') {
        const embed = this.createBaseEmbed(title, description, COLORS.INFO);
        options.forEach((option, index) => {
            embed.addFields({
                name: `${index + 1}. ${option.label}`,
                value: option.description || 'Немає опису',
                inline: false
            });
        });
        const row = new discord_js_1.ActionRowBuilder();
        options.forEach((option, index) => {
            row.addComponents(new discord_js_1.ButtonBuilder()
                .setCustomId(option.id)
                .setLabel(`${index + 1}`)
                .setStyle(discord_js_1.ButtonStyle.Primary));
        });
        return { embed, row };
    }
    /**
     * Обробка інтерактивних компонентів
     */
    static async handleInteraction(interaction, timeout = 60000) {
        try {
            const response = await interaction.awaitMessageComponent({
                filter: (i) => i.user.id === interaction.user.id,
                time: timeout,
                componentType: discord_js_1.ComponentType.Button
            });
            return response;
        }
        catch (error) {
            logger_1.default.error(`Interaction timeout or error: ${(error instanceof Error) ? error.message : String(error)}`);
            return null;
        }
    }
}
exports.UIHelper = UIHelper;
//# sourceMappingURL=uiHelpers.js.map