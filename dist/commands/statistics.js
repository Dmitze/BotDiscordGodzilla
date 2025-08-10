"use strict";
/**
 * Команда для роботи зі статистикою та складними формулами Google Sheets
 * Підтримує підрахунок по парних/непарних стовпцях, агрегацію по аркушах
 * TypeScript версія 3.0.0
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const discord_js_1 = require("discord.js");
const GoogleService_1 = require("@/services/GoogleService");
const AIService_1 = require("@/services/AIService");
const security_1 = require("@/utils/security");
const logger_1 = __importDefault(require("@/utils/logger"));
const uiHelpers_1 = require("@/utils/uiHelpers");
const formatters_1 = require("@/utils/formatters");
class StatisticsCommand {
    constructor() {
        this.name = 'statistics';
        this.description = 'Отримання статистики з Google Sheets з підтримкою складних формул';
        this.usage = '/statistics <операція> <аркуші> [опції]';
        this.googleService = new GoogleService_1.GoogleService();
        this.aiService = new AIService_1.AIService();
    }
    /**
     * Створення команди
     */
    getCommandData() {
        return new discord_js_1.SlashCommandBuilder()
            .setName(this.name)
            .setDescription(this.description)
            .addStringOption(option => option.setName('operation')
            .setDescription('Тип операції для статистики')
            .setRequired(true)
            .addChoices({ name: 'Сума', value: 'sum' }, { name: 'Середнє', value: 'average' }, { name: 'Кількість', value: 'count' }, { name: 'Максимум', value: 'max' }, { name: 'Мінімум', value: 'min' }, { name: 'Парні стовпці', value: 'even_columns' }, { name: 'Непарні стовпці', value: 'odd_columns' }, { name: 'Складена формула', value: 'complex_formula' }))
            .addStringOption(option => option.setName('sheets')
            .setDescription('Аркуші для аналізу (через кому)')
            .setRequired(true))
            .addStringOption(option => option.setName('range')
            .setDescription('Діапазон даних (наприклад: H6:AB6)')
            .setRequired(false))
            .addStringOption(option => option.setName('column_type')
            .setDescription('Тип стовпців для аналізу')
            .setRequired(false)
            .addChoices({ name: 'Всі', value: 'all' }, { name: 'Парні', value: 'even' }, { name: 'Непарні', value: 'odd' }))
            .addStringOption(option => option.setName('group_by')
            .setDescription('Групування за стовпцем')
            .setRequired(false))
            .addStringOption(option => option.setName('filters')
            .setDescription('Фільтри у форматі JSON')
            .setRequired(false))
            .addStringOption(option => option.setName('custom_formula')
            .setDescription('Власна формула для аналізу')
            .setRequired(false));
    }
    /**
     * Виконання команди
     */
    async execute(interaction) {
        const startTime = performance.now();
        try {
            logger_1.default.info('Початок виконання команди statistics', {
                user: interaction.user.tag,
                userId: interaction.user.id,
                guildId: interaction.guildId,
            });
            // Валідація опцій
            const options = this.extractOptions(interaction);
            const validation = (0, security_1.validateCommandOptions)(options, this.getValidationSchema());
            if (!validation.isValid) {
                await interaction.reply({
                    content: `❌ Помилка валідації: ${validation.errors.join(', ')}`,
                    ephemeral: true
                });
                return;
            }
            // Дефірування відповіді
            await interaction.deferReply();
            // Отримання статистики
            const result = await this.getStatistics(options);
            // Створення відповіді
            const embed = this.createStatisticsEmbed(result, options);
            const buttons = this.createActionButtons(result, options);
            const duration = performance.now() - startTime;
            logger_1.default.info(`Команда statistics виконана за ${duration.toFixed(2)}ms`, {
                user: interaction.user.tag,
                operation: options.operation,
                sheets: options.sheets.length,
                result: result.total,
            });
            await interaction.editReply({
                embeds: [embed],
                components: buttons ? [buttons] : undefined,
            });
        }
        catch (error) {
            const duration = performance.now() - startTime;
            logger_1.default.error(`Помилка команди statistics після ${duration.toFixed(2)}ms:`, error);
            await this.handleError(interaction, error);
        }
    }
    /**
     * Витягування опцій з interaction
     */
    extractOptions(interaction) {
        const operation = interaction.options.getString('operation', true);
        const sheetsInput = interaction.options.getString('sheets', true);
        const range = interaction.options.getString('range') || 'H6:AB6';
        const columnType = interaction.options.getString('column_type') || 'all';
        const groupBy = interaction.options.getString('group_by');
        const filtersInput = interaction.options.getString('filters');
        const customFormula = interaction.options.getString('custom_formula');
        // Санітизація вхідних даних
        const sheets = (0, security_1.sanitizeInput)(sheetsInput, 'command').sanitizedValue?.split(',').map(s => s.trim()) || [];
        const filters = filtersInput ? JSON.parse((0, security_1.sanitizeInput)(filtersInput, 'command').sanitizedValue || '{}') : {};
        return {
            sheets,
            range,
            columnType,
            operation: operation,
            groupBy,
            filters,
            customFormula: customFormula ? (0, security_1.sanitizeInput)(customFormula, 'command').sanitizedValue : undefined,
        };
    }
    /**
     * Схема валідації
     */
    getValidationSchema() {
        return {
            sheets: {
                required: true,
                type: 'object',
                minLength: 1,
            },
            range: {
                required: true,
                type: 'string',
                pattern: /^[A-Z]+\d+:[A-Z]+\d+$/,
            },
            operation: {
                required: true,
                type: 'string',
                enum: ['sum', 'average', 'count', 'max', 'min', 'even_columns', 'odd_columns', 'complex_formula'],
            },
        };
    }
    /**
     * Отримання статистики
     */
    async getStatistics(config) {
        const startTime = performance.now();
        try {
            logger_1.default.debug('Початок отримання статистики', { config });
            let total = 0;
            const breakdown = {};
            // Обробка різних типів операцій
            switch (config.operation) {
                case 'even_columns':
                case 'odd_columns':
                    total = await this.calculateColumnStatistics(config, config.operation === 'even_columns');
                    break;
                case 'complex_formula':
                    total = await this.executeComplexFormula(config);
                    break;
                default:
                    total = await this.calculateBasicStatistics(config);
                    break;
            }
            // Групування результатів
            if (config.groupBy) {
                breakdown[config.groupBy] = total;
            }
            else {
                breakdown['Загальна сума'] = total;
            }
            const processingTime = performance.now() - startTime;
            return {
                total,
                breakdown,
                summary: this.generateSummary(total, config),
                timestamp: new Date(),
                processingTime,
            };
        }
        catch (error) {
            logger_1.default.error('Помилка отримання статистики:', error);
            throw error;
        }
    }
    /**
     * Розрахунок статистики по парних/непарних стовпцях
     */
    async calculateColumnStatistics(config, isEven) {
        let total = 0;
        for (const sheetName of config.sheets) {
            try {
                const data = await this.googleService.getSheetData(sheetName, config.range);
                if (!data || !data.values || data.values.length === 0) {
                    logger_1.default.warn(`Немає даних в аркуші ${sheetName}`);
                    continue;
                }
                const row = data.values[0]; // Перший рядок
                const startCol = this.getColumnIndex(config.range.split(':')[0]);
                const endCol = this.getColumnIndex(config.range.split(':')[1]);
                for (let col = startCol; col <= endCol; col++) {
                    const isEvenColumn = col % 2 === 0;
                    if (isEven ? isEvenColumn : !isEvenColumn) {
                        const value = parseFloat(row[col - startCol] || '0');
                        if (!isNaN(value)) {
                            total += value;
                        }
                    }
                }
                logger_1.default.debug(`Оброблено аркуш ${sheetName}`, { total, isEven });
            }
            catch (error) {
                logger_1.default.error(`Помилка обробки аркуша ${sheetName}:`, error);
            }
        }
        return total;
    }
    /**
     * Виконання складних формул
     */
    async executeComplexFormula(config) {
        if (!config.customFormula) {
            throw new Error('Власна формула не надана');
        }
        try {
            // Аналіз формули за допомогою AI
            const analysis = await this.aiService.analyzeFormula(config.customFormula);
            logger_1.default.info('AI аналіз формули завершено', { analysis });
            // Виконання формули через Google Sheets API
            const result = await this.googleService.executeFormula(config.customFormula);
            return parseFloat(result) || 0;
        }
        catch (error) {
            logger_1.default.error('Помилка виконання складной формули:', error);
            throw new Error('Не вдалося виконати складну формулу');
        }
    }
    /**
     * Розрахунок базової статистики
     */
    async calculateBasicStatistics(config) {
        let total = 0;
        let count = 0;
        for (const sheetName of config.sheets) {
            try {
                const data = await this.googleService.getSheetData(sheetName, config.range);
                if (!data || !data.values)
                    continue;
                for (const row of data.values) {
                    for (const cell of row) {
                        const value = parseFloat(cell || '0');
                        if (!isNaN(value)) {
                            switch (config.operation) {
                                case 'sum':
                                    total += value;
                                    break;
                                case 'average':
                                    total += value;
                                    count++;
                                    break;
                                case 'count':
                                    if (value > 0)
                                        count++;
                                    break;
                                case 'max':
                                    total = Math.max(total, value);
                                    break;
                                case 'min':
                                    total = total === 0 ? value : Math.min(total, value);
                                    break;
                            }
                        }
                    }
                }
            }
            catch (error) {
                logger_1.default.error(`Помилка обробки аркуша ${sheetName}:`, error);
            }
        }
        return config.operation === 'average' ? (count > 0 ? total / count : 0) :
            config.operation === 'count' ? count : total;
    }
    /**
     * Отримання індексу стовпця
     */
    getColumnIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + (column.charCodeAt(i) - 64);
        }
        return index;
    }
    /**
     * Генерація підсумку
     */
    generateSummary(total, config) {
        const operationNames = {
            sum: 'сума',
            average: 'середнє',
            count: 'кількість',
            max: 'максимум',
            min: 'мінімум',
            even_columns: 'сума парних стовпців',
            odd_columns: 'сума непарних стовпців',
            complex_formula: 'результат формули',
        };
        return `**${operationNames[config.operation]}**: ${formatters_1.DataFormatters.formatNumber(total)}`;
    }
    /**
     * Створення embed для відповіді
     */
    createStatisticsEmbed(result, config) {
        const embed = uiHelpers_1.UIHelper.createBaseEmbed()
            .setTitle('📊 Статистика Google Sheets')
            .setColor('#00ff00')
            .setTimestamp(result.timestamp);
        // Основна інформація
        embed.addFields({ name: '📈 Результат', value: result.summary, inline: true }, { name: '⏱️ Час обробки', value: `${result.processingTime.toFixed(2)}ms`, inline: true }, { name: '📋 Аркуші', value: config.sheets.length.toString(), inline: true });
        // Детальна розбивка
        if (Object.keys(result.breakdown).length > 1) {
            const breakdownText = Object.entries(result.breakdown)
                .map(([key, value]) => `**${key}**: ${formatters_1.DataFormatters.formatNumber(value)}`)
                .join('\n');
            embed.addFields({ name: '📊 Детальна розбивка', value: breakdownText });
        }
        // Додаткова інформація
        embed.addFields({ name: '🔧 Операція', value: config.operation, inline: true }, { name: '📏 Діапазон', value: config.range, inline: true }, { name: '📊 Тип стовпців', value: config.columnType, inline: true });
        // Фільтри
        if (Object.keys(config.filters).length > 0) {
            const filtersText = Object.entries(config.filters)
                .map(([key, value]) => `**${key}**: ${value}`)
                .join('\n');
            embed.addFields({ name: '🔍 Фільтри', value: filtersText });
        }
        return embed;
    }
    /**
     * Створення кнопок дій
     */
    createActionButtons(result, config) {
        const row = new discord_js_1.ActionRowBuilder();
        // Кнопка експорту
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(`export_stats_${Date.now()}`)
            .setLabel('📊 Експорт')
            .setStyle(discord_js_1.ButtonStyle.Primary));
        // Кнопка детального аналізу
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(`analyze_stats_${Date.now()}`)
            .setLabel('🔍 Аналіз')
            .setStyle(discord_js_1.ButtonStyle.Secondary));
        // Кнопка оновлення
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(`refresh_stats_${Date.now()}`)
            .setLabel('🔄 Оновити')
            .setStyle(discord_js_1.ButtonStyle.Success));
        return row;
    }
    /**
     * Обробка помилок
     */
    async handleError(interaction, error) {
        const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
        try {
            if (interaction.deferred) {
                await interaction.editReply({
                    content: `❌ Помилка отримання статистики: ${errorMessage}`,
                });
            }
            else {
                await interaction.reply({
                    content: `❌ Помилка отримання статистики: ${errorMessage}`,
                    ephemeral: true,
                });
            }
        }
        catch (replyError) {
            logger_1.default.error('Помилка відповіді на помилку:', replyError);
        }
    }
    /**
     * Отримання назви команди
     */
    getName() {
        return this.name;
    }
    /**
     * Отримання опису команди
     */
    getDescription() {
        return this.description;
    }
}
exports.default = StatisticsCommand;
//# sourceMappingURL=statistics.js.map