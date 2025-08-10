"use strict";
/**
 * Рефакторований базовий клас для команд Discord бота
 * Використовує модульну архітектуру для кращої підтримки
 * Версія 4.0.0 - Модульна архітектура
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.BaseCommand = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("@/utils/logger"));
const CommandValidator_1 = __importDefault(require("./modules/CommandValidator"));
const CommandMetrics_1 = __importDefault(require("./modules/CommandMetrics"));
// Константи конфігурації
const COMMAND_CONFIG = {
    DEFAULT_COOLDOWN: 3000,
    MAX_EXECUTION_TIME: 30000,
    MAX_RETRIES: 3,
    RETRY_DELAY: 1000,
};
class BaseCommand {
    constructor(commandData, config) {
        this.cooldowns = new Map();
        this.isShuttingDown = false;
        this.config = config;
        this.name = commandData.name;
        this.description = commandData.description;
        this.category = commandData.category || 'Загальні';
        this.usage = commandData.usage || `/${commandData.name}`;
        this.examples = commandData.examples || [];
        this.permissions = commandData.permissions || [];
        this.cooldown = commandData.cooldown || COMMAND_CONFIG.DEFAULT_COOLDOWN;
        // Створення SlashCommandBuilder
        this.data = new discord_js_1.SlashCommandBuilder()
            .setName(commandData.name)
            .setDescription(commandData.description);
        // Налаштування дозволів
        if (commandData.defaultMemberPermissions) {
            this.data.setDefaultMemberPermissions(commandData.defaultMemberPermissions);
        }
        if (commandData.dmPermission !== undefined) {
            this.data.setDMPermission(commandData.dmPermission);
        }
        // Додавання опцій
        if (commandData.options) {
            this.addOptions(commandData.options);
        }
        // Ініціалізація статистики
        this.stats = {
            commandName: this.name,
            executionCount: 0,
            successCount: 0,
            errorCount: 0,
            averageExecutionTime: 0,
            totalExecutionTime: 0,
            lastExecuted: 0,
            cacheHits: 0,
            cacheMisses: 0
        };
        // Ініціалізація модулів
        this.validator = new CommandValidator_1.default();
        this.metrics = new CommandMetrics_1.default();
        logger_1.default.debug(`✅ Команда "${this.name}" ініціалізована`);
    }
    /**
     * Головна точка входу для виконання команди
     */
    async handleInteraction(interaction) {
        const startTime = Date.now();
        let success = false;
        let error;
        try {
            // Перевірка cooldown
            if (this.isOnCooldown(interaction.user.id)) {
                const remainingTime = this.getRemainingCooldown(interaction.user.id);
                await this.sendCooldownMessage(interaction, remainingTime);
                this.metrics.recordCooldownHit(this.name);
                return;
            }
            // Валідація команди
            const validationResult = await this.validateInteraction(interaction);
            if (!validationResult.isValid) {
                await this.sendValidationError(interaction, validationResult);
                return;
            }
            // Встановлення cooldown
            this.setCooldown(interaction.user.id);
            // Виконання команди з retry логікою
            await this.executeWithRetry({
                interaction,
                startTime,
                retryCount: 0,
                validationResult
            });
            success = true;
        }
        catch (err) {
            error = err instanceof Error ? err.message : String(err);
            logger_1.default.error(`❌ Помилка виконання команди "${this.name}":`, {
                userId: interaction.user.id,
                error: error,
                duration: Date.now() - startTime
            });
            await this.handleExecutionError(interaction, err);
        }
        finally {
            // Запис метрик
            const duration = Date.now() - startTime;
            this.updateStats(duration, success);
            this.metrics.recordExecution(this.name, interaction.user.id, duration, success, { error });
        }
    }
    /**
     * Валідація взаємодії
     */
    async validateInteraction(interaction, customRules) {
        try {
            // Базова валідація
            const baseValidation = await this.validator.validateCommand(interaction, customRules);
            // Кастомна валідація команди
            const customValidation = await this.customValidation(interaction);
            // Об'єднання результатів
            const combinedErrors = [...baseValidation.errors, ...customValidation.errors];
            const combinedWarnings = [...baseValidation.warnings, ...customValidation.warnings];
            return {
                isValid: combinedErrors.length === 0,
                errors: combinedErrors,
                warnings: combinedWarnings,
                sanitizedValues: {
                    ...baseValidation.sanitizedValues,
                    ...customValidation.sanitizedValues
                }
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка валідації команди:', error);
            return {
                isValid: false,
                errors: ['Внутрішня помилка валідації'],
                warnings: []
            };
        }
    }
    /**
     * Кастомна валідація для конкретної команди
     */
    async customValidation(interaction) {
        // Базова реалізація - може бути перевизначена в дочірніх класах
        return {
            isValid: true,
            errors: [],
            warnings: []
        };
    }
    /**
     * Виконання з повторними спробами
     */
    async executeWithRetry(options) {
        const { interaction, retryCount = 0 } = options;
        try {
            await this.execute(options);
        }
        catch (error) {
            if (retryCount < COMMAND_CONFIG.MAX_RETRIES && this.shouldRetry(error)) {
                logger_1.default.warn(`🔄 Повторна спроба ${retryCount + 1}/${COMMAND_CONFIG.MAX_RETRIES} для команди "${this.name}"`);
                await new Promise(resolve => setTimeout(resolve, COMMAND_CONFIG.RETRY_DELAY * (retryCount + 1)));
                await this.executeWithRetry({
                    ...options,
                    retryCount: retryCount + 1
                });
            }
            else {
                throw error;
            }
        }
    }
    /**
     * Перевірка чи потрібно повторити виконання
     */
    shouldRetry(error) {
        if (error instanceof Error) {
            // Повторюємо для тимчасових помилок мережі
            return error.message.includes('timeout') ||
                error.message.includes('network') ||
                error.message.includes('ECONNRESET') ||
                error.message.includes('rate limit');
        }
        return false;
    }
    /**
     * Управління cooldown
     */
    isOnCooldown(userId) {
        const userCooldown = this.cooldowns.get(userId);
        return userCooldown ? Date.now() < userCooldown : false;
    }
    setCooldown(userId) {
        this.cooldowns.set(userId, Date.now() + this.cooldown);
        // Автоматичне видалення після закінчення cooldown
        setTimeout(() => {
            this.cooldowns.delete(userId);
        }, this.cooldown);
    }
    getRemainingCooldown(userId) {
        const userCooldown = this.cooldowns.get(userId);
        return userCooldown ? Math.max(0, userCooldown - Date.now()) : 0;
    }
    /**
     * Відправка повідомлень про помилки
     */
    async sendCooldownMessage(interaction, remainingTime) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor(0xFFA500)
            .setTitle('⏱️ Cooldown')
            .setDescription(`Зачекайте ще ${Math.ceil(remainingTime / 1000)} секунд перед наступним використанням команди.`)
            .setTimestamp();
        if (interaction.replied || interaction.deferred) {
            await interaction.followUp({ embeds: [embed], ephemeral: true });
        }
        else {
            await interaction.reply({ embeds: [embed], ephemeral: true });
        }
    }
    async sendValidationError(interaction, validation) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor(0xFF0000)
            .setTitle('❌ Помилка валідації')
            .setDescription(validation.errors.join('\n'))
            .setTimestamp();
        if (validation.warnings.length > 0) {
            embed.addFields({
                name: '⚠️ Попередження',
                value: validation.warnings.join('\n'),
                inline: false
            });
        }
        if (interaction.replied || interaction.deferred) {
            await interaction.followUp({ embeds: [embed], ephemeral: true });
        }
        else {
            await interaction.reply({ embeds: [embed], ephemeral: true });
        }
    }
    async handleExecutionError(interaction, error) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor(0xFF0000)
            .setTitle('❌ Помилка виконання')
            .setDescription('Виникла помилка під час виконання команди. Спробуйте пізніше.')
            .setTimestamp();
        // В development режимі показуємо деталі помилки
        if (this.config.environment === 'development' && error instanceof Error) {
            embed.addFields({
                name: 'Деталі помилки',
                value: error.message.substring(0, 1000),
                inline: false
            });
        }
        try {
            if (interaction.replied || interaction.deferred) {
                await interaction.followUp({ embeds: [embed], ephemeral: true });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (replyError) {
            logger_1.default.error('❌ Не вдалося відправити повідомлення про помилку:', replyError);
        }
    }
    /**
     * Оновлення статистики
     */
    updateStats(executionTime, success) {
        this.stats.executionCount++;
        this.stats.lastExecuted = Date.now();
        this.stats.totalExecutionTime += executionTime;
        this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.executionCount;
        if (success) {
            this.stats.successCount++;
        }
        else {
            this.stats.errorCount++;
        }
    }
    /**
     * Додавання опцій до команди
     */
    addOptions(options) {
        options.forEach(option => {
            switch (option.type) {
                case 'string':
                    this.data.addStringOption(opt => {
                        opt.setName(option.name)
                            .setDescription(option.description)
                            .setRequired(option.required || false);
                        if (option.choices) {
                            opt.addChoices(...option.choices);
                        }
                        return opt;
                    });
                    break;
                case 'integer':
                    this.data.addIntegerOption(opt => {
                        opt.setName(option.name)
                            .setDescription(option.description)
                            .setRequired(option.required || false);
                        if (option.min_value !== undefined)
                            opt.setMinValue(option.min_value);
                        if (option.max_value !== undefined)
                            opt.setMaxValue(option.max_value);
                        return opt;
                    });
                    break;
                case 'boolean':
                    this.data.addBooleanOption(opt => {
                        return opt.setName(option.name)
                            .setDescription(option.description)
                            .setRequired(option.required || false);
                    });
                    break;
                case 'user':
                    this.data.addUserOption(opt => {
                        return opt.setName(option.name)
                            .setDescription(option.description)
                            .setRequired(option.required || false);
                    });
                    break;
                case 'attachment':
                    this.data.addAttachmentOption(opt => {
                        return opt.setName(option.name)
                            .setDescription(option.description)
                            .setRequired(option.required || false);
                    });
                    break;
            }
        });
    }
    /**
     * Отримання статистики команди
     */
    getStats() {
        return { ...this.stats };
    }
    /**
     * Скидання статистики
     */
    resetStats() {
        this.stats = {
            commandName: this.name,
            executionCount: 0,
            successCount: 0,
            errorCount: 0,
            averageExecutionTime: 0,
            totalExecutionTime: 0,
            lastExecuted: 0,
            cacheHits: 0,
            cacheMisses: 0
        };
    }
    /**
     * Створення стандартного embed відповіді
     */
    createEmbed(title, description, color = 0x00AE86) {
        return new discord_js_1.EmbedBuilder()
            .setColor(color)
            .setTitle(title)
            .setDescription(description)
            .setTimestamp()
            .setFooter({
            text: `${this.name} | Discord AI Assistant Bot`,
            iconURL: 'https://cdn.discordapp.com/embed/avatars/0.png'
        });
    }
    /**
     * Перевірка дозволів
     */
    hasPermission(interaction, permission) {
        if (!interaction.guild || !interaction.member)
            return false;
        const member = interaction.member;
        if ('permissions' in member) {
            return member.permissions.has(permission);
        }
        return false;
    }
    /**
     * Shutdown hook для очищення ресурсів
     */
    shutdown() {
        this.isShuttingDown = true;
        this.cooldowns.clear();
        logger_1.default.debug(`🛑 Команда "${this.name}" зупинена`);
    }
}
exports.BaseCommand = BaseCommand;
exports.default = BaseCommand;
//# sourceMappingURL=BaseCommandRefactored.js.map