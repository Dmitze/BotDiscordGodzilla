"use strict";
/**
 * Базовий абстрактний клас для всіх команд Discord бота
 * Забезпечує уніфіковану структуру та типізацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.BaseCommand = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("@/utils/logger"));
const security_1 = require("@/utils/security");
// Константи для конфігурації команд
const COMMAND_CONFIG = {
    DEFAULT_COOLDOWN: 3000, // 3 секунди
    MAX_COOLDOWN: 30000, // 30 секунд
    MIN_COOLDOWN: 1000, // 1 секунда
    MAX_EXECUTION_TIME: 30000, // 30 секунд
    MAX_RETRIES: 3,
    RETRY_DELAY: 1000, // 1 секунда
    CACHE_SIZE: 1000,
    CLEANUP_INTERVAL: 5 * 60 * 1000, // 5 хвилин
};
class BaseCommand {
    constructor(name, description, config, options = {}, builder) {
        this.cooldowns = new Map();
        this.executionCache = new Map();
        this.errorCount = new Map();
        this.lastExecution = new Map();
        this.isShuttingDown = false;
        this.name = name;
        this.description = description;
        this.config = config;
        this.category = options.category || 'general';
        this.usage = options.usage || `/${name}`;
        this.examples = options.examples || [];
        this.permissions = options.permissions || [];
        this.cooldown = this.validateCooldown(options.cooldown || COMMAND_CONFIG.DEFAULT_COOLDOWN);
        // Створення SlashCommandBuilder
        this.data = new discord_js_1.SlashCommandBuilder()
            .setName(name)
            .setDescription(description);
        // Додавання опцій через builder функцію
        if (builder) {
            try {
                builder(this.data);
            }
            catch (error) {
                logger_1.default.error(`Помилка створення builder для команди ${name}:`, error);
                throw new Error(`Помилка створення команди: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
            }
        }
        // Встановлення дозволів
        if (options.defaultMemberPermissions) {
            this.data.setDefaultMemberPermissions(options.defaultMemberPermissions);
        }
        if (options.dmPermission !== undefined) {
            this.data.setDMPermission(options.dmPermission);
        }
        this.stats = {
            service: `Command:${name}`,
            uptime: 0,
            requests: 0,
            errors: 0,
            totalExecutions: 0,
            successfulExecutions: 0,
            failedExecutions: 0,
            averageExecutionTime: 0,
            totalExecutionTime: 0,
            cacheHits: 0,
            cacheMisses: 0,
            retries: 0,
        };
        // Запуск періодичного очищення
        this.startCleanupInterval();
        logger_1.default.info(`Команда ${name} ініціалізована`, {
            category: this.category,
            cooldown: this.cooldown,
            permissions: this.permissions,
        });
    }
    /**
     * Виконання команди з детальним логуванням та обробкою помилок
     */
    async execute(options) {
        const startTime = performance.now();
        const userId = options.interaction.user.id;
        const userTag = options.interaction.user.tag;
        try {
            // Перевірка стану команди
            if (this.isShuttingDown) {
                await this.handleShutdownError(options.interaction);
                return;
            }
            // Валідація вхідних даних
            const validation = await this.validateExecution(options);
            if (!validation.isValid) {
                await this.handleValidationError(options.interaction, validation.errors);
                return;
            }
            // Перевірка cooldown
            if (this.isOnCooldown(userId)) {
                await this.handleCooldown(options.interaction);
                return;
            }
            // Перевірка кешу
            const cacheKey = this.generateCacheKey(options);
            const cachedResult = this.getCachedResult(cacheKey);
            if (cachedResult) {
                await this.handleCachedResult(options.interaction, cachedResult);
                return;
            }
            // Встановлення cooldown
            this.setCooldown(userId);
            // Логування початку виконання
            this.logCommandStart(options.interaction);
            // Виконання команди з retry логікою
            const result = await this.executeWithRetry(options);
            // Кешування результату
            this.cacheResult(cacheKey, result);
            // Оновлення статистики
            const duration = performance.now() - startTime;
            this.updateStats(true, duration);
            // Логування успішного завершення
            this.logCommandSuccess(options.interaction, duration);
            // Оновлення часу останнього виконання
            this.lastExecution.set(userId, Date.now());
        }
        catch (error) {
            // Оновлення статистики помилок
            const duration = performance.now() - startTime;
            this.updateStats(false, duration);
            // Збільшення лічильника помилок
            this.incrementErrorCount(userId);
            // Логування помилки
            this.logCommandError(options.interaction, error);
            // Обробка помилки
            await this.handleError(options.interaction, error);
        }
    }
    /**
     * Виконання команди з retry логікою
     */
    async executeWithRetry(options) {
        let lastError = null;
        for (let attempt = 1; attempt <= COMMAND_CONFIG.MAX_RETRIES; attempt++) {
            try {
                const result = await this.onExecute(options);
                this.stats.retries += attempt - 1;
                return result;
            }
            catch (error) {
                lastError = error instanceof Error ? error : new Error(String(error));
                if (attempt < COMMAND_CONFIG.MAX_RETRIES) {
                    logger_1.default.warn(`Спроба ${attempt} команди ${this.name} невдала, повтор...`, {
                        error: lastError.message,
                        attempt,
                        maxRetries: COMMAND_CONFIG.MAX_RETRIES,
                    });
                    await new Promise(resolve => setTimeout(resolve, COMMAND_CONFIG.RETRY_DELAY * attempt));
                }
            }
        }
        throw lastError || new Error('Всі спроби виконання команди невдалі');
    }
    /**
     * Валідація виконання команди
     */
    async validateExecution(options) {
        const errors = [];
        const warnings = [];
        try {
            // Перевірка користувача
            if (!options.interaction.user) {
                errors.push('Користувач не знайдено');
            }
            // Перевірка сервера (якщо потрібно)
            if (!options.interaction.guild && !this.data.dmPermission) {
                errors.push('Команда доступна тільки на сервері');
            }
            // Перевірка дозволів
            if (this.permissions.length > 0) {
                const member = options.interaction.member;
                if (member && 'permissions' in member) {
                    const hasPermission = this.permissions.some(permission => member.permissions.has(permission));
                    if (!hasPermission) {
                        errors.push(`Необхідні дозволи: ${this.permissions.join(', ')}`);
                    }
                }
            }
            // Санітизація опцій
            if (options.options) {
                const sanitizedOptions = {};
                for (const [key, value] of Object.entries(options.options)) {
                    if (typeof value === 'string') {
                        const sanitized = (0, security_1.sanitizeInput)(value, 'command');
                        if (sanitized.isValid) {
                            sanitizedOptions[key] = sanitized.sanitizedValue;
                            if (sanitized.warnings.length > 0) {
                                warnings.push(...sanitized.warnings.map(w => `${key}: ${w}`));
                            }
                        }
                        else {
                            errors.push(...sanitized.errors.map(e => `${key}: ${e}`));
                        }
                    }
                    else {
                        sanitizedOptions[key] = value;
                    }
                }
                options.options = sanitizedOptions;
            }
            return {
                isValid: errors.length === 0,
                errors,
                warnings,
                sanitizedOptions: options.options,
            };
        }
        catch (error) {
            logger_1.default.error('Помилка валідації команди:', error);
            return {
                isValid: false,
                errors: ['Помилка валідації команди'],
                warnings: [],
            };
        }
    }
    /**
     * Обробка автодоповнення з детальним логуванням
     */
    async autocomplete(options) {
        const startTime = performance.now();
        try {
            logger_1.default.debug(`Автодоповнення для команди ${this.name}`, {
                user: options.interaction.user.tag,
                query: options.query,
            });
            await this.onAutocomplete(options);
            const duration = performance.now() - startTime;
            logger_1.default.debug(`Автодоповнення ${this.name} завершено за ${duration.toFixed(2)}ms`);
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.logAutocompleteError(options.interaction, error);
            await this.handleAutocompleteError(options.interaction, error);
        }
    }
    /**
     * Обробка компонентів з детальним логуванням
     */
    async handleComponent(options) {
        const startTime = performance.now();
        try {
            logger_1.default.debug(`Обробка компонента для команди ${this.name}`, {
                user: options.interaction.user.tag,
                componentType: options.componentType,
                customId: options.interaction.customId,
            });
            await this.onComponent(options);
            const duration = performance.now() - startTime;
            logger_1.default.debug(`Компонент ${this.name} оброблено за ${duration.toFixed(2)}ms`);
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.logComponentError(options.interaction, error);
            await this.handleComponentError(options.interaction, error);
        }
    }
    /**
     * Обробка автодоповнення (опціонально)
     */
    async onAutocomplete(options) {
        // Базова реалізація - нічого не робить
    }
    /**
     * Обробка компонентів (опціонально)
     */
    async onComponent(options) {
        // Базова реалізація - нічого не робить
    }
    /**
     * Валідація cooldown
     */
    validateCooldown(cooldown) {
        if (cooldown < COMMAND_CONFIG.MIN_COOLDOWN) {
            logger_1.default.warn(`Cooldown для команди ${this.name} занадто малий, встановлюю мінімальний`);
            return COMMAND_CONFIG.MIN_COOLDOWN;
        }
        if (cooldown > COMMAND_CONFIG.MAX_COOLDOWN) {
            logger_1.default.warn(`Cooldown для команди ${this.name} занадто великий, встановлюю максимальний`);
            return COMMAND_CONFIG.MAX_COOLDOWN;
        }
        return cooldown;
    }
    /**
     * Перевірка cooldown
     */
    isOnCooldown(userId) {
        const cooldownTime = this.cooldowns.get(userId);
        if (!cooldownTime)
            return false;
        return Date.now() < cooldownTime;
    }
    /**
     * Встановлення cooldown
     */
    setCooldown(userId) {
        this.cooldowns.set(userId, Date.now() + this.cooldown);
    }
    /**
     * Отримання часу cooldown
     */
    getCooldownTime(userId) {
        const cooldownTime = this.cooldowns.get(userId);
        if (!cooldownTime)
            return 0;
        return Math.max(0, cooldownTime - Date.now());
    }
    /**
     * Генерація ключа кешу
     */
    generateCacheKey(options) {
        const userId = options.interaction.user.id;
        const optionsHash = JSON.stringify(options.options || {});
        return `${this.name}:${userId}:${optionsHash}`;
    }
    /**
     * Отримання кешованого результату
     */
    getCachedResult(cacheKey) {
        const cached = this.executionCache.get(cacheKey);
        if (cached && Date.now() - cached.timestamp < 300000) { // 5 хвилин
            this.stats.cacheHits++;
            return cached.result;
        }
        this.stats.cacheMisses++;
        return null;
    }
    /**
     * Кешування результату
     */
    cacheResult(cacheKey, result) {
        this.executionCache.set(cacheKey, {
            result,
            timestamp: Date.now(),
        });
        // Обмеження розміру кешу
        if (this.executionCache.size > COMMAND_CONFIG.CACHE_SIZE) {
            const oldestKey = this.executionCache.keys().next().value;
            this.executionCache.delete(oldestKey);
        }
    }
    /**
     * Збільшення лічильника помилок
     */
    incrementErrorCount(userId) {
        const currentCount = this.errorCount.get(userId) || 0;
        this.errorCount.set(userId, currentCount + 1);
    }
    /**
     * Обробка cooldown
     */
    async handleCooldown(interaction) {
        const remainingTime = this.getCooldownTime(interaction.user.id);
        const seconds = Math.ceil(remainingTime / 1000);
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#FF6B6B')
            .setTitle('⏰ Cooldown активний')
            .setDescription(`Спробуйте ще раз через **${seconds} секунд**`)
            .addFields({ name: 'Команда', value: this.name, inline: true }, { name: 'Залишилось', value: `${seconds}с`, inline: true })
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (error) {
            logger_1.default.error('Помилка відправки cooldown повідомлення:', error);
        }
    }
    /**
     * Обробка кешованого результату
     */
    async handleCachedResult(interaction, result) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#4CAF50')
            .setTitle('⚡ Кешований результат')
            .setDescription('Результат завантажено з кешу')
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (error) {
            logger_1.default.error('Помилка відправки кешованого результату:', error);
        }
    }
    /**
     * Обробка помилки валідації
     */
    async handleValidationError(interaction, errors) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#FF9800')
            .setTitle('⚠️ Помилка валідації')
            .setDescription('Виправте наступні помилки:')
            .addFields(errors.map(error => ({ name: '❌', value: error, inline: false })))
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (error) {
            logger_1.default.error('Помилка відправки повідомлення про валідацію:', error);
        }
    }
    /**
     * Обробка помилки зупинки
     */
    async handleShutdownError(interaction) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#FF6B6B')
            .setTitle('🛑 Команда недоступна')
            .setDescription('Бот знаходиться в процесі зупинки')
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (error) {
            logger_1.default.error('Помилка відправки повідомлення про зупинку:', error);
        }
    }
    /**
     * Обробка помилок
     */
    async handleError(interaction, error) {
        const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
        const errorStack = error instanceof Error ? error.stack : undefined;
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#FF6B6B')
            .setTitle('❌ Помилка виконання команди')
            .setDescription(`**Помилка:** ${errorMessage}`)
            .addFields({ name: 'Команда', value: this.name, inline: true }, { name: 'Користувач', value: interaction.user.tag, inline: true })
            .setTimestamp();
        if (errorStack) {
            embed.addFields({ name: 'Деталі', value: `\`\`\`${errorStack.substring(0, 1000)}...\`\`\``, inline: false });
        }
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (replyError) {
            logger_1.default.error('Помилка відправки повідомлення про помилку:', replyError);
        }
    }
    /**
     * Обробка помилок автодоповнення
     */
    async handleAutocompleteError(interaction, error) {
        try {
            await interaction.respond([
                { name: 'Помилка завантаження', value: 'error' }
            ]);
        }
        catch (replyError) {
            logger_1.default.error('Помилка відповіді автодоповнення:', replyError);
        }
    }
    /**
     * Обробка помилок компонентів
     */
    async handleComponentError(interaction, error) {
        const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#FF6B6B')
            .setTitle('❌ Помилка обробки компонента')
            .setDescription(`**Помилка:** ${errorMessage}`)
            .addFields({ name: 'Команда', value: this.name, inline: true }, { name: 'Компонент', value: interaction.customId, inline: true })
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [embed] });
            }
            else {
                await interaction.reply({ embeds: [embed], ephemeral: true });
            }
        }
        catch (replyError) {
            logger_1.default.error('Помилка відправки повідомлення про помилку компонента:', replyError);
        }
    }
    /**
     * Логування початку команди
     */
    logCommandStart(interaction) {
        logger_1.default.info(`🚀 Команда ${this.name} виконується`, {
            user: interaction.user.tag,
            userId: interaction.user.id,
            guildId: interaction.guildId,
            channelId: interaction.channelId,
            options: interaction.options.data,
        });
    }
    /**
     * Логування успішного завершення
     */
    logCommandSuccess(interaction, duration) {
        logger_1.default.info(`✅ Команда ${this.name} успішно виконана`, {
            user: interaction.user.tag,
            duration: `${duration.toFixed(2)}ms`,
            performance: duration > 5000 ? 'slow' : duration > 1000 ? 'medium' : 'fast',
        });
    }
    /**
     * Логування помилки команди
     */
    logCommandError(interaction, error) {
        logger_1.default.error(`❌ Помилка команди ${this.name}`, {
            user: interaction.user.tag,
            userId: interaction.user.id,
            error: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined,
        });
    }
    /**
     * Логування помилки автодоповнення
     */
    logAutocompleteError(interaction, error) {
        logger_1.default.error(`❌ Помилка автодоповнення команди ${this.name}`, {
            user: interaction.user.tag,
            error: error instanceof Error ? error.message : String(error),
        });
    }
    /**
     * Логування помилки компонента
     */
    logComponentError(interaction, error) {
        logger_1.default.error(`❌ Помилка компонента команди ${this.name}`, {
            user: interaction.user.tag,
            customId: interaction.customId,
            error: error instanceof Error ? error.message : String(error),
        });
    }
    /**
     * Оновлення статистики
     */
    updateStats(success, duration) {
        this.stats.totalExecutions++;
        this.stats.totalExecutionTime += duration;
        this.stats.uptime = Date.now() - this.stats.uptime;
        if (success) {
            this.stats.successfulExecutions++;
        }
        else {
            this.stats.failedExecutions++;
        }
        this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.totalExecutions;
    }
    /**
     * Запуск періодичного очищення
     */
    startCleanupInterval() {
        setInterval(() => {
            this.cleanupExpiredData();
        }, COMMAND_CONFIG.CLEANUP_INTERVAL);
    }
    /**
     * Очищення застарілих даних
     */
    cleanupExpiredData() {
        const now = Date.now();
        let cleanedCooldowns = 0;
        let cleanedCache = 0;
        let cleanedErrors = 0;
        // Очищення cooldowns
        for (const [userId, cooldownTime] of this.cooldowns.entries()) {
            if (now > cooldownTime) {
                this.cooldowns.delete(userId);
                cleanedCooldowns++;
            }
        }
        // Очищення кешу
        for (const [key, cached] of this.executionCache.entries()) {
            if (now - cached.timestamp > 300000) { // 5 хвилин
                this.executionCache.delete(key);
                cleanedCache++;
            }
        }
        // Очищення помилок (старіші 1 години)
        for (const [userId, errorTime] of this.errorCount.entries()) {
            if (now - errorTime > 3600000) {
                this.errorCount.delete(userId);
                cleanedErrors++;
            }
        }
        if (cleanedCooldowns > 0 || cleanedCache > 0 || cleanedErrors > 0) {
            logger_1.default.debug(`Очищено застарілі дані команди ${this.name}`, {
                cooldowns: cleanedCooldowns,
                cache: cleanedCache,
                errors: cleanedErrors,
            });
        }
    }
    /**
     * Отримання статистики команди
     */
    getCommandStats() {
        return { ...this.stats };
    }
    /**
     * Очищення cooldowns
     */
    clearCooldowns() {
        this.cooldowns.clear();
        logger_1.default.debug(`Cooldowns команди ${this.name} очищено`);
    }
    /**
     * Health check
     */
    async healthCheck() {
        const successRate = this.stats.totalExecutions > 0
            ? (this.stats.successfulExecutions / this.stats.totalExecutions) * 100
            : 0;
        const isHealthy = successRate > 80 && this.stats.averageExecutionTime < 10000;
        return {
            healthy: isHealthy,
            service: this.name,
            details: {
                totalExecutions: this.stats.totalExecutions,
                successRate: `${successRate.toFixed(2)}%`,
                averageExecutionTime: `${this.stats.averageExecutionTime.toFixed(2)}ms`,
                activeCooldowns: this.cooldowns.size,
                cacheSize: this.executionCache.size,
                errorCount: this.errorCount.size,
            },
        };
    }
    /**
     * Завершення роботи
     */
    async shutdown() {
        this.isShuttingDown = true;
        this.clearCooldowns();
        this.executionCache.clear();
        this.errorCount.clear();
        this.lastExecution.clear();
        logger_1.default.info(`Команда ${this.name} зупинена`);
    }
    /**
     * Отримання статистики
     */
    getStats() {
        return { ...this.stats };
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
    /**
     * Отримання даних команди для реєстрації в Discord
     */
    getData() {
        return this.data;
    }
    /**
     * Отримання допомоги по команді
     */
    getHelp() {
        return `**Команда:** ${this.name}
**Опис:** ${this.description}
**Використання:** ${this.usage}
**Категорія:** ${this.category}
**Cooldown:** ${this.cooldown / 1000}с
${this.examples.length > 0 ? `**Приклади:**\n${this.examples.map(ex => `\`${ex}\``).join('\n')}` : ''}`;
    }
}
exports.BaseCommand = BaseCommand;
//# sourceMappingURL=BaseCommand.js.map