"use strict";
/**
 * Основний клас Discord бота
 * Управляє всіма компонентами та сервісами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.Bot = void 0;
const discord_js_1 = require("discord.js");
const ServiceContainer_1 = require("./ServiceContainer");
const BaseService_1 = require("./BaseService");
const CommandManager_1 = require("./CommandManager");
const ErrorHandler_1 = require("./ErrorHandler");
const EventManager_1 = require("./EventManager");
const ServiceManager_1 = require("./ServiceManager");
const logger_1 = __importDefault(require("@/utils/logger"));
// Константи для конфігурації бота
const BOT_CONSTANTS = {
    READY_TIMEOUT: 30000, // 30 секунд
    COMMAND_TIMEOUT: 15000, // 15 секунд
    MAX_RECONNECT_ATTEMPTS: 5,
    RECONNECT_DELAY: 5000, // 5 секунд
    HEALTH_CHECK_INTERVAL: 60000, // 1 хвилина
    MAX_MEMORY_USAGE: 512 * 1024 * 1024, // 512MB
    COMMAND_RATE_LIMIT: 10, // команд за хвилину
    INTERACTION_RATE_LIMIT: 50, // interactions за хвилину
};
class Bot extends BaseService_1.BaseService {
    constructor(config) {
        super('DiscordBot', config);
        this.commands = new discord_js_1.Collection();
        this.isReady = false;
        this.isConnecting = false;
        this.reconnectAttempts = 0;
        this.healthCheckInterval = null;
        this.lastInteractionTime = new Date();
        this.rateLimitMap = new Map();
        this.slowCommandThreshold = 3000; // 3 секунди
        // Ініціалізація статистики
        this.stats = {
            uptime: 0,
            commands: 0,
            interactions: 0,
            errors: 0,
            reconnects: 0,
            lastActivity: new Date(),
            memory: process.memoryUsage(),
            rateLimitHits: 0,
            slowCommands: 0,
        };
        // Створення Discord клієнта з розширеними intents
        this.client = new discord_js_1.Client({
            intents: [
                discord_js_1.GatewayIntentBits.Guilds,
                discord_js_1.GatewayIntentBits.GuildMessages,
                discord_js_1.GatewayIntentBits.MessageContent,
                discord_js_1.GatewayIntentBits.GuildMembers,
                discord_js_1.GatewayIntentBits.DirectMessages,
                discord_js_1.GatewayIntentBits.GuildPresences,
            ],
            failIfNotExists: false,
            retryLimit: 3,
            ws: {
                properties: {
                    browser: 'Discord AI Assistant Bot',
                },
            },
        });
        // Ініціалізація менеджерів та сервісів
        this.serviceContainer = new ServiceContainer_1.ServiceContainer(config);
        this.commandManager = new CommandManager_1.CommandManager(this.client, config);
        this.errorHandler = new ErrorHandler_1.ErrorHandler(this.serviceContainer);
        this.eventManager = new EventManager_1.EventManager(this);
        this.serviceManager = new ServiceManager_1.ServiceManager(this);
        // Налаштування обробників подій
        this.setupEventHandlers();
        logger_1.default.info('🤖 Екземпляр Discord бота створено');
    }
    /**
     * Ініціалізація бота з детальним логуванням
     */
    async onInitialize() {
        const startTime = Date.now();
        try {
            logger_1.default.info('🚀 Початок ініціалізації Discord бота...');
            // Перевірка системних ресурсів
            await this.checkSystemResources();
            // Ініціалізація обробника помилок
            logger_1.default.info('🛡️ Ініціалізація обробника помилок...');
            await this.errorHandler.initialize();
            // Ініціалізація менеджера подій
            logger_1.default.info('📡 Ініціалізація менеджера подій...');
            await this.eventManager.initialize();
            // Ініціалізація сервісів
            logger_1.default.info('🔧 Ініціалізація сервісів...');
            await this.serviceContainer.initialize();
            // Ініціалізація менеджера сервісів
            logger_1.default.info('⚙️ Ініціалізація менеджера сервісів...');
            await this.serviceManager.initialize();
            // Ініціалізація менеджера команд
            logger_1.default.info('📝 Ініціалізація менеджера команд...');
            await this.commandManager.initialize();
            // Підключення до Discord
            logger_1.default.info('🔌 Підключення до Discord...');
            await this.connectToDiscord();
            // Очікування готовності клієнта
            logger_1.default.info('⏳ Очікування готовності клієнта...');
            await this.waitForReady();
            // Запуск health check
            this.startHealthCheck();
            const initDuration = Date.now() - startTime;
            logger_1.default.info(`✅ Discord бот успішно ініціалізовано за ${initDuration}ms`);
            // Логування статистики запуску
            this.logStartupStats();
        }
        catch (error) {
            const initDuration = Date.now() - startTime;
            logger_1.default.error(`❌ Помилка ініціалізації бота після ${initDuration}ms:`, error);
            // Спроба очищення ресурсів
            await this.cleanupOnError();
            throw new Error(`Помилка ініціалізації бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Завершення роботи бота з детальним логуванням
     */
    async onShutdown() {
        const shutdownStartTime = Date.now();
        try {
            logger_1.default.info('🛑 Початок завершення роботи Discord бота...');
            // Зупинка health check
            this.stopHealthCheck();
            // Завершення менеджера подій
            logger_1.default.info('📡 Завершення менеджера подій...');
            await this.eventManager.shutdown();
            // Завершення менеджера команд
            logger_1.default.info('📝 Завершення менеджера команд...');
            await this.commandManager.shutdown();
            // Завершення менеджера сервісів
            logger_1.default.info('⚙️ Завершення менеджера сервісів...');
            await this.serviceManager.shutdown();
            // Завершення сервісів
            logger_1.default.info('🔧 Завершення сервісів...');
            await this.serviceContainer.shutdown();
            // Завершення обробника помилок
            logger_1.default.info('🛡️ Завершення обробника помилок...');
            await this.errorHandler.shutdown();
            // Відключення від Discord
            logger_1.default.info('🔌 Відключення від Discord...');
            this.client.destroy();
            const shutdownDuration = Date.now() - shutdownStartTime;
            logger_1.default.info(`✅ Discord бот успішно завершено за ${shutdownDuration}ms`);
        }
        catch (error) {
            const shutdownDuration = Date.now() - shutdownStartTime;
            logger_1.default.error(`❌ Помилка завершення бота після ${shutdownDuration}ms:`, error);
            // Примусова зупинка при помилці
            logger_1.default.warn('🔄 Примусова зупинка Discord клієнта...');
            this.client.destroy();
            throw new Error(`Помилка завершення бота: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
        }
    }
    /**
     * Health check бота з розширеною інформацією
     */
    async onHealthCheck() {
        try {
            const isConnected = this.client.isReady();
            const servicesHealth = await this.serviceContainer.getHealthStatus();
            const commandsHealth = await this.commandManager.getHealthStatus();
            const allServicesHealthy = Object.values(servicesHealth).every(health => health.healthy);
            const allCommandsHealthy = Object.values(commandsHealth).every(health => health.healthy);
            const healthy = isConnected && allServicesHealthy && allCommandsHealthy;
            return {
                healthy,
                service: this.name,
                details: {
                    connected: isConnected,
                    ready: this.isReady,
                    services: servicesHealth,
                    commands: commandsHealth,
                    stats: this.getStats(),
                    uptime: this.getStats().uptime,
                    memory: process.memoryUsage(),
                    lastActivity: this.lastInteractionTime,
                    rateLimitHits: this.stats.rateLimitHits,
                    slowCommands: this.stats.slowCommands,
                },
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка health check бота:', error);
            return {
                healthy: false,
                service: this.name,
                error: `Health check failed: ${error instanceof Error ? error.message : 'Невідома помилка'}`,
            };
        }
    }
    /**
     * Отримання детальної статистики бота
     */
    onGetStats() {
        this.stats.uptime = Date.now() - this.startTime;
        this.stats.memory = process.memoryUsage();
        return { ...this.stats };
    }
    /**
     * Перевірка системних ресурсів
     */
    async checkSystemResources() {
        try {
            logger_1.default.info('🔍 Перевірка системних ресурсів бота...');
            const memoryUsage = process.memoryUsage();
            const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
            if (heapUsedMB > 200) {
                logger_1.default.warn(`⚠️ Високе використання пам'яті бота: ${Math.round(heapUsedMB)}MB`);
            }
            // Перевірка доступності мережі
            try {
                const testUrl = 'https://discord.com/api/v10/gateway';
                const response = await fetch(testUrl, { method: 'HEAD' });
                if (!response.ok) {
                    logger_1.default.warn('⚠️ Проблеми з підключенням до Discord API');
                }
            }
            catch (networkError) {
                logger_1.default.warn('⚠️ Проблеми з мережевим підключенням:', networkError);
            }
            logger_1.default.info('✅ Системні ресурси бота перевірено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка перевірки системних ресурсів бота:', error);
            throw error;
        }
    }
    /**
     * Підключення до Discord з обробкою помилок
     */
    async connectToDiscord() {
        if (this.isConnecting) {
            logger_1.default.warn('⚠️ Вже виконується підключення до Discord');
            return;
        }
        this.isConnecting = true;
        try {
            logger_1.default.info('🔌 Спроба підключення до Discord...');
            await this.client.login(this.config.discord.token);
            logger_1.default.info('✅ Успішно підключено до Discord');
            // Скидання лічильника спроб перепідключення
            this.reconnectAttempts = 0;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка підключення до Discord:', error);
            if (this.reconnectAttempts < BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
                this.reconnectAttempts++;
                logger_1.default.info(`🔄 Спроба перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS}...`);
                setTimeout(() => {
                    this.connectToDiscord();
                }, BOT_CONSTANTS.RECONNECT_DELAY);
            }
            else {
                throw new Error(`Не вдалося підключитися до Discord після ${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} спроб`);
            }
        }
        finally {
            this.isConnecting = false;
        }
    }
    /**
     * Налаштування обробників подій з детальним логуванням
     */
    setupEventHandlers() {
        // Ready event
        this.client.on(discord_js_1.Events.ClientReady, () => {
            this.isReady = true;
            this.stats.lastActivity = new Date();
            logger_1.default.info(`🤖 Бот ${this.client.user?.tag} готовий до роботи`);
            logger_1.default.info(`📊 Статистика: ${this.client.guilds.cache.size} серверів, ${this.client.channels.cache.size} каналів`);
            // Встановлення статусу бота
            this.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
        });
        // Interaction event
        this.client.on(discord_js_1.Events.InteractionCreate, async (interaction) => {
            this.stats.interactions++;
            this.lastInteractionTime = new Date();
            try {
                // Перевірка rate limit
                if (this.isRateLimited(interaction.user?.id || 'unknown')) {
                    logger_1.default.warn(`⚠️ Rate limit для користувача ${interaction.user?.id}`);
                    await this.handleRateLimit(interaction);
                    return;
                }
                if (interaction.isCommand()) {
                    await this.handleCommand(interaction);
                }
                else if (interaction.isButton()) {
                    await this.handleButtonInteraction(interaction);
                }
                else if (interaction.isSelectMenu()) {
                    await this.handleSelectMenuInteraction(interaction);
                }
            }
            catch (error) {
                this.stats.errors++;
                logger_1.default.error('❌ Помилка обробки interaction:', error);
                await this.handleInteractionError(interaction, error);
            }
        });
        // Error event
        this.client.on(discord_js_1.Events.Error, (error) => {
            this.stats.errors++;
            logger_1.default.error('❌ Помилка Discord клієнта:', error);
            // Спроба перепідключення при критичних помилках
            if (this.shouldReconnect(error)) {
                this.scheduleReconnect();
            }
        });
        // Disconnect event
        this.client.on(discord_js_1.Events.Disconnect, (event) => {
            this.isReady = false;
            logger_1.default.warn('🔌 Discord клієнт відключено:', event);
            // Автоматичне перепідключення
            this.scheduleReconnect();
        });
        // Reconnecting event
        this.client.on(discord_js_1.Events.Reconnecting, () => {
            this.stats.reconnects++;
            logger_1.default.info('🔄 Discord клієнт перепідключається...');
        });
        // Guild events
        this.client.on(discord_js_1.Events.GuildCreate, (guild) => {
            logger_1.default.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`);
        });
        this.client.on(discord_js_1.Events.GuildDelete, (guild) => {
            logger_1.default.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`);
        });
        logger_1.default.info('✅ Обробники подій Discord налаштовано');
    }
    /**
     * Обробка команд з детальним логуванням
     */
    async handleCommand(interaction) {
        const startTime = Date.now();
        const commandName = interaction.commandName;
        const userId = interaction.user.id;
        const guildId = interaction.guildId;
        logger_1.default.info(`📝 Обробка команди: ${commandName} від користувача ${userId} в сервері ${guildId}`);
        try {
            const command = this.commands.get(commandName);
            if (!command) {
                logger_1.default.warn(`⚠️ Команда не знайдена: ${commandName}`);
                await interaction.reply({
                    content: '❌ Команда не знайдена або не зареєстрована',
                    ephemeral: true
                });
                return;
            }
            // Встановлення таймауту для команди
            const commandTimeout = setTimeout(() => {
                logger_1.default.warn(`⏰ Таймаут команди: ${commandName}`);
            }, BOT_CONSTANTS.COMMAND_TIMEOUT);
            await command.execute(interaction);
            clearTimeout(commandTimeout);
            this.stats.commands++;
            const duration = Date.now() - startTime;
            // Логування повільних команд
            if (duration > this.slowCommandThreshold) {
                this.stats.slowCommands++;
                logger_1.default.warn(`🐌 Повільна команда ${commandName}: ${duration}ms`);
            }
            logger_1.default.info(`✅ Команда ${commandName} виконана за ${duration}ms`);
        }
        catch (error) {
            const duration = Date.now() - startTime;
            logger_1.default.error(`❌ Помилка виконання команди ${commandName} після ${duration}ms:`, error);
            throw error;
        }
    }
    /**
     * Обробка кнопкових interactions
     */
    async handleButtonInteraction(interaction) {
        logger_1.default.debug(`🔘 Обробка кнопкового interaction: ${interaction.customId}`);
        // Тут можна додати логіку обробки кнопок
    }
    /**
     * Обробка select menu interactions
     */
    async handleSelectMenuInteraction(interaction) {
        logger_1.default.debug(`📋 Обробка select menu interaction: ${interaction.customId}`);
        // Тут можна додати логіку обробки select menu
    }
    /**
     * Обробка помилок interactions
     */
    async handleInteractionError(interaction, error) {
        const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
        try {
            if (interaction.isRepliable()) {
                if (interaction.replied || interaction.deferred) {
                    await interaction.editReply({
                        content: `❌ Помилка виконання: ${errorMessage}`
                    });
                }
                else {
                    await interaction.reply({
                        content: `❌ Помилка виконання: ${errorMessage}`,
                        ephemeral: true
                    });
                }
            }
        }
        catch (replyError) {
            logger_1.default.error('❌ Помилка відповіді на помилку interaction:', replyError);
        }
    }
    /**
     * Перевірка rate limit
     */
    isRateLimited(userId) {
        const now = Date.now();
        const userLimit = this.rateLimitMap.get(userId);
        if (!userLimit) {
            this.rateLimitMap.set(userId, { count: 1, resetTime: now + 60000 });
            return false;
        }
        if (now > userLimit.resetTime) {
            this.rateLimitMap.set(userId, { count: 1, resetTime: now + 60000 });
            return false;
        }
        if (userLimit.count >= BOT_CONSTANTS.COMMAND_RATE_LIMIT) {
            this.stats.rateLimitHits++;
            return true;
        }
        userLimit.count++;
        return false;
    }
    /**
     * Обробка rate limit
     */
    async handleRateLimit(interaction) {
        try {
            if (interaction.isRepliable()) {
                await interaction.reply({
                    content: '⚠️ Забагато запитів. Спробуйте пізніше.',
                    ephemeral: true
                });
            }
        }
        catch (error) {
            logger_1.default.error('❌ Помилка обробки rate limit:', error);
        }
    }
    /**
     * Очікування готовності клієнта з таймаутом
     */
    waitForReady() {
        return new Promise((resolve, reject) => {
            if (this.isReady) {
                resolve();
                return;
            }
            const timeout = setTimeout(() => {
                reject(new Error(`Таймаут очікування готовності клієнта (${BOT_CONSTANTS.READY_TIMEOUT}ms)`));
            }, BOT_CONSTANTS.READY_TIMEOUT);
            this.client.once(discord_js_1.Events.ClientReady, () => {
                clearTimeout(timeout);
                resolve();
            });
        });
    }
    /**
     * Перевірка чи потрібно перепідключення
     */
    shouldReconnect(error) {
        const reconnectErrors = [
            'ECONNRESET',
            'ENOTFOUND',
            'ETIMEDOUT',
            'ECONNREFUSED',
            'WebSocket connection was closed',
        ];
        return reconnectErrors.some(errorType => error.message.includes(errorType) || error.name.includes(errorType));
    }
    /**
     * Планування перепідключення
     */
    scheduleReconnect() {
        if (this.reconnectAttempts >= BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS) {
            logger_1.default.error(`❌ Досягнуто максимальну кількість спроб перепідключення (${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS})`);
            return;
        }
        this.reconnectAttempts++;
        logger_1.default.info(`🔄 Планування перепідключення ${this.reconnectAttempts}/${BOT_CONSTANTS.MAX_RECONNECT_ATTEMPTS} через ${BOT_CONSTANTS.RECONNECT_DELAY}ms`);
        setTimeout(() => {
            this.connectToDiscord();
        }, BOT_CONSTANTS.RECONNECT_DELAY);
    }
    /**
     * Запуск health check
     */
    startHealthCheck() {
        this.healthCheckInterval = setInterval(async () => {
            try {
                const health = await this.onHealthCheck();
                if (!health.healthy) {
                    logger_1.default.warn('⚠️ Health check виявив проблеми:', health);
                }
            }
            catch (error) {
                logger_1.default.error('❌ Помилка health check:', error);
            }
        }, BOT_CONSTANTS.HEALTH_CHECK_INTERVAL);
        logger_1.default.info('🏥 Health check запущено');
    }
    /**
     * Зупинка health check
     */
    stopHealthCheck() {
        if (this.healthCheckInterval) {
            clearInterval(this.healthCheckInterval);
            this.healthCheckInterval = null;
            logger_1.default.info('🏥 Health check зупинено');
        }
    }
    /**
     * Очищення ресурсів при помилці
     */
    async cleanupOnError() {
        try {
            logger_1.default.info('🧹 Очищення ресурсів при помилці...');
            this.stopHealthCheck();
            if (this.client) {
                this.client.destroy();
            }
            logger_1.default.info('✅ Ресурси очищено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка очищення ресурсів:', error);
        }
    }
    /**
     * Логування статистики запуску
     */
    logStartupStats() {
        try {
            const stats = this.getStats();
            logger_1.default.info('📊 Статистика запуску бота:', {
                uptime: `${Math.round(stats.uptime / 1000)}s`,
                commands: stats.commands,
                interactions: stats.interactions,
                errors: stats.errors,
                reconnects: stats.reconnects,
                memory: {
                    rss: `${Math.round(stats.memory.rss / 1024 / 1024)}MB`,
                    heapUsed: `${Math.round(stats.memory.heapUsed / 1024 / 1024)}MB`,
                },
                rateLimitHits: stats.rateLimitHits,
                slowCommands: stats.slowCommands,
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка логування статистики запуску:', error);
        }
    }
    /**
     * Реєстрація команди
     */
    registerCommand(command) {
        try {
            const commandName = command.getName();
            this.commands.set(commandName, command);
            logger_1.default.debug(`✅ Команда зареєстрована: ${commandName}`);
        }
        catch (error) {
            logger_1.default.error('❌ Помилка реєстрації команди:', error);
        }
    }
    /**
     * Отримання всіх команд
     */
    getCommands() {
        return this.commands;
    }
    /**
     * Перевірка чи бот готовий
     */
    isBotReady() {
        return this.isReady && this.client.isReady();
    }
    /**
     * Отримання детальної статистики
     */
    getDetailedStats() {
        return {
            ...this.getStats(),
            isReady: this.isReady,
            isConnecting: this.isConnecting,
            reconnectAttempts: this.reconnectAttempts,
        };
    }
}
exports.Bot = Bot;
//# sourceMappingURL=Bot.js.map