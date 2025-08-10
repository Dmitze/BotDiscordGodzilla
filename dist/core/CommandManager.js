"use strict";
/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CommandManager = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("@/utils/logger"));
// Імпорт всіх команд
const SearchCommand_1 = require("@/commands/SearchCommand");
const PerformanceCommand_1 = require("@/commands/PerformanceCommand");
const AIAssistantCommand_1 = require("@/commands/AIAssistantCommand");
const DocumentsCommand_1 = require("@/commands/DocumentsCommand");
const FileManagerCommand_1 = require("@/commands/FileManagerCommand");
const OperationsCommand_1 = require("@/commands/OperationsCommand");
const AnalyticsCommand_1 = require("@/commands/AnalyticsCommand");
const EnhancedSearchCommand_1 = require("@/commands/EnhancedSearchCommand");
class CommandManager {
    constructor(bot, config) {
        this.bot = bot;
        this.config = config;
        this.commands = new discord_js_1.Collection();
        this.commandCategories = new Map();
        this.stats = {
            totalCommands: 0,
            categories: 0,
            commandsByCategory: {},
            lastUsed: new Date()
        };
    }
    /**
     * Ініціалізація менеджера команд
     */
    async initialize() {
        try {
            console.log('📋 Ініціалізація менеджера команд...');
            // Завантаження команд
            await this.loadCommands();
            // Реєстрація обробників подій
            this.registerEventHandlers();
            console.log(`✅ Завантажено ${this.commands.size} команд`);
        }
        catch (error) {
            console.error('❌ Помилка ініціалізації менеджера команд:', error);
            throw error;
        }
    }
    /**
     * Завантаження всіх команд
     */
    async loadCommands() {
        try {
            // Створюємо екземпляри всіх команд
            const commandInstances = [
                new SearchCommand_1.SearchCommand(this.config),
                new PerformanceCommand_1.PerformanceCommand(this.config),
                new AIAssistantCommand_1.AIAssistantCommand(this.config),
                new DocumentsCommand_1.DocumentsCommand(this.config),
                new FileManagerCommand_1.FileManagerCommand(this.config),
                new OperationsCommand_1.OperationsCommand(this.config),
                new AnalyticsCommand_1.AnalyticsCommand(this.config),
                new EnhancedSearchCommand_1.EnhancedSearchCommand(this.config)
            ];
            // Реєструємо команди
            for (const command of commandInstances) {
                if (this.validateCommand(command)) {
                    const commandName = command.getName();
                    this.commands.set(commandName, command);
                    // Категоризація команд
                    const category = this.getCommandCategory(command);
                    if (!this.commandCategories.has(category)) {
                        this.commandCategories.set(category, []);
                    }
                    this.commandCategories.get(category).push(commandName);
                    console.log(`📝 Завантажено команду: ${commandName} (${category})`);
                }
            }
            // Оновлюємо статистику
            this.updateStats();
        }
        catch (error) {
            console.error('❌ Помилка завантаження команд:', error);
            throw error;
        }
    }
    /**
     * Валідація команди
     */
    validateCommand(command) {
        if (!command.getName()) {
            console.warn('Команда не має назви');
            return false;
        }
        if (!command.getDescription()) {
            console.warn(`Команда ${command.getName()} не має опису`);
            return false;
        }
        return true;
    }
    /**
     * Визначення категорії команди
     */
    getCommandCategory(command) {
        const name = command.getName();
        if (name.includes('пошук') || name.includes('search')) {
            return 'Пошук';
        }
        if (name.includes('продуктивність') || name.includes('performance')) {
            return 'Моніторинг';
        }
        if (name.includes('ai') || name.includes('асистент')) {
            return 'AI';
        }
        if (name.includes('документи') || name.includes('documents')) {
            return 'Документи';
        }
        if (name.includes('файли') || name.includes('file')) {
            return 'Файли';
        }
        if (name.includes('операції') || name.includes('operations')) {
            return 'Операції';
        }
        if (name.includes('аналітика') || name.includes('analytics')) {
            return 'Аналітика';
        }
        return 'Інші';
    }
    /**
     * Реєстрація обробників подій
     */
    registerEventHandlers() {
        this.bot.on('interactionCreate', async (interaction) => {
            if (interaction.isChatInputCommand()) {
                await this.handleCommand(interaction);
            }
        });
    }
    /**
     * Обробка команди
     */
    async handleCommand(interaction) {
        try {
            const commandName = interaction.commandName;
            const command = this.commands.get(commandName);
            if (!command) {
                await interaction.reply({
                    content: '❌ Команда не знайдена',
                    ephemeral: true
                });
                return;
            }
            // Оновлюємо статистику
            this.stats.lastUsed = new Date();
            // Перевірка прав доступу
            const hasPermission = await this.checkPermissions(interaction, command);
            if (!hasPermission) {
                await interaction.reply({
                    content: '❌ Недостатньо прав для виконання цієї команди',
                    ephemeral: true
                });
                return;
            }
            // Виконання команди
            await command.execute({
                interaction
            });
            console.log(`✅ Команда ${commandName} виконана користувачем ${interaction.user.tag}`);
        }
        catch (error) {
            console.error(`❌ Помилка виконання команди ${interaction.commandName}:`, error);
            const errorMessage = '❌ Помилка при виконанні команди. Спробуйте ще раз або зверніться до адміністратора.';
            if (interaction.replied || interaction.deferred) {
                await interaction.editReply({ content: errorMessage });
            }
            else {
                await interaction.reply({ content: errorMessage, ephemeral: true });
            }
        }
    }
    /**
     * Перевірка прав доступу
     */
    async checkPermissions(interaction, command) {
        try {
            // Імпорт PermissionManager
            const { PermissionManager } = await Promise.resolve().then(() => __importStar(require('./PermissionManager')));
            const permissionManager = new PermissionManager(this.config);
            // Перевірка прав доступу
            const result = await permissionManager.checkPermission(interaction.user, interaction.member, interaction.commandName, interaction.channelId);
            // Якщо доступ заборонено, відправляємо повідомлення користувачу
            if (!result.allowed) {
                const embed = this.createPermissionDeniedEmbed(result);
                await interaction.reply({ embeds: [embed], ephemeral: true });
                logger_1.default.security('command_access_denied', interaction.user.id, {
                    command: interaction.commandName,
                    reason: result.reason,
                    userLevel: result.userLevel,
                    guildId: interaction.guildId,
                    channelId: interaction.channelId
                });
                return false;
            }
            // Логування успішного доступу
            logger_1.default.info('✅ Команда дозволена', {
                userId: interaction.user.id,
                command: interaction.commandName,
                userLevel: result.userLevel,
                remainingUses: result.remainingUses
            });
            return true;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка перевірки прав доступу:', error);
            // У разі помилки дозволяємо виконання для базових команд
            const allowedCommands = ['пошук', 'довідка', 'статус'];
            return allowedCommands.includes(interaction.commandName);
        }
    }
    /**
     * Створення embed повідомлення про відмову доступу
     */
    createPermissionDeniedEmbed(result) {
        return new discord_js_1.EmbedBuilder()
            .setColor(0xFF0000)
            .setTitle('🚫 Доступ заборонено')
            .setDescription(`Вам заборонено використовувати цю команду.\n\n**Причина:** ${result.reason}`)
            .addFields([
            {
                name: '📊 Ваш рівень доступу',
                value: `${result.userLevel} (${['Заборонений', 'Користувач', 'Довірений', 'Модератор', 'Адміністратор', 'Власник'][result.userLevel]})`,
                inline: true
            },
            {
                name: '🔄 Використання за день',
                value: result.remainingUses ? `Залишилось: ${result.remainingUses}` : 'Інформація недоступна',
                inline: true
            },
            {
                name: '📞 Зв\'яжіться з адміністратором',
                value: 'Якщо вважаєте, що це помилка, зверніться до адміністрації сервера.',
                inline: false
            }
        ])
            .setFooter({ text: 'Discord AI Assistant Bot - Security System' })
            .setTimestamp();
    }
    /**
     * Отримання команди за назвою
     */
    getCommand(name) {
        return this.commands.get(name);
    }
    /**
     * Отримання всіх команд
     */
    getAllCommands() {
        return this.commands;
    }
    /**
     * Отримання команд за категорією
     */
    getCommandsByCategory(category) {
        return this.commandCategories.get(category) || [];
    }
    /**
     * Отримання всіх категорій
     */
    getCategories() {
        return Array.from(this.commandCategories.keys());
    }
    /**
     * Отримання статистики
     */
    getStats() {
        return { ...this.stats };
    }
    /**
     * Оновлення статистики
     */
    updateStats() {
        this.stats.totalCommands = this.commands.size;
        this.stats.categories = this.commandCategories.size;
        this.stats.commandsByCategory = {};
        for (const [category, commands] of this.commandCategories.entries()) {
            this.stats.commandsByCategory[category] = commands.length;
        }
    }
    /**
     * Отримання даних для реєстрації команд в Discord
     */
    getCommandsData() {
        return Array.from(this.commands.values()).map(command => command.getData());
    }
    /**
     * Перезавантаження команд
     */
    async reloadCommands() {
        console.log('🔄 Перезавантаження команд...');
        this.commands.clear();
        this.commandCategories.clear();
        await this.loadCommands();
        console.log(`✅ Перезавантажено ${this.commands.size} команд`);
    }
}
exports.CommandManager = CommandManager;
//# sourceMappingURL=CommandManager.js.map