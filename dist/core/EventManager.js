"use strict";
/**
 * Event Manager для Discord бота
 * Централізована обробка Discord подій
 * TypeScript версія
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const logger_1 = __importDefault(require("../utils/logger"));
class EventManager {
    constructor(bot) {
        this.bot = bot;
        this.events = new Map();
        this.isActive = false;
    }
    /**
     * Ініціалізація менеджера подій
     */
    async initialize() {
        try {
            logger_1.default.info('📡 Ініціалізація менеджера подій...');
            // Реєстрація стандартних подій
            this.registerDefaultEvents();
            this.isActive = true;
            logger_1.default.info('✅ Менеджер подій ініціалізовано');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка ініціалізації менеджера подій:', error);
            throw error;
        }
    }
    /**
     * Реєстрація стандартних подій
     */
    registerDefaultEvents() {
        // Ready event
        this.registerEvent('ready', () => {
            logger_1.default.info(`🤖 Бот ${this.bot.client.user?.tag} готовий до роботи!`);
            this.bot.client.user?.setActivity('ЗСУ Документи', { type: 3 }); // WATCHING
        });
        // Error event
        this.registerEvent('error', (error) => {
            logger_1.default.error('Discord клієнт помилка:', error);
        });
        // Warn event
        this.registerEvent('warn', (warning) => {
            logger_1.default.warn('Discord клієнт попередження:', warning);
        });
        // Disconnect event
        this.registerEvent('disconnect', () => {
            logger_1.default.warn('Discord клієнт відключено');
        });
        // Reconnecting event
        this.registerEvent('reconnecting', () => {
            logger_1.default.info('Discord клієнт перепідключається...');
        });
        // Guild Create event
        this.registerEvent('guildCreate', (guild) => {
            logger_1.default.info(`📥 Бот додано на сервер: ${guild.name} (${guild.id})`);
        });
        // Guild Delete event
        this.registerEvent('guildDelete', (guild) => {
            logger_1.default.info(`📤 Бот видалено з сервера: ${guild.name} (${guild.id})`);
        });
        // Message Create event (для логування)
        this.registerEvent('messageCreate', (message) => {
            if (message.author.bot)
                return;
            logger_1.default.debug(`💬 Повідомлення від ${message.author.tag}: ${message.content.substring(0, 50)}...`);
        });
        logger_1.default.debug('✅ Стандартні події зареєстровано');
    }
    /**
     * Реєстрація події
     */
    registerEvent(eventName, handler) {
        try {
            if (this.events.has(eventName)) {
                logger_1.default.warn(`Подія "${eventName}" вже зареєстрована, перезаписуємо`);
            }
            // Обгортка для обробки помилок
            const wrappedHandler = async (...args) => {
                try {
                    await handler(...args);
                }
                catch (error) {
                    logger_1.default.error(`Помилка обробки події "${eventName}":`, error);
                }
            };
            this.events.set(eventName, wrappedHandler);
            this.bot.client.on(eventName, wrappedHandler);
            logger_1.default.debug(`✅ Подія "${eventName}" зареєстрована`);
        }
        catch (error) {
            logger_1.default.error(`Помилка реєстрації події "${eventName}":`, error);
        }
    }
    /**
     * Видалення події
     */
    removeEvent(eventName) {
        try {
            const handler = this.events.get(eventName);
            if (handler) {
                this.bot.client.off(eventName, handler);
                this.events.delete(eventName);
                logger_1.default.debug(`✅ Подія "${eventName}" видалена`);
            }
        }
        catch (error) {
            logger_1.default.error(`Помилка видалення події "${eventName}":`, error);
        }
    }
    /**
     * Отримання списку зареєстрованих подій
     */
    getRegisteredEvents() {
        return Array.from(this.events.keys());
    }
    /**
     * Перевірка активності
     */
    isActive() {
        return this.isActive;
    }
    /**
     * Завершення роботи
     */
    async shutdown() {
        logger_1.default.info('🛑 Завершення роботи менеджера подій...');
        try {
            // Видалення всіх подій
            for (const eventName of this.events.keys()) {
                this.removeEvent(eventName);
            }
            this.isActive = false;
            logger_1.default.info('✅ Менеджер подій завершено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка завершення менеджера подій:', error);
        }
    }
}
exports.default = EventManager;
//# sourceMappingURL=EventManager.js.map