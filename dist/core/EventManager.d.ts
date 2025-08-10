/**
 * Event Manager для Discord бота
 * Централізована обробка Discord подій
 * TypeScript версія
 */
import { Client } from 'discord.js';
interface Bot {
    client: Client;
}
type EventHandler = (...args: any[]) => Promise<void> | void;
declare class EventManager {
    private bot;
    private events;
    private isActive;
    constructor(bot: Bot);
    /**
     * Ініціалізація менеджера подій
     */
    initialize(): Promise<void>;
    /**
     * Реєстрація стандартних подій
     */
    private registerDefaultEvents;
    /**
     * Реєстрація події
     */
    registerEvent(eventName: string, handler: EventHandler): void;
    /**
     * Видалення події
     */
    removeEvent(eventName: string): void;
    /**
     * Отримання списку зареєстрованих подій
     */
    getRegisteredEvents(): string[];
    /**
     * Перевірка активності
     */
    isActive(): boolean;
    /**
     * Завершення роботи
     */
    shutdown(): Promise<void>;
}
export default EventManager;
//# sourceMappingURL=EventManager.d.ts.map