/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */
import { Collection } from 'discord.js';
import type { BotConfig } from '@/types';
import { BaseCommand } from '@/commands/BaseCommand';
interface CommandStats {
    totalCommands: number;
    categories: number;
    commandsByCategory: Record<string, number>;
    lastUsed: Date;
}
export declare class CommandManager {
    private bot;
    private config;
    private commands;
    private commandCategories;
    private stats;
    constructor(bot: any, config: BotConfig);
    /**
     * Ініціалізація менеджера команд
     */
    initialize(): Promise<void>;
    /**
     * Завантаження всіх команд
     */
    private loadCommands;
    /**
     * Валідація команди
     */
    private validateCommand;
    /**
     * Визначення категорії команди
     */
    private getCommandCategory;
    /**
     * Реєстрація обробників подій
     */
    private registerEventHandlers;
    /**
     * Обробка команди
     */
    private handleCommand;
    /**
     * Перевірка прав доступу
     */
    private checkPermissions;
    /**
     * Створення embed повідомлення про відмову доступу
     */
    private createPermissionDeniedEmbed;
    /**
     * Отримання команди за назвою
     */
    getCommand(name: string): BaseCommand | undefined;
    /**
     * Отримання всіх команд
     */
    getAllCommands(): Collection<string, BaseCommand>;
    /**
     * Отримання команд за категорією
     */
    getCommandsByCategory(category: string): string[];
    /**
     * Отримання всіх категорій
     */
    getCategories(): string[];
    /**
     * Отримання статистики
     */
    getStats(): CommandStats;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Отримання даних для реєстрації команд в Discord
     */
    getCommandsData(): any[];
    /**
     * Перезавантаження команд
     */
    reloadCommands(): Promise<void>;
}
export {};
//# sourceMappingURL=CommandManager.d.ts.map