/**
 * Базовий абстрактний клас для всіх команд Discord бота
 * Забезпечує уніфіковану структуру та типізацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import { SlashCommandBuilder, ChatInputCommandInteraction, AutocompleteInteraction, MessageComponentInteraction } from 'discord.js';
import type { BotConfig, CommandOptions, CommandStats, CommandContext, HealthStatus } from '@/types';
export interface CommandData {
    name: string;
    description: string;
    options?: any[];
    defaultMemberPermissions?: string | number;
    dmPermission?: boolean;
    cooldown?: number;
    permissions?: string[];
    category?: string;
    usage?: string;
    examples?: string[];
}
export interface CommandExecuteOptions {
    interaction: ChatInputCommandInteraction;
    context?: CommandContext;
    options?: CommandOptions;
    startTime?: number;
    retryCount?: number;
}
export interface CommandAutocompleteOptions {
    interaction: AutocompleteInteraction;
    context?: CommandContext;
    query?: string;
}
export interface CommandComponentOptions {
    interaction: MessageComponentInteraction;
    context?: CommandContext;
    componentType?: 'button' | 'select' | 'modal';
}
export interface CommandValidationResult {
    isValid: boolean;
    errors: string[];
    warnings: string[];
    sanitizedOptions?: any;
}
export declare abstract class BaseCommand {
    readonly data: SlashCommandBuilder;
    readonly name: string;
    readonly description: string;
    readonly category: string;
    readonly usage: string;
    readonly examples: string[];
    readonly permissions: string[];
    readonly cooldown: number;
    protected stats: CommandStats;
    protected cooldowns: Map<string, number>;
    protected executionCache: Map<string, {
        result: any;
        timestamp: number;
    }>;
    protected errorCount: Map<string, number>;
    protected lastExecution: Map<string, number>;
    protected readonly config: BotConfig;
    protected isShuttingDown: boolean;
    constructor(name: string, description: string, config: BotConfig, options?: Partial<CommandData>, builder?: (builder: SlashCommandBuilder) => SlashCommandBuilder);
    /**
     * Виконання команди з детальним логуванням та обробкою помилок
     */
    execute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Виконання команди з retry логікою
     */
    private executeWithRetry;
    /**
     * Валідація виконання команди
     */
    private validateExecution;
    /**
     * Обробка автодоповнення з детальним логуванням
     */
    autocomplete(options: CommandAutocompleteOptions): Promise<void>;
    /**
     * Обробка компонентів з детальним логуванням
     */
    handleComponent(options: CommandComponentOptions): Promise<void>;
    /**
     * Абстрактний метод виконання команди
     */
    protected abstract onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Обробка автодоповнення (опціонально)
     */
    protected onAutocomplete(options: CommandAutocompleteOptions): Promise<void>;
    /**
     * Обробка компонентів (опціонально)
     */
    protected onComponent(options: CommandComponentOptions): Promise<void>;
    /**
     * Валідація cooldown
     */
    private validateCooldown;
    /**
     * Перевірка cooldown
     */
    protected isOnCooldown(userId: string): boolean;
    /**
     * Встановлення cooldown
     */
    protected setCooldown(userId: string): void;
    /**
     * Отримання часу cooldown
     */
    protected getCooldownTime(userId: string): number;
    /**
     * Генерація ключа кешу
     */
    private generateCacheKey;
    /**
     * Отримання кешованого результату
     */
    private getCachedResult;
    /**
     * Кешування результату
     */
    private cacheResult;
    /**
     * Збільшення лічильника помилок
     */
    private incrementErrorCount;
    /**
     * Обробка cooldown
     */
    protected handleCooldown(interaction: ChatInputCommandInteraction): Promise<void>;
    /**
     * Обробка кешованого результату
     */
    private handleCachedResult;
    /**
     * Обробка помилки валідації
     */
    private handleValidationError;
    /**
     * Обробка помилки зупинки
     */
    private handleShutdownError;
    /**
     * Обробка помилок
     */
    protected handleError(interaction: ChatInputCommandInteraction, error: unknown): Promise<void>;
    /**
     * Обробка помилок автодоповнення
     */
    protected handleAutocompleteError(interaction: AutocompleteInteraction, error: unknown): Promise<void>;
    /**
     * Обробка помилок компонентів
     */
    protected handleComponentError(interaction: MessageComponentInteraction, error: unknown): Promise<void>;
    /**
     * Логування початку команди
     */
    protected logCommandStart(interaction: ChatInputCommandInteraction): void;
    /**
     * Логування успішного завершення
     */
    protected logCommandSuccess(interaction: ChatInputCommandInteraction, duration: number): void;
    /**
     * Логування помилки команди
     */
    protected logCommandError(interaction: ChatInputCommandInteraction, error: unknown): void;
    /**
     * Логування помилки автодоповнення
     */
    protected logAutocompleteError(interaction: AutocompleteInteraction, error: unknown): void;
    /**
     * Логування помилки компонента
     */
    protected logComponentError(interaction: MessageComponentInteraction, error: unknown): void;
    /**
     * Оновлення статистики
     */
    protected updateStats(success: boolean, duration: number): void;
    /**
     * Запуск періодичного очищення
     */
    private startCleanupInterval;
    /**
     * Очищення застарілих даних
     */
    private cleanupExpiredData;
    /**
     * Отримання статистики команди
     */
    getCommandStats(): CommandStats;
    /**
     * Очищення cooldowns
     */
    clearCooldowns(): void;
    /**
     * Health check
     */
    healthCheck(): Promise<HealthStatus>;
    /**
     * Завершення роботи
     */
    shutdown(): Promise<void>;
    /**
     * Отримання статистики
     */
    getStats(): CommandStats;
    /**
     * Отримання назви команди
     */
    getName(): string;
    /**
     * Отримання опису команди
     */
    getDescription(): string;
    /**
     * Отримання даних команди для реєстрації в Discord
     */
    getData(): SlashCommandBuilder;
    /**
     * Отримання допомоги по команді
     */
    getHelp(): string;
}
//# sourceMappingURL=BaseCommand.d.ts.map