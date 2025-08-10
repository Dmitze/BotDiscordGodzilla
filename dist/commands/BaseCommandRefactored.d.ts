/**
 * Рефакторований базовий клас для команд Discord бота
 * Використовує модульну архітектуру для кращої підтримки
 * Версія 4.0.0 - Модульна архітектура
 */
import { SlashCommandBuilder, ChatInputCommandInteraction, EmbedBuilder, AutocompleteInteraction, MessageComponentInteraction } from 'discord.js';
import type { BotConfig, CommandStats, CommandContext } from '@/types';
import CommandValidator, { type ValidationResult, type ValidationRules } from './modules/CommandValidator';
import CommandMetricsCollector from './modules/CommandMetrics';
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
    startTime?: number;
    retryCount?: number;
    validationResult?: ValidationResult;
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
    protected readonly config: BotConfig;
    protected isShuttingDown: boolean;
    protected validator: CommandValidator;
    protected metrics: CommandMetricsCollector;
    constructor(commandData: CommandData, config: BotConfig);
    /**
     * Головна точка входу для виконання команди
     */
    handleInteraction(interaction: ChatInputCommandInteraction): Promise<void>;
    /**
     * Валідація взаємодії
     */
    protected validateInteraction(interaction: ChatInputCommandInteraction, customRules?: ValidationRules): Promise<ValidationResult>;
    /**
     * Кастомна валідація для конкретної команди
     */
    protected customValidation(interaction: ChatInputCommandInteraction): Promise<ValidationResult>;
    /**
     * Виконання з повторними спробами
     */
    private executeWithRetry;
    /**
     * Перевірка чи потрібно повторити виконання
     */
    protected shouldRetry(error: unknown): boolean;
    /**
     * Управління cooldown
     */
    protected isOnCooldown(userId: string): boolean;
    protected setCooldown(userId: string): void;
    protected getRemainingCooldown(userId: string): number;
    /**
     * Відправка повідомлень про помилки
     */
    protected sendCooldownMessage(interaction: ChatInputCommandInteraction, remainingTime: number): Promise<void>;
    protected sendValidationError(interaction: ChatInputCommandInteraction, validation: ValidationResult): Promise<void>;
    protected handleExecutionError(interaction: ChatInputCommandInteraction, error: unknown): Promise<void>;
    /**
     * Оновлення статистики
     */
    protected updateStats(executionTime: number, success: boolean): void;
    /**
     * Додавання опцій до команди
     */
    private addOptions;
    /**
     * Отримання статистики команди
     */
    getStats(): CommandStats;
    /**
     * Скидання статистики
     */
    resetStats(): void;
    /**
     * Створення стандартного embed відповіді
     */
    protected createEmbed(title: string, description: string, color?: number): EmbedBuilder;
    /**
     * Перевірка дозволів
     */
    protected hasPermission(interaction: ChatInputCommandInteraction, permission: string): boolean;
    /**
     * Shutdown hook для очищення ресурсів
     */
    shutdown(): void;
    abstract execute(options: CommandExecuteOptions): Promise<void>;
    handleAutocomplete?(interaction: AutocompleteInteraction): Promise<void>;
    handleComponent?(interaction: MessageComponentInteraction): Promise<void>;
}
export default BaseCommand;
//# sourceMappingURL=BaseCommandRefactored.d.ts.map