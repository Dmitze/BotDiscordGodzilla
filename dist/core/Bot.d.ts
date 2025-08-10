/**
 * Основний клас Discord бота
 * Управляє всіма компонентами та сервісами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import { Client, Collection } from 'discord.js';
import type { BotConfig, BaseCommand } from '@/types';
import { ServiceContainer } from './ServiceContainer';
import { BaseService as BaseServiceClass } from './BaseService';
import { CommandManager } from './CommandManager';
import { ErrorHandler } from './ErrorHandler';
import { EventManager } from './EventManager';
import { ServiceManager } from './ServiceManager';
interface BotStats {
    uptime: number;
    commands: number;
    interactions: number;
    errors: number;
    reconnects: number;
    lastActivity: Date;
    memory: NodeJS.MemoryUsage;
    rateLimitHits: number;
    slowCommands: number;
}
export declare class Bot extends BaseServiceClass {
    readonly client: Client;
    readonly serviceContainer: ServiceContainer;
    readonly commandManager: CommandManager;
    readonly errorHandler: ErrorHandler;
    readonly eventManager: EventManager;
    readonly serviceManager: ServiceManager;
    private commands;
    private isReady;
    private isConnecting;
    private reconnectAttempts;
    private stats;
    private healthCheckInterval;
    private lastInteractionTime;
    private rateLimitMap;
    private slowCommandThreshold;
    constructor(config: BotConfig);
    /**
     * Ініціалізація бота з детальним логуванням
     */
    protected onInitialize(): Promise<void>;
    /**
     * Завершення роботи бота з детальним логуванням
     */
    protected onShutdown(): Promise<void>;
    /**
     * Health check бота з розширеною інформацією
     */
    protected onHealthCheck(): Promise<{
        healthy: boolean;
        service: string;
        error?: string;
        details?: Record<string, unknown>;
    }>;
    /**
     * Отримання детальної статистики бота
     */
    protected onGetStats(): BotStats;
    /**
     * Перевірка системних ресурсів
     */
    private checkSystemResources;
    /**
     * Підключення до Discord з обробкою помилок
     */
    private connectToDiscord;
    /**
     * Налаштування обробників подій з детальним логуванням
     */
    private setupEventHandlers;
    /**
     * Обробка команд з детальним логуванням
     */
    private handleCommand;
    /**
     * Обробка кнопкових interactions
     */
    private handleButtonInteraction;
    /**
     * Обробка select menu interactions
     */
    private handleSelectMenuInteraction;
    /**
     * Обробка помилок interactions
     */
    private handleInteractionError;
    /**
     * Перевірка rate limit
     */
    private isRateLimited;
    /**
     * Обробка rate limit
     */
    private handleRateLimit;
    /**
     * Очікування готовності клієнта з таймаутом
     */
    private waitForReady;
    /**
     * Перевірка чи потрібно перепідключення
     */
    private shouldReconnect;
    /**
     * Планування перепідключення
     */
    private scheduleReconnect;
    /**
     * Запуск health check
     */
    private startHealthCheck;
    /**
     * Зупинка health check
     */
    private stopHealthCheck;
    /**
     * Очищення ресурсів при помилці
     */
    private cleanupOnError;
    /**
     * Логування статистики запуску
     */
    private logStartupStats;
    /**
     * Реєстрація команди
     */
    registerCommand(command: BaseCommand): void;
    /**
     * Отримання всіх команд
     */
    getCommands(): Collection<string, BaseCommand>;
    /**
     * Перевірка чи бот готовий
     */
    isBotReady(): boolean;
    /**
     * Отримання детальної статистики
     */
    getDetailedStats(): BotStats & {
        isReady: boolean;
        isConnecting: boolean;
        reconnectAttempts: number;
    };
}
export {};
//# sourceMappingURL=Bot.d.ts.map