/**
 * Клас для управління конфігурацією додатку
 * Завантажує та валідує налаштування з змінних середовища
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { BotConfig } from '@/types';
export declare class Config {
    private static instance;
    private static readonly configCache;
    /**
     * Завантаження конфігурації з змінних середовища (Singleton pattern)
     */
    static load(): BotConfig;
    /**
     * Завантаження Discord конфігурації
     */
    private static loadDiscordConfig;
    /**
     * Парсинг Discord intents
     */
    private static parseIntents;
    /**
     * Завантаження Google конфігурації
     */
    private static loadGoogleConfig;
    /**
     * Завантаження Google credentials з файлу або змінних середовища
     */
    private static loadGoogleCredentials;
    /**
     * Завантаження AI конфігурації
     */
    private static loadAIConfig;
    /**
     * Завантаження Redis конфігурації
     */
    private static loadRedisConfig;
    /**
     * Завантаження Metrics конфігурації
     */
    private static loadMetricsConfig;
    /**
     * Завантаження Security конфігурації
     */
    private static loadSecurityConfig;
    /**
     * Завантаження Performance конфігурації
     */
    private static loadPerformanceConfig;
    /**
     * Завантаження Logging конфігурації
     */
    private static loadLoggingConfig;
    /**
     * Валідація числових значень
     */
    private static validateNumber;
    /**
     * Валідація конфігурації
     */
    private static validate;
    /**
     * Логування підсумку конфігурації
     */
    private static logConfigurationSummary;
    /**
     * Отримання обов'язкової змінної середовища
     */
    private static getRequiredEnv;
    /**
     * Отримання змінної середовища з значенням за замовчуванням
     */
    private static getEnv;
    /**
     * Очищення кешу конфігурації
     */
    static clearCache(): void;
    /**
     * Перезавантаження конфігурації
     */
    static reload(): BotConfig;
}
//# sourceMappingURL=Config.d.ts.map