/**
 * Розширена система безпеки для Discord AI Assistant Bot
 * Валідація, санітизація та захист від атак
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { SecurityEvent, SecurityValidationResult } from '@/types';
export interface SecurityStats {
    totalValidations: number;
    successfulValidations: number;
    failedValidations: number;
    suspiciousActivities: number;
    rateLimitHits: number;
    blacklistHits: number;
    xssAttempts: number;
    sqlInjectionAttempts: number;
    lastSecurityEvent?: SecurityEvent;
    averageValidationTime: number;
    totalValidationTime: number;
}
export interface RateLimitInfo {
    count: number;
    resetTime: number;
    lastRequest: number;
}
export declare class SecurityManager {
    private static instance;
    private stats;
    private rateLimitMap;
    private blacklistCache;
    private suspiciousActivities;
    private _isInitialized;
    constructor();
    /**
     * Ініціалізація системи безпеки
     */
    private initialize;
    /**
     * Завантаження чорного списку
     */
    private loadBlacklist;
    /**
     * Запуск періодичних завдань
     */
    private startPeriodicTasks;
    /**
     * Валідація та санітизація введення
     */
    validateInput(input: string, context?: {
        userId?: string;
        guildId?: string;
        channelId?: string;
        commandName?: string;
        inputType?: 'command' | 'message' | 'url' | 'file';
    }): SecurityValidationResult;
    /**
     * Перевірка на XSS атаки
     */
    private checkForXSS;
    /**
     * Перевірка на SQL ін'єкції
     */
    private checkForSQLInjection;
    /**
     * Перевірка чорного списку
     */
    private checkBlacklist;
    /**
     * Санітизація введення
     */
    private sanitizeInput;
    /**
     * Перевірка rate limit
     */
    checkRateLimit(userId: string): {
        allowed: boolean;
        remaining: number;
        resetTime: number;
    };
    /**
     * Валідація URL
     */
    validateUrl(url: string): SecurityValidationResult;
    /**
     * Запис події безпеки
     */
    private recordSecurityEvent;
    /**
     * Визначення серйозності події
     */
    private determineEventSeverity;
    /**
     * Очищення rate limit кешу
     */
    private cleanupRateLimitCache;
    /**
     * Очищення підозрілої активності
     */
    private cleanupSuspiciousActivities;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Отримання статистики безпеки
     */
    getStats(): SecurityStats;
    /**
     * Отримання підозрілої активності
     */
    getSuspiciousActivities(): SecurityEvent[];
    /**
     * Очищення ресурсів
     */
    cleanup(): void;
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized(): boolean;
}
export declare const securityManager: SecurityManager;
export declare const validateInput: (input: string, context?: {
    userId?: string;
    guildId?: string;
    channelId?: string;
    commandName?: string;
    inputType?: "command" | "message" | "url" | "file";
}) => SecurityValidationResult;
export declare const checkRateLimit: (userId: string) => {
    allowed: boolean;
    remaining: number;
    resetTime: number;
};
export declare const validateUrl: (url: string) => SecurityValidationResult;
export declare const getSecurityStats: () => SecurityStats;
export declare const getSuspiciousActivities: () => SecurityEvent[];
export declare const cleanupSecurityManager: () => void;
export declare const sanitizeInput: (input: string) => string;
export declare const validateCommandOptions: (options: any) => SecurityValidationResult;
//# sourceMappingURL=security.d.ts.map