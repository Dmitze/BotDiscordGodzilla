/**
 * Система управління правами доступу для Discord AI Assistant Bot
 * Забезпечує гранульний контроль доступу до команд та функцій
 * Версія 1.0.0 - Нова реалізація
 */
import { GuildMember, User, PermissionResolvable } from 'discord.js';
import type { BotConfig } from '@/types';
export interface PermissionConfig {
    command: string;
    requiredRoles: string[];
    requiredPermissions: PermissionResolvable[];
    minUserLevel: UserLevel;
    allowedChannels?: string[];
    deniedChannels?: string[];
    cooldown?: number;
    maxUsesPerDay?: number;
}
export declare enum UserLevel {
    BANNED = 0,
    USER = 1,
    TRUSTED = 2,
    MODERATOR = 3,
    ADMIN = 4,
    OWNER = 5
}
export interface PermissionCheckResult {
    allowed: boolean;
    reason?: string;
    userLevel: UserLevel;
    hasRequiredRoles: boolean;
    hasRequiredPermissions: boolean;
    canUseInChannel: boolean;
    remainingUses?: number;
}
export declare class PermissionManager {
    private static instance;
    private permissionConfigs;
    private userCache;
    private commandUsage;
    private isInitialized;
    constructor(_config: BotConfig);
    /**
     * Ініціалізація системи прав
     */
    private initialize;
    /**
     * Завантаження конфігурацій прав для команд
     */
    private loadPermissionConfigs;
    /**
     * Головна функція перевірки прав доступу
     */
    checkPermission(user: User, member: GuildMember | null, commandName: string, channelId?: string): Promise<PermissionCheckResult>;
    /**
     * Отримання інформації про користувача
     */
    private getUserInfo;
    /**
     * Визначення рівня користувача
     */
    private determineUserLevel;
    /**
     * Перевірка ролей
     */
    private checkRequiredRoles;
    /**
     * Перевірка дозволів Discord
     */
    private checkRequiredPermissions;
    /**
     * Перевірка дозволів для каналу
     */
    private checkChannelPermissions;
    /**
     * Перевірка використання команди
     */
    private checkCommandUsage;
    /**
     * Запис використання команди
     */
    private recordCommandUsage;
    /**
     * Створення повідомлення про відмову
     */
    private buildDenialReason;
    /**
     * Логування подій безпеки
     */
    private logSecurityEvent;
    /**
     * Очищення найстаріших записів з кешу
     */
    private cleanupOldestCacheEntries;
    /**
     * Запуск періодичного очищення кешу
     */
    private startCacheCleanup;
    /**
     * Додавання нової конфігурації прав
     */
    addPermissionConfig(config: PermissionConfig): void;
    /**
     * Отримання статистики системи прав
     */
    getStats(): {
        cachedUsers: number;
        commandConfigs: number;
        dailyUsage: number;
        isInitialized: boolean;
    };
    /**
     * Очищення ресурсів
     */
    cleanup(): void;
}
export default PermissionManager;
//# sourceMappingURL=PermissionManager.d.ts.map