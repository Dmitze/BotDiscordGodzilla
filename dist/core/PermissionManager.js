"use strict";
/**
 * Система управління правами доступу для Discord AI Assistant Bot
 * Забезпечує гранульний контроль доступу до команд та функцій
 * Версія 1.0.0 - Нова реалізація
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.PermissionManager = exports.UserLevel = void 0;
const discord_js_1 = require("discord.js");
const logger_1 = __importDefault(require("@/utils/logger"));
// Константи для системи прав
const PERMISSION_CONSTANTS = {
    CACHE_TTL: 300000, // 5 хвилин
    MAX_CACHE_ENTRIES: 1000,
    PERMISSION_TIMEOUT: 5000, // 5 секунд
    ADMIN_BYPASS: true,
};
var UserLevel;
(function (UserLevel) {
    UserLevel[UserLevel["BANNED"] = 0] = "BANNED";
    UserLevel[UserLevel["USER"] = 1] = "USER";
    UserLevel[UserLevel["TRUSTED"] = 2] = "TRUSTED";
    UserLevel[UserLevel["MODERATOR"] = 3] = "MODERATOR";
    UserLevel[UserLevel["ADMIN"] = 4] = "ADMIN";
    UserLevel[UserLevel["OWNER"] = 5] = "OWNER";
})(UserLevel || (exports.UserLevel = UserLevel = {}));
class PermissionManager {
    constructor(_config) {
        this.permissionConfigs = new Map();
        this.userCache = new Map();
        this.commandUsage = new Map();
        this.isInitialized = false;
        if (PermissionManager.instance) {
            return PermissionManager.instance;
        }
        PermissionManager.instance = this;
        this.initialize();
    }
    /**
     * Ініціалізація системи прав
     */
    initialize() {
        try {
            logger_1.default.info('🔐 Ініціалізація системи прав доступу...');
            // Завантаження конфігурацій прав для команд
            this.loadPermissionConfigs();
            // Запуск періодичного очищення кешу
            this.startCacheCleanup();
            this.isInitialized = true;
            logger_1.default.info('✅ Система прав доступу ініціалізована');
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка ініціалізації системи прав: ${error instanceof Error ? error.message : String(error)}`);
            throw error;
        }
    }
    /**
     * Завантаження конфігурацій прав для команд
     */
    loadPermissionConfigs() {
        // Конфігурації для різних команд
        const configs = [
            {
                command: 'пошук',
                requiredRoles: ['Bot User'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.ViewChannel],
                minUserLevel: UserLevel.USER,
                cooldown: 5000,
                maxUsesPerDay: 100
            },
            {
                command: 'ai-асистент',
                requiredRoles: ['Bot User'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.ViewChannel],
                minUserLevel: UserLevel.TRUSTED,
                cooldown: 10000,
                maxUsesPerDay: 50
            },
            {
                command: 'файли',
                requiredRoles: ['Bot User', 'File Manager'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.ViewChannel, discord_js_1.PermissionFlagsBits.AttachFiles],
                minUserLevel: UserLevel.TRUSTED,
                cooldown: 3000,
                maxUsesPerDay: 200
            },
            {
                command: 'аналітика',
                requiredRoles: ['Analytics User'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.ViewChannel],
                minUserLevel: UserLevel.TRUSTED,
                cooldown: 15000,
                maxUsesPerDay: 20
            },
            {
                command: 'операції',
                requiredRoles: ['Admin', 'Operations'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.ManageGuild],
                minUserLevel: UserLevel.MODERATOR,
                cooldown: 30000,
                maxUsesPerDay: 10
            },
            {
                command: 'продуктивність',
                requiredRoles: ['Admin'],
                requiredPermissions: [discord_js_1.PermissionFlagsBits.Administrator],
                minUserLevel: UserLevel.ADMIN,
                cooldown: 60000,
                maxUsesPerDay: 5
            }
        ];
        configs.forEach(config => {
            this.permissionConfigs.set(config.command, config);
        });
        logger_1.default.info(`📋 Завантажено ${configs.length} конфігурацій прав доступу`);
    }
    /**
     * Головна функція перевірки прав доступу
     */
    async checkPermission(user, member, commandName, channelId) {
        try {
            const startTime = Date.now();
            // Отримання конфігурації команди
            const permConfig = this.permissionConfigs.get(commandName);
            if (!permConfig) {
                logger_1.default.warn(`⚠️ Конфігурація прав для команди "${commandName}" не знайдена`);
                return {
                    allowed: true, // За замовчуванням дозволяємо невідомі команди
                    reason: 'Конфігурація не знайдена',
                    userLevel: UserLevel.USER,
                    hasRequiredRoles: true,
                    hasRequiredPermissions: true,
                    canUseInChannel: true
                };
            }
            // Якщо member відсутній, команда виконується в DM
            if (!member) {
                return {
                    allowed: false,
                    reason: 'Команда доступна тільки на сервері',
                    userLevel: UserLevel.USER,
                    hasRequiredRoles: false,
                    hasRequiredPermissions: false,
                    canUseInChannel: false
                };
            }
            // Отримання інформації про користувача з кешу або Discord API
            const userInfo = await this.getUserInfo(user.id, member);
            // Перевірка рівня користувача
            if (userInfo.userLevel < permConfig.minUserLevel) {
                this.logSecurityEvent('insufficient_user_level', user.id, {
                    command: commandName,
                    userLevel: userInfo.userLevel,
                    requiredLevel: permConfig.minUserLevel
                });
                return {
                    allowed: false,
                    reason: `Недостатній рівень користувача (потрібен ${UserLevel[permConfig.minUserLevel]})`,
                    userLevel: userInfo.userLevel,
                    hasRequiredRoles: false,
                    hasRequiredPermissions: false,
                    canUseInChannel: true
                };
            }
            // Перевірка ролей
            const hasRequiredRoles = this.checkRequiredRoles(member, permConfig.requiredRoles);
            // Перевірка дозволів Discord
            const hasRequiredPermissions = this.checkRequiredPermissions(member, permConfig.requiredPermissions);
            // Перевірка каналу
            const canUseInChannel = this.checkChannelPermissions(channelId, permConfig);
            // Перевірка використання команди (rate limiting + daily limits)
            const usageCheck = this.checkCommandUsage(user.id, commandName, permConfig);
            // Загальна перевірка
            const allowed = hasRequiredRoles &&
                hasRequiredPermissions &&
                canUseInChannel &&
                usageCheck.allowed;
            const resultBase = {
                allowed,
                userLevel: userInfo.userLevel,
                hasRequiredRoles,
                hasRequiredPermissions,
                canUseInChannel,
                remainingUses: usageCheck.remainingUses
            };
            if (!allowed) {
                const denialObj = {
                    hasRequiredRoles,
                    hasRequiredPermissions,
                    canUseInChannel,
                    usageAllowed: usageCheck.allowed
                };
                if (usageCheck.reason !== undefined) {
                    denialObj.usageReason = usageCheck.reason;
                }
                const reasonStr = this.buildDenialReason(denialObj);
                resultBase.reason = reasonStr;
            }
            const result = resultBase;
            // Логування результату
            const duration = Date.now() - startTime;
            if (allowed) {
                this.recordCommandUsage(user.id, commandName);
                logger_1.default.debug(`✅ Доступ дозволено user=${user.id} command=${commandName} duration=${duration}ms level=${UserLevel[userInfo.userLevel]}`);
            }
            else {
                this.logSecurityEvent('access_denied', user.id, {
                    command: commandName,
                    reason: result.reason ?? '',
                    duration: `${duration}ms`
                });
            }
            return result;
        }
        catch (error) {
            logger_1.default.error(`❌ Помилка перевірки прав доступу: user=${user.id} command=${commandName} error=${error instanceof Error ? error.message : String(error)}`);
            // У разі помилки дозволяємо доступ для базових команд
            return {
                allowed: ['пошук', 'довідка'].includes(commandName),
                reason: 'Помилка системи прав доступу',
                userLevel: UserLevel.USER,
                hasRequiredRoles: false,
                hasRequiredPermissions: false,
                canUseInChannel: true
            };
        }
    }
    /**
     * Отримання інформації про користувача
     */
    async getUserInfo(userId, member) {
        const cacheKey = `${member.guild.id}:${userId}`;
        const cached = this.userCache.get(cacheKey);
        // Перевірка кешу
        if (cached && Date.now() - cached.timestamp < PERMISSION_CONSTANTS.CACHE_TTL) {
            cached.lastActivity = Date.now();
            return cached;
        }
        // Отримання інформації з Discord
        const userLevel = this.determineUserLevel(member);
        const roles = member.roles.cache.map(role => role.name);
        const permissions = member.permissions.toArray();
        const userInfo = {
            userId,
            guildId: member.guild.id,
            userLevel,
            roles,
            permissions,
            timestamp: Date.now(),
            lastActivity: Date.now()
        };
        // Кешування з обмеженням розміру
        if (this.userCache.size >= PERMISSION_CONSTANTS.MAX_CACHE_ENTRIES) {
            this.cleanupOldestCacheEntries();
        }
        this.userCache.set(cacheKey, userInfo);
        return userInfo;
    }
    /**
     * Визначення рівня користувача
     */
    determineUserLevel(member) {
        // Перевірка на власника сервера
        if (member.guild.ownerId === member.id) {
            return UserLevel.OWNER;
        }
        // Перевірка на адміністратора
        if (member.permissions.has(discord_js_1.PermissionFlagsBits.Administrator)) {
            return UserLevel.ADMIN;
        }
        // Перевірка на модератора
        if (member.permissions.has([
            discord_js_1.PermissionFlagsBits.ManageMessages,
            discord_js_1.PermissionFlagsBits.ManageChannels,
            discord_js_1.PermissionFlagsBits.ManageRoles
        ])) {
            return UserLevel.MODERATOR;
        }
        // Перевірка на довірених користувачів
        const trustedRoles = ['Trusted', 'VIP', 'Premium', 'Verified'];
        if (member.roles.cache.some(role => trustedRoles.includes(role.name))) {
            return UserLevel.TRUSTED;
        }
        // Перевірка на заборонених користувачів
        const bannedRoles = ['Banned', 'Muted', 'Restricted'];
        if (member.roles.cache.some(role => bannedRoles.includes(role.name))) {
            return UserLevel.BANNED;
        }
        return UserLevel.USER;
    }
    /**
     * Перевірка ролей
     */
    checkRequiredRoles(member, requiredRoles) {
        if (requiredRoles.length === 0)
            return true;
        // Адміністратори завжди мають доступ
        if (PERMISSION_CONSTANTS.ADMIN_BYPASS &&
            member.permissions.has(discord_js_1.PermissionFlagsBits.Administrator)) {
            return true;
        }
        const memberRoles = member.roles.cache.map(role => role.name);
        return requiredRoles.some(requiredRole => memberRoles.includes(requiredRole));
    }
    /**
     * Перевірка дозволів Discord
     */
    checkRequiredPermissions(member, requiredPermissions) {
        if (requiredPermissions.length === 0)
            return true;
        return member.permissions.has(requiredPermissions);
    }
    /**
     * Перевірка дозволів для каналу
     */
    checkChannelPermissions(channelId, config) {
        if (!channelId)
            return true;
        // Перевірка дозволених каналів
        if (config.allowedChannels && config.allowedChannels.length > 0) {
            return config.allowedChannels.includes(channelId);
        }
        // Перевірка заборонених каналів
        if (config.deniedChannels && config.deniedChannels.length > 0) {
            return !config.deniedChannels.includes(channelId);
        }
        return true;
    }
    /**
     * Перевірка використання команди
     */
    checkCommandUsage(userId, command, config) {
        const usageKey = `${userId}:${command}`;
        const now = Date.now();
        const [today = ''] = new Date().toISOString().split('T');
        // Перевірка cooldown
        const lastUsage = this.commandUsage.get(usageKey);
        if (lastUsage && config.cooldown && (now - lastUsage.lastUse) < config.cooldown) {
            const remainingTime = config.cooldown - (now - lastUsage.lastUse);
            return {
                allowed: false,
                reason: `Cooldown: ${Math.ceil(remainingTime / 1000)} секунд`
            };
        }
        // Перевірка денного ліміту
        if (config.maxUsesPerDay) {
            if (lastUsage && lastUsage.date === today && lastUsage.uses >= config.maxUsesPerDay) {
                return {
                    allowed: false,
                    reason: `Досягнуто денний ліміт (${config.maxUsesPerDay})`
                };
            }
            const remainingUses = config.maxUsesPerDay - (lastUsage?.uses || 0);
            return {
                allowed: true,
                remainingUses
            };
        }
        return { allowed: true };
    }
    /**
     * Запис використання команди
     */
    recordCommandUsage(userId, command) {
        const usageKey = `${userId}:${command}`;
        const now = Date.now();
        const [today = ''] = new Date().toISOString().split('T');
        const existing = this.commandUsage.get(usageKey);
        if (existing && existing.date === today) {
            existing.uses++;
            existing.lastUse = now;
        }
        else {
            this.commandUsage.set(usageKey, {
                userId,
                command,
                uses: 1,
                date: today,
                lastUse: now
            });
        }
    }
    /**
     * Створення повідомлення про відмову
     */
    buildDenialReason(checks) {
        const reasons = [];
        if (!checks.hasRequiredRoles) {
            reasons.push('відсутні необхідні ролі');
        }
        if (!checks.hasRequiredPermissions) {
            reasons.push('недостатні дозволи Discord');
        }
        if (!checks.canUseInChannel) {
            reasons.push('команда недоступна в цьому каналі');
        }
        if (!checks.usageAllowed && checks.usageReason) {
            reasons.push(checks.usageReason);
        }
        return reasons.join(', ');
    }
    /**
     * Логування подій безпеки
     */
    logSecurityEvent(eventType, userId, data) {
        const details = JSON.stringify({ ...data, timestamp: new Date().toISOString() });
        if (typeof logger_1.default.security === 'function') {
            logger_1.default.security(`🔒 ${eventType} user=${userId} details=${details}`);
        }
        else {
            logger_1.default.warn(`🔒 ${eventType} user=${userId} details=${details}`);
        }
    }
    /**
     * Очищення найстаріших записів з кешу
     */
    cleanupOldestCacheEntries() {
        const entries = Array.from(this.userCache.entries());
        entries.sort((a, b) => a[1].lastActivity - b[1].lastActivity);
        const toDelete = entries.slice(0, Math.floor(PERMISSION_CONSTANTS.MAX_CACHE_ENTRIES * 0.1));
        toDelete.forEach(([key]) => this.userCache.delete(key));
        logger_1.default.debug(`🧹 Очищено ${toDelete.length} застарілих записів з кешу прав`);
    }
    /**
     * Запуск періодичного очищення кешу
     */
    startCacheCleanup() {
        setInterval(() => {
            const now = Date.now();
            let cleanedCount = 0;
            // Очищення кешу користувачів
            for (const [key, entry] of this.userCache.entries()) {
                if (now - entry.lastActivity > PERMISSION_CONSTANTS.CACHE_TTL * 2) {
                    this.userCache.delete(key);
                    cleanedCount++;
                }
            }
            // Очищення використання команд (тільки старі дні)
            const [today = ''] = new Date().toISOString().split('T');
            for (const [key, usage] of this.commandUsage.entries()) {
                if (usage.date !== today && now - usage.lastUse > 24 * 60 * 60 * 1000) {
                    this.commandUsage.delete(key);
                }
            }
            if (cleanedCount > 0) {
                logger_1.default.debug(`🧹 Очищено кеш прав: ${cleanedCount} записів`);
            }
        }, PERMISSION_CONSTANTS.CACHE_TTL);
    }
    /**
     * Додавання нової конфігурації прав
     */
    addPermissionConfig(config) {
        this.permissionConfigs.set(config.command, config);
        logger_1.default.info(`➕ Додано конфігурацію прав для команди "${config.command}"`);
    }
    /**
     * Отримання статистики системи прав
     */
    getStats() {
        const [today = ''] = new Date().toISOString().split('T');
        const dailyUsage = Array.from(this.commandUsage.values())
            .filter(usage => usage.date === today)
            .reduce((sum, usage) => sum + usage.uses, 0);
        return {
            cachedUsers: this.userCache.size,
            commandConfigs: this.permissionConfigs.size,
            dailyUsage,
            isInitialized: this.isInitialized
        };
    }
    /**
     * Очищення ресурсів
     */
    cleanup() {
        this.userCache.clear();
        this.commandUsage.clear();
        this.permissionConfigs.clear();
        this.isInitialized = false;
        logger_1.default.info('🧹 PermissionManager ресурси очищено');
    }
}
exports.PermissionManager = PermissionManager;
PermissionManager.instance = null;
exports.default = PermissionManager;
//# sourceMappingURL=PermissionManager.js.map