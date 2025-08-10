/**
 * Система управління правами доступу для Discord AI Assistant Bot
 * Забезпечує гранульний контроль доступу до команд та функцій
 * Версія 1.0.0 - Нова реалізація
 */

import { GuildMember, User, PermissionResolvable, PermissionFlagsBits } from 'discord.js';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

// Константи для системи прав
const PERMISSION_CONSTANTS = {
  CACHE_TTL: 300000, // 5 хвилин
  MAX_CACHE_ENTRIES: 1000,
  PERMISSION_TIMEOUT: 5000, // 5 секунд
  ADMIN_BYPASS: true,
} as const;

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

export enum UserLevel {
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

interface UserPermissionCache {
  userId: string;
  guildId: string;
  userLevel: UserLevel;
  roles: string[];
  permissions: string[];
  timestamp: number;
  lastActivity: number;
}

interface CommandUsage {
  userId: string;
  command: string;
  uses: number;
  date: string;
  lastUse: number;
}

export class PermissionManager {
  private static instance: PermissionManager | null = null;
  private permissionConfigs = new Map<string, PermissionConfig>();
  private userCache = new Map<string, UserPermissionCache>();
  private commandUsage = new Map<string, CommandUsage>();
  private isInitialized = false;

  constructor(_config: BotConfig) {
    if (PermissionManager.instance) {
      return PermissionManager.instance;
    }
    
    PermissionManager.instance = this;
    this.initialize();
  }

  /**
   * Ініціалізація системи прав
   */
  private initialize(): void {
    try {
      logger.info('🔐 Ініціалізація системи прав доступу...');
      
      // Завантаження конфігурацій прав для команд
      this.loadPermissionConfigs();
      
      // Запуск періодичного очищення кешу
      this.startCacheCleanup();
      
      this.isInitialized = true;
      logger.info('✅ Система прав доступу ініціалізована');
    } catch (error) {
      logger.error(`❌ Помилка ініціалізації системи прав: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Завантаження конфігурацій прав для команд
   */
  private loadPermissionConfigs(): void {
    // Конфігурації для різних команд
    const configs: PermissionConfig[] = [
      {
        command: 'пошук',
        requiredRoles: ['Bot User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.USER,
        cooldown: 5000,
        maxUsesPerDay: 100
      },
      {
        command: 'ai-асистент',
        requiredRoles: ['Bot User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 10000,
        maxUsesPerDay: 50
      },
      {
        command: 'файли',
        requiredRoles: ['Bot User', 'File Manager'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.AttachFiles],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 3000,
        maxUsesPerDay: 200
      },
      {
        command: 'аналітика',
        requiredRoles: ['Analytics User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 15000,
        maxUsesPerDay: 20
      },
      {
        command: 'операції',
        requiredRoles: ['Admin', 'Operations'],
        requiredPermissions: [PermissionFlagsBits.ManageGuild],
        minUserLevel: UserLevel.MODERATOR,
        cooldown: 30000,
        maxUsesPerDay: 10
      },
      {
        command: 'продуктивність',
        requiredRoles: ['Admin'],
        requiredPermissions: [PermissionFlagsBits.Administrator],
        minUserLevel: UserLevel.ADMIN,
        cooldown: 60000,
        maxUsesPerDay: 5
      }
    ];

    configs.forEach(config => {
      this.permissionConfigs.set(config.command, config);
    });

    logger.info(`📋 Завантажено ${configs.length} конфігурацій прав доступу`);
  }

  /**
   * Головна функція перевірки прав доступу
   */
  public async checkPermission(
    user: User,
    member: GuildMember | null,
    commandName: string,
    channelId?: string
  ): Promise<PermissionCheckResult> {
    try {
      const startTime = Date.now();
      
      // Отримання конфігурації команди
      const permConfig = this.permissionConfigs.get(commandName);
      if (!permConfig) {
        logger.warn(`⚠️ Конфігурація прав для команди "${commandName}" не знайдена`);
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
      } as PermissionCheckResult;
      if (!allowed) {
        const denialObj: {
          hasRequiredRoles: boolean;
          hasRequiredPermissions: boolean;
          canUseInChannel: boolean;
          usageAllowed: boolean;
          // do not include usageReason key when undefined (exactOptionalPropertyTypes)
          usageReason?: string;
        } = {
          hasRequiredRoles,
          hasRequiredPermissions,
          canUseInChannel,
          usageAllowed: usageCheck.allowed
        };
        if (usageCheck.reason !== undefined) {
          denialObj.usageReason = usageCheck.reason;
        }
        const reasonStr = this.buildDenialReason(denialObj);
        (resultBase as any).reason = reasonStr;
      }
      const result: PermissionCheckResult = resultBase;

      // Логування результату
      const duration = Date.now() - startTime;
      if (allowed) {
        this.recordCommandUsage(user.id, commandName);
        logger.debug(`✅ Доступ дозволено user=${user.id} command=${commandName} duration=${duration}ms level=${UserLevel[userInfo.userLevel]}`);
      } else {
        this.logSecurityEvent('access_denied', user.id, {
          command: commandName,
          reason: result.reason ?? '',
          duration: `${duration}ms`
        });
      }

      return result;
    } catch (error) {
      logger.error(`❌ Помилка перевірки прав доступу: user=${user.id} command=${commandName} error=${error instanceof Error ? error.message : String(error)}`);
      
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
  private async getUserInfo(userId: string, member: GuildMember): Promise<UserPermissionCache> {
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

    const userInfo: UserPermissionCache = {
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
  private determineUserLevel(member: GuildMember): UserLevel {
    // Перевірка на власника сервера
    if (member.guild.ownerId === member.id) {
      return UserLevel.OWNER;
    }

    // Перевірка на адміністратора
    if (member.permissions.has(PermissionFlagsBits.Administrator)) {
      return UserLevel.ADMIN;
    }

    // Перевірка на модератора
    if (member.permissions.has([
      PermissionFlagsBits.ManageMessages,
      PermissionFlagsBits.ManageChannels,
      PermissionFlagsBits.ManageRoles
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
  private checkRequiredRoles(member: GuildMember, requiredRoles: string[]): boolean {
    if (requiredRoles.length === 0) return true;
    
    // Адміністратори завжди мають доступ
    if (PERMISSION_CONSTANTS.ADMIN_BYPASS && 
        member.permissions.has(PermissionFlagsBits.Administrator)) {
      return true;
    }

    const memberRoles = member.roles.cache.map(role => role.name);
    return requiredRoles.some(requiredRole => memberRoles.includes(requiredRole));
  }

  /**
   * Перевірка дозволів Discord
   */
  private checkRequiredPermissions(member: GuildMember, requiredPermissions: PermissionResolvable[]): boolean {
    if (requiredPermissions.length === 0) return true;
    
    return member.permissions.has(requiredPermissions);
  }

  /**
   * Перевірка дозволів для каналу
   */
  private checkChannelPermissions(channelId: string | undefined, config: PermissionConfig): boolean {
    if (!channelId) return true;
    
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
  private checkCommandUsage(userId: string, command: string, config: PermissionConfig): {
    allowed: boolean;
    reason?: string;
    remainingUses?: number;
  } {
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
  private recordCommandUsage(userId: string, command: string): void {
    const usageKey = `${userId}:${command}`;
    const now = Date.now();
    const [today = ''] = new Date().toISOString().split('T');
    
    const existing = this.commandUsage.get(usageKey);
    
    if (existing && existing.date === today) {
      existing.uses++;
      existing.lastUse = now;
    } else {
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
  private buildDenialReason(checks: {
    hasRequiredRoles: boolean;
    hasRequiredPermissions: boolean;
    canUseInChannel: boolean;
    usageAllowed: boolean;
    usageReason?: string;
  }): string {
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
  private logSecurityEvent(eventType: string, userId: string, data: Record<string, unknown>): void {
    const details = JSON.stringify({ ...data, timestamp: new Date().toISOString() });
    if (typeof (logger as any).security === 'function') {
      (logger as any).security(`🔒 ${eventType} user=${userId} details=${details}`);
    } else {
      logger.warn(`🔒 ${eventType} user=${userId} details=${details}`);
    }
  }

  /**
   * Очищення найстаріших записів з кешу
   */
  private cleanupOldestCacheEntries(): void {
    const entries = Array.from(this.userCache.entries());
    entries.sort((a, b) => a[1].lastActivity - b[1].lastActivity);
    
    const toDelete = entries.slice(0, Math.floor(PERMISSION_CONSTANTS.MAX_CACHE_ENTRIES * 0.1));
    toDelete.forEach(([key]) => this.userCache.delete(key));
    
    logger.debug(`🧹 Очищено ${toDelete.length} застарілих записів з кешу прав`);
  }

  /**
   * Запуск періодичного очищення кешу
   */
  private startCacheCleanup(): void {
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
        logger.debug(`🧹 Очищено кеш прав: ${cleanedCount} записів`);
      }
    }, PERMISSION_CONSTANTS.CACHE_TTL);
  }

  /**
   * Додавання нової конфігурації прав
   */
  public addPermissionConfig(config: PermissionConfig): void {
    this.permissionConfigs.set(config.command, config);
    logger.info(`➕ Додано конфігурацію прав для команди "${config.command}"`);
  }

  /**
   * Отримання статистики системи прав
   */
  public getStats(): {
    cachedUsers: number;
    commandConfigs: number;
    dailyUsage: number;
    isInitialized: boolean;
  } {
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
  public cleanup(): void {
    this.userCache.clear();
    this.commandUsage.clear();
    this.permissionConfigs.clear();
    this.isInitialized = false;
    logger.info('🧹 PermissionManager ресурси очищено');
  }
}

export default PermissionManager;