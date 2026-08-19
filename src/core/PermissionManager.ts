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

// Нові інтерфейси для детального контролю доступу
export interface ResourcePermission {
  resourceId: string;
  resourceType: 'document' | 'folder' | 'command' | 'feature';
  actions: ('read' | 'write' | 'delete' | 'share' | 'admin')[];
  conditions?: PermissionCondition[];
}

export interface PermissionCondition {
  type: 'time' | 'ip' | 'userAttribute' | 'custom';
  operator: 'equals' | 'notEquals' | 'greaterThan' | 'lessThan' | 'contains' | 'between';
  value: string | number | string[] | number[];
  attribute?: string; // For userAttribute type
}

export interface DetailedPermissionConfig extends PermissionConfig {
  resourcePermissions?: ResourcePermission[];
  attributes?: Record<string, any>; // Custom attributes for advanced permissions
}

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
  OWNER = 5,
}

export interface PermissionCheckResult {
  allowed: boolean;
  reason?: string;
  userLevel: UserLevel;
  hasRequiredRoles: boolean;
  hasRequiredPermissions: boolean;
  canUseInChannel: boolean;
  remainingUses?: number;
  // Нові поля для детального контролю
  resourceAccess?: {
    resourceId: string;
    allowed: boolean;
    reason?: string;
  }[];
}

interface UserPermissionCache {
  userId: string;
  guildId: string;
  userLevel: UserLevel;
  roles: string[];
  permissions: string[];
  timestamp: number;
  lastActivity: number;
  // Нові поля для детального контролю
  resourcePermissions?: Map<string, ResourcePermission>;
  attributes?: Record<string, any>;
}

interface CommandUsage {
  userId: string;
  command: string;
  uses: number;
  date: string;
  lastUse: number;
}

// Новий інтерфейс для документів
export interface DocumentAccessControl {
  documentId: string;
  allowedUsers: string[]; // Discord user IDs
  allowedRoles: string[]; // Role names
  permissions: ('view' | 'edit' | 'delete' | 'share' | 'download')[];
  expiration?: Date;
  conditions?: PermissionCondition[];
}

export class PermissionManager {
  private static instance: PermissionManager | null = null;
  private permissionConfigs = new Map<string, DetailedPermissionConfig>();
  private userCache = new Map<string, UserPermissionCache>();
  private commandUsage = new Map<string, CommandUsage>();
  private documentAccessControls = new Map<string, DocumentAccessControl>();
  private isInitialized = false;

  constructor(_config: BotConfig) {
    if (PermissionManager.instance) {
      // If existing instance was cleaned up in previous test run, re-initialize it
      if (!(PermissionManager.instance).isInitialized) {
        (PermissionManager.instance).initialize();
      }
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
    const configs: DetailedPermissionConfig[] = [
      {
        command: 'пошук',
        requiredRoles: ['Bot User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.USER,
        cooldown: 5000,
        maxUsesPerDay: 100,
        resourcePermissions: [
          {
            resourceId: 'global-search',
            resourceType: 'feature',
            actions: ['read'],
          }
        ]
      },
      {
        command: 'ai-асистент',
        requiredRoles: ['Bot User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 10000,
        maxUsesPerDay: 50,
        resourcePermissions: [
          {
            resourceId: 'ai-assistant',
            resourceType: 'feature',
            actions: ['read', 'write'],
          }
        ]
      },
      {
        command: 'файли',
        requiredRoles: ['Bot User', 'File Manager'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.AttachFiles],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 3000,
        maxUsesPerDay: 200,
        resourcePermissions: [
          {
            resourceId: 'file-management',
            resourceType: 'feature',
            actions: ['read', 'write', 'delete'],
          }
        ]
      },
      {
        command: 'аналітика',
        requiredRoles: ['Analytics User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.TRUSTED,
        cooldown: 15000,
        maxUsesPerDay: 20,
        resourcePermissions: [
          {
            resourceId: 'analytics',
            resourceType: 'feature',
            actions: ['read'],
          }
        ]
      },
      {
        command: 'операції',
        requiredRoles: ['Admin', 'Operations'],
        requiredPermissions: [PermissionFlagsBits.ManageGuild],
        minUserLevel: UserLevel.MODERATOR,
        cooldown: 30000,
        maxUsesPerDay: 10,
        resourcePermissions: [
          {
            resourceId: 'admin-operations',
            resourceType: 'feature',
            actions: ['read', 'write', 'delete', 'admin'],
          }
        ]
      },
      {
        command: 'продуктивність',
        requiredRoles: ['Admin'],
        requiredPermissions: [PermissionFlagsBits.Administrator],
        minUserLevel: UserLevel.ADMIN,
        cooldown: 60000,
        maxUsesPerDay: 5,
        resourcePermissions: [
          {
            resourceId: 'performance-monitoring',
            resourceType: 'feature',
            actions: ['read', 'admin'],
          }
        ]
      },
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
    channelId?: string,
    resourceChecks?: { resourceId: string; action: string }[]
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
          requiredLevel: permConfig.minUserLevel,
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
      const hasRequiredPermissions = this.checkRequiredPermissions(
        member,
        permConfig.requiredPermissions
      );

      // Перевірка каналу
      const canUseInChannel = this.checkChannelPermissions(channelId, permConfig);

      // Перевірка використання команди (rate limiting + daily limits)
      const usageCheck = this.checkCommandUsage(user.id, commandName, permConfig);

      // Initialize allowed variable
      let allowed = hasRequiredRoles && hasRequiredPermissions && canUseInChannel && usageCheck.allowed;

      // Перевірка доступу до ресурсів, якщо потрібно
      let resourceAccess: {
        resourceId: string;
        allowed: boolean;
        reason?: string;
      }[] | undefined = undefined;
      
      if (resourceChecks && resourceChecks.length > 0) {
        const resourceResults = await this.checkResourcePermissions(user.id, resourceChecks, userInfo);
        resourceAccess = resourceResults;
        const allResourcesAllowed = resourceResults.every(result => result.allowed);
        allowed = allowed && allResourcesAllowed;
      }

      const resultBase: PermissionCheckResult = {
        allowed,
        userLevel: userInfo.userLevel,
        hasRequiredRoles,
        hasRequiredPermissions,
        canUseInChannel
      };
      
      // Only add optional properties if they're not undefined
      if (usageCheck.remainingUses !== undefined) {
        resultBase.remainingUses = usageCheck.remainingUses;
      }
      
      if (resourceAccess !== undefined) {
        resultBase.resourceAccess = resourceAccess;
      }

      if (!allowed) {
        const denialObj: {
          hasRequiredRoles: boolean;
          hasRequiredPermissions: boolean;
          canUseInChannel: boolean;
          usageAllowed: boolean;
          resourceAccessAllowed: boolean;
          // do not include usageReason key when undefined (exactOptionalPropertyTypes)
          usageReason?: string;
        } = {
          hasRequiredRoles,
          hasRequiredPermissions,
          canUseInChannel,
          usageAllowed: usageCheck.allowed,
          resourceAccessAllowed: resourceAccess ? resourceAccess.every(ra => ra.allowed) : true
        };
        if (usageCheck.reason !== undefined) {
          denialObj.usageReason = usageCheck.reason;
        }
        const reasonStr = this.buildDenialReason(denialObj);
        resultBase.reason = reasonStr;
      }

      // Логування результату
      const duration = Date.now() - startTime;
      if (allowed) {
        this.recordCommandUsage(user.id, commandName);
        logger.debug(`✅ Доступ дозволено user=${user.id} command=${commandName} duration=${duration}ms level=${UserLevel[userInfo.userLevel]}`);
      } else {
        this.logSecurityEvent('access_denied', user.id, {
          command: commandName,
          reason: resultBase.reason ?? '',
          duration: `${duration}ms`
        });
      }

      return resultBase;
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
    // Support both discord.js Collection and plain Map used in tests
    const roles = Array.from((member.roles.cache as any).values()).map((role: any) => role.name);
    const permissions = member.permissions.toArray();

    // Отримання детальних дозволів користувача
    const resourcePermissions = await this.getUserResourcePermissions(userId, member);

    const userInfo: UserPermissionCache = {
      userId,
      guildId: member.guild.id,
      userLevel,
      roles,
      permissions,
      timestamp: Date.now(),
      lastActivity: Date.now(),
      resourcePermissions,
      attributes: await this.getUserAttributes(userId, member),
    };

    // Кешування з обмеженням розміру
    if (this.userCache.size >= PERMISSION_CONSTANTS.MAX_CACHE_ENTRIES) {
      this.cleanupOldestCacheEntries();
    }

    this.userCache.set(cacheKey, userInfo);
    return userInfo;
  }

  /**
   * Отримання детальних дозволів користувача для ресурсів
   */
  private async getUserResourcePermissions(_userId: string, _member: GuildMember): Promise<Map<string, ResourcePermission>> {
    // В реальній реалізації тут би була логіка отримання детальних дозволів користувача
    // з бази даних або іншого джерела
    
    // Для демонстрації повертаємо порожню мапу
    return new Map();
  }

  /**
   * Отримання атрибутів користувача
   */
  private async getUserAttributes(userId: string, member: GuildMember): Promise<Record<string, any>> {
    // В реальній реалізації тут би була логіка отримання атрибутів користувача
    // з бази даних або іншого джерела
    
    return {
      userId,
      guildId: member.guild.id,
      joinedAt: member.joinedAt?.toISOString(),
      premiumSince: member.premiumSince?.toISOString(),
      isOwner: member.guild.ownerId === userId,
      isAdministrator: member.permissions.has(PermissionFlagsBits.Administrator),
    };
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
    if (
      member.permissions.has([
        PermissionFlagsBits.ManageMessages,
        PermissionFlagsBits.ManageChannels,
        PermissionFlagsBits.ManageRoles,
      ])
    ) {
      return UserLevel.MODERATOR;
    }

    // Перевірка на довірених користувачів
    const trustedRoles = ['Trusted', 'VIP', 'Premium', 'Verified'];
    if (Array.from((member.roles.cache as any).values()).some((role: any) => trustedRoles.includes(role.name))) {
      return UserLevel.TRUSTED;
    }

    // Перевірка на заборонених користувачів
    const bannedRoles = ['Banned', 'Muted', 'Restricted'];
    if (Array.from((member.roles.cache as any).values()).some((role: any) => bannedRoles.includes(role.name))) {
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
    if (
      PERMISSION_CONSTANTS.ADMIN_BYPASS &&
      member.permissions.has(PermissionFlagsBits.Administrator)
    ) {
      return true;
    }

    const memberRoles = Array.from((member.roles.cache as any).values()).map((role: any) => role.name);
    return requiredRoles.some(requiredRole => memberRoles.includes(requiredRole));
  }

  /**
   * Перевірка дозволів Discord
   */
  private checkRequiredPermissions(
    member: GuildMember,
    requiredPermissions: PermissionResolvable[]
  ): boolean {
    if (requiredPermissions.length === 0) return true;
    // Адміністратори завжди мають доступ до перевірки дозволів
    if (
      PERMISSION_CONSTANTS.ADMIN_BYPASS &&
      member.permissions.has(PermissionFlagsBits.Administrator)
    ) {
      return true;
    }

    return member.permissions.has(requiredPermissions);
  }

  /**
   * Перевірка дозволів для каналу
   */
  private checkChannelPermissions(
    channelId: string | undefined,
    config: PermissionConfig
  ): boolean {
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
  private checkCommandUsage(
    userId: string,
    command: string,
    config: PermissionConfig
  ): {
    allowed: boolean;
    reason?: string;
    remainingUses?: number;
  } {
    const usageKey = `${userId}:${command}`;
    const now = Date.now();
    const [today = ''] = new Date().toISOString().split('T');

    // Перевірка cooldown
    const lastUsage = this.commandUsage.get(usageKey);
    if (lastUsage && config.cooldown && now - lastUsage.lastUse < config.cooldown) {
      const remainingTime = config.cooldown - (now - lastUsage.lastUse);
      return {
        allowed: false,
        reason: `Cooldown: ${Math.ceil(remainingTime / 1000)} секунд`,
      };
    }

    // Перевірка денного ліміту
    if (config.maxUsesPerDay) {
      if (lastUsage && lastUsage.date === today && lastUsage.uses >= config.maxUsesPerDay) {
        return {
          allowed: false,
          reason: `Досягнуто денний ліміт (${config.maxUsesPerDay})`,
        };
      }

      const remainingUses = config.maxUsesPerDay - (lastUsage?.uses || 0);
      return {
        allowed: true,
        remainingUses,
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
        lastUse: now,
      });
    }
  }

  /**
   * Перевірка доступу до ресурсів
   */
  private async checkResourcePermissions(
    userId: string,
    resourceChecks: { resourceId: string; action: string }[],
    userInfo: UserPermissionCache
  ): Promise<{ resourceId: string; allowed: boolean; reason?: string }[]> {
    const results: { resourceId: string; allowed: boolean; reason?: string }[] = [];

    for (const check of resourceChecks) {
      const { resourceId, action } = check;
      let allowed = false;
      let reason = '';

      try {
        // Перевірка документних дозволів
        if (this.documentAccessControls.has(resourceId)) {
          const accessControl = this.documentAccessControls.get(resourceId)!;
          allowed = this.checkDocumentAccess(userId, userInfo, accessControl, action as any);
          if (!allowed) {
            reason = `Недостатні дозволи для дії "${action}" над ресурсом "${resourceId}"`;
          }
        } 
        // Перевірка загальних ресурсних дозволів
        else if (userInfo.resourcePermissions?.has(resourceId)) {
          const resourcePerm = userInfo.resourcePermissions.get(resourceId)!;
          allowed = resourcePerm.actions.includes(action as any);
          if (!allowed) {
            reason = `Недостатні дозволи для дії "${action}" над ресурсом "${resourceId}"`;
          }
        } 
        // Якщо немає спеціальних дозволів, дозволяємо доступ для адміністраторів
        else if (userInfo.userLevel >= UserLevel.ADMIN) {
          allowed = true;
        } 
        // Для інших користувачів відмовляємо в доступі
        else {
          allowed = false;
          reason = `Немає дозволів для доступу до ресурсу "${resourceId}"`;
        }

        results.push({
          resourceId,
          allowed,
          reason: allowed ? undefined : reason,
        } as {
          resourceId: string;
          allowed: boolean;
          reason?: string;
        });
      } catch (error) {
        logger.error(`Помилка перевірки доступу до ресурсу: ${error}`);
        results.push({
          resourceId,
          allowed: false,
          reason: `Помилка перевірки доступу: ${error instanceof Error ? error.message : String(error)}`,
        });
      }
    }

    return results;
  }

  /**
   * Перевірка доступу до документа
   */
  private checkDocumentAccess(
    userId: string,
    userInfo: UserPermissionCache,
    accessControl: DocumentAccessControl,
    action: 'view' | 'edit' | 'delete' | 'share' | 'download'
  ): boolean {
    // Перевірка терміну дії
    if (accessControl.expiration && new Date() > accessControl.expiration) {
      return false;
    }

    // Перевірка дозволів для конкретної дії
    if (!accessControl.permissions.includes(action)) {
      return false;
    }

    // Перевірка дозволених користувачів
    if (accessControl.allowedUsers.length > 0 && !accessControl.allowedUsers.includes(userId)) {
      // Якщо є список дозволених користувачів і поточний користувач не в ньому, відмовити
      return false;
    }

    // Перевірка дозволених ролей
    if (accessControl.allowedRoles.length > 0) {
      const hasAllowedRole = accessControl.allowedRoles.some(role => 
        userInfo.roles.includes(role)
      );
      if (!hasAllowedRole) {
        // Якщо є список дозволених ролей і користувач не має жодної з них, відмовити
        return false;
      }
    }

    // Перевірка умов доступу
    if (accessControl.conditions && accessControl.conditions.length > 0) {
      for (const condition of accessControl.conditions) {
        if (!this.evaluateCondition(condition, userInfo)) {
          return false;
        }
      }
    }

    // Якщо всі перевірки пройшли, дозволити доступ
    return true;
  }

  /**
   * Оцінка умови доступу
   */
  private evaluateCondition(condition: PermissionCondition, userInfo: UserPermissionCache): boolean {
    try {
      switch (condition.type) {
        case 'time': {
          // Перевірка часу доступу
          const now = new Date().getTime();
          if (condition.operator === 'between' && Array.isArray(condition.value) && condition.value.length === 2) {
            const [start, end] = condition.value as number[];
            if (start !== undefined && end !== undefined) {
              return now >= start && now <= end;
            }
          }
          return false;
        }

        case 'userAttribute':
          // Перевірка атрибутів користувача
          if (condition.attribute && userInfo.attributes) {
            const userValue = userInfo.attributes[condition.attribute];
            const conditionValue = condition.value;
            
            switch (condition.operator) {
              case 'equals':
                return userValue === conditionValue;
              case 'notEquals':
                return userValue !== conditionValue;
              case 'greaterThan':
                return typeof userValue === 'number' && typeof conditionValue === 'number' && userValue > conditionValue;
              case 'lessThan':
                return typeof userValue === 'number' && typeof conditionValue === 'number' && userValue < conditionValue;
              case 'contains':
                if (Array.isArray(userValue)) {
                  return userValue.includes(conditionValue as string);
                }
                return false;
            }
          }
          return false;

        default:
          // Для інших типів умов повертаємо false
          return false;
      }
    } catch (error) {
      logger.warn(`Помилка оцінки умови доступу: ${error}`);
      return false;
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
    resourceAccessAllowed: boolean;
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

    if (!checks.resourceAccessAllowed) {
      reasons.push('недостатні дозволи для доступу до ресурсу');
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
    // Avoid background timers in test environment to prevent open handle issues in Jest
    if (process.env['NODE_ENV'] === 'test' || process.env['JEST_WORKER_ID']) {
      return;
    }
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
  public addPermissionConfig(config: DetailedPermissionConfig): void {
    this.permissionConfigs.set(config.command, config);
    logger.info(`➕ Додано конфігурацію прав для команди "${config.command}"`);
  }

  /**
   * Встановлення контролю доступу до документа
   */
  public setDocumentAccessControl(documentId: string, accessControl: DocumentAccessControl): void {
    this.documentAccessControls.set(documentId, accessControl);
    logger.info(`🔒 Встановлено контроль доступу до документа ${documentId}`);
  }

  /**
   * Видалення контролю доступу до документа
   */
  public removeDocumentAccessControl(documentId: string): boolean {
    const result = this.documentAccessControls.delete(documentId);
    if (result) {
      logger.info(`🔓 Видалено контроль доступу до документа ${documentId}`);
    }
    return result;
  }

  /**
   * Отримання контролю доступу до документа
   */
  public getDocumentAccessControl(documentId: string): DocumentAccessControl | undefined {
    return this.documentAccessControls.get(documentId);
  }

  /**
   * Отримання статистики системи прав
   */
  public getStats(): {
    cachedUsers: number;
    commandConfigs: number;
    documentAccessControls: number;
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
      documentAccessControls: this.documentAccessControls.size,
      dailyUsage,
      isInitialized: this.isInitialized,
    };
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    this.userCache.clear();
    this.commandUsage.clear();
    this.permissionConfigs.clear();
    this.documentAccessControls.clear();
    this.isInitialized = false;
    logger.info('🧹 PermissionManager ресурси очищено');
  }
}

export default PermissionManager;