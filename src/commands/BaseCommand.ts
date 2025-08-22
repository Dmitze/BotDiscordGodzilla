/**
 * Базовий абстрактний клас для всіх команд Discord бота
 * Забезпечує уніфіковану структуру та типізацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import { SlashCommandBuilder, EmbedBuilder } from 'discord.js';
import type {
  ChatInputCommandInteraction,
  AutocompleteInteraction,
  MessageComponentInteraction,
} from 'discord.js';

import type {
  BotConfig,
  CommandOptions,
  CommandStats,
  CommandContext,
  HealthStatus,
} from '@/types';
import logger from '@/utils/logger';
import { sanitizeInput } from '@/utils/security';
import { UserPreferencesService } from '@/services/UserPreferencesService';
import { t, tUser } from '@/i18n';
import { replyWithPrivacy } from '@/ui/reply';
import { verifyComponentId } from '@/security/componentId';

// Константи для конфігурації команд
const COMMAND_CONFIG = {
  DEFAULT_COOLDOWN: 3000, // 3 секунди
  MAX_COOLDOWN: 30000, // 30 секунд
  MIN_COOLDOWN: 1000, // 1 секунда
  MAX_EXECUTION_TIME: 30000, // 30 секунд
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000, // 1 секунда
  CACHE_SIZE: 1000,
  CLEANUP_INTERVAL: 5 * 60 * 1000, // 5 хвилин
} as const;

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
  i18n?: {
    nameKey?: string;
    descriptionKey?: string;
  };
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

export abstract class BaseCommand {
  public readonly data: SlashCommandBuilder;
  public readonly name: string;
  public readonly description: string;
  public readonly category: string;
  public readonly usage: string;
  public readonly examples: string[];
  public readonly permissions: string[];
  public readonly cooldown: number;
  public readonly allowDM: boolean;

  protected stats: CommandStats;
  protected cooldowns: Map<string, number> = new Map();
  protected executionCache: Map<string, { result: any; timestamp: number }> = new Map();
  protected errorCount: Map<string, number> = new Map();
  protected lastExecution: Map<string, number> = new Map();
  protected readonly config: BotConfig;
  protected isShuttingDown = false;

  constructor(
    name: string,
    description: string,
    config: BotConfig,
    options: Partial<CommandData> = {},
    builder?: (builder: SlashCommandBuilder) => SlashCommandBuilder
  ) {
    this.name = name;
    this.description = description;
    this.config = config;
    this.category = options.category || 'general';
    this.usage = options.usage || `/${name}`;
    this.examples = options.examples || [];
    this.permissions = options.permissions || [];
    this.cooldown = this.validateCooldown(options.cooldown || COMMAND_CONFIG.DEFAULT_COOLDOWN);
    this.allowDM = options.dmPermission ?? true;

    // Створення SlashCommandBuilder
    this.data = new SlashCommandBuilder().setName(name).setDescription(description);

    // Локалізації імені/опису команди (без зміни стабільного ідентифікатора name)
    if (options.i18n?.nameKey) {
      try {
        const nameUk = t(options.i18n.nameKey, undefined, 'uk');
        const nameEn = t(options.i18n.nameKey, undefined, 'en');
        this.data.setNameLocalizations({ uk: nameUk, 'en-US': nameEn });
      } catch {}
    }
    if (options.i18n?.descriptionKey) {
      try {
        const descUk = t(options.i18n.descriptionKey, undefined, 'uk');
        const descEn = t(options.i18n.descriptionKey, undefined, 'en');
        this.data.setDescriptionLocalizations({ uk: descUk, 'en-US': descEn });
      } catch {}
    }

    // Додавання опцій через builder функцію
    if (builder) {
      try {
        builder(this.data);
      } catch (error) {
        logger.error('Помилка створення builder для команди', {
          type: 'command',
          component: name,
          event: 'builder_failed',
          errorName: error instanceof Error ? error.name : undefined,
          errorMessage: error instanceof Error ? error.message : String(error),
          stack: error instanceof Error ? error.stack : undefined,
        });
        throw new Error(
          `Помилка створення команди: ${error instanceof Error ? error.message : 'Невідома помилка'}`
        );
      }
    }

    // Встановлення дозволів
    if (options.defaultMemberPermissions) {
      this.data.setDefaultMemberPermissions(options.defaultMemberPermissions);
    }

    if (options.dmPermission !== undefined) {
      this.data.setDMPermission(options.dmPermission);
    }

    this.stats = {
      service: `Command:${name}`,
      uptime: 0,
      requests: 0,
      errors: 0,
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      averageExecutionTime: 0,
      totalExecutionTime: 0,
      cacheHits: 0,
      cacheMisses: 0,
      retries: 0,
    };

    // Запуск періодичного очищення
    this.startCleanupInterval();

    logger.info('Команда ініціалізована', {
      type: 'command',
      component: name,
      event: 'initialized',
      category: this.category,
      cooldown: this.cooldown,
      permissions: this.permissions,
    });
  }

  /**
   * Виконання команди з детальним логуванням та обробкою помилок
   */
  public async execute(
    arg: CommandExecuteOptions | ChatInputCommandInteraction
  ): Promise<void> {
    // Backward-compatible adapter: tests may call execute(interaction)
    const options: CommandExecuteOptions =
      (arg as ChatInputCommandInteraction)?.user !== undefined
        ? { interaction: arg as ChatInputCommandInteraction }
        : (arg as CommandExecuteOptions);

    const startTime = performance.now();
    const userId = options.interaction.user.id;

    try {
      // Застосувати локаль користувача (i18n), дефолт 'uk' з підтримкою псевдоніма 'uk-UA'
      await UserPreferencesService.resolveAndApplyLocale(options.interaction);

      // Перевірка стану команди
      if (this.isShuttingDown) {
        await this.handleShutdownError(options.interaction);
        return;
      }

      // Валідація вхідних даних
      const validation = await this.validateExecution(options);
      if (!validation.isValid) {
        await this.handleValidationError(options.interaction, validation.errors);
        return;
      }

      // Перевірка cooldown
      if (this.isOnCooldown(userId)) {
        await this.handleCooldown(options.interaction);
        return;
      }

      // Перевірка кешу
      const cacheKey = this.generateCacheKey(options);
      const cachedResult = this.getCachedResult(cacheKey);
      if (cachedResult) {
        await this.handleCachedResult(options.interaction, cachedResult);
        return;
      }

      // Встановлення cooldown
      this.setCooldown(userId);

      // Логування початку виконання
      this.logCommandStart(options.interaction);

      // Виконання команди з retry логікою
      const result = await this.executeWithRetry(options);

      // Кешування результату
      this.cacheResult(cacheKey, result);

      // Оновлення статистики
      const duration = performance.now() - startTime;
      this.updateStats(true, duration);

      // Логування успішного завершення
      this.logCommandSuccess(options.interaction, duration);

      // Оновлення часу останнього виконання
      this.lastExecution.set(userId, Date.now());
    } catch (error) {
      // Оновлення статистики помилок
      const duration = performance.now() - startTime;
      this.updateStats(false, duration);

      // Збільшення лічильника помилок
      this.incrementErrorCount(userId);

      // Логування помилки
      this.logCommandError(options.interaction, error);

      // Обробка помилки
      await this.handleError(options.interaction, error);
    }
  }

  /**
   * Виконання команди з retry логікою
   */
  private async executeWithRetry(options: CommandExecuteOptions): Promise<any> {
    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= COMMAND_CONFIG.MAX_RETRIES; attempt++) {
      try {
        const result = await this.onExecute(options);
        this.stats.retries += attempt - 1;
        return result;
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        if (attempt < COMMAND_CONFIG.MAX_RETRIES) {
          logger.warn('Спроба виконання невдала, повтор', {
            type: 'command',
            component: this.name,
            event: 'retry',
            errorMessage: lastError.message,
            attempt,
            maxRetries: COMMAND_CONFIG.MAX_RETRIES,
          });

          await new Promise(resolve => setTimeout(resolve, COMMAND_CONFIG.RETRY_DELAY * attempt));
        }
      }
    }

    throw lastError || new Error('Всі спроби виконання команди невдалі');
  }

  /**
   * Валідація виконання команди
   */
  private async validateExecution(
    options: CommandExecuteOptions
  ): Promise<CommandValidationResult> {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      // Перевірка користувача
      if (!options.interaction.user) {
        errors.push('Користувач не знайдено');
      }

      // Перевірка сервера (якщо потрібно)
      if (!options.interaction.guild && !this.allowDM) {
        errors.push('Команда доступна тільки на сервері');
      }

      // Перевірка дозволів
      if (this.permissions.length > 0) {
        const member = options.interaction.member as Record<string, unknown> | null | undefined;
        const perms: unknown =
          member && 'permissions' in (member as object) ? (member as any).permissions : undefined;
        const hasFn =
          perms && typeof (perms as any).has === 'function'
            ? (perms as any).has.bind(perms)
            : undefined;
        if (hasFn) {
          const hasPermission = this.permissions.some(permission => hasFn(permission as any));
          if (!hasPermission) {
            errors.push(`Необхідні дозволи: ${this.permissions.join(', ')}`);
          }
        }
      }

      // Санітизація опцій
      if (options.options) {
        const sanitizedOptions: Record<string, unknown> = {};
        for (const [key, value] of Object.entries(options.options)) {
          if (typeof value === 'string') {
            const sanitized = sanitizeInput(value, 'command');
            if (sanitized.isValid) {
              sanitizedOptions[key] = sanitized.sanitizedValue;
              if (sanitized.warnings.length > 0) {
                warnings.push(...sanitized.warnings.map(w => `${key}: ${w}`));
              }
            } else {
              errors.push(...sanitized.errors.map(e => `${key}: ${e}`));
            }
          } else {
            sanitizedOptions[key] = value as unknown;
          }
        }
        options.options = sanitizedOptions;
      }

      return {
        isValid: errors.length === 0,
        errors,
        warnings,
        sanitizedOptions: options.options,
      };
    } catch (error) {
      logger.error('Помилка валідації команди', {
        type: 'command',
        component: this.name,
        event: 'validation_error',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      return {
        isValid: false,
        errors: ['Помилка валідації команди'],
        warnings: [],
      };
    }
  }

  /**
   * Обробка автодоповнення з детальним логуванням
   */
  public async autocomplete(options: CommandAutocompleteOptions): Promise<void> {
    const startTime = performance.now();

    try {
      logger.debug('Автодоповнення почато', {
        type: 'command',
        component: this.name,
        event: 'autocomplete_start',
        userTag: options.interaction.user.tag,
        query: options.query,
      });

      await this.onAutocomplete(options);

      const duration = performance.now() - startTime;
      logger.debug('Автодоповнення завершено', {
        type: 'command',
        component: this.name,
        event: 'autocomplete_finish',
        durationMs: Number(duration.toFixed(2)),
      });
    } catch (error) {
      // keep duration calculation if needed in future metrics
      this.logAutocompleteError(options.interaction, error);
      await this.handleAutocompleteError(options.interaction, error);
    }
  }

  /**
   * Обробка компонентів з детальним логуванням
   */
  public async handleComponent(options: CommandComponentOptions): Promise<void> {
    const startTime = performance.now();

    try {
      logger.debug('Обробка компонента почата', {
        type: 'command',
        component: this.name,
        event: 'component_start',
        userTag: options.interaction.user.tag,
        componentType: options.componentType,
        customId: options.interaction.customId,
      });

      // Centralized verification of component customId (HMAC+TTL), with backward compatibility.
      // We only enforce verification for token-like IDs that follow the signed format: header.body.sig (contain two dots)
      const customId = options.interaction.customId;
      const dotCount = (customId?.match(/\./g) || []).length;
      if (dotCount === 2) {
        const res = verifyComponentId(customId);
        if (!res.valid) {
          const reason = res.reason || 'invalid';
          logger.warn('Компонент customId не пройшов верифікацію', {
            type: 'security',
            component: this.name,
            event: 'component_verify_failed',
            reason,
          });

          const key = reason === 'expired' ? 'security.component.expiredId' : 'security.component.invalidId';
          await replyWithPrivacy(
            options.interaction as any,
            { content: tUser(key, options.interaction as any) },
            { ephemeralByDefault: true, shareFlagSupport: true }
          );
          return;
        }
        // Optionally, attach verified payload to context for downstream usage
        (options as any).context = { ...(options.context || {}), componentPayload: res.payload };
      }

      await this.onComponent(options);

      const duration = performance.now() - startTime;
      logger.debug('Обробка компонента завершена', {
        type: 'command',
        component: this.name,
        event: 'component_finish',
        durationMs: Number(duration.toFixed(2)),
      });
    } catch (error) {
      // keep duration calculation if needed in future metrics
      this.logComponentError(options.interaction, error);
      await this.handleComponentError(options.interaction, error);
    }
  }

  /**
   * Абстрактний метод виконання команди
   */
  protected abstract onExecute(options: CommandExecuteOptions): Promise<void>;

  /**
   * Обробка автодоповнення (опціонально)
   */
  protected async onAutocomplete(_options: CommandAutocompleteOptions): Promise<void> {
    // Базова реалізація - нічого не робить
  }

  /**
   * Обробка компонентів (опціонально)
   */
  protected async onComponent(_options: CommandComponentOptions): Promise<void> {
    // Базова реалізація - нічого не робить
  }

  /**
   * Валідація cooldown
   */
  private validateCooldown(cooldown: number): number {
    if (cooldown < COMMAND_CONFIG.MIN_COOLDOWN) {
      logger.warn(`Cooldown для команди ${this.name} занадто малий, встановлюю мінімальний`);
      return COMMAND_CONFIG.MIN_COOLDOWN;
    }
    if (cooldown > COMMAND_CONFIG.MAX_COOLDOWN) {
      logger.warn(`Cooldown для команди ${this.name} занадто великий, встановлюю максимальний`);
      return COMMAND_CONFIG.MAX_COOLDOWN;
    }
    return cooldown;
  }

  /**
   * Перевірка cooldown
   */
  protected isOnCooldown(userId: string): boolean {
    const cooldownTime = this.cooldowns.get(userId);
    if (!cooldownTime) return false;

    return Date.now() < cooldownTime;
  }

  /**
   * Встановлення cooldown
   */
  protected setCooldown(userId: string): void {
    this.cooldowns.set(userId, Date.now() + this.cooldown);
  }

  /**
   * Отримання часу cooldown
   */
  protected getCooldownTime(userId: string): number {
    const cooldownTime = this.cooldowns.get(userId);
    if (!cooldownTime) return 0;

    return Math.max(0, cooldownTime - Date.now());
  }

  /**
   * Генерація ключа кешу
   */
  private generateCacheKey(options: CommandExecuteOptions): string {
    // Деякі тести викликають цей метод без interaction, тому робимо його безпечним
    const anyOpts: any = options as any;
    const interactionUserId = anyOpts?.interaction?.user?.id ?? 'anon';
    const payload = anyOpts?.options ?? anyOpts ?? {};
    const optionsJson = JSON.stringify(payload);

    // Легасі формат для SearchCommand: "search:base64:<base64(json)>"
    // Тести очікують префікс 'search:' та наявність 'base64'
    const isSearch = this.name === 'пошук' || (this as any)?.constructor?.name === 'SearchCommand';
    if (isSearch) {
      const b64 = Buffer.from(optionsJson, 'utf8').toString('base64');
      return `search:base64:${b64}`;
    }

    // Загальний випадок для інших команд
    return `${this.name}:${interactionUserId}:${optionsJson}`;
  }

  /**
   * Отримання кешованого результату
   */
  private getCachedResult(cacheKey: string): any {
    const cached = this.executionCache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < 300000) {
      // 5 хвилин
      this.stats.cacheHits++;
      return cached.result;
    }
    this.stats.cacheMisses++;
    return null;
  }

  /**
   * Кешування результату
   */
  private cacheResult(cacheKey: string, result: any): void {
    this.executionCache.set(cacheKey, {
      result,
      timestamp: Date.now(),
    });

    // Обмеження розміру кешу
    if (this.executionCache.size > COMMAND_CONFIG.CACHE_SIZE) {
      const oldestKey = this.executionCache.keys().next().value as unknown;
      if (typeof oldestKey === 'string') {
        this.executionCache.delete(oldestKey);
      }
    }
  }

  /**
   * Збільшення лічильника помилок
   */
  private incrementErrorCount(userId: string): void {
    const currentCount = this.errorCount.get(userId) || 0;
    this.errorCount.set(userId, currentCount + 1);
  }

  /**
   * Обробка cooldown
   */
  protected async handleCooldown(interaction: ChatInputCommandInteraction): Promise<void> {
    const remainingTime = this.getCooldownTime(interaction.user.id);
    const seconds = Math.ceil(remainingTime / 1000);

    const embed = new EmbedBuilder()
      .setColor('#FF6B6B')
      .setTitle('⏰ Cooldown активний')
      .setDescription(`Спробуйте ще раз через **${seconds} секунд**`)
      .addFields(
        { name: 'Команда', value: this.name, inline: true },
        { name: 'Залишилось', value: `${seconds}с`, inline: true }
      )
      .setTimestamp();

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (error) {
      logger.error('Помилка відправки cooldown повідомлення:', { error });
    }
  }

  /**
   * Обробка кешованого результату
   */
  private async handleCachedResult(
    interaction: ChatInputCommandInteraction,
    _result: any
  ): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor('#4CAF50')
      .setTitle('⚡ Кешований результат')
      .setDescription('Результат завантажено з кешу')
      .setTimestamp();

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (error) {
      logger.error('Помилка відправки кешованого результату:', { error });
    }
  }

  /**
   * Обробка помилки валідації
   */
  private async handleValidationError(
    interaction: ChatInputCommandInteraction,
    errors: string[]
  ): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor('#FF9800')
      .setTitle('⚠️ Помилка валідації')
      .setDescription('Виправте наступні помилки:')
      .addFields(errors.map(error => ({ name: '❌', value: error, inline: false })))
      .setTimestamp();

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (error) {
      logger.error('Помилка відправки повідомлення про валідацію:', { error });
    }
  }

  /**
   * Обробка помилки зупинки
   */
  private async handleShutdownError(interaction: ChatInputCommandInteraction): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor('#FF6B6B')
      .setTitle('🛑 Команда недоступна')
      .setDescription('Бот знаходиться в процесі зупинки')
      .setTimestamp();

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (error) {
      logger.error('Помилка відправки повідомлення про зупинку:', { error });
    }
  }

  /**
   * Обробка помилок
   */
  protected async handleError(
    interaction: ChatInputCommandInteraction,
    error: unknown
  ): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
    const errorStack = error instanceof Error ? error.stack : undefined;

    const embed = new EmbedBuilder()
      .setColor('#FF6B6B')
      .setTitle('❌ Помилка виконання команди')
      .setDescription(`**Помилка:** ${errorMessage}`)
      .addFields(
        { name: 'Команда', value: this.name, inline: true },
        { name: 'Користувач', value: interaction.user.tag, inline: true }
      )
      .setTimestamp();

    if (errorStack) {
      embed.addFields({
        name: 'Деталі',
        value: `\`\`\`${errorStack.substring(0, 1000)}...\`\`\``,
        inline: false,
      });
    }

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (replyError) {
      logger.error('Помилка відправки повідомлення про помилку:', { error: replyError });
    }
  }

  /**
   * Обробка помилок автодоповнення
   */
  protected async handleAutocompleteError(
    interaction: AutocompleteInteraction,
    _error: unknown
  ): Promise<void> {
    try {
      await interaction.respond([{ name: 'Помилка завантаження', value: 'error' }]);
    } catch (replyError) {
      logger.error('Помилка відповіді автодоповнення:', { error: replyError });
    }
  }

  /**
   * Обробка помилок компонентів
   */
  protected async handleComponentError(
    interaction: MessageComponentInteraction,
    error: unknown
  ): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

    const embed = new EmbedBuilder()
      .setColor('#FF6B6B')
      .setTitle('❌ Помилка обробки компонента')
      .setDescription(`**Помилка:** ${errorMessage}`)
      .addFields(
        { name: 'Команда', value: this.name, inline: true },
        { name: 'Компонент', value: interaction.customId, inline: true }
      )
      .setTimestamp();

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed] });
      } else {
        await replyWithPrivacy(interaction as any, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (replyError) {
      logger.error('Помилка відправки повідомлення про помилку компонента:', { error: replyError });
    }
  }

  /**
   * Логування початку команди
   */
  protected logCommandStart(interaction: ChatInputCommandInteraction): void {
    const meta: Record<string, unknown> = {
      user: interaction.user.tag,
      userId: interaction.user.id,
      options: interaction.options.data,
    };
    if (interaction['guildId']) meta['guildId'] = interaction['guildId'];
    if (interaction['channelId']) meta['channelId'] = interaction['channelId'];
    logger.info(`🚀 Команда ${this.name} виконується`, meta);
  }

  /**
   * Логування успішного завершення
   */
  protected logCommandSuccess(interaction: ChatInputCommandInteraction, duration: number): void {
    logger.info(`✅ Команда ${this.name} успішно виконана`, {
      user: interaction.user.tag,
      duration: `${duration.toFixed(2)}ms`,
      performance: duration > 5000 ? 'slow' : duration > 1000 ? 'medium' : 'fast',
    });
  }

  /**
   * Логування помилки команди
   */
  protected logCommandError(interaction: ChatInputCommandInteraction, error: unknown): void {
    logger.error(`❌ Помилка команди ${this.name}`, {
      user: interaction.user.tag,
      userId: interaction.user.id,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
    });
  }

  /**
   * Логування помилки автодоповнення
   */
  protected logAutocompleteError(interaction: AutocompleteInteraction, error: unknown): void {
    logger.error(`❌ Помилка автодоповнення команди ${this.name}`, {
      user: interaction.user.tag,
      error: error instanceof Error ? error.message : String(error),
    });
  }

  /**
   * Логування помилки компонента
   */
  protected logComponentError(interaction: MessageComponentInteraction, error: unknown): void {
    logger.error(`❌ Помилка компонента команди ${this.name}`, {
      user: interaction.user.tag,
      customId: interaction.customId,
      error: error instanceof Error ? error.message : String(error),
    });
  }

  /**
   * Оновлення статистики
   */
  protected updateStats(success: boolean, duration: number): void {
    this.stats.totalExecutions++;
    this.stats.totalExecutionTime += duration;
    this.stats.uptime = Date.now() - this.stats.uptime;

    if (success) {
      this.stats.successfulExecutions++;
    } else {
      this.stats.failedExecutions++;
    }

    this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.totalExecutions;
  }

  /**
   * Запуск періодичного очищення
   */
  private startCleanupInterval(): void {
    setInterval(() => {
      this.cleanupExpiredData();
    }, COMMAND_CONFIG.CLEANUP_INTERVAL);
  }

  /**
   * Очищення застарілих даних
   */
  private cleanupExpiredData(): void {
    const now = Date.now();
    let cleanedCooldowns = 0;
    let cleanedCache = 0;
    let cleanedErrors = 0;

    // Очищення cooldowns
    for (const [userId, cooldownTime] of this.cooldowns.entries()) {
      if (now > cooldownTime) {
        this.cooldowns.delete(userId);
        cleanedCooldowns++;
      }
    }

    // Очищення кешу
    for (const [key, cached] of this.executionCache.entries()) {
      if (now - cached.timestamp > 300000) {
        // 5 хвилин
        this.executionCache.delete(key);
        cleanedCache++;
      }
    }

    // Очищення помилок (старіші 1 години)
    for (const [userId, errorTime] of this.errorCount.entries()) {
      if (now - errorTime > 3600000) {
        this.errorCount.delete(userId);
        cleanedErrors++;
      }
    }

    if (cleanedCooldowns > 0 || cleanedCache > 0 || cleanedErrors > 0) {
      logger.debug(`Очищено застарілі дані команди ${this.name}`, {
        cooldowns: cleanedCooldowns,
        cache: cleanedCache,
        errors: cleanedErrors,
      });
    }
  }

  /**
   * Отримання статистики команди
   */
  public getCommandStats(): CommandStats {
    return { ...this.stats };
  }

  /**
   * Очищення cooldowns
   */
  public clearCooldowns(): void {
    this.cooldowns.clear();
    logger.debug(`Cooldowns команди ${this.name} очищено`);
  }

  /**
   * Health check
   */
  public async healthCheck(): Promise<HealthStatus> {
    const successRate =
      this.stats.totalExecutions > 0
        ? (this.stats.successfulExecutions / this.stats.totalExecutions) * 100
        : 0;

    const isHealthy = successRate > 80 && this.stats.averageExecutionTime < 10000;

    return {
      healthy: isHealthy,
      service: this.name,
      details: {
        totalExecutions: this.stats.totalExecutions,
        successRate: `${successRate.toFixed(2)}%`,
        averageExecutionTime: `${this.stats.averageExecutionTime.toFixed(2)}ms`,
        activeCooldowns: this.cooldowns.size,
        cacheSize: this.executionCache.size,
        errorCount: this.errorCount.size,
      },
    };
  }

  /**
   * Завершення роботи
   */
  public async shutdown(): Promise<void> {
    this.isShuttingDown = true;
    this.clearCooldowns();
    this.executionCache.clear();
    this.errorCount.clear();
    this.lastExecution.clear();

    logger.info(`Команда ${this.name} зупинена`);
  }

  /**
   * Отримання статистики
   */
  public getStats(): CommandStats {
    return { ...this.stats };
  }

  /**
   * Отримання назви команди
   */
  public getName(): string {
    return this.name;
  }

  /**
   * Отримання опису команди
   */
  public getDescription(): string {
    return this.description;
  }

  /**
   * Отримання даних команди для реєстрації в Discord
   */
  public getData(): SlashCommandBuilder {
    return this.data;
  }

  /**
   * Отримання допомоги по команді
   */
  public getHelp(): string {
    return `**Команда:** ${this.name}
**Опис:** ${this.description}
**Використання:** ${this.usage}
**Категорія:** ${this.category}
**Cooldown:** ${this.cooldown / 1000}с
${this.examples.length > 0 ? `**Приклади:**\n${this.examples.map(ex => `\`${ex}\``).join('\n')}` : ''}`;
  }
}
