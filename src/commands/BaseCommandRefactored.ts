/**
 * Рефакторований базовий клас для команд Discord бота
 * Використовує модульну архітектуру для кращої підтримки
 * Версія 4.0.0 - Модульна архітектура
 */

import type {
  ChatInputCommandInteraction,
  AutocompleteInteraction,
  MessageComponentInteraction} from 'discord.js';
import {
  SlashCommandBuilder,
  EmbedBuilder
} from 'discord.js';

import type { BotConfig, CommandStats, CommandContext } from '@/types';

import logger from '@/utils/logger';
import CommandValidator, {
  type ValidationResult,
  type ValidationRules,
} from './modules/CommandValidator';
import CommandMetricsCollector from './modules/CommandMetrics';

// Константи конфігурації
const COMMAND_CONFIG = {
  DEFAULT_COOLDOWN: 3000,
  MAX_EXECUTION_TIME: 30000,
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000,
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
}

export interface CommandExecuteOptions {
  interaction: ChatInputCommandInteraction;
  context?: CommandContext;
  startTime?: number;
  retryCount?: number;
  validationResult?: ValidationResult;
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

  protected stats: CommandStats;
  protected cooldowns: Map<string, number> = new Map();
  protected readonly config: BotConfig;
  protected isShuttingDown = false;
  protected readonly startedAt: number = Date.now();

  // Модульні компоненти
  protected validator: CommandValidator;
  protected metrics: CommandMetricsCollector;

  constructor(commandData: CommandData, config: BotConfig) {
    this.config = config;
    this.name = commandData.name;
    this.description = commandData.description;
    this.category = commandData.category || 'Загальні';
    this.usage = commandData.usage || `/${commandData.name}`;
    this.examples = commandData.examples || [];
    this.permissions = commandData.permissions || [];
    this.cooldown = commandData.cooldown || COMMAND_CONFIG.DEFAULT_COOLDOWN;

    // Створення SlashCommandBuilder
    this.data = new SlashCommandBuilder()
      .setName(commandData.name)
      .setDescription(commandData.description);

    // Налаштування дозволів
    if (commandData.defaultMemberPermissions) {
      this.data.setDefaultMemberPermissions(commandData.defaultMemberPermissions);
    }

    if (commandData.dmPermission !== undefined) {
      this.data.setDMPermission(commandData.dmPermission);
    }

    // Додавання опцій
    if (commandData.options) {
      this.addOptions(commandData.options);
    }

    // Ініціалізація статистики (узгоджено з CommandStats)
    this.stats = {
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      averageExecutionTime: 0,
      totalExecutionTime: 0,
      cacheHits: 0,
      cacheMisses: 0,
      retries: 0,
      service: this.name,
      uptime: 0,
      requests: 0,
      errors: 0,
    };

    // Ініціалізація модулів
    this.validator = new CommandValidator();
    this.metrics = new CommandMetricsCollector();

    logger.debug('Команда ініціалізована', {
      type: 'command',
      component: this.name,
      event: 'initialized',
    });
  }

  /**
   * Головна точка входу для виконання команди
   */
  public async handleInteraction(interaction: ChatInputCommandInteraction): Promise<void> {
    const startTime = Date.now();
    let success = false;
    let error: string | undefined;

    try {
      // Перевірка cooldown
      if (this.isOnCooldown(interaction.user.id)) {
        const remainingTime = this.getRemainingCooldown(interaction.user.id);
        await this.sendCooldownMessage(interaction, remainingTime);
        this.metrics.recordCooldownHit(this.name);
        return;
      }

      // Валідація команди
      const validationResult = await this.validateInteraction(interaction);
      if (!validationResult.isValid) {
        await this.sendValidationError(interaction, validationResult);
        return;
      }

      // Встановлення cooldown
      this.setCooldown(interaction.user.id);

      // Виконання команди з retry логікою
      await this.executeWithRetry({
        interaction,
        startTime,
        retryCount: 0,
        validationResult,
      });

      success = true;
    } catch (err) {
      error = err instanceof Error ? err.message : String(err);
      logger.error('Помилка виконання команди', {
        type: 'command',
        component: this.name,
        event: 'execute_error',
        userId: interaction.user.id,
        errorMessage: error,
        durationMs: Date.now() - startTime,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
      });

      await this.handleExecutionError(interaction, err);
    } finally {
      // Запис метрик
      const duration = Date.now() - startTime;
      this.updateStats(duration, success);
      if (error) {
        this.metrics.recordExecution(this.name, interaction.user.id, duration, success, { error });
      } else {
        this.metrics.recordExecution(this.name, interaction.user.id, duration, success, {});
      }
    }
  }

  /**
   * Валідація взаємодії
   */
  protected async validateInteraction(
    interaction: ChatInputCommandInteraction,
    customRules?: ValidationRules
  ): Promise<ValidationResult> {
    try {
      // Базова валідація
      const baseValidation = await this.validator.validateCommand(interaction, customRules);

      // Кастомна валідація команди
      const customValidation = await this.customValidation(interaction);

      // Об'єднання результатів
      const combinedErrors = [...baseValidation.errors, ...customValidation.errors];
      const combinedWarnings = [...baseValidation.warnings, ...customValidation.warnings];

      return {
        isValid: combinedErrors.length === 0,
        errors: combinedErrors,
        warnings: combinedWarnings,
        sanitizedValues: {
          ...baseValidation.sanitizedValues,
          ...customValidation.sanitizedValues,
        },
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
        errors: ['Внутрішня помилка валідації'],
        warnings: [],
      };
    }
  }

  /**
   * Кастомна валідація для конкретної команди
   */
  protected async customValidation(
    _interaction: ChatInputCommandInteraction
  ): Promise<ValidationResult> {
    // Базова реалізація - може бути перевизначена в дочірніх класах
    return {
      isValid: true,
      errors: [],
      warnings: [],
    };
  }

  /**
   * Виконання з повторними спробами
   */
  private async executeWithRetry(options: CommandExecuteOptions): Promise<void> {
    const { retryCount = 0 } = options;

    try {
      await this.execute(options);
    } catch (error) {
      if (retryCount < COMMAND_CONFIG.MAX_RETRIES && this.shouldRetry(error)) {
        logger.warn('Повторна спроба виконання команди', {
          type: 'command',
          component: this.name,
          event: 'retry',
          attempt: retryCount + 1,
          maxAttempts: COMMAND_CONFIG.MAX_RETRIES,
        });

        await new Promise(resolve =>
          setTimeout(resolve, COMMAND_CONFIG.RETRY_DELAY * (retryCount + 1))
        );

        await this.executeWithRetry({
          ...options,
          retryCount: retryCount + 1,
        });
      } else {
        throw error;
      }
    }
  }

  /**
   * Перевірка чи потрібно повторити виконання
   */
  protected shouldRetry(error: unknown): boolean {
    if (error instanceof Error) {
      // Повторюємо для тимчасових помилок мережі
      return (
        error.message.includes('timeout') ||
        error.message.includes('network') ||
        error.message.includes('ECONNRESET') ||
        error.message.includes('rate limit')
      );
    }
    return false;
  }

  /**
   * Управління cooldown
   */
  protected isOnCooldown(userId: string): boolean {
    const userCooldown = this.cooldowns.get(userId);
    return userCooldown ? Date.now() < userCooldown : false;
  }

  protected setCooldown(userId: string): void {
    this.cooldowns.set(userId, Date.now() + this.cooldown);

    // Автоматичне видалення після закінчення cooldown
    setTimeout(() => {
      this.cooldowns.delete(userId);
    }, this.cooldown);
  }

  protected getRemainingCooldown(userId: string): number {
    const userCooldown = this.cooldowns.get(userId);
    return userCooldown ? Math.max(0, userCooldown - Date.now()) : 0;
  }

  /**
   * Відправка повідомлень про помилки
   */
  protected async sendCooldownMessage(
    interaction: ChatInputCommandInteraction,
    remainingTime: number
  ): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor(0xffa500)
      .setTitle('⏱️ Cooldown')
      .setDescription(
        `Зачекайте ще ${Math.ceil(remainingTime / 1000)} секунд перед наступним використанням команди.`
      )
      .setTimestamp();

    if (interaction.replied || interaction.deferred) {
      await interaction.followUp({ embeds: [embed], ephemeral: true });
    } else {
      await interaction.reply({ embeds: [embed], ephemeral: true });
    }
  }

  protected async sendValidationError(
    interaction: ChatInputCommandInteraction,
    validation: ValidationResult
  ): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor(0xff0000)
      .setTitle('❌ Помилка валідації')
      .setDescription(validation.errors.join('\n'))
      .setTimestamp();

    if (validation.warnings.length > 0) {
      embed.addFields({
        name: '⚠️ Попередження',
        value: validation.warnings.join('\n'),
        inline: false,
      });
    }

    if (interaction.replied || interaction.deferred) {
      await interaction.followUp({ embeds: [embed], ephemeral: true });
    } else {
      await interaction.reply({ embeds: [embed], ephemeral: true });
    }
  }

  protected async handleExecutionError(
    interaction: ChatInputCommandInteraction,
    error: unknown
  ): Promise<void> {
    const embed = new EmbedBuilder()
      .setColor(0xff0000)
      .setTitle('❌ Помилка виконання')
      .setDescription('Виникла помилка під час виконання команди. Спробуйте пізніше.')
      .setTimestamp();

    // В development режимі показуємо деталі помилки
    if (this.config.logging?.level === 'debug' && error instanceof Error) {
      embed.addFields({
        name: 'Деталі помилки',
        value: error.message.substring(0, 1000),
        inline: false,
      });
    }

    try {
      if (interaction.replied || interaction.deferred) {
        await interaction.followUp({ embeds: [embed], ephemeral: true });
      } else {
        await interaction.reply({ embeds: [embed], ephemeral: true });
      }
    } catch (replyError) {
      logger.error('Не вдалося відправити повідомлення про помилку', {
        type: 'command',
        component: this.name,
        event: 'error_reply_failed',
        errorName: replyError instanceof Error ? replyError.name : undefined,
        errorMessage: replyError instanceof Error ? replyError.message : String(replyError),
        stack: replyError instanceof Error ? replyError.stack : undefined,
      });
    }
  }

  /**
   * Оновлення статистики
   */
  protected updateStats(executionTime: number, success: boolean): void {
    this.stats.totalExecutions++;
    this.stats.requests++;
    this.stats.totalExecutionTime += executionTime;
    this.stats.averageExecutionTime = this.stats.totalExecutionTime / this.stats.totalExecutions;

    if (success) {
      this.stats.successfulExecutions++;
    } else {
      this.stats.failedExecutions++;
      this.stats.errors++;
    }
  }

  /**
   * Додавання опцій до команди
   */
  private addOptions(options: any[]): void {
    options.forEach(option => {
      switch (option.type) {
        case 'string':
          this.data.addStringOption(opt => {
            opt
              .setName(option.name)
              .setDescription(option.description)
              .setRequired(option.required || false);

            if (option.choices) {
              opt.addChoices(...option.choices);
            }

            return opt;
          });
          break;

        case 'integer':
          this.data.addIntegerOption(opt => {
            opt
              .setName(option.name)
              .setDescription(option.description)
              .setRequired(option.required || false);

            if (option.min_value !== undefined) opt.setMinValue(option.min_value);
            if (option.max_value !== undefined) opt.setMaxValue(option.max_value);

            return opt;
          });
          break;

        case 'boolean':
          this.data.addBooleanOption(opt => {
            return opt
              .setName(option.name)
              .setDescription(option.description)
              .setRequired(option.required || false);
          });
          break;

        case 'user':
          this.data.addUserOption(opt => {
            return opt
              .setName(option.name)
              .setDescription(option.description)
              .setRequired(option.required || false);
          });
          break;

        case 'attachment':
          this.data.addAttachmentOption(opt => {
            return opt
              .setName(option.name)
              .setDescription(option.description)
              .setRequired(option.required || false);
          });
          break;
      }
    });
  }

  /**
   * Отримання статистики команди
   */
  public getStats(): CommandStats {
    return { ...this.stats, uptime: Date.now() - this.startedAt } as CommandStats;
  }

  /**
   * Скидання статистики
   */
  public resetStats(): void {
    this.stats = {
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      averageExecutionTime: 0,
      totalExecutionTime: 0,
      cacheHits: 0,
      cacheMisses: 0,
      retries: 0,
      service: this.name,
      uptime: 0,
      requests: 0,
      errors: 0,
    };
  }

  /**
   * Створення стандартного embed відповіді
   */
  protected createEmbed(
    title: string,
    description: string,
    color: number = 0x00ae86
  ): EmbedBuilder {
    return new EmbedBuilder()
      .setColor(color)
      .setTitle(title)
      .setDescription(description)
      .setTimestamp()
      .setFooter({
        text: `${this.name} | Discord AI Assistant Bot`,
        iconURL: 'https://cdn.discordapp.com/embed/avatars/0.png',
      });
  }

  /**
   * Перевірка дозволів
   */
  protected hasPermission(interaction: ChatInputCommandInteraction, permission: string): boolean {
    if (!interaction.guild || !interaction.member) return false;

    const member: any = interaction.member as any;
    const perms: any = member?.permissions;
    return Boolean(perms && typeof perms.has === 'function' && perms.has(permission as any));
  }

  /**
   * Shutdown hook для очищення ресурсів
   */
  public shutdown(): void {
    this.isShuttingDown = true;
    this.cooldowns.clear();
    logger.debug(`🛑 Команда "${this.name}" зупинена`);
  }

  // Абстрактні методи які повинні бути реалізовані в дочірніх класах
  abstract execute(options: CommandExecuteOptions): Promise<void>;

  // Опціональні методи для розширення функціональності
  async handleAutocomplete?(interaction: AutocompleteInteraction): Promise<void>;
  async handleComponent?(interaction: MessageComponentInteraction): Promise<void>;
}

export default BaseCommand;
