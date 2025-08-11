/**
 * Валідатор команд Discord бота
 * Централізована логіка валідації та санітизації
 * Версія 1.0.0 - Виокремлено з BaseCommand
 */

import type { ChatInputCommandInteraction } from 'discord.js';
import type { LogMeta } from '@/types';
import { sanitizeInput } from '@/utils/security';

import logger from '@/utils/logger';

export interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  sanitizedOptions?: any;
  sanitizedValues?: Record<string, unknown>;
}

export interface ValidationRules {
  maxStringLength?: number;
  maxNumberValue?: number;
  minNumberValue?: number;
  requiredFields?: string[];
  allowedValues?: Record<string, unknown[]>;
  customValidators?: Array<(value: unknown, field: string) => ValidationResult>;
}

export class CommandValidator {
  private static instance: CommandValidator | null = null;

  constructor() {
    if (CommandValidator.instance) {
      return CommandValidator.instance;
    }
    CommandValidator.instance = this;
  }

  /**
   * Головна функція валідації команди
   */
  public async validateCommand(
    interaction: ChatInputCommandInteraction,
    rules?: ValidationRules
  ): Promise<ValidationResult> {
    try {
      const errors: string[] = [];
      const warnings: string[] = [];
      const sanitizedValues: Record<string, unknown> = {};

      // Валідація базових параметрів
      if (!interaction.commandName) {
        errors.push('Назва команди відсутня');
      }

      if (!interaction.user) {
        errors.push('Користувач не ідентифікований');
      }

      // Валідація опцій команди
      const optionsResult = this.validateOptions(interaction, rules);
      errors.push(...optionsResult.errors);
      warnings.push(...optionsResult.warnings);
      Object.assign(sanitizedValues, optionsResult.sanitizedValues);

      // Валідація дозволів користувача
      const permissionsResult = this.validateUserPermissions(interaction);
      errors.push(...permissionsResult.errors);
      warnings.push(...permissionsResult.warnings);

      // Валідація контексту виконання
      const contextResult = this.validateExecutionContext(interaction);
      errors.push(...contextResult.errors);
      warnings.push(...contextResult.warnings);

      const isValid = errors.length === 0;

      // Логування результату валідації
      if (!isValid) {
        logger.warn('⚠️ Валідація команди невдала', { command: interaction.commandName, userId: interaction.user?.id, errors, warnings } as LogMeta);
      } else if (warnings.length > 0) {
        logger.debug('ℹ️ Валідація команди з попередженнями', { command: interaction.commandName, userId: interaction.user?.id, warnings } as LogMeta);
      }

      return {
        isValid,
        errors,
        warnings,
        sanitizedValues
      };
    } catch (error) {
      logger.error('❌ Помилка валідації команди:', { error } as LogMeta);
      return {
        isValid: false,
        errors: ['Внутрішня помилка валідації'],
        warnings: []
      };
    }
  }

  /**
   * Валідація опцій команди
   */
  private validateOptions(
    interaction: ChatInputCommandInteraction,
    rules?: ValidationRules
  ): ValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];
    const sanitizedValues: Record<string, unknown> = {};

    try {
      // Отримання всіх опцій
      const options = interaction.options.data;

      for (const option of options) {
        const { name, value, type: _type } = option;

        // Перевірка обов'язкових полів
        if (rules?.requiredFields?.includes(name) && (!value || value === '')) {
          errors.push(`Поле '${name}' є обов'язковим`);
          continue;
        }

        // Санітизація значень
        if (typeof value === 'string') {
          const sanitized = this.sanitizeStringValue(value, rules);
          sanitizedValues[name] = sanitized.value;
          errors.push(...sanitized.errors);
          warnings.push(...sanitized.warnings);
        } else if (typeof value === 'number') {
          const validated = this.validateNumberValue(value, name, rules);
          sanitizedValues[name] = validated.value;
          errors.push(...validated.errors);
          warnings.push(...validated.warnings);
        } else {
          sanitizedValues[name] = value;
        }

        // Перевірка дозволених значень
        if (rules?.allowedValues?.[name]) {
          if (!rules.allowedValues[name].includes(value)) {
            errors.push(`Недозволене значення для поля '${name}': ${value}`);
          }
        }

        // Користувацькі валідатори
        if (rules?.customValidators) {
          for (const validator of rules.customValidators) {
            const result = validator(value, name);
            errors.push(...result.errors);
            warnings.push(...result.warnings);
          }
        }
      }

      return {
        isValid: errors.length === 0,
        errors,
        warnings,
        sanitizedValues
      };
    } catch (error) {
      return {
        isValid: false,
        errors: ['Помилка валідації опцій'],
        warnings: []
      };
    }
  }

  /**
   * Санітизація строкових значень
   */
  private sanitizeStringValue(
    value: string,
    rules?: ValidationRules
  ): {
    value: string;
    errors: string[];
    warnings: string[];
  } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Санітизація через security utils
    const sanitized = sanitizeInput(value);

    // Перевірка довжини
    if (rules?.maxStringLength && sanitized.length > rules.maxStringLength) {
      errors.push(`Текст занадто довгий (макс. ${rules.maxStringLength} символів)`);
      const truncated = sanitized.substring(0, rules.maxStringLength);
      warnings.push('Текст було обрізано');
      return { value: truncated, errors, warnings };
    }

    // Перевірка на підозрілий контент
    if (this.containsSuspiciousContent(sanitized)) {
      warnings.push('Виявлено потенційно небезпечний контент');
    }

    return { value: sanitized, errors, warnings };
  }

  /**
   * Валідація числових значень
   */
  private validateNumberValue(
    value: number,
    fieldName: string,
    rules?: ValidationRules
  ): {
    value: number;
    errors: string[];
    warnings: string[];
  } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Перевірка діапазону
    if (rules?.minNumberValue !== undefined && value < rules.minNumberValue) {
      errors.push(`Значення '${fieldName}' занадто мале (мін. ${rules.minNumberValue})`);
    }

    if (rules?.maxNumberValue !== undefined && value > rules.maxNumberValue) {
      errors.push(`Значення '${fieldName}' занадто велике (макс. ${rules.maxNumberValue})`);
    }

    // Перевірка на розумність значення
    if (value < 0 && fieldName.includes('count')) {
      warnings.push(`Від'ємне значення для лічильника: ${fieldName}`);
    }

    return { value, errors, warnings };
  }

  /**
   * Валідація дозволів користувача
   */
  private validateUserPermissions(interaction: ChatInputCommandInteraction): ValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      // Перевірка що користувач існує
      if (!interaction.user) {
        errors.push('Користувач не ідентифікований');
        return { isValid: false, errors, warnings };
      }

      // Перевірка що команда виконується на сервері (якщо потрібно)
      if (!interaction.guild) {
        warnings.push('Команда виконується поза сервером');
      }

      // Перевірка member об'єкта
      if (interaction.guild && !interaction.member) {
        errors.push('Не вдалося отримати інформацію про учасника сервера');
      }

      return {
        isValid: errors.length === 0,
        errors,
        warnings
      };
    } catch (error) {
      return {
        isValid: false,
        errors: ['Помилка валідації дозволів'],
        warnings: []
      };
    }
  }

  /**
   * Валідація контексту виконання
   */
  private validateExecutionContext(interaction: ChatInputCommandInteraction): ValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      // Перевірка каналу
      if (!interaction.channel) {
        errors.push('Канал недоступний');
      }

      // Перевірка що interaction не застарілий
      const interactionAge = Date.now() - interaction.createdTimestamp;
      if (interactionAge > 15 * 60 * 1000) { // 15 хвилин
        warnings.push('Interaction застарілий');
      }

      // Перевірка що bot має дозволи у каналі
      if (interaction.guild && interaction.channel) {
        const botMember = interaction.guild.members.me;
        if (botMember && 'permissionsFor' in interaction.channel) {
          const permissions = interaction.channel.permissionsFor(botMember);
          if (!permissions?.has(['SendMessages', 'ViewChannel'])) {
            errors.push('Бот не має необхідних дозволів у цьому каналі');
          }
        }
      }

      return {
        isValid: errors.length === 0,
        errors,
        warnings
      };
    } catch (error) {
      return {
        isValid: false,
        errors: ['Помилка валідації контексту'],
        warnings: []
      };
    }
  }

  /**
   * Перевірка на підозрілий контент
   */
  private containsSuspiciousContent(text: string): boolean {
    const suspiciousPatterns = [
      /discord\.gg\/[a-zA-Z0-9]+/gi, // Discord invite links
      /https?:\/\/[^\s]+/gi, // URLs
      /@everyone|@here/gi, // Mass mentions
      /\b(free|nitro|giveaway)\b/gi, // Suspicious keywords
      /<[@#!&][0-9]+>/gi, // Discord mentions
    ];

    return suspiciousPatterns.some(pattern => pattern.test(text));
  }

  /**
   * Валідація з кастомними правилами
   */
  public async validateWithRules(
    interaction: ChatInputCommandInteraction,
    rules: ValidationRules
  ): Promise<ValidationResult> {
    return this.validateCommand(interaction, rules);
  }

  /**
   * Швидка валідація без складних перевірок
   */
  public quickValidate(interaction: ChatInputCommandInteraction): ValidationResult {
    const errors: string[] = [];

    if (!interaction.commandName) errors.push('Назва команди відсутня');
    if (!interaction.user) errors.push('Користувач не ідентифікований');
    if (!interaction.channel) errors.push('Канал недоступний');

    return {
      isValid: errors.length === 0,
      errors,
      warnings: []
    };
  }

  /**
   * Отримання статистики валідації
   */
  public getValidationStats(): {
    totalValidations: number;
    successfulValidations: number;
    failedValidations: number;
  } {
    // TODO: Реалізувати збір статистики
    return {
      totalValidations: 0,
      successfulValidations: 0,
      failedValidations: 0
    };
  }
}

export default CommandValidator;