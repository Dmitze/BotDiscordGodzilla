/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */

import { SlashCommandBuilder } from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';

interface AIQueryResult {
  response: string;
  confidence: number;
  action?: string;
  actionData?: {
    type: string;
    format?: string;
  };
}

interface ValidationResult {
  isValid: boolean;
  errors: string[];
}

interface ValidationSchema {
  [key: string]: {
    required: boolean;
    type: string;
    maxLength?: number;
    sanitize?: string;
  };
}

export class AIAssistantCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'ai',
      '🤖 AI-асистент для роботи з Google Sheets',
      config,
      (builder: any) => {
        return builder
          .addStringOption((option: any) =>
            option
              .setName('запит')
              .setDescription(
                'Що ви хочете зробити? (наприклад: "знайди товари iPhone", "проаналізуй залишки")'
              )
              .setRequired(true)
              .setMaxLength(1000)
          )
          .addStringOption((option: any) =>
            option
              .setName('контекст')
              .setDescription('Додатковий контекст для AI')
              .setRequired(false)
              .setMaxLength(500)
          );
      }
    );
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      // Перевірка прав доступу
      const hasAccess = await this.checkPermission(interaction);
      if (!hasAccess) {
        return;
      }

      // Валідація параметрів
      const commandOptions = {
        query: interaction.options.getString('запит'),
        context: interaction.options.getString('контекст'),
      };

      const validation = this.validateCommandOptions(commandOptions);
      if (!validation.isValid) {
        await interaction.reply({
          content: `❌ Помилка валідації:\n${validation.errors.join('\n')}`,
          ephemeral: true,
        });
        return;
      }

      // Логування події
      this.logSecurityEvent('ai_command_executed', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        query: commandOptions.query,
        context: commandOptions.context,
      });

      // Відповідь про обробку
      await interaction.deferReply();

      // Обробка запиту через AI
      const result = await this.processAIQuery(
        interaction.user.id,
        commandOptions.query || '',
        commandOptions.context
      );

      // Формування відповіді
      let response = `🤖 **AI-асистент**\n\n`;

      if (result.confidence < 0.7) {
        response += `⚠️ **Низька впевненість** (${Math.round(result.confidence * 100)}%)\n`;
      }

      response += `**Ваш запит:** ${commandOptions.query}\n\n`;
      response += `**Відповідь:**\n${result.response}`;

      // Додавання контексту якщо є
      if (commandOptions.context) {
        response += `\n\n**Контекст:** ${commandOptions.context}`;
      }

      // Додавання додаткової інформації
      if (result.actionData) {
        response += `\n\n**Дія:** ${result.actionData.type}`;
        if (result.actionData.format) {
          response += ` (формат: ${result.actionData.format})`;
        }
      }

      // Відправка відповіді
      await interaction.editReply({
        content: response,
        ephemeral: false,
      });

      // Логування успішного виконання
      console.log(`AI command executed successfully for ${interaction.user.tag}`, {
        action: result.action,
        confidence: result.confidence,
        hasActionData: !!result.actionData,
      });
    } catch (error) {
      console.error('AI Assistant command error:', error);

      const errorMessage =
        '❌ Помилка при обробці AI-запиту. Спробуйте ще раз або зверніться до адміністратора.';

      if (interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermission(interaction: any): Promise<boolean> {
    // TODO: Реалізувати перевірку прав доступу
    // Тимчасова реалізація - дозволяємо всім
    return true;
  }

  /**
   * Валідація параметрів команди
   */
  private validateCommandOptions(options: Record<string, any>): ValidationResult {
    const validationSchema: ValidationSchema = {
      query: {
        required: true,
        type: 'string',
        maxLength: 1000,
        sanitize: 'ai_prompt',
      },
      context: {
        required: false,
        type: 'string',
        maxLength: 500,
        sanitize: 'ai_prompt',
      },
    };

    const errors: string[] = [];

    for (const [key, schema] of Object.entries(validationSchema)) {
      const value = options[key];

      if (schema.required && !value) {
        errors.push(`${key} є обов'язковим`);
        continue;
      }

      if (value && typeof value !== schema.type) {
        errors.push(`${key} має бути типу ${schema.type}`);
        continue;
      }

      if (value && schema.maxLength && value.length > schema.maxLength) {
        errors.push(`${key} не може бути довшим за ${schema.maxLength} символів`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
    };
  }

  /**
   * Логування події безпеки
   */
  private logSecurityEvent(eventType: string, data: Record<string, any>): void {
    // TODO: Реалізувати логування подій безпеки
    console.log(`Security event: ${eventType}`, data);
  }

  /**
   * Обробка AI запиту
   */
  private async processAIQuery(userId: string, query: string, context?: string): Promise<AIQueryResult> {
    // TODO: Інтеграція з AI сервісом
    // Тимчасова реалізація
    const response = `Це тимчасова відповідь AI на запит: "${query}"`;
    
    return {
      response,
      confidence: 0.8,
      action: 'search',
      actionData: {
        type: 'search',
        format: 'text',
      },
    };
  }
} 