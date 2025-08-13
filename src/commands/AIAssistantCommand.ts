/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';

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
  private readonly googleService: GoogleService | undefined;

  constructor(config: BotConfig, googleService?: GoogleService) {
    super('ai', '🤖 AI-асистент для роботи з Google Sheets', config, {}, (builder: any) => {
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
    });

    this.googleService = googleService;
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
        command: this.name,
        query: commandOptions.query,
        context: commandOptions.context,
        guildId: interaction.guildId,
        channelId: interaction.channelId,
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
      logger.info('AI command executed successfully', {
        type: 'command',
        command: this.name,
        component: 'AIAssistantCommand.onExecute',
        userTag: interaction.user.tag,
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        action: result.action,
        confidence: result.confidence,
        hasActionData: !!result.actionData,
      });
    } catch (error) {
      logger.error('AI Assistant command error', {
        type: 'command',
        command: this.name,
        component: 'AIAssistantCommand.onExecute',
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

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
    try {
      const { PermissionManager } = await import('../core/PermissionManager');
      const permissionManager = new PermissionManager(this.config);
      const result = await permissionManager.checkPermission(
        interaction.user,
        interaction.member ?? null,
        interaction.commandName,
        interaction.channelId
      );

      if (!result.allowed) {
        this.logSecurityEvent('command_access_denied', {
          severity: 'medium',
          reason: result.reason,
          userLevel: result.userLevel,
          userId: interaction.user.id,
          ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
          channelId: interaction.channelId,
          command: interaction.commandName,
        });

        await interaction.reply({
          content: `🚫 Доступ заборонено: ${result.reason}`,
          ephemeral: true,
        });
        return false;
      }

      logger.info('✅ Доступ до AI-команди дозволено', {
        type: 'security',
        event: 'permission_granted',
        component: 'AIAssistantCommand.checkPermission',
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        userLevel: result.userLevel,
        remainingUses: result.remainingUses,
      });
      return true;
    } catch (error) {
      logger.error('❌ Помилка перевірки прав у AI-команді', {
        type: 'security',
        component: 'AIAssistantCommand.checkPermission',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        userId: interaction.user?.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        severity: 'high',
      });
      // За замовчуванням не пускати у разі помилки
      await interaction.reply({ content: '❌ Помилка перевірки прав доступу', ephemeral: true });
      return false;
    }
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
    logger.security(eventType, data['userId'] || 'unknown', {
      type: 'security',
      component: 'AIAssistantCommand',
      command: this.name,
      ...data,
    });
  }

  /**
   * Обробка AI запиту
   */
  private async processAIQuery(
    _userId: string,
    query: string,
    _context?: string
  ): Promise<AIQueryResult> {
    const q = (query || '').toLowerCase();

    // Проста детекція наміру: показати таблиці Google
    const intentListSheets =
      /таблиц|таблицы|лист(ы|и)?|sheets?|список.*таблиц|какие.*таблиц|google\s*диск|google\s*sheets/.test(
        q
      ) && /какие|покажи|список|list|что|найд/i.test(q);

    if (intentListSheets) {
      try {
        if (!this.googleService) {
          return {
            response:
              'GoogleService не доступний для цієї команди. Перевірте ініціалізацію сервісів або конфігурацію.',
            confidence: 0.6,
            action: 'list_sheets',
            actionData: { type: 'list', format: 'text' },
          };
        }

        const folderId = this.config.google.driveFolderId;
        if (!folderId) {
          return {
            response:
              'Не налаштовано google.driveFolderId у конфігурації. Додайте ID каталогу з таблицями.',
            confidence: 0.6,
            action: 'list_sheets',
            actionData: { type: 'list', format: 'text' },
          };
        }

        const files = await this.googleService.listDriveFilesInFolder(folderId, {
          recursive: true,
          type: 'sheet',
          limit: 50,
          maxDepth: 3,
        });

        if (!files.length) {
          return {
            response: 'Таблиці не знайдені у вказаній папці Google Drive.',
            confidence: 0.9,
            action: 'list_sheets',
            actionData: { type: 'list', format: 'text' },
          };
        }

        const lines = files.slice(0, 20).map((f, idx) => `${idx + 1}. ${f.name} (id: ${f.id})`);
        const more = files.length > 20 ? `\n… та ще ${files.length - 20}` : '';

        return {
          response: `Знайдено ${files.length} таблиць:\n\n${lines.join('\n')}${more}`,
          confidence: 0.95,
          action: 'list_sheets',
          actionData: { type: 'list', format: 'text' },
        };
      } catch (error) {
        logger.error('Помилка отримання списку таблиць', {
          type: 'command',
          component: 'AIAssistantCommand.processAIQuery',
          error: error instanceof Error ? error.message : String(error),
        });
        return {
          response:
            'Сталася помилка при отриманні списку таблиць Google. Спробуйте пізніше або перевірте доступи.',
          confidence: 0.4,
          action: 'list_sheets',
          actionData: { type: 'list', format: 'text' },
        };
      }
    }

    // Базова тимчасова відповідь
    const response = `Це тимчасова відповідь AI на запит: "${query}"`;
    return {
      response,
      confidence: 0.8,
      action: 'search',
      actionData: { type: 'search', format: 'text' },
    };
  }
}
