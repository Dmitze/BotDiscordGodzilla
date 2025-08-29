/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

import type { GoogleService } from '@/services/GoogleService';
import type {
  SlashCommandBuilder,
  SlashCommandStringOption,
  ChatInputCommandInteraction,
  GuildMember,
} from 'discord.js';
import { AnalyticsService } from '@/services/AnalyticsService';
import { t } from '@/i18n';
import {
  tokenizeName,
  findMonthNumber,
  isImageMime,
  isDocLikeMime,
  ensureDriveIndex,
  readGoogleSheet,
  readExcelBuffer,
} from '@/commands/modules/aiAssistant/helpers';

interface AIQueryResult {
  response: string;
  confidence: number;
  action?: string;
  actionData?: {
    type: string;
    format?: string;
  };
}

interface UserContext {
  query: string;
  response: string;
  timestamp: number;
  fileIds?: string[];
  action?: string;
}

interface UserSession {
  userId: string;
  contexts: UserContext[];
  lastActivity: number;
}

type AICommandOptions = {
  query: string | null;
  context: string | null;
};

type DriveIndexedFile = {
  id?: string;
  name?: string;
  mimeType?: string;
};

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
  
  // Статичне сховище контексту для всіх користувачів
  private static userSessions = new Map<string, UserSession>();
  private static readonly MAX_CONTEXTS = 20;
  private static readonly SESSION_TIMEOUT = 2 * 60 * 60 * 1000; // 2 години

  constructor(config: BotConfig, googleService?: GoogleService) {
    super('ai', t('ai.command.description'), config, {
      i18n: { nameKey: 'commands.ai.name', descriptionKey: 'ai.command.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('запит')
          .setDescription(t('ai.opt.query.description'))
          .setRequired(true)
          .setMaxLength(1000)
      );
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('контекст')
          .setDescription(t('ai.opt.context.description'))
          .setRequired(false)
          .setMaxLength(500)
      );
      return builder;
    });

    this.googleService = googleService;
  }

  /**
   * Отримання контексту користувача
   */
  private static getUserSession(userId: string): UserSession {
    const now = Date.now();
    let session = AIAssistantCommand.userSessions.get(userId);
    
    if (!session || (now - session.lastActivity) > AIAssistantCommand.SESSION_TIMEOUT) {
      const isNewSession = !session;
      session = {
        userId,
        contexts: [],
        lastActivity: now
      };
      AIAssistantCommand.userSessions.set(userId, session);
      
      if (isNewSession) {
        logger.debug('Нова сесія створена', { userId });
      } else {
        logger.debug('Сесія оновлена (минула час очікування)', { userId });
      }
    } else {
      session.lastActivity = now;
    }
    
    return session;
  }

  /**
   * Додавання нового контексту
   */
  private static addContext(userId: string, query: string, response: string, action?: string, fileIds?: string[]): void {
    const session = AIAssistantCommand.getUserSession(userId);
    
    const context: UserContext = {
      query,
      response,
      timestamp: Date.now(),
      ...(action && { action }),
      ...(fileIds && { fileIds })
    };
    
    session.contexts.push(context);
    
    // Обмежуємо кількість контекстів
    if (session.contexts.length > AIAssistantCommand.MAX_CONTEXTS) {
      session.contexts = session.contexts.slice(-AIAssistantCommand.MAX_CONTEXTS);
    }
    
    logger.debug('Контекст додано', {
      userId,
      contextCount: session.contexts.length,
      queryPreview: query.substring(0, 50)
    });
  }

  /**
   * Отримання контексту для промпта
   */
  private static getContextForPrompt(userId: string, currentQuery: string): string {
    const session = AIAssistantCommand.getUserSession(userId);
    
    if (session.contexts.length === 0) {
      logger.debug('Немає контексту для користувача', { userId });
      return '';
    }
    
    // Отримуємо більше контекстів для кращої повноти відповіді (останні 5 контекстів)
    const recentContexts = session.contexts.slice(-5);
    
    let contextText = '\n\n📎 ПОВНИЙ КОНТЕКСТ ПОПЕРЕДНІХ ЗАПИТІВ:\n';
    
    recentContexts.forEach((ctx, index) => {
      const timeAgo = Math.round((Date.now() - ctx.timestamp) / 1000 / 60); // хвилин
      contextText += `\n--- Контекст ${index + 1} (${timeAgo}хв тому) ---\n`;
      contextText += `Запит: "${ctx.query}"\n`;
      contextText += `Відповідь: "${ctx.response}"\n`;
      if (ctx.fileIds && ctx.fileIds.length > 0) {
        contextText += `Файли: ${ctx.fileIds.join(', ')}\n`;
      }
      if (ctx.action) {
        contextText += `Дія: ${ctx.action}\n`;
      }
    });
    
    // Перевіряємо, чи поточний запит посилається на контекст
    const contextReferences = [
      /а як щодо/i, /у тому ж файлі/i, /далі/i, /а тепер/i,
      /там само/i, /тією ж таблицю/i, /раніше казав/i,
      /продовж/i, /доповн/i, /ще/i, /більш/i, /детальн/i
    ];
    
    const hasContextReference = contextReferences.some(pattern => pattern.test(currentQuery));
    
    if (hasContextReference) {
      contextText += '\nℹ️ Користувач посилається на попередні дані. Використовуй весь доступний контекст для повної відповіді!\n';
      logger.debug('Виявлено посилання на контекст', { userId, query: currentQuery.substring(0, 50) });
    }
    
    logger.debug('Контекст підготовлено для промпта', {
      userId,
      contextCount: recentContexts.length,
      hasContextReference,
      contextLength: contextText.length
    });
    
    return contextText;
  }

  /**
   * Очищення старих сесій
   */
  private static cleanupOldSessions(): void {
    const now = Date.now();
    const toDelete: string[] = [];
    
    for (const [userId, session] of AIAssistantCommand.userSessions.entries()) {
      if ((now - session.lastActivity) > AIAssistantCommand.SESSION_TIMEOUT) {
        toDelete.push(userId);
      }
    }
    
    toDelete.forEach(userId => AIAssistantCommand.userSessions.delete(userId));
    
    if (toDelete.length > 0) {
      logger.debug(`Очищено ${toDelete.length} старих сесій`);
    }
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      // Перевіряємо стан інтеракції перед defer
      if (!interaction.deferred && !interaction.replied) {
        await interaction.deferReply({ ephemeral: true });
      }

      // Перевірка прав доступу (після defer)
      const hasAccess = await this.checkPermission(interaction);
      if (!hasAccess) {
        // checkPermission already handled the response
        return;
      }

      // Валідація параметрів
      const commandOptions: AICommandOptions = {
        query: interaction.options.getString('запит'),
        context: interaction.options.getString('контекст'),
      };

      const validation = this.validateCommandOptions(commandOptions);
      if (!validation.isValid) {
        if (interaction.deferred) {
          await interaction.editReply({
            content: t('ai.validation.failed', { errors: validation.errors.join('\n') }),
          });
        } else if (!interaction.replied) {
          await interaction.reply({
            content: t('ai.validation.failed', { errors: validation.errors.join('\n') }),
            ephemeral: true
          });
        }
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

      // Обробка запиту через AI
      const result = await this.processAIQuery(
        interaction.user.id,
        commandOptions.query || '',
        commandOptions.context ?? undefined,
        interaction
      );

      // Формування відповіді
      const response = this.buildResponse(result, commandOptions);

      // Відправка відповіді
      if (interaction.deferred) {
        await interaction.editReply({ content: response });
      } else if (!interaction.replied) {
        await interaction.reply({ content: response, ephemeral: true });
      }

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

      const errorMessage = t('ai.error.process');

      // Перевіряємо, чи інтеракція ще дійсна
      if (error instanceof Error && 
          (error.message.includes('Unknown interaction') || 
           error.message.includes('Interaction has already been acknowledged'))) {
        // Інтеракція прострочена або вже оброблена
        logger.warn('ℹ️ Інтеракція прострочена або вже оброблена, пропускаємо відповідь');
        return;
      }

      try {
        if (interaction.deferred) {
          await interaction.editReply({ content: errorMessage });
        } else if (!interaction.replied) {
          await interaction.reply({ content: errorMessage, ephemeral: true });
        }
      } catch (replyError) {
        // Помилка при відповіді - логуємо і продовжуємо
        logger.warn('ℹ️ Не вдалося відповісти на інтеракцію', {
          error: replyError instanceof Error ? replyError.message : String(replyError)
        });
      }
    }
  }

  private buildResponse(result: AIQueryResult, commandOptions: AICommandOptions): string {
    // Формуємо більш структуровану відповідь з кращим форматуванням
    let response = `🤖 **AI Асистент - Повна Відповідь**\n\n`;

    // Додаємо індикатор впевненості з кращим форматуванням
    const confidencePercent = Math.round(result.confidence * 100);
    if (result.confidence < 0.7) {
      response += `⚠️ **Низька впевненість** (${confidencePercent}%)\n\n`;
    } else if (result.confidence < 0.9) {
      response += `✅ **Середня впевненість** (${confidencePercent}%)\n\n`;
    } else {
      response += `🟢 **Висока впевненість** (${confidencePercent}%)\n\n`;
    }

    // Додаємо оригінальне запитання
    response += `**🔍 Запит:**\n${String(commandOptions.query)}\n\n`;
    
    // Додаємо відповідь з кращим форматуванням та повною інформацією
    response += `**💬 Повна Відповідь:**\n${result.response}\n\n`;

    // Додаємо контекст, якщо він є
    if (commandOptions.context) {
      response += `**📎 Контекст:**\n${String(commandOptions.context)}\n\n`;
    }

    // Додаємо інформацію про дію з кращим форматуванням
    if (result.action) {
      response += `**🔧 Дія:** ${result.action}\n`;
    }

    if (result.actionData) {
      response += `**📄 Тип даних:** ${result.actionData.type}\n`;
      if (result.actionData.format) {
        response += `**📊 Формат:** ${result.actionData.format}\n`;
      }
    }

    // Додаємо рекомендації щодо подальших дій з кращим форматуванням
    response += `\n---\n`;
    response += `💡 **Рекомендації:**\n`;
    response += `• Якщо відповідь не повна, уточніть запит\n`;
    response += `• Використовуйте контекст для посилання на попередні запити\n`;
    response += `• Для складних запитів розбийте їх на кілька простих\n`;
    response += `• Якщо потрібна інформація з конкретного документа, вкажіть його назву\n\n`;
    
    // Додаємо інформацію про джерела, якщо вони є
    if (result.response.includes('Джерело:') || result.response.includes('джерело')) {
      response += `📚 **Джерела:**\n`;
      response += `Відповідь базується на інформації з Google Диску та інших джерел.\n\n`;
    }

    // Обмежуємо довжину відповіді до 2000 символів, щоб уникнути помилок Discord
    if (response.length > 1900) {
      response = response.substring(0, 1900) + '\n\n... [Відповідь обрізана через обмеження Discord]\n\n';
    }

    return response;
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermission(interaction: ChatInputCommandInteraction): Promise<boolean> {
    try {
      const { PermissionManager } = await import('../core/PermissionManager');
      const permissionManager = new PermissionManager(this.config);
      const member: GuildMember | null = interaction.member instanceof (await import('discord.js')).GuildMember
        ? (interaction.member)
        : null;
      const result = await permissionManager.checkPermission(
        interaction.user,
        member,
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

        if (interaction.deferred) {
          await interaction.editReply({ content: t('ai.error.accessDenied', { reason: result.reason || '' }) });
        } else if (!interaction.replied) {
          await interaction.reply({ content: t('ai.error.accessDenied', { reason: result.reason || '' }), ephemeral: true });
        }
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
      if (interaction.deferred) {
        await interaction.editReply({ content: t('ai.error.permissionCheck') });
      } else if (!interaction.replied) {
        await interaction.reply({ content: t('ai.error.permissionCheck'), ephemeral: true });
      }
      return false;
    }
  }

  /**
   * Валідація параметрів команди
   */
  private validateCommandOptions(options: AICommandOptions): ValidationResult {
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
      const value = (options as Record<string, unknown>)[key];

      if (schema.required && !value) {
        errors.push(`${key} є обов'язковим`);
        continue;
      }

      if (value && typeof value !== schema.type) {
        errors.push(`${key} має бути типу ${schema.type}`);
        continue;
      }

      if (value && schema.maxLength && String(value).length > schema.maxLength) {
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
  private logSecurityEvent(eventType: string, data: Record<string, unknown>): void {
    logger.security(eventType, (data['userId'] as string) || 'unknown', {
      type: 'security',
      component: 'AIAssistantCommand',
      command: this.name,
      ...data,
    });
  }

  /**
   * Обробка AI запиту: визначення наміру та виконання дій
   */
  private async processAIQuery(
    userId: string,
    query: string,
    _context?: string,
    interaction?: ChatInputCommandInteraction
  ): Promise<AIQueryResult> {
    // Очищуємо старі сесії перед обробкою
    AIAssistantCommand.cleanupOldSessions();
    
    // Отримуємо контекст користувача
    const contextText = AIAssistantCommand.getContextForPrompt(userId, query);
    
    const handlers = this.getQueryHandlers(query);
    const normalized = this.normalizeQuery(query);
    const handled = await this.runHandlers(handlers, normalized);
    
    let result: AIQueryResult;
    if (handled) {
      result = handled;
    } else {
      result = await this.buildDefaultAIQueryResult(query, interaction, contextText);
    }
    
    // Зберігаємо контекст після обробки
    AIAssistantCommand.addContext(
      userId,
      query,
      result.response,
      result.action,
      result.actionData?.type === 'files' ? (result.actionData as any).fileIds : undefined
    );
    
    return result;
  }

  /**
   * Повертає впорядкований список обробників, які намагаються розпізнати намір запиту
   */
  private getQueryHandlers(query: string): ReadonlyArray<(q: string) => Promise<AIQueryResult | null>> {
    return [
      (q) => this.tryOcrImage(query, q),
      (q) => this.tryTableAnalytics(query, q),
      (q) => this.tryExtractText(query, q),
      (q) => this.tryAnalyzeBuses(query, q),
      (q) => this.tryListSheets(query, q),
    ] as const;
  }

  /**
   * Нормалізує вхідний текст запиту для порівнянь
   */
  private normalizeQuery(query: string): string {
    return (query || '').toLowerCase();
  }

  /**
   * Послідовно виконує обробники та повертає перший успішний результат
   */
  private async runHandlers(
    handlers: ReadonlyArray<(q: string) => Promise<AIQueryResult | null>>,
    normalizedQuery: string
  ): Promise<AIQueryResult | null> {
    for (const handle of handlers) {
      const res = await handle(normalizedQuery);
      if (res) return res;
    }
    return null;
  }

  /**
   * Базова відповідь за замовчуванням, якщо намір не розпізнано
   */
  private async buildDefaultAIQueryResult(query: string, interaction?: ChatInputCommandInteraction, contextText?: string): Promise<AIQueryResult> {
    try {
      // Спробуємо використати AI-сервіс, якщо він є
      const aiService = (interaction?.client as any)?.serviceContainer?.get?.('ai');
      if (aiService && typeof aiService.processNaturalLanguageQuery === 'function') {
        const userId = interaction?.user?.id || 'unknown';
        const aiResponse = await aiService.processNaturalLanguageQuery(userId, query, {
          source: 'discord_command',
          timestamp: Date.now(),
        });

        return {
          response: aiResponse.content,
          confidence: 0.9,
          action: 'ai_response',
          actionData: { type: 'ai_response', format: 'text' },
        };
      }

      // Якщо AI-сервіс недоступний — використовуємо Ollama напряму з покращеним промптом
      logger.info('Using Ollama fallback for AI query', { query: query.substring(0, 100) });

      const ollamaResponse = await fetch('http://localhost:11434/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'llama3.2', // ← зміни на свою модель, якщо потрібно
          prompt: `📚 Ти — офіційний AI-асистент "GodzillaBot" для військових, держслужбовців та адміністративних працівників України. Ти маєш високий рівень професійної ерудиції, відмінно володієш українською мовою та дотримуєшся офіційно-ділового стилю.

🎯 ГОЛОВНІ ПРИНЦИПИ РОБОТИ:

1. 🔄 **ПІДТРИМКА ДІАЛОГУ**: Запам'ятовуй контекст, посилайся на попередні відповіді.
2. 💾 **КЕШУВАННЯ**: Посилайся на відомі дані з попередніх запитів.
3. 📁 **RAG З GOOGLE ДИСКУ**: Перевіряй наявність документів, посилайся на них з ID і назвою.
4. 🇺🇦 **МОВА**: Лише чиста українська, офіційно-діловий стиль.
5. 🧠 **ПРОФЕСІЙНІСТЬ**: Точність, структура, конкретні посилання.

📌 **ФОРМАТ ВІДПОВІДІ:**
- Повна відповідь (без скорочень)
- Джерело (якщо є)
- Практичні рекомендації
- Конкретні приклади (якщо можливо)

🔥 ТИ — ЕКСПЕРТ, НЕ ПОМИЛЯЄШСЯ У ТЕРМІНАХ, ДОТРИМУЄШСЯ УКРАЇНСЬКОЇ МОВИ.${contextText || ''}

🔥 ЗАПИТАННЯ: ${query}

💬 ВІДПОВІДЬ УКРАЇНСЬКОЮ (ПОВНА, ДЕТАЛЬНА, БЕЗ СКОРОЧЕНЬ):`,  
          stream: false,
        }),
      });

      if (!ollamaResponse.ok) {
        throw new Error(`Ollama error: ${await ollamaResponse.text()}`);
      }

      const data = await ollamaResponse.json() as { response?: string; };

      // Покращення відповіді - додавання структури
      let enhancedResponse = data.response || 'Отримано відповідь, але вона порожня.';
      
      // Якщо відповідь занадто коротка, намагаємось отримати більш детальну
      if (enhancedResponse.length < 300) {
        const detailedResponse = await fetch('http://localhost:11434/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            model: 'llama3.2',
            prompt: `Розшир детальну відповідь на запитання: ${query}
            
            Попередня відповідь: ${enhancedResponse}
            
            Надай більш повну та детальну інформацію з конкретними прикладами та рекомендаціями.
            
            Форматуй відповідь як:
            1. Основна інформація
            2. Джерела (якщо є)
            3. Практичні рекомендації
            4. Приклади (якщо можливо)`,
            stream: false,
          }),
        });
        
        if (detailedResponse.ok) {
          const detailedData = await detailedResponse.json() as { response?: string; };
          if (detailedData.response && detailedData.response.length > enhancedResponse.length) {
            enhancedResponse = detailedData.response;
          }
        }
      }

      // Додаткове покращення - структурування відповіді
      const structuredResponse = await fetch('http://localhost:11434/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'llama3.2',
          prompt: `Структуруй наступну відповідь у форматі:
          
          🔍 Основна інформація:
          [основна відповідь з деталями]
          
          📚 Джерело:
          [джерело, якщо є]
          
          💡 Рекомендації:
          [практичні рекомендації]
          
          📋 Приклади:
          [конкретні приклади, якщо можливо]
          
          Відповідь для структурування: ${enhancedResponse}`,
          stream: false,
        }),
      });
      
      if (structuredResponse.ok) {
        const structuredData = await structuredResponse.json() as { response?: string; };
        if (structuredData.response) {
          enhancedResponse = structuredData.response;
        }
      }

      return {
        response: enhancedResponse,
        confidence: 0.85,
        action: 'ollama_fallback',
        actionData: { type: 'ai_response', format: 'text' },
      };
    } catch (error) {
      logger.warn('Failed to get response from Ollama or AI service', {
        error: error instanceof Error ? error.message : String(error),
        query: query.substring(0, 50),
      });

      // Остаточний fallback
      return {
        response: `❌ Не вдалося отримати відповідь від AI. Ваш запит: "${query}"

Покращена відповідь:
Для отримання повної відповіді на ваше запитання, рекомендую:
1. Уточнити формулювання запиту
2. Додати контекст або конкретні деталі
3. Перевірити доступність AI-сервісу

Якщо проблема повторюється, зверніться до адміністратора системи.`,
        confidence: 0.5,
        action: 'error_fallback',
        actionData: { type: 'text', format: 'text' },
      };
    }
  }

  private async tryOcrImage(query: string, q: string): Promise<AIQueryResult | null> {
    const intent = /(картин|изображен|image|photo|png|jpg|jpeg)/i.test(q) && /(ocr|текст|прочитай|извле(ки|чи))/i.test(q);
    if (!intent) return null;
    if (!this.googleService) {
      return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
    }
    const folderId = this.config.google.driveFolderId;
    if (!folderId) {
      return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
    }
    const nameQuery = tokenizeName(query, 5);
    const index = await ensureDriveIndex(this.googleService, folderId);
    const qlc = (s: string) => s.toLowerCase();
    const matchesName = (name?: string) => !nameQuery || qlc(name || '').includes(qlc(nameQuery));
    const candidates = (index || []).filter((f: unknown) => {
      const file = f as DriveIndexedFile;
      return isImageMime(file.mimeType) && matchesName(file.name);
    }) as DriveIndexedFile[];
    if (!candidates.length) {
      return { response: t('ai.ocr.noImages'), confidence: 0.85, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
    }
    for (const f of candidates.slice(0, 5)) {
      try {
        const text = await this.googleService.extractTextFromImage(f);
        if (!text.trim()) continue;
        const preview = text.length > 1500 ? text.slice(0, 1500) + '…' : text;
        return { response: t('ai.ocr.result', { name: String(f.name ?? ''), id: String(f.id ?? ''), preview }), confidence: 0.9, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
      } catch (e) {
        logger.warn('OCR error', { type: 'command', component: 'AIAssistantCommand.tryOcrImage', fileId: f.id, err: e instanceof Error ? e.message : String(e) });
      }
    }
    return { response: t('ai.ocr.cannotRead'), confidence: 0.7, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
  }

  private async tryTableAnalytics(query: string, q: string): Promise<AIQueryResult | null> {
    const intent = /(группируй|сгруппируй|групу(ва|пу)й|посчитай|підрахуй)/i.test(q) && /(статус|status)/i.test(q);
    if (!intent) return null;
    if (!this.googleService) {
      return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
    }
    const folderId = this.config.google.driveFolderId;
    if (!folderId) {
      return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
    }
    const nameQuery = tokenizeName(query, 5);
    const monthNum = findMonthNumber(q);
    const files = await this.googleService.listDriveFilesInFolder(folderId, { recursive: true, type: 'any', limit: 100, maxDepth: -1, ...(nameQuery ? { query: nameQuery } : {}) });
    const tableLike = files.filter(f => {
      const mt = (f.mimeType || '');
      return mt === 'application/vnd.google-apps.spreadsheet' || mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mt === 'application/vnd.ms-excel';
    });
    if (!tableLike.length) {
      return { response: t('ai.analytics.noTables'), confidence: 0.85, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
    }
    const analytics = new AnalyticsService();
    for (const f of tableLike.slice(0, 5)) {
      try {
        const mt = f.mimeType || '';
        let rows: Array<Record<string, unknown>> = [];
        if (mt === 'application/vnd.google-apps.spreadsheet') rows = await readGoogleSheet(this.googleService, f.id!);
        else rows = readExcelBuffer(await this.googleService.downloadDriveFile(f.id!));
        if (!rows.length) continue;
        const schema = analytics.inferSchema(rows);
        const statusKey = schema.find(k => /статус|status/i.test(k)) || schema[0];
        const dateKey = schema.find(k => /дата|date/i.test(k));
        let filtered = rows;
        if (monthNum && dateKey) {
          const toDate = (v: unknown): Date | null => {
            if (v instanceof Date && !isNaN(+v)) return v;
            if (typeof v === 'string' || typeof v === 'number') {
              const d = new Date(v);
              return isNaN(+d) ? null : d;
            }
            return null;
          };
          filtered = rows.filter(r => {
            const v = (r)[dateKey];
            const d = toDate(v);
            return d instanceof Date && !isNaN(+d) && d.getMonth() + 1 === monthNum;
          });
        }
        if (!statusKey) continue;
        const groups = analytics.groupBy(filtered, [statusKey]);
        const lines: string[] = [];
        for (const [gk, arr] of Object.entries(groups)) {
          const cnt = (arr as unknown[]).length;
          lines.push(`${gk || '—'}: ${cnt}`);
        }
        const head = `Файл: ${f.name} (id: ${f.id})`;
        return { response: head + '\n' + lines.join('\n'), confidence: 0.9, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
      } catch (e) {
        logger.warn('Analytics failed for file', { type: 'command', component: 'AIAssistantCommand.tryTableAnalytics', fileId: f.id, err: e instanceof Error ? e.message : String(e) });
      }
    }
    return { response: 'Не вдалося виконати аналітику: дані порожні або структура невідома.', confidence: 0.7, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
  }

  private async tryExtractText(query: string, q: string): Promise<AIQueryResult | null> {
    const intent = /(pdf|word|docx|docs?|документ|файл)/i.test(q) && /(покажи|выведи|витягни|извле(ки|чи)|текст)/i.test(q);
    if (!intent) return null;
    if (!this.googleService) {
      return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
    }
    const folderId = this.config.google.driveFolderId;
    if (!folderId) {
      return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
    }
    const nameQuery = tokenizeName(query, 5);
    const index = await ensureDriveIndex(this.googleService, folderId);
    const qlc = (s: string) => s.toLowerCase();
    const matchesName = (name?: string) => !nameQuery || qlc(name || '').includes(qlc(nameQuery));
    const candidates = (index || []).filter((f: unknown) => {
      const file = f as DriveIndexedFile;
      return isDocLikeMime(file.mimeType) && matchesName(file.name);
    }) as DriveIndexedFile[];
    if (!candidates.length) {
      return { response: 'Не знайдено відповідних документів (Docs/Word/PDF) за вашим описом.', confidence: 0.85, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
    }
    for (const f of candidates.slice(0, 5)) {
      try {
        const text = await this.googleService.extractTextFromFile(f);
        if (!text.trim()) continue;
        const preview = text.length > 1500 ? text.slice(0, 1500) + '…' : text;
        return { response: `Файл: ${String(f.name)} (id: ${String(f.id)})\n\n${preview}`, confidence: 0.9, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
      } catch (e) {
        logger.warn('Не вдалося витягти текст з документу', { type: 'command', component: 'AIAssistantCommand.tryExtractText', fileId: f.id, fileName: f.name, error: e instanceof Error ? e.message : String(e) });
      }
    }
    return { response: 'Не вдалося витягти текст: документи порожні або формат не підтримується. Уточніть назву файла або надішліть приклад.', confidence: 0.7, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
  }

  private async tryAnalyzeBuses(query: string, q: string): Promise<AIQueryResult | null> {
    const intent = /автобус|bus/.test(q) && /(сколько|скiльки|скільки|осталось|залишил(о|ось)|бг|остат)/.test(q);
    if (!intent) return null;
    if (!this.googleService) {
      return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
    }
    const nameQuery = tokenizeName(query, 4);
    const folderId = this.config.google.driveFolderId;
    if (!folderId) {
      return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
    }
    const baseOpts: { recursive?: boolean; type?: 'sheet' | 'folder' | 'any'; query?: string; limit?: number; pageToken?: string; maxDepth?: number } = { recursive: true, type: 'any', limit: 100, maxDepth: -1 };
    if (nameQuery) baseOpts.query = nameQuery;
    const files = await this.googleService.listDriveFilesInFolder(folderId, baseOpts);
    const candidates = files.filter(f => {
      const mt = f.mimeType || '';
      return mt === 'application/vnd.google-apps.spreadsheet' || mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mt === 'application/vnd.ms-excel';
    });
    if (!candidates.length) {
      return { response: 'Не знайдено придатних таблиць (Google Sheets/Excel) за вашим описом.', confidence: 0.85, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
    }
    const countBuses = (rows: Array<Record<string, any>>): number => {
      if (!rows.length) return 0;
      const norm = (s: unknown) => String(s ?? '').toLowerCase();
      const keys = Object.keys(rows[0] || {});
      const typeKey = keys.find(k => /(тип|вид|категор|vehicle|type|category)/i.test(k)) || (keys[0] as string | undefined);
      const statusKey = keys.find(k => /(статус|state|status)/i.test(k)) || ((keys[1] as string | undefined) ?? (keys[0] as string | undefined));
      if (!typeKey || !statusKey) return 0;
      let count = 0;
      for (const r of rows) {
        if (/автобус|bus/i.test(norm(r[typeKey])) && /(бг|остат|остал|залиш|в наличии|на складе)/i.test(norm(r[statusKey]))) count++;
      }
      return count;
    };
    for (const f of candidates.slice(0, 5)) {
      try {
        const mt = f.mimeType || '';
        let rows: Array<Record<string, any>> = [];
        if (mt === 'application/vnd.google-apps.spreadsheet') rows = await readGoogleSheet(this.googleService, f.id!, 'A1:Z1000');
        else rows = readExcelBuffer(await this.googleService.downloadDriveFile(f.id!)) as Array<Record<string, any>>;
        if (!rows.length) continue;
        const total = countBuses(rows);
        return { response: `Файл: ${f.name} (id: ${f.id})\nРезультат: автобусів зі статусом БГ/залишок — ${total} шт.`, confidence: 0.92, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
      } catch (e) {
        logger.warn('Не вдалося обробити файл-кандидат', { type: 'command', component: 'AIAssistantCommand.tryAnalyzeBuses', fileId: (f as any).id, fileName: (f as any).name, error: e instanceof Error ? e.message : String(e) });
      }
    }
    return { response: 'Не вдалося виконати аналіз: таблиці порожні або структура невідома.', confidence: 0.7, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
  }

  private async tryListSheets(_query: string, q: string): Promise<AIQueryResult | null> {
    const intent = /таблиц|таблицы|лист(ы|и)?|sheets?|список.*таблиц|какие.*таблиц|google\s*диск|google\s*sheets/.test(q) && /какие|покажи|список|list|что|найд/i.test(q);
    if (!intent) return null;
    try {
      if (!this.googleService) {
        return { response: 'GoogleService не доступний для цієї команди. Перевірте ініціалізацію сервісів або конфігурацію.', confidence: 0.6, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
      }
      const folderId = this.config.google.driveFolderId;
      if (!folderId) {
        return { response: 'Не налаштовано google.driveFolderId у конфігурації. Додайте ID каталогу з таблицями.', confidence: 0.6, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
      }
      const files = await this.googleService.listDriveFilesInFolder(folderId, { recursive: true, type: 'sheet', limit: 50, maxDepth: -1 });
      if (!files.length) {
        const spreadsheetId = this.config.google.spreadsheetId;
        if (spreadsheetId) {
          try {
            const sheetTitles = await this.googleService.listSheets(spreadsheetId);
            if (sheetTitles && sheetTitles.length >= 0) {
              return { response: 'Таблиці не знайдені у вказаній папці Google Drive. Проте доступ до таблиці з конфігурації працює. Переконайтесь, що потрібні файли знаходяться у цій папці або змініть GOOGLE_DRIVE_FOLDER_ID на правильну папку.', confidence: 0.92, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
            }
          } catch {
            return { response: 'Не вдалось отримати доступ до таблиць: папка порожня або недоступна, а також немає доступу до таблиці з конфігурації. Перевірте, що ви надали доступ сервісному акаунту та що файли знаходяться у вказаній папці.', confidence: 0.7, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
          }
        }
        return { response: 'Таблиці не знайдені у вказаній папці Google Drive.', confidence: 0.9, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
      }
      const lines = files.slice(0, 20).map((f, idx) => {
        const mime = f.mimeType || '';
        const label = mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ? 'Excel (.xlsx)' : mime === 'application/vnd.ms-excel' ? 'Excel (.xls)' : 'Google Sheets';
        return `${idx + 1}. ${f.name} [${label}] (id: ${f.id})`;
      });
      const more = files.length > 20 ? `\n… та ще ${files.length - 20}` : '';
      return { response: `Знайдено ${files.length} таблиць/Excel-файлів:\n\n${lines.join('\n')}${more}`, confidence: 0.95, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
    } catch (error) {
      logger.error('Помилка отримання списку таблиць', { type: 'command', component: 'AIAssistantCommand.tryListSheets', error: error instanceof Error ? error.message : String(error) });
      return { response: 'Сталася помилка при отриманні списку таблиць Google. Спробуйте пізніше або перевірте доступи.', confidence: 0.4, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
    }
  }
}
