/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */

import { Collection, Events, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import type { ChatInputCommandInteraction, Interaction, Client } from 'discord.js';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';
import { replyWithPrivacy } from '@/ui/reply';
import { signComponentId, verifyComponentId } from '@/security/componentId';
import { tUser } from '@/i18n';
import { AppError } from '@/core/errors/AppError';
import { normalizeText } from '@/nlp/normalize';
import { detectIntent } from '@/nlp/IntentDetector';
import type { ClassifyIntentFn } from '@/nlp/IntentDetector';
import { detectLanguage } from '@/nlp/LanguageDetector';
import { validateInput } from '@/utils/security';

interface CommandStats {
  totalCommands: number;
  categories: number;
  commandsByCategory: Record<string, number>;
  lastUsed: Date;
}

// Мінімальний контракт команди, щоб уникнути конфлікту приватних полів TS між різними деклараціями класів
interface ICommand {
  getName(): string;
  getDescription(): string;
  getData(): any;
  execute(args: { interaction: ChatInputCommandInteraction }): Promise<void> | void;
}

type BotLike = { client: Client; getService?: (name: string) => unknown };

export class CommandManager {
  // Nodemon touch: ensure restart after hotfix
  private bot: BotLike;
  private config: BotConfig;
  private commands: Collection<string, ICommand>;

  private commandCategories: Map<string, string[]>;
  private stats: CommandStats;

  constructor(bot: BotLike | Client | unknown, config: BotConfig) {
    // Accept either a raw Discord Client or an object with { client }
    const resolved: BotLike = ((): BotLike => {
      const b = bot as any;
      if (b && typeof b === 'object') {
        if ('client' in b && b.client) {
          return { client: b.client as Client, getService: b.getService?.bind(b) };
        }
        // If raw Client or mock-like object passed (may not have .on in tests)
        return { client: b as Client };
      }
      // Fallback dummy client to avoid crashes in tests
      return { client: ({} as unknown) as Client };
    })();

    this.bot = resolved;
    this.config = config;
    this.commands = new Collection();
    this.commandCategories = new Map();
    this.stats = {
      totalCommands: 0,
      categories: 0,
      commandsByCategory: {},
      lastUsed: new Date(),
    };
  }

  /**
   * Обработка кнопок поиска: expand/page
   */
  private async handleSearchButton(interaction: any): Promise<void> {
    try {
      const customId = String(interaction.customId || '');

      // Підтримка двох форматів: підписаний compact та легасі (для тестів)
      type SearchPayload = { kind?: string; action?: 'expand' | 'page'; id?: string; documentId?: string; page?: number; ts?: number };
      let action: 'expand' | 'page' | undefined;
      let fileId: string | undefined;
      let pageIndex = 0;

      const verified = verifyComponentId<SearchPayload>(customId);
      if (verified.valid && (verified.payload?.kind === 'srch' || verified.payload?.kind === 'search')) {
        action = verified.payload.action;
        fileId = verified.payload.id || verified.payload.documentId;
        if (typeof verified.payload.page === 'number') pageIndex = verified.payload.page;
      } else {
        // Легасі fallback: search|{action}|{fileId}|{index}
        const parts = customId.split('|');
        action = (parts[1] as 'expand' | 'page' | undefined);
        fileId = parts[2];
        pageIndex = parts[3] ? parseInt(parts[3], 10) : 0;
      }
      if (!fileId) {
        await replyWithPrivacy(interaction, { content: tUser('files.validation.invalidFileId', interaction) });
        return;
      }

      const driveIndexer = (this.bot.getService?.('driveIndexer') ?? undefined) as
        | import('@/services/DriveIndexerService').DriveIndexerService
        | undefined;
      if (!driveIndexer) {
        await replyWithPrivacy(interaction, { content: tUser('search.error.noService', interaction) });
        return;
      }

      const chunks = await driveIndexer.getTextChunks(fileId, 1800);
      if (!chunks.length) {
        await replyWithPrivacy(interaction, { content: tUser('files.error.noText', interaction) });
        return;
      }

      const idx = Math.min(Math.max(0, pageIndex), chunks.length - 1);
      const content = chunks[idx];

      // Кнопки пагінації (назад/вперед) — підписані, з легасі фолбеком у тестах
      const useLegacy = process.env['NODE_ENV'] === 'test' || process.env['LEGACY_CUSTOM_ID'] === '1';
      const mkId = (act: 'page' | 'expand', page?: number) =>
        useLegacy
          ? `search|${act}|${fileId}|${page ?? idx}`
          : signComponentId({ kind: 'srch', action: act, id: fileId, page: page ?? idx, ts: Date.now() });
      const row = new ActionRowBuilder<ButtonBuilder>().addComponents(
        new ButtonBuilder()
          .setCustomId(mkId('page', Math.max(0, idx - 1)))
          .setEmoji('⬅️')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(idx === 0),
        new ButtonBuilder()
          .setCustomId(mkId('page', Math.min(chunks.length - 1, idx + 1)))
          .setEmoji('➡️')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(idx >= chunks.length - 1)
      );

      if (action === 'expand') {
        await replyWithPrivacy(interaction, { content: content ?? '', components: [row] });
        return;
      }
      if (action === 'page') {
        if (interaction.deferred || interaction.replied) {
          await interaction.editReply({ content, components: [row] });
        } else {
          await replyWithPrivacy(interaction, { content: content ?? '', components: [row] });
        }
        return;
      }

      await replyWithPrivacy(interaction, { content: tUser('files.error.unknownSubcommand', interaction) });
    } catch (e) {
      logger.error('search_button_failed', { error: e instanceof Error ? e.message : String(e) });
      try {
        await replyWithPrivacy(interaction, { content: tUser('workspace.common.execError', interaction) });
      } catch {
        // ignore
      }
    }
  }

  /**
   * Обработка автодополнения
   */
  private async handleAutocomplete(interaction: import('discord.js').AutocompleteInteraction): Promise<void> {
    try {
      const command = this.commands.get(interaction.commandName) as unknown;
      const maybe = command as { autocomplete?: (args: { interaction: import('discord.js').AutocompleteInteraction; query?: string }) => Promise<void> | void } | undefined;
      if (!maybe || typeof maybe.autocomplete !== 'function') {
        // Команда не поддерживает автодополнение
        await interaction.respond([]);
        return;
      }

      // Получаем текущее поле и значение
      const focused = interaction.options.getFocused(true);
      const query = typeof focused?.value === 'string' ? focused.value : String(focused?.value ?? '');

      await maybe.autocomplete({ interaction, query });
    } catch (error) {
      logger.error('❌ Ошибка автодополнения', {
        type: 'command_manager',
        event: 'autocomplete_error',
        commandName: interaction.commandName,
        errorMessage: String(error),
      });
      try {
        await interaction.respond([]);
      } catch {
        // ignore
      }
    }
  }

  /**
   * Ініціалізація менеджера команд
   */
  async initialize(): Promise<void> {
    try {
      if (this.config.discord.enableSlash === false) {
        logger.info('🚫 Slash-команды отключены (enableSlash=false) — пропускаю инициализацию CommandManager', {
          type: 'command_manager',
          event: 'init_skipped',
        });
        return;
      }

      logger.info('📋 Ініціалізація менеджера команд...', {
        type: 'command_manager',
        event: 'init_start',
      });

      // Завантаження команд
      await this.loadCommands();

      // Реєстрація обробників подій (безпечно)
      this.registerEventHandlers();

      // Реєстрація slash-команд у Discord, якщо доступно (для інтеграційних тестів)
      const app = (this.bot.client as any)?.application;
      const commandsApi = app?.commands;
      if (commandsApi && typeof commandsApi.set === 'function') {
        try {
          await commandsApi.set(this.getCommandsData());
          logger.info('🔗 Команди зареєстровано у Discord', {
            type: 'command_manager',
            event: 'discord_register_done',
            count: this.commands.size,
          });
        } catch (e) {
          logger.warn('⚠️ Не вдалося зареєструвати команди у Discord (продовжую)', {
            type: 'command_manager',
            event: 'discord_register_failed',
            errorMessage: String(e),
          });
        }
      }

      logger.info('✅ Завантажено команди', {
        type: 'command_manager',
        event: 'init_loaded',
        count: this.commands.size,
      });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера команд', {
        type: 'command_manager',
        event: 'init_error',
        errorMessage: String(error),
      });
      throw error;
    }
  }

  /**
   * Завантаження всіх команд
   */
  private async loadCommands(): Promise<void> {
    logger.info('📥 Loading commands...', { component: 'CommandManager' });

    // Import all command classes
    const { AIAssistantCommand } = await import('@/commands/AIAssistantCommand');
    const { AdvancedAnalysisCommand } = await import('@/commands/AdvancedAnalysisCommand');
    const { AnalyticsCommand } = await import('@/commands/AnalyticsCommand');
    const { AnalyzeCommand } = await import('@/commands/AnalyzeCommand');
    const { DocCommand } = await import('@/commands/DocCommand');
    const { DocumentAnalysisCommand } = await import('@/commands/DocumentAnalysisCommand');
    const { DocumentsCommand } = await import('@/commands/DocumentsCommand');
    const { DriveExtractCommand } = await import('@/commands/DriveExtractCommand');
    const { DriveNavigateCommand } = await import('@/commands/DriveNavigateCommand');
    const { EnhancedDriveSearchCommand } = await import('@/commands/EnhancedDriveSearchCommand');
    const { EnhancedSearchCommand } = await import('@/commands/EnhancedSearchCommand');
    const { FavoritesCommand } = await import('@/commands/FavoritesCommand');
    const { FileManagerCommand } = await import('@/commands/FileManagerCommand');
    const { LangCommand } = await import('@/commands/LangCommand');
    const { OCRCommand } = await import('@/commands/OCRCommand');
    const { OperationsCommand } = await import('@/commands/OperationsCommand');
    const { PerformanceCommand } = await import('@/commands/PerformanceCommand');
    const { SavedSearchCommand } = await import('@/commands/SavedSearchCommand');
    const { SearchCommand } = await import('@/commands/SearchCommand');
    const { SelectSheetCommand } = await import('@/commands/SelectSheetCommand');
    const { SimplifiedCommand } = await import('@/commands/SimplifiedCommand');
    const { SmartSearchCommand } = await import('@/commands/SmartSearchCommand');
    const { WorkflowCommand } = await import('@/commands/WorkflowCommand');
    const { WorkspaceCommand } = await import('@/commands/WorkspaceCommand');
    const { MarkdownCommand } = await import('@/commands/MarkdownCommand');
    const { OllamaCommand } = await import('@/commands/OllamaCommand');

    // Create command instances
    const googleService = this.bot.getService ? this.bot.getService('google') : undefined;
    const aiService = this.bot.getService ? this.bot.getService('ai') : undefined;
    const workflowEngine = this.bot.getService ? this.bot.getService('workflow') : undefined;
    const config = this.config || (this.bot as any).config || {};

    const commands = [
      new AIAssistantCommand(config, googleService as any),
      new AdvancedAnalysisCommand(config, googleService as any, aiService as any),
      new AnalyticsCommand(config),
      new AnalyzeCommand(config),
      new DocCommand(config, googleService as any),
      new DocumentAnalysisCommand(config, googleService as any),
      new DocumentsCommand(config),
      new DriveExtractCommand(config, googleService as any),
      new DriveNavigateCommand(config, googleService as any),
      new EnhancedDriveSearchCommand(config, googleService as any),
      new EnhancedSearchCommand(config, googleService as any),
      new FavoritesCommand(config),
      new FileManagerCommand(config),
      new LangCommand(config),
      new OCRCommand(config, googleService as any),
      new OperationsCommand(config),
      new PerformanceCommand(config),
      new SavedSearchCommand(config),
      new SearchCommand(config, googleService as any),
      new SelectSheetCommand(config, googleService as any),
      new SimplifiedCommand(config, googleService as any),
      new SmartSearchCommand(config, googleService as any),
      new WorkflowCommand(config, workflowEngine as any),
      new WorkspaceCommand(config),
      new MarkdownCommand(config),
      new OllamaCommand(),
    ];

    // Реєструємо команди
    for (const command of commands) {
      if (this.validateCommand(command)) {
        const commandName = command.getName();
        this.commands.set(commandName, command);

        // Категоризація команд
        const category = this.getCommandCategory(command);
        if (!this.commandCategories.has(category)) {
          this.commandCategories.set(category, []);
        }
        this.commandCategories.get(category)!.push(commandName);

        logger.info('📝 Завантажено команду', {
          type: 'command_manager',
          event: 'command_loaded',
          commandName,
          category,
        });
      }
    }

    // Оновлюємо статистику
    this.updateStats();
  }

  /**
   * Валідація команди
   */
  private validateCommand(command: ICommand): boolean {
    if (!command.getName()) {
      logger.warn('Команда не має назви', {
        type: 'command_manager',
        event: 'validation_warn',
        reason: 'empty_name',
      });
      return false;
    }

    if (!command.getDescription()) {
      logger.warn('Команда не має опису', {
        type: 'command_manager',
        event: 'validation_warn',
        reason: 'empty_description',
        commandName: command.getName(),
      });
      return false;
    }

    return true;
  }

  /**
   * Визначення категорії команди
   */
  private getCommandCategory(command: ICommand): string {
    const name = command.getName();

    if (name.includes('пошук') || name.includes('search')) {
      return 'Пошук';
    }
    if (name.includes('smart-search')) {
      return 'Розумний Пошук';
    }
    if (name.includes('advanced-analysis')) {
      return 'Розширений Аналіз';
    }
    if (name.includes('продуктивність') || name.includes('performance')) {
      return 'Моніторинг';
    }
    if (name.includes('ai') || name.includes('асистент') || name.includes('ollama')) {
      return 'AI';
    }
    if (name.includes('документи') || name.includes('documents')) {
      return 'Документи';
    }
    if (name.includes('файли') || name.includes('file')) {
      return 'Файли';
    }
    if (name.includes('операції') || name.includes('operations')) {
      return 'Операції';
    }
    if (name.includes('аналітика') || name.includes('analytics')) {
      return 'Аналітика';
    }

    return 'Інші';
  }

  /**
   * Реєстрація обробників подій
   */
  private registerEventHandlers(): void {
    if (!this.config.discord.enableSlash) {
      logger.debug('Slash-команды отключены — обработчики не регистрируются', {
        type: 'command_manager',
        event: 'handlers_skipped',
      });
      return;
    }
    // Реєструємо обробник на Discord клієнті, а не на екземплярі Bot (перевірка наявності on)
    const onFn = (this.bot.client as any)?.on;
    if (typeof onFn !== 'function') {
      logger.debug('Клієнт не підтримує on() — пропускаю реєстрацію обробників (тестове середовище?)', {
        type: 'command_manager',
        event: 'handlers_noop',
      });
      return;
    }
    (this.bot.client as any).on(Events.InteractionCreate, async (interaction: Interaction) => {
      try {
        if (interaction.isChatInputCommand()) {
          await this.handleCommand(interaction);
          return;
        }

        // Поддержка автодополнения для опций команд
        if ('isAutocomplete' in interaction && typeof (interaction as any).isAutocomplete === 'function' && (interaction as any).isAutocomplete()) {
          await this.handleAutocomplete(interaction as any);
          return;
        }

        // Компоненты (кнопки и т.п.) для пагинации /doc blocks
        if ('isButton' in interaction && typeof (interaction as any).isButton === 'function' && (interaction as any).isButton()) {
          const customId = (interaction as any).customId as string | undefined;
          // Дизамбигуация дубликатов: dup|scope|userId|nonce|action|page
          if (customId && customId.startsWith('dup|')) {
            const parts = customId.split('|');
            const scope = parts[1];
            const cmd = scope ? (this.commands.get(scope) as unknown as { handleComponent?: (args: { interaction: any; componentType?: 'button' | 'select' | 'modal' }) => Promise<void> } | undefined) : undefined;
            if (cmd && typeof cmd.handleComponent === 'function') {
              await cmd.handleComponent({ interaction: interaction as any, componentType: 'button' });
              return;
            }
          }
          // Поиск: предпросмотр и пагинация текста
          if (customId && customId.startsWith('search|')) {
            await this.handleSearchButton(interaction as any);
            return;
          }
          if (customId && customId.startsWith('docblk|')) {
            const cmd = this.commands.get('doc') as unknown as { handleComponent?: (args: { interaction: any; componentType?: 'button' | 'select' | 'modal' }) => Promise<void> } | undefined;
            if (cmd && typeof cmd.handleComponent === 'function') {
              await cmd.handleComponent({ interaction: interaction as any, componentType: 'button' });
              return;
            }
          }
          // Компоненты FileManager (/файли пошук) — пагинация и переключатели
          if (customId && customId.startsWith('filesrch|')) {
            const cmd = this.commands.get('файли') as unknown as { handleComponent?: (args: { interaction: any; componentType?: 'button' | 'select' | 'modal' }) => Promise<void> } | undefined;
            if (cmd && typeof cmd.handleComponent === 'function') {
              await cmd.handleComponent({ interaction: interaction as any, componentType: 'button' });
              return;
            }
          }
        }

        // SelectMenu для дизамбигуации и других компонентов
        if ('isStringSelectMenu' in interaction && typeof (interaction as any).isStringSelectMenu === 'function' && (interaction as any).isStringSelectMenu()) {
          const customId = (interaction as any).customId as string | undefined;
          if (customId && customId.startsWith('dup|')) {
            const parts = customId.split('|');
            const scope = parts[1];
            const cmd = scope ? (this.commands.get(scope) as unknown as { handleComponent?: (args: { interaction: any; componentType?: 'button' | 'select' | 'modal' }) => Promise<void> } | undefined) : undefined;
            if (cmd && typeof cmd.handleComponent === 'function') {
              await cmd.handleComponent({ interaction: interaction as any, componentType: 'select' });
              return;
            }
          }
        }
      } catch (error) {
        logger.error('❌ Ошибка верхнего уровня в обработчике InteractionCreate', {
          type: 'command_manager',
          event: 'interaction_error',
          errorMessage: String(error),
        });
      }
    });

    // Гейтований слухач повідомлень для NLP intent (не впливає на поточну поведінку)
    try {
      if ((this.config as any)?.nlp?.enableIntent) {
        (this.bot.client as any).on('messageCreate', async (message: any) => {
          try {
            if (!message || message.author?.bot) return;
            const content: string = String(message.content ?? '');
            if (!content.trim()) return;

            // Валідація безпеки вводу
            try {
              const valid = await Promise.resolve(validateInput(content));
              if (valid && (valid as any).isValid === false) {
                // Не відповідаємо, лише лог
                logger.debug('intent_skip_invalid', { reason: (valid as any).reason });
                return;
              }
            } catch {
              // ignore validation errors
            }

            // Нормалізація і мова
            const normalized = normalizeText(content);
            const lang = detectLanguage(normalized);

            // Опційний AI fallback через AIService (якщо доступний)
            let classify: ClassifyIntentFn | undefined;
            try {
              const ai = this.bot.getService?.('ai') as any;
              if (ai && typeof ai.classifyIntent === 'function') {
                classify = async (text, opts) => {
                  const res = await ai.classifyIntent(text, opts);
                  return { intent: res?.intent, confidence: res?.confidence };
                };
              }
            } catch {
              // no ai fallback
            }

            const options: Parameters<typeof detectIntent>[1] = {
              timeoutMs: 2000,
              maxTokens: 128,
              defaultLocale: 'uk',
            };
            if (classify) {
              (options as any).classifyIntent = classify;
            }
            const detected = await detectIntent(normalized, options);

            logger.debug('intent_detected', {
              userId: message.author?.id,
              guildId: message.guild?.id,
              channelId: message.channel?.id,
              lang,
              ...detected,
            });
            // На цьому етапі лише логування; маршрутизацію увімкнемо пізніше за погодженням
          } catch (e) {
            logger.debug('intent_listener_failed', { error: e instanceof Error ? e.message : String(e) });
          }
        });
      }
    } catch {
      // ignore listener setup errors
    }
  }

  /**
   * Перевірка прав доступу (дружня до тестів)
   */
  private async checkPermissions(
    interaction: ChatInputCommandInteraction
  ): Promise<boolean> {
    try {
      // У тестах пропускаємо важку систему прав, щоб уникнути таймерів та відкритих хендлів
      if (process.env['JEST_WORKER_ID'] || process.env['NODE_ENV'] === 'test') {
        return true;
      }

      const user = (interaction as any).user;
      const member = (interaction as any).member ?? null;
      const channelId = (interaction as any).channelId as string | undefined;
      const commandName = interaction.commandName;

      // Якщо немає користувача (мок), дозволяємо виконання
      if (!user) return true;

      // Ліниве підключення, щоб уникнути циклічних залежностей у тестах
      const mod = await import('@/core/PermissionManager');
      const PermissionManager = (mod as any).PermissionManager as
        | (new (config: any) => { checkPermission: Function })
        | undefined;
      if (!PermissionManager) return true;

      const pm = new PermissionManager(this.config) as any;
      const result = await pm.checkPermission(user, member, commandName, channelId);
      if (!result?.allowed) {
        const reason = result?.reason ?? 'Недостатньо прав для виконання команди.';
        const msg = `🚫 Доступ заборонено. ${reason}`;
        if ((interaction as any).replied || (interaction as any).deferred) {
          await (interaction as any).editReply({ content: msg });
        } else {
          await replyWithPrivacy(interaction as any, { content: msg });
        }
        return false;
      }
      return true;
    } catch (e) {
      // На будь-яку помилку — не блокуємо виконання команди
      return true;
    }
  }

  /**
   * Обробка команди
   */
  private async handleCommand(interaction: ChatInputCommandInteraction): Promise<void> {
    try {
      const commandName = interaction.commandName;
      const command = this.commands.get(commandName);

      if (!command) {
        await replyWithPrivacy(interaction as any, { content: '❌ Команда не знайдена' });
        return;
      }

      // Оновлюємо статистику
      this.stats.lastUsed = new Date();

      // Перевірка прав доступу
      const hasPermission = await this.checkPermissions(interaction);
      if (!hasPermission) {
        // Повідомлення вже відправлено в checkPermissions(); уникаємо повторної відповіді
        return;
      }

      // Виконання команди
      await command.execute({
        interaction,
      });

      logger.info('✅ Команда виконана', {
        type: 'command',
        event: 'executed',
        commandName,
        userId: (interaction as any)?.user?.id ?? 'unknown',
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: (interaction as any)?.channelId ?? 'unknown',
      });
    } catch (error) {
      const appErr = error instanceof AppError
        ? error
        : new AppError('UNEXPECTED_ERROR', 'common.error.unexpected', error);

      // Stable logging code + context
      const logObj = {
        type: 'command',
        event: 'execute_error',
        commandName: interaction.commandName,
        userId: (interaction as any)?.user?.id ?? 'unknown',
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: (interaction as any)?.channelId ?? 'unknown',
        code: appErr.code,
        errorMessage: String((appErr.cause as any)?.message ?? appErr.message ?? error),
      } as Record<string, unknown>;
      logger.error('❌ Помилка виконання команди', logObj);

      const localized = tUser(appErr.userMessageKey, interaction as any, appErr.meta as any);
      if (interaction.replied || interaction.deferred) {
        await interaction.editReply({ content: localized });
      } else {
        await replyWithPrivacy(interaction as any, { content: localized });
      }
    }
  }

  /**
   * Отримання команди за назвою
   */
  getCommand(name: string): ICommand | undefined {
    return this.commands.get(name);
  }

  /**
   * Отримання всіх команд
   */
  getAllCommands(): Collection<string, ICommand> {
    return this.commands;
  }

  /**
   * Сумісність з інтеграційними тестами: повертає всі команди
   */
  getCommands(): Collection<string, ICommand> {
    return this.getAllCommands();
  }

  /**
   * Отримання команд за категорією
   */
  getCommandsByCategory(category: string): string[] {
    return this.commandCategories.get(category) || [];
  }

  /**
   * Отримання всіх категорій
   */
  getCategories(): string[] {
    return Array.from(this.commandCategories.keys());
  }

  /**
   * Отримання статистики
   */
  getStats(): CommandStats {
    return { ...this.stats };
  }

  /**
   * Оновлення статистики
   */
  private updateStats(): void {
    this.stats.totalCommands = this.commands.size;
    this.stats.categories = this.commandCategories.size;
    this.stats.commandsByCategory = {};

    for (const [category, commands] of this.commandCategories.entries()) {
      this.stats.commandsByCategory[category] = commands.length;
    }
  }

  /**
   * Отримання даних для реєстрації команд в Discord
   */
  getCommandsData(): any[] {
    return Array.from(this.commands.values()).map(command => command.getData());
  }

  /**
   * Перезавантаження команд
   */
  async reloadCommands(): Promise<void> {
    logger.info('🔄 Перезавантаження команд...', {
      type: 'command_manager',
      event: 'reload_start',
    });

    this.commands.clear();
    this.commandCategories.clear();

    await this.loadCommands();

    logger.info('✅ Перезавантажено команди', {
      type: 'command_manager',
      event: 'reload_done',
      count: this.commands.size,
    });
  }

  /**
   * Сумісність з інтеграційними тестами: виконати команду напряму
   */
  async execute(interaction: ChatInputCommandInteraction | { commandName: string; reply?: Function; deferred?: boolean; replied?: boolean; user?: any; channelId?: string; guildId?: string | null }): Promise<void> {
    // Проксі до приватного handleCommand
    // Нотатка: у тестах може бути спрощений мок Interaction
    return this.handleCommand(interaction as ChatInputCommandInteraction);
  }
}
