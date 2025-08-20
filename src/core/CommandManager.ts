/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */

import { Collection, EmbedBuilder, Events, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import type { ChatInputCommandInteraction, GuildMember, Interaction, Client } from 'discord.js';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import { replyWithPrivacy } from '@/ui/reply';
import { tUser } from '@/i18n';

// Імпорт всіх команд
import { SearchCommand } from '@/commands/SearchCommand';
import { PerformanceCommand } from '@/commands/PerformanceCommand';
import { AIAssistantCommand } from '@/commands/AIAssistantCommand';
import { DocumentsCommand } from '@/commands/DocumentsCommand';
import { FileManagerCommand } from '@/commands/FileManagerCommand';
import { OperationsCommand } from '@/commands/OperationsCommand';
import { AnalyticsCommand } from '@/commands/AnalyticsCommand';
import { EnhancedSearchCommand } from '@/commands/EnhancedSearchCommand';
import { SelectSheetCommand } from '@/commands/SelectSheetCommand';
import { OCRCommand } from '@/commands/OCRCommand';
import { DriveExtractCommand } from '@/commands/DriveExtractCommand';
import { DocCommand } from '@/commands/DocCommand';
import { WorkspaceCommand } from '@/commands/WorkspaceCommand';
import { LangCommand } from '@/commands/LangCommand';
import { AnalyzeCommand } from '@/commands/AnalyzeCommand';
import { FavoritesCommand } from '@/commands/FavoritesCommand';
import { SavedSearchCommand } from '@/commands/SavedSearchCommand';

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

  constructor(bot: BotLike, config: BotConfig) {
    this.bot = bot;
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
      // Форматы: search|expand|{fileId}  или  search|page|{fileId}|{index}
      const parts = customId.split('|');
      const action = parts[1];
      const fileId = parts[2];
      const pageIndex = parts[3] ? parseInt(parts[3], 10) : 0;
      if (!fileId) {
        await replyWithPrivacy(interaction as any, { content: tUser('files.validation.invalidFileId', interaction) });
        return;
      }

      const driveIndexer = (this.bot.getService?.('driveIndexer') ?? undefined) as
        | import('@/services/DriveIndexerService').DriveIndexerService
        | undefined;
      if (!driveIndexer) {
        await replyWithPrivacy(interaction as any, { content: tUser('search.error.noService', interaction) });
        return;
      }

      const chunks = await driveIndexer.getTextChunks(fileId, 1800);
      if (!chunks.length) {
        await replyWithPrivacy(interaction as any, { content: tUser('files.error.noText', interaction) });
        return;
      }

      const idx = Math.min(Math.max(0, pageIndex), chunks.length - 1);
      const content = chunks[idx];

      const row = new ActionRowBuilder<ButtonBuilder>().addComponents(
        new ButtonBuilder()
          .setCustomId(`search|page|${fileId}|${Math.max(0, idx - 1)}`)
          .setLabel('⬅️')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(idx === 0),
        new ButtonBuilder()
          .setCustomId(`search|page|${fileId}|${Math.min(chunks.length - 1, idx + 1)}`)
          .setLabel('➡️')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(idx >= chunks.length - 1)
      );

      if (action === 'expand') {
        await replyWithPrivacy(interaction as any, { content: content ?? '', components: [row] });
        return;
      }
      if (action === 'page') {
        if (interaction.deferred || interaction.replied) {
          await interaction.editReply({ content, components: [row] });
        } else {
          await replyWithPrivacy(interaction as any, { content: content ?? '', components: [row] });
        }
        return;
      }

      await replyWithPrivacy(interaction as any, { content: tUser('files.error.unknownSubcommand', interaction) });
    } catch (e) {
      logger.error('search_button_failed', { error: e instanceof Error ? e.message : String(e) });
      try {
        await replyWithPrivacy(interaction as any, { content: tUser('workspace.common.execError', interaction) });
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
      if (!this.config.discord.enableSlash) {
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

      // Реєстрація обробників подій
      this.registerEventHandlers();

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
    try {
      // keep async semantics for future IO; satisfies lint rule
      await Promise.resolve();
      // Отримуємо сервіси через Bot.getService() (проксі до ServiceManager/ServiceContainer)
      // Це гарантує доступ до сервісів, створених у ServiceManager
      const googleService = (this.bot?.getService?.('google') ?? undefined) as
        | GoogleService
        | undefined;
      const sheetsContext = (this.bot?.getService?.('sheetsContext') ?? undefined) as
        | SheetsContextService
        | undefined;

      // Створюємо екземпляри всіх команд
      const commandInstances = [
        new SearchCommand(this.config, googleService),
        new PerformanceCommand(this.config),
        new AIAssistantCommand(this.config, googleService),
        new DocumentsCommand(this.config),
        new FileManagerCommand(this.config),
        new OperationsCommand(this.config),
        new AnalyticsCommand(this.config),
        new EnhancedSearchCommand(this.config, googleService),
        new SelectSheetCommand(this.config, googleService, sheetsContext),
        new OCRCommand(this.config, googleService),
        new DriveExtractCommand(this.config, googleService),
        new DocCommand(this.config, googleService),
        new WorkspaceCommand(this.config),
        new FavoritesCommand(this.config),
        new SavedSearchCommand(this.config),
        new LangCommand(this.config),
        new AnalyzeCommand(this.config),
      ];

      // Реєструємо команди
      for (const command of commandInstances) {
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
    } catch (error) {
      logger.error('❌ Помилка завантаження команд', {
        type: 'command_manager',
        event: 'load_error',
        errorMessage: String(error),
      });
      throw error;
    }
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
    if (name.includes('продуктивність') || name.includes('performance')) {
      return 'Моніторинг';
    }
    if (name.includes('ai') || name.includes('асистент')) {
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
    // Реєструємо обробник на Discord клієнті, а не на екземплярі Bot
    this.bot.client.on(Events.InteractionCreate, async (interaction: Interaction) => {
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
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
      });
    } catch (error) {
      logger.error('❌ Помилка виконання команди', {
        type: 'command',
        event: 'execute_error',
        commandName: interaction.commandName,
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        errorMessage: String(error),
      });

      const errorMessage =
        '❌ Помилка при виконанні команди. Спробуйте ще раз або зверніться до адміністратора.';

      if (interaction.replied || interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await replyWithPrivacy(interaction as any, { content: errorMessage });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermissions(interaction: ChatInputCommandInteraction): Promise<boolean> {
    try {
      // Імпорт PermissionManager
      const { PermissionManager } = await import('./PermissionManager');
      const permissionManager = new PermissionManager(this.config);

      // Перевірка прав доступу
      const result = await permissionManager.checkPermission(
        interaction.user,
        interaction.member as GuildMember | null,
        interaction.commandName,
        interaction.channelId
      );

      // Якщо доступ заборонено, відправляємо повідомлення користувачу
      if (!result.allowed) {
        const embed = this.createPermissionDeniedEmbed(result);
        await interaction.reply({ embeds: [embed], ephemeral: true });

        logger.security('command_access_denied', interaction.user.id, {
          type: 'security',
          event: 'command_access_denied',
          severity: 'medium',
          commandName: interaction.commandName,
          reason: result.reason,
          userLevel: result.userLevel,
          ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
          channelId: interaction.channelId,
          userId: interaction.user.id,
        });

        return false;
      }

      // Логування успішного доступу
      logger.info('✅ Команда дозволена', {
        type: 'command',
        event: 'permission_granted',
        userId: interaction.user.id,
        commandName: interaction.commandName,
        userLevel: result.userLevel,
        remainingUses: result.remainingUses,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
      });

      return true;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка перевірки прав доступу', {
          type: 'security',
          event: 'permission_check_error',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          commandName: interaction.commandName,
          userId: interaction.user.id,
          ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
          channelId: interaction.channelId,
          severity: 'high',
        });
      } else {
        logger.error('❌ Помилка перевірки прав доступу', {
          type: 'security',
          event: 'permission_check_error',
          commandName: interaction.commandName,
          userId: interaction.user.id,
          ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
          channelId: interaction.channelId,
          severity: 'high',
          errorMessage: String(error),
        });
      }

      // У разі помилки дозволяємо виконання для базових команд
      const allowedCommands = ['пошук', 'довідка', 'статус'];
      return allowedCommands.includes(interaction.commandName);
    }
  }

  /**
   * Створення embed повідомлення про відмову доступу
   */
  private createPermissionDeniedEmbed(result: any): EmbedBuilder {
    return new EmbedBuilder()
      .setColor(0xff0000)
      .setTitle('🚫 Доступ заборонено')
      .setDescription(`Вам заборонено використовувати цю команду.\n\n**Причина:** ${result.reason}`)
      .addFields([
        {
          name: '📊 Ваш рівень доступу',
          value: `${result.userLevel} (${['Заборонений', 'Користувач', 'Довірений', 'Модератор', 'Адміністратор', 'Власник'][result.userLevel]})`,
          inline: true,
        },
        {
          name: '🔄 Використання за день',
          value: result.remainingUses
            ? `Залишилось: ${result.remainingUses}`
            : 'Інформація недоступна',
          inline: true,
        },
        {
          name: "📞 Зв'яжіться з адміністратором",
          value: 'Якщо вважаєте, що це помилка, зверніться до адміністрації сервера.',
          inline: false,
        },
      ])
      .setFooter({ text: 'Discord AI Assistant Bot - Security System' })
      .setTimestamp();
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
}
