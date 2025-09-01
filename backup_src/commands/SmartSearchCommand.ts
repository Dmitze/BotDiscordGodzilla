/**
 * 🔍 Команда розумного пошуку документів
 * Smart Search Command
 */

import {
  SlashCommandBuilder,
  SlashCommandStringOption,
  SlashCommandBooleanOption,
  SlashCommandIntegerOption,
  ChatInputCommandInteraction,
  EmbedBuilder,
  ButtonBuilder,
  ActionRowBuilder,
  ButtonStyle
} from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { SmartSearchEngine } from '@/services/SmartSearchEngine';

import logger from '@/utils/logger';

export class SmartSearchCommand extends BaseCommand {
  constructor(
    config: BotConfig,
    private smartSearch?: SmartSearchEngine
  ) {
    super('smart-search', 'Розумний пошук документів з AI', config, {
      i18n: { nameKey: 'commands.smart_search.name', descriptionKey: 'commands.smart_search.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('query')
          .setDescription('Пошуковий запит (природна мова)')
          .setRequired(true)
          .setMaxLength(500)
      );

      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('document_type')
          .setDescription('Тип документів для пошуку')
          .setRequired(false)
          .addChoices(
            { name: 'Всі типи', value: 'all' },
            { name: 'Google Docs', value: 'application/vnd.google-apps.document' },
            { name: 'Google Sheets', value: 'application/vnd.google-apps.spreadsheet' },
            { name: 'PDF файли', value: 'application/pdf' },
            { name: 'Зображення', value: 'image/*' },
            { name: 'Презентації', value: 'application/vnd.google-apps.presentation' }
          )
      );

      builder.addBooleanOption((option: SlashCommandBooleanOption) =>
        option
          .setName('semantic_search')
          .setDescription('Використовувати семантичний пошук')
          .setRequired(false)
      );

      builder.addBooleanOption((option: SlashCommandBooleanOption) =>
        option
          .setName('fuzzy_match')
          .setDescription('Нечіткий пошук (виправлення помилок)')
          .setRequired(false)
      );

      builder.addIntegerOption((option: SlashCommandIntegerOption) =>
        option
          .setName('limit')
          .setDescription('Максимальна кількість результатів')
          .setRequired(false)
          .setMinValue(1)
          .setMaxValue(50)
      );

      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('date_range')
          .setDescription('Діапазон дат (наприклад: last_week, last_month, last_year)')
          .setRequired(false)
          .addChoices(
            { name: 'Останній тиждень', value: 'last_week' },
            { name: 'Останній місяць', value: 'last_month' },
            { name: 'Останні 3 місяці', value: 'last_3_months' },
            { name: 'Останній рік', value: 'last_year' }
          )
      );

      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('sort_by')
          .setDescription('Сортування результатів')
          .setRequired(false)
          .addChoices(
            { name: 'За релевантністю', value: 'relevance' },
            { name: 'За датою зміни', value: 'date' },
            { name: 'За назвою', value: 'name' },
            { name: 'За розміром', value: 'size' }
          )
      );

      return builder;
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    if (!this.smartSearch) {
      await interaction.reply({
        content: '❌ Сервіс розумного пошуку недоступний',
        ephemeral: true
      });
      return;
    }

    try {
      await interaction.deferReply();

      const query = interaction.options.getString('query', true);
      const documentType = interaction.options.getString('document_type');
      const semanticSearch = interaction.options.getBoolean('semantic_search') ?? true;
      const fuzzyMatch = interaction.options.getBoolean('fuzzy_match') ?? true;
      const limit = interaction.options.getInteger('limit') ?? 10;
      const dateRange = interaction.options.getString('date_range');
      const sortBy = interaction.options.getString('sort_by') ?? 'relevance';

      // Побудова пошукового запиту
      const searchQuery = {
        text: query,
        filters: this.buildFilters(documentType, dateRange),
        options: {
          limit,
          semanticSearch,
          fuzzyMatch,
          language: 'uk' as const,
          sortBy: sortBy as any,
          sortOrder: 'desc' as const,
          includeContent: true
        }
      };

      // Виконання пошуку
      const searchStart = Date.now();
      const { results, insight } = await this.smartSearch.search(searchQuery);
      const searchTime = Date.now() - searchStart;

      // Створення відповіді
      if (results.length === 0) {
        await this.handleNoResults(interaction, query, insight);
        return;
      }

      // Створення embed з результатами
      const embed = this.createResultsEmbed(query, results, insight, searchTime);
      
      // Створення кнопок навігації
      const components = this.createNavigationComponents(results, 0);

      await interaction.editReply({
        embeds: [embed],
        components
      });

      // Збереження результатів для навігації
      this.saveSearchResults(interaction.user.id, results, insight);

      logger.info('Розумний пошук завершено', {
        component: 'SmartSearchCommand',
        query,
        resultsCount: results.length,
        searchTime,
        userId: interaction.user.id
      });

    } catch (error) {
      logger.error('Помилка розумного пошуку', {
        component: 'SmartSearchCommand',
        userId: interaction.user.id,
        error: error instanceof Error ? error.message : String(error)
      });

      await interaction.editReply({
        content: '❌ Помилка під час виконання пошуку'
      });
    }
  }

  /**
   * Побудова фільтрів пошуку
   */
  private buildFilters(documentType?: string | null, dateRange?: string | null): any {
    const filters: any = {};

    if (documentType && documentType !== 'all') {
      filters.mimeType = [documentType];
    }

    if (dateRange) {
      const now = new Date();
      let fromDate: Date;

      switch (dateRange) {
        case 'last_week':
          fromDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          break;
        case 'last_month':
          fromDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
          break;
        case 'last_3_months':
          fromDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
          break;
        case 'last_year':
          fromDate = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000);
          break;
        default:
          return filters;
      }

      filters.dateRange = { from: fromDate, to: now };
    }

    return filters;
  }

  /**
   * Створення embed з результатами пошуку
   */
  private createResultsEmbed(query: string, results: any[], insight: any, searchTime: number): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setTitle('🔍 Результати розумного пошуку')
      .setDescription(`**Запит:** ${query}`)
      .setColor(0x0099ff)
      .setFooter({ 
        text: `Знайдено ${results.length} результатів за ${searchTime}мс` 
      })
      .setTimestamp();

    // Додавання перших 5 результатів
    const displayResults = results.slice(0, 5);
    
    for (let i = 0; i < displayResults.length; i++) {
      const result = displayResults[i];
      const relevanceBar = this.createRelevanceBar(result.relevanceScore);
      
      embed.addFields({
        name: `${i + 1}. ${result.name}`,
        value: [
          `📊 Релевантність: ${relevanceBar} (${Math.round(result.relevanceScore * 100)}%)`,
          `🔗 ID: \`${result.fileId}\``,
          result.summary ? `📝 ${result.summary.substring(0, 100)}...` : '',
          result.lastModified ? `📅 Змінено: <t:${Math.floor(result.lastModified.getTime() / 1000)}:R>` : ''
        ].filter(Boolean).join('\n'),
        inline: false
      });
    }

    // Додавання insights
    if (insight.suggestions && insight.suggestions.length > 0) {
      embed.addFields({
        name: '💡 Пропозиції',
        value: insight.suggestions.slice(0, 3).map((s: string, i: number) => `${i + 1}. ${s}`).join('\n'),
        inline: false
      });
    }

    // Додавання категорій
    if (insight.categories && Object.keys(insight.categories).length > 0) {
      const categories = Object.entries(insight.categories)
        .map(([cat, count]) => `${cat}: ${count}`)
        .join(', ');
      
      embed.addFields({
        name: '📂 Категорії',
        value: categories,
        inline: false
      });
    }

    return embed;
  }

  /**
   * Створення компонентів навігації
   */
  private createNavigationComponents(results: any[], currentPage: number): ActionRowBuilder<ButtonBuilder>[] {
    const itemsPerPage = 5;
    const totalPages = Math.ceil(results.length / itemsPerPage);
    
    if (totalPages <= 1) return [];

    const components: ActionRowBuilder<ButtonBuilder>[] = [];
    
    // Кнопки навігації по сторінках
    const navigationRow = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('search_prev')
          .setLabel('← Попередня')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(currentPage === 0),
        
        new ButtonBuilder()
          .setCustomId('search_info')
          .setLabel(`${currentPage + 1}/${totalPages}`)
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(true),
        
        new ButtonBuilder()
          .setCustomId('search_next')
          .setLabel('Наступна →')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(currentPage >= totalPages - 1)
      );

    // Кнопки дій
    const actionRow = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('search_analyze')
          .setLabel('🧠 Аналізувати вибране')
          .setStyle(ButtonStyle.Primary),
        
        new ButtonBuilder()
          .setCustomId('search_download')
          .setLabel('📥 Завантажити')
          .setStyle(ButtonStyle.Success),
        
        new ButtonBuilder()
          .setCustomId('search_refine')
          .setLabel('🔧 Уточнити пошук')
          .setStyle(ButtonStyle.Secondary)
      );

    components.push(navigationRow, actionRow);
    return components;
  }

  /**
   * Обробка відсутності результатів
   */
  private async handleNoResults(interaction: ChatInputCommandInteraction, query: string, insight: any): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('🔍 Результати пошуку')
      .setDescription(`Не знайдено документів для запиту: **${query}**`)
      .setColor(0xff9900)
      .setTimestamp();

    // Додавання пропозицій
    if (insight.suggestions && insight.suggestions.length > 0) {
      embed.addFields({
        name: '💡 Спробуйте',
        value: insight.suggestions.slice(0, 5).map((s: string, i: number) => `${i + 1}. ${s}`).join('\n'),
        inline: false
      });
    }

    // Кнопка для нового пошуку
    const retryButton = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('search_new')
          .setLabel('🔍 Новий пошук')
          .setStyle(ButtonStyle.Primary),
        
        new ButtonBuilder()
          .setCustomId('search_help')
          .setLabel('❓ Допомога')
          .setStyle(ButtonStyle.Secondary)
      );

    await interaction.editReply({
      embeds: [embed],
      components: [retryButton]
    });
  }

  /**
   * Створення шкали релевантності
   */
  private createRelevanceBar(score: number): string {
    const bars = Math.round(score * 10);
    const filled = '█'.repeat(bars);
    const empty = '░'.repeat(10 - bars);
    return `${filled}${empty}`;
  }

  /**
   * Збереження результатів пошуку для навігації
   */
  private saveSearchResults(userId: string, results: any[], insight: any): void {
    // Тимчасове збереження в пам'яті
    // В продакшені краще використовувати Redis або базу даних
    const key = `search_${userId}`;
    (global as any).searchCache = (global as any).searchCache || new Map();
    (global as any).searchCache.set(key, {
      results,
      insight,
      timestamp: Date.now()
    });
    
    // Очищення старих результатів (через 30 хвилин)
    setTimeout(() => {
      (global as any).searchCache?.delete(key);
    }, 30 * 60 * 1000);
  }
}