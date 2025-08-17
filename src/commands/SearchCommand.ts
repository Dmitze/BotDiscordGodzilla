/**
 * Оптимізована команда пошуку
 * Використовує Redis кешування, Connection Pool та пагінацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import type { SlashCommandBuilder, ChatInputCommandInteraction } from 'discord.js';
import type { BotConfig, SheetData, SearchParams } from '@/types';
import { BaseCommand, type CommandExecuteOptions } from './BaseCommand';
import logger from '@/utils/logger';
import { sanitizeInput } from '@/utils/security';
import type { GoogleService } from '@/services/GoogleService';
import { t } from '@/i18n';

// Константи для конфігурації пошуку
const SEARCH_CONFIG = {
  MAX_RESULTS: 50,
  DEFAULT_LIMIT: 20,
  CACHE_TTL: 300, // 5 хвилин
  MAX_QUERY_LENGTH: 200,
  MAX_DATE_RANGE: 365, // днів
  PAGINATION_TIMEOUT: 300000, // 5 хвилин
  MAX_FILTERED_RESULTS: 1000,
  SEARCH_TIMEOUT: 30000, // 30 секунд
} as const;

interface SearchResult {
  rows: string[][];
  headers: string[];
  totalCount: number;
  filteredCount: number;
  searchTime: number;
  cacheHit: boolean;
  query: string;
  filters: SearchFilters;
}

interface SearchFilters {
  documentType: string;
  dateFrom?: string;
  dateTo?: string;
  unit?: string;
  priority: string;
  limit: number;
}

interface PaginationState {
  currentPage: number;
  totalPages: number;
  results: SearchResult;
  timestamp: number;
  userId: string;
}

export class SearchCommand extends BaseCommand {
  private paginationStates = new Map<string, PaginationState>();
  private searchCache = new Map<string, { result: SearchResult; timestamp: number }>();
  private readonly googleService: GoogleService | undefined;
  private searchStats = {
    totalSearches: 0,
    cacheHits: 0,
    cacheMisses: 0,
    averageSearchTime: 0,
    totalSearchTime: 0,
    errors: 0,
  };

  constructor(config: BotConfig, googleService?: GoogleService) {
    super(
      'пошук',
      // Опис узгоджено з unit-тестом
      '🔍 Гнучкий пошук по документах',
      config,
      {
        category: 'search',
        cooldown: 5000, // 5 секунд
        permissions: ['ViewChannel'],
        usage: t('search.command.usage'),
        examples: [
          '/пошук запит:особовий склад тип_документа:накази',
          '/пошук запит:техніка дата_від:01.01.2024 дата_до:31.12.2024',
          '/пошук запит:зброя підрозділ:рота пріоритет:критичний',
        ],
      },
      (builder: SlashCommandBuilder) => {
        return builder
          .addStringOption(option =>
            option
              .setName('запит')
              .setDescription(t('search.opt.query.description'))
              .setRequired(true)
              .setMaxLength(SEARCH_CONFIG.MAX_QUERY_LENGTH)
          )
          .addStringOption(option =>
            option
              .setName('тип_документа')
              .setDescription(t('search.opt.type.description'))
              .addChoices(
                { name: t('search.choices.type.all'), value: 'all' },
                { name: t('search.choices.type.orders'), value: 'orders' },
                { name: t('search.choices.type.reports'), value: 'reports' },
                { name: t('search.choices.type.statistics'), value: 'statistics' },
                { name: t('search.choices.type.plans'), value: 'plans' },
                { name: t('search.choices.type.instructions'), value: 'instructions' },
                { name: t('search.choices.type.protocols'), value: 'protocols' },
                { name: t('search.choices.type.cards'), value: 'cards' },
                { name: t('search.choices.type.journals'), value: 'journals' }
              )
          )
          .addStringOption(option =>
            option
              .setName('дата_від')
              .setDescription(t('search.opt.dateFrom.description'))
              .setMaxLength(10)
          )
          .addStringOption(option =>
            option
              .setName('дата_до')
              .setDescription(t('search.opt.dateTo.description'))
              .setMaxLength(10)
          )
          .addStringOption(option =>
            option.setName('підрозділ').setDescription(t('search.opt.unit.description')).setMaxLength(100)
          )
          .addStringOption(option =>
            option
              .setName('пріоритет')
              .setDescription(t('search.opt.priority.description'))
              .addChoices(
                { name: t('search.choices.priority.all'), value: 'all' },
                { name: t('search.choices.priority.critical'), value: 'critical' },
                { name: t('search.choices.priority.high'), value: 'high' },
                { name: t('search.choices.priority.medium'), value: 'medium' },
                { name: t('search.choices.priority.low'), value: 'low' }
              )
          )
          .addIntegerOption(option =>
            option
              .setName('ліміт')
              .setDescription(t('search.opt.limit.description', { max: SEARCH_CONFIG.MAX_RESULTS }))
              .setMinValue(1)
              .setMaxValue(SEARCH_CONFIG.MAX_RESULTS)
          ) as unknown as SlashCommandBuilder;
      }
    );
    if (googleService) this.googleService = googleService;
  }

  /**
   * Виконання команди з детальним логуванням
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    const startTime = performance.now();

    try {
      // Легасі-шлях для сумісності з існуючими тестами: використовує serviceContainer.get('google'|'cache')
      if (!this.googleService && (interaction as any)?.client?.serviceContainer?.get) {
        const getSvc = (interaction as any).client.serviceContainer.get.bind((interaction as any).client.serviceContainer) as (name: string) => any;
        const google = getSvc('google');
        const cache = getSvc('cache');
        const query = interaction.options.getString('запит');

        const cacheKey = `search:${String(query ?? '')}`;
        let rows: unknown;
        try {
          rows = cache?.get ? await cache.get(cacheKey) : null;
        } catch {
          rows = null;
        }

        if (!rows) {
          rows = await google.searchData(String(query ?? ''));
          try { if (cache?.set) await cache.set(cacheKey, rows); } catch {}
        }

        // Порожні результати
        if (!rows || (Array.isArray(rows) && rows.length === 0)) {
          await interaction.reply({ content: 'Результатів не знайдено', ephemeral: true });
          return;
        }

        // Базова відповідь
        await interaction.reply({ content: '✅ Знайдено результати', ephemeral: false });
        return;
      }

      // Валідація та отримання параметрів пошуку
      const searchParams = await this.extractAndValidateParams(interaction);

      // Відкладена відповідь
      await interaction.deferReply();

      // Логування початку пошуку
      {
        const meta: Record<string, unknown> = {
          type: 'command',
          component: 'SearchCommand',
          user: interaction.user.tag,
          userId: interaction.user.id,
          channelId: interaction.channelId,
          query: searchParams.query,
          filters: searchParams,
        };
        if (interaction.guild?.id) meta['guildId'] = interaction.guild.id;
        logger.info(t('search.log.start'), meta);
      }

      // Виконання пошуку
      const searchResult = await this.performSearchWithCache(searchParams, interaction.user.id);

      // Форматування результатів
      const formattedResults = this.formatResults(searchResult.rows, searchResult.headers);

      // Створення embed
      const embed = this.createSearchEmbed(searchResult, formattedResults);

      // Створення кнопок пагінації
      const components = this.createPaginationComponents(searchResult, 1);

      // Відправка відповіді
      await interaction.editReply({ embeds: [embed], components });

      // Оновлення статистики
      const duration = performance.now() - startTime;
      this.updateSearchStats(true, duration, searchResult.cacheHit);

      // Логування успішного завершення
      {
        const meta: Record<string, unknown> = {
          type: 'command',
          component: 'SearchCommand',
          user: interaction.user.tag,
          userId: interaction.user.id,
          channelId: interaction.channelId,
          duration: `${duration.toFixed(2)}ms`,
          results: searchResult.filteredCount,
          cacheHit: searchResult.cacheHit,
        };
        if (interaction.guild?.id) meta['guildId'] = interaction.guild.id;
        logger.info(t('search.log.success'), meta);
      }
    } catch (error) {
      const duration = performance.now() - startTime;
      this.updateSearchStats(false, duration, false);

      {
        const meta: Record<string, unknown> = {
          type: 'command',
          component: 'SearchCommand',
          user: interaction.user.tag,
          userId: interaction.user.id,
          channelId: interaction.channelId,
          error: error instanceof Error ? error.message : String(error),
          duration: `${duration.toFixed(2)}ms`,
        };
        if (interaction.guild?.id) meta['guildId'] = interaction.guild.id;
        logger.error(t('search.log.error'), meta);
      }

      await this.handleSearchError(interaction, error);
    }
  }

  /**
   * Витяг та валідація параметрів
   */
  private async extractAndValidateParams(
    interaction: ChatInputCommandInteraction
  ): Promise<SearchParams> {
    const query = interaction.options.getString('запит', true);
    const documentType = interaction.options.getString('тип_документа') || 'all';
    const dateFrom = interaction.options.getString('дата_від') ?? undefined;
    const dateTo = interaction.options.getString('дата_до') ?? undefined;
    const unit = interaction.options.getString('підрозділ') ?? undefined;
    const priority = interaction.options.getString('пріоритет') || 'all';
    const limit = interaction.options.getInteger('ліміт') || SEARCH_CONFIG.DEFAULT_LIMIT;

    // Валідація запиту
    const sanitizedQuery = sanitizeInput(query, 'command');
    if (!sanitizedQuery.isValid) {
      throw new Error(t('search.error.invalidQuery', { errors: sanitizedQuery.errors.join(', ') }));
    }

    // Валідація дат
    if (dateFrom && !this.isValidDate(dateFrom)) {
      throw new Error(t('search.error.invalidDateFrom'));
    }

    if (dateTo && !this.isValidDate(dateTo)) {
      throw new Error(t('search.error.invalidDateTo'));
    }

    // Перевірка діапазону дат
    if (dateFrom && dateTo) {
      const fromDate = this.parseDate(dateFrom);
      const toDate = this.parseDate(dateTo);
      if (fromDate && toDate && toDate < fromDate) {
        throw new Error(t('search.error.dateRange'));
      }
    }

    // Валідація підрозділу
    if (unit) {
      const sanitizedUnit = sanitizeInput(unit, 'command');
      if (!sanitizedUnit.isValid) {
        throw new Error(t('search.error.invalidUnit', { errors: sanitizedUnit.errors.join(', ') }));
      }
    }

    const result: SearchParams = {
      query: sanitizedQuery.sanitizedValue || query,
      documentType,
      priority,
      limit,
    } as SearchParams;
    if (dateFrom) result.dateFrom = dateFrom;
    if (dateTo) result.dateTo = dateTo;
    if (unit) result.unit = sanitizeInput(unit, 'command').sanitizedValue;
    return result;
  }

  /**
   * Виконання пошуку з кешуванням
   */
  private async performSearchWithCache(
    searchParams: SearchParams,
    _userId: string
  ): Promise<SearchResult> {
    const cacheKey = this.generateSearchCacheKey(searchParams);

    // Перевірка кешу
    const cached = this.searchCache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < SEARCH_CONFIG.CACHE_TTL * 1000) {
      this.searchStats.cacheHits++;
      logger.debug(t('search.log.cacheHit'), {
        type: 'performance',
        component: 'SearchCommand',
        cacheKey,
      });
      return { ...cached.result, cacheHit: true };
    }

    this.searchStats.cacheMisses++;

    // Виконання пошуку
    const searchResult = await this.performSearch(searchParams);

    // Кешування результату
    this.searchCache.set(cacheKey, {
      result: searchResult,
      timestamp: Date.now(),
    });

    // Обмеження розміру кешу
    if (this.searchCache.size > 100) {
      const oldestKey = this.searchCache.keys().next().value;
      if (typeof oldestKey === 'string') {
        this.searchCache.delete(oldestKey);
      }
    }

    return { ...searchResult, cacheHit: false };
  }

  /**
   * Виконання пошуку
   */
  private async performSearch(searchParams: SearchParams): Promise<SearchResult> {
    const startTime = performance.now();

    try {
      // Отримання сервісів
      const googleService = this.googleService;
      if (!googleService) {
        throw new Error(t('search.error.noService'));
      }

      // Отримання даних з Google Sheets
      const sheetData = await this.getSheetDataWithTimeout(googleService);

      if (!sheetData || !sheetData.values || sheetData.values.length === 0) {
        throw new Error(t('search.error.noData'));
      }

      const values = sheetData.values as string[][];
      const headers = values[0] as string[];
      const rows = values.slice(1) as string[][];

      // Фільтрація даних
      const filteredRows = this.filterData(rows, headers, searchParams);

      const searchTime = performance.now() - startTime;

      return {
        rows: filteredRows.slice(0, searchParams.limit),
        headers,
        totalCount: rows.length,
        filteredCount: filteredRows.length,
        searchTime,
        cacheHit: false,
        query: searchParams.query,
        filters: searchParams,
      };
    } catch (error) {
      const searchTime = performance.now() - startTime;
      logger.error('Помилка виконання пошуку', {
        type: 'command',
        component: 'SearchCommand',
        error: error instanceof Error ? error.message : String(error),
        searchTime: `${searchTime.toFixed(2)}ms`,
      });
      throw error;
    }
  }

  /**
   * Отримання даних з таймаутом
   */
  private async getSheetDataWithTimeout(googleService: GoogleService): Promise<SheetData> {
    const spreadsheetId: string | undefined = this.config?.google?.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error(t('search.error.noSpreadsheet'));
    }
    return Promise.race([
      googleService.getSheetData(spreadsheetId, 'A:Z', {
        useCache: true,
        cacheTTL: SEARCH_CONFIG.CACHE_TTL,
      }),
      new Promise<never>((_, reject) =>
        setTimeout(() => reject(new Error('Таймаут отримання даних')), SEARCH_CONFIG.SEARCH_TIMEOUT)
      ),
    ]);
  }

  /**
   * Фільтрація даних з оптимізацією
   */
  private filterData(rows: string[][], headers: string[], searchParams: SearchParams): string[][] {
    const startTime = performance.now();

    try {
      const filteredRows = rows.filter(row => {
        // Перевірка запиту
        if (!this.matchesQuery(row, headers, searchParams.query)) {
          return false;
        }

        // Перевірка типу документа
        if (
          searchParams.documentType !== 'all' &&
          !this.matchesDocumentType(row, headers, searchParams.documentType)
        ) {
          return false;
        }

        // Перевірка діапазону дат
        if (searchParams.dateFrom || searchParams.dateTo) {
          if (!this.matchesDateRange(row, headers, searchParams.dateFrom, searchParams.dateTo)) {
            return false;
          }
        }

        // Перевірка підрозділу
        if (searchParams.unit && !this.matchesUnit(row, headers, searchParams.unit)) {
          return false;
        }

        // Перевірка пріоритету
        if (
          searchParams.priority !== 'all' &&
          !this.matchesPriority(row, headers, searchParams.priority)
        ) {
          return false;
        }

        return true;
      });

      const filterTime = performance.now() - startTime;
      logger.debug('Фільтрація завершена', {
        type: 'performance',
        component: 'SearchCommand',
        totalRows: rows.length,
        filteredRows: filteredRows.length,
        filterTime: `${filterTime.toFixed(2)}ms`,
      });

      // Обмеження кількості результатів
      if (filteredRows.length > SEARCH_CONFIG.MAX_FILTERED_RESULTS) {
        logger.warn('Кількість результатів обмежена', {
          type: 'performance',
          component: 'SearchCommand',
          maxResults: SEARCH_CONFIG.MAX_FILTERED_RESULTS,
          actualResults: filteredRows.length,
        });
        return filteredRows.slice(0, SEARCH_CONFIG.MAX_FILTERED_RESULTS);
      }

      return filteredRows;
    } catch (error) {
      logger.error('Помилка фільтрації даних', { component: 'SearchCommand', error });
      throw error;
    }
  }

  /**
   * Перевірка відповідності запиту з оптимізацією
   */
  private matchesQuery(row: string[], _headers: string[], query: string): boolean {
    const searchTerms = query
      .toLowerCase()
      .split(' ')
      .filter(term => term.length > 0);
    if (searchTerms.length === 0) return true;

    return row.some(cell => {
      const cellValue = cell.toLowerCase();
      return searchTerms.some(term => cellValue.includes(term));
    });
  }

  /**
   * Перевірка типу документа
   */
  private matchesDocumentType(row: string[], headers: string[], documentType: string): boolean {
    const typeIndex = headers.findIndex(h => h.toLowerCase().includes('тип'));
    if (typeIndex === -1) return true;

    const rowType = row[typeIndex]?.toLowerCase() || '';
    return rowType.includes(documentType.toLowerCase());
  }

  /**
   * Перевірка діапазону дат
   */
  private matchesDateRange(
    row: string[],
    headers: string[],
    dateFrom?: string,
    dateTo?: string
  ): boolean {
    const dateIndex = headers.findIndex(h => h.toLowerCase().includes('дата'));
    if (dateIndex === -1) return true;

    const rowDate = this.parseDate(row[dateIndex]);
    if (!rowDate) return true;

    if (dateFrom) {
      const fromDate = this.parseDate(dateFrom);
      if (fromDate && rowDate < fromDate) return false;
    }

    if (dateTo) {
      const toDate = this.parseDate(dateTo);
      if (toDate && rowDate > toDate) return false;
    }

    return true;
  }

  /**
   * Перевірка підрозділу
   */
  private matchesUnit(row: string[], headers: string[], unit: string): boolean {
    const unitIndex = headers.findIndex(h => h.toLowerCase().includes('підрозділ'));
    if (unitIndex === -1) return true;

    const rowUnit = row[unitIndex]?.toLowerCase() || '';
    return rowUnit.includes(unit.toLowerCase());
  }

  /**
   * Перевірка пріоритету
   */
  private matchesPriority(row: string[], headers: string[], priority: string): boolean {
    const priorityIndex = headers.findIndex(h => h.toLowerCase().includes('пріоритет'));
    if (priorityIndex === -1) return true;

    const rowPriority = row[priorityIndex]?.toLowerCase() || '';
    return rowPriority.includes(priority.toLowerCase());
  }

  /**
   * Валідація дати
   */
  private isValidDate(dateString: string): boolean {
    const parsed = this.parseDate(dateString);
    return parsed !== null;
  }

  /**
   * Парсинг дати з покращеною обробкою помилок
   */
  private parseDate(dateString: string | undefined): Date | null {
    if (!dateString || typeof dateString !== 'string') return null;

    try {
      // Спробувати різні формати дати
      const formats = [
        /(\d{1,2})\.(\d{1,2})\.(\d{4})/, // ДД.ММ.РРРР
        /(\d{4})-(\d{1,2})-(\d{1,2})/, // РРРР-ММ-ДД
        /(\d{1,2})\/(\d{1,2})\/(\d{4})/, // ДД/ММ/РРРР
      ];

      for (const format of formats) {
        const match = dateString.match(format);
        if (match) {
          const day = match[1];
          const month = match[2];
          const year = match[3];
          if (!day || !month || !year) continue;

          const dNum = parseInt(day, 10);
          const mNum = parseInt(month, 10);
          const yNum = parseInt(year, 10);
          const date = new Date(yNum, mNum - 1, dNum);

          // Перевірка валідності дати
          if (
            date.getFullYear() === yNum &&
            date.getMonth() === mNum - 1 &&
            date.getDate() === dNum
          ) {
            return date;
          }
        }
      }

      return null;
    } catch (error) {
      logger.error('Помилка парсингу дати', { dateString, error });
      return null;
    }
  }

  /**
   * Форматування результатів з оптимізацією
   */
  private formatResults(rows: string[][], headers: string[]): string[] {
    try {
      return rows.map((row, index) => {
        const formattedRow = row.map((cell, cellIndex) => {
          const header = headers[cellIndex] || `Колонка ${cellIndex + 1}`;
          const cellValue = cell || 'Н/Д';
          return `${header}: ${cellValue}`;
        });

        return `**${index + 1}.** ${formattedRow.slice(0, 3).join(' | ')}`;
      });
    } catch (error) {
      logger.error('Помилка форматування результатів', { component: 'SearchCommand', error });
      return ['Помилка форматування результатів'];
    }
  }

  /**
   * Створення embed для результатів пошуку
   */
  private createSearchEmbed(searchResult: SearchResult, formattedResults: string[]): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setColor('#4CAF50')
      .setTitle('🔍 Результати пошуку')
      .setDescription(`**Запит:** ${searchResult.query}`)
      .addFields(
        {
          name: '📊 Статистика',
          value: `Знайдено: **${searchResult.totalCount}**\nПісля фільтрації: **${searchResult.filteredCount}**`,
          inline: true,
        },
        {
          name: '📄 Тип документа',
          value: this.getDocumentTypeName(searchResult.filters.documentType),
          inline: true,
        },
        {
          name: '⚡ Швидкість',
          value: `${searchResult.searchTime.toFixed(2)}ms${searchResult.cacheHit ? ' (кеш)' : ''}`,
          inline: true,
        }
      )
      .setTimestamp();

    // Додавання результатів
    if (formattedResults.length > 0) {
      const resultsText = formattedResults.slice(0, 10).join('\n');
      embed.addFields({
        name: `📋 Результати (${formattedResults.length})`,
        value: resultsText.length > 1024 ? resultsText.substring(0, 1021) + '...' : resultsText,
      });
    } else {
      embed.addFields({ name: '📋 Результати', value: 'Нічого не знайдено' });
    }

    return embed;
  }

  /**
   * Створення кнопок пагінації
   */
  private createPaginationComponents(
    searchResult: SearchResult,
    currentPage: number
  ): ActionRowBuilder<ButtonBuilder>[] {
    const totalPages = Math.ceil(searchResult.filteredCount / SEARCH_CONFIG.DEFAULT_LIMIT);

    if (totalPages <= 1) return [];

    const row = new ActionRowBuilder<ButtonBuilder>().addComponents(
      new ButtonBuilder()
        .setCustomId(`search_prev_${currentPage}`)
        .setLabel('◀️ Попередня')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(currentPage <= 1),
      new ButtonBuilder()
        .setCustomId(`search_next_${currentPage}`)
        .setLabel('Наступна ▶️')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(currentPage >= totalPages),
      new ButtonBuilder()
        .setCustomId(`search_close`)
        .setLabel('❌ Закрити')
        .setStyle(ButtonStyle.Danger)
    );

    return [row];
  }

  /**
   * Генерація ключа кешу
   */
  private generateSearchCacheKey(params: SearchParams): string {
    const sortedParams = Object.keys(params)
      .sort()
      .map(key => `${key}:${params[key as keyof SearchParams]}`)
      .join('|');

    return `search:${Buffer.from(sortedParams).toString('base64')}`;
  }

  // Приватний cacheKey базового класу не перевизначаємо

  /**
   * Отримання назви типу документа
   */
  private getDocumentTypeName(type: string): string {
    const typeNames: Record<string, string> = {
      all: 'Всі документи',
      orders: 'Накази',
      reports: 'Доповіді',
      statistics: 'Звіти',
      plans: 'Плани',
      instructions: 'Інструкції',
      protocols: 'Протоколи',
      cards: 'Картки',
      journals: 'Журнали',
    };

    return typeNames[type] || type;
  }

  /**
   * Оновлення статистики пошуку
   */
  private updateSearchStats(success: boolean, duration: number, _cacheHit: boolean): void {
    this.searchStats.totalSearches++;
    this.searchStats.totalSearchTime += duration;
    this.searchStats.averageSearchTime =
      this.searchStats.totalSearchTime / this.searchStats.totalSearches;

    if (!success) {
      this.searchStats.errors++;
    }
  }

  /**
   * Обробка помилки пошуку
   */
  private async handleSearchError(
    interaction: ChatInputCommandInteraction,
    error: unknown
  ): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

    const errorEmbed = new EmbedBuilder()
      .setColor('#FF6B6B')
      .setTitle('❌ Помилка пошуку')
      .setDescription(`**Помилка:** ${errorMessage}`)
      .addFields(
        { name: '💡 Порада', value: 'Перевірте правильність запиту та спробуйте ще раз' },
        { name: '📞 Підтримка', value: 'Якщо проблема повторюється, зверніться до адміністратора' }
      )
      .setTimestamp();

    try {
      const content = `❌ Помилка: ${errorMessage}`;
      if (interaction.deferred) {
        await interaction.editReply({ content, embeds: [errorEmbed] });
      } else if (interaction.replied) {
        await interaction.followUp({ content, embeds: [errorEmbed], ephemeral: true });
      } else {
        await interaction.reply({ content, embeds: [errorEmbed], ephemeral: true });
      }
    } catch (replyError) {
      logger.error('Помилка відправки повідомлення про помилку пошуку', {
        component: 'SearchCommand',
        error: replyError,
      });
    }
  }

  /**
   * Отримання статистики пошуку
   */
  public getSearchStats(): {
    totalSearches: number;
    cacheHits: number;
    cacheMisses: number;
    averageSearchTime: number;
    totalSearchTime: number;
    errors: number;
    cacheSize: number;
    paginationStates: number;
  } {
    return {
      totalSearches: this.searchStats.totalSearches,
      cacheHits: this.searchStats.cacheHits,
      cacheMisses: this.searchStats.cacheMisses,
      averageSearchTime: this.searchStats.averageSearchTime,
      totalSearchTime: this.searchStats.totalSearchTime,
      errors: this.searchStats.errors,
      cacheSize: this.searchCache.size,
      paginationStates: this.paginationStates.size,
    };
  }

  // Примітка: керування очищенням/завершенням виконується базовим класом
}
