/**
 * 🔍 Покращений пошук з діапазонами та сортуванням
 * Розширені можливості пошуку та фільтрації даних
 */

import { 
  EmbedBuilder, 
  ActionRowBuilder, 
  ButtonBuilder, 
  ButtonStyle,
  StringSelectMenuBuilder,
  StringSelectMenuOptionBuilder,
  ChatInputCommandInteraction
} from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

interface SearchFilters {
  name?: string | undefined;
  client?: string | undefined;
  series?: string | undefined;
  priceFrom?: number | undefined;
  priceTo?: number | undefined;
  quantityFrom?: number | undefined;
  quantityTo?: number | undefined;
  sortBy: string;
  sortOrder: string;
}

interface SearchResult {
  row: string[];
  score: number;
}

export class EnhancedSearchCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'розширений_пошук',
      '🔍 Покращений пошук з діапазонами та сортуванням',
      config,
      {},
      (builder: any) => {
        return builder
          .addStringOption((option: any) =>
            option
              .setName('номенклатура')
              .setDescription('Назва товару для пошуку')
              .setRequired(false)
              .setMaxLength(100)
          )
          .addStringOption((option: any) =>
            option
              .setName('контрагент')
              .setDescription('Назва контрагента')
              .setRequired(false)
              .setMaxLength(100)
          )
          .addStringOption((option: any) =>
            option
              .setName('серія')
              .setDescription('Серія товару')
              .setRequired(false)
              .setMaxLength(50)
          )
          .addNumberOption((option: any) =>
            option
              .setName('ціна_від')
              .setDescription('Мінімальна ціна')
              .setRequired(false)
              .setMinValue(0)
          )
          .addNumberOption((option: any) =>
            option
              .setName('ціна_до')
              .setDescription('Максимальна ціна')
              .setRequired(false)
              .setMinValue(0)
          )
          .addNumberOption((option: any) =>
            option
              .setName('кількість_від')
              .setDescription('Мінімальна кількість')
              .setRequired(false)
              .setMinValue(0)
          )
          .addNumberOption((option: any) =>
            option
              .setName('кількість_до')
              .setDescription('Максимальна кількість')
              .setRequired(false)
              .setMinValue(0)
          )
          .addStringOption((option: any) =>
            option
              .setName('сортування')
              .setDescription('Поле для сортування')
              .setRequired(false)
              .addChoices(
                { name: 'Назва', value: 'назва' },
                { name: 'Ціна', value: 'ціна' },
                { name: 'Кількість', value: 'кількість' },
                { name: 'Контрагент', value: 'контрагент' },
                { name: 'Серія', value: 'серія' }
              )
          )
          .addStringOption((option: any) =>
            option
              .setName('порядок')
              .setDescription('Порядок сортування')
              .setRequired(false)
              .addChoices(
                { name: 'За зростанням', value: 'asc' },
                { name: 'За спаданням', value: 'desc' }
              )
          );
      }
    );
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    await interaction.deferReply();

    try {
      // Отримуємо параметри пошуку
      const filters = this.extractFilters(interaction);
      
      // Отримуємо дані з Google Sheets
      const sheetData = await this.getSheetData();
      if (!sheetData || sheetData.length === 0) {
        return interaction.editReply('❌ Немає даних для пошуку');
      }

      const headers = sheetData[0] || [];
      const data = sheetData.slice(1) || [];

      // Виконуємо пошук
      const results = this.performSearch(data, headers, filters);
      
      if (results.length === 0) {
        return interaction.editReply('🔍 Нічого не знайдено за вказаними критеріями');
      }

      // Сортуємо результати
      const sortedResults = this.sortResults(results, headers, filters.sortBy, filters.sortOrder);

      // Створюємо embed з результатами
      const embed = this.createResultsEmbed(sortedResults, headers, filters);
      
      // Створюємо кнопки для навігації
      const components = this.createNavigationComponents(sortedResults.length);

      await interaction.editReply({
        embeds: [embed],
        components: components
      });

    } catch (error) {
      logger.error('Помилка покращеного пошуку', {
        error: error instanceof Error ? error.message : String(error),
        userId: interaction.user?.id,
      });
      await interaction.editReply('❌ Помилка при виконанні пошуку');
    }
  }

  /**
   * Отримання даних з Google Sheets
   */
  private async getSheetData(): Promise<string[][]> {
    // TODO: Інтеграція з Google Sheets API
    // Тимчасова реалізація з моковими даними
    return [
      ['назва', 'контрагент', 'серія', 'ціна', 'кількість'],
      ['Товар 1', 'Контрагент А', 'Серія 1', '100', '50'],
      ['Товар 2', 'Контрагент Б', 'Серія 2', '200', '30'],
      ['Товар 3', 'Контрагент А', 'Серія 1', '150', '25'],
    ];
  }

  /**
   * Витягування фільтрів з interaction
   */
  private extractFilters(interaction: ChatInputCommandInteraction): SearchFilters {
    const filters: SearchFilters = {
      name: interaction.options.getString('номенклатура') || undefined,
      client: interaction.options.getString('контрагент') || undefined,
      series: interaction.options.getString('серія') || undefined,
      priceFrom: interaction.options.getNumber('ціна_від') || undefined,
      priceTo: interaction.options.getNumber('ціна_до') || undefined,
      quantityFrom: interaction.options.getNumber('кількість_від') || undefined,
      quantityTo: interaction.options.getNumber('кількість_до') || undefined,
      sortBy: interaction.options.getString('сортування') || 'назва',
      sortOrder: interaction.options.getString('порядок') || 'asc'
    };

    return filters;
  }

  /**
   * Виконання пошуку з фільтрами
   */
  private performSearch(data: string[][], headers: string[], filters: SearchFilters): SearchResult[] {
    const results: SearchResult[] = [];

    for (const row of data) {
      let matches = true;
      let score = 0;

      // Фільтр по назві
      if (filters.name) {
        const nameIndex = this.getColumnIndex(headers, 'назва');
        if (nameIndex !== -1 && row[nameIndex]) {
          if (row[nameIndex]?.toLowerCase().includes(filters.name.toLowerCase())) {
            score += 10;
          } else {
            matches = false;
          }
        }
      }

      // Фільтр по контрагенту
      if (filters.client) {
        const clientIndex = this.getColumnIndex(headers, 'контрагент');
        if (clientIndex !== -1 && row[clientIndex]) {
          if (row[clientIndex]?.toLowerCase().includes(filters.client.toLowerCase())) {
            score += 5;
          } else {
            matches = false;
          }
        }
      }

      // Фільтр по серії
      if (filters.series) {
        const seriesIndex = this.getColumnIndex(headers, 'серія');
        if (seriesIndex !== -1 && row[seriesIndex]) {
          if (row[seriesIndex]?.toLowerCase().includes(filters.series.toLowerCase())) {
            score += 3;
          } else {
            matches = false;
          }
        }
      }

      // Фільтр по ціні
      if (filters.priceFrom || filters.priceTo) {
        const priceIndex = this.getColumnIndex(headers, 'ціна');
        if (priceIndex !== -1 && row[priceIndex]) {
          const price = parseFloat(row[priceIndex] || '0');
          if (filters.priceFrom && price < filters.priceFrom) {
            matches = false;
          }
          if (filters.priceTo && price > filters.priceTo) {
            matches = false;
          }
        }
      }

      // Фільтр по кількості
      if (filters.quantityFrom || filters.quantityTo) {
        const quantityIndex = this.getColumnIndex(headers, 'кількість');
        if (quantityIndex !== -1 && row[quantityIndex]) {
          const quantity = parseFloat(row[quantityIndex] || '0');
          if (filters.quantityFrom && quantity < filters.quantityFrom) {
            matches = false;
          }
          if (filters.quantityTo && quantity > filters.quantityTo) {
            matches = false;
          }
        }
      }

      if (matches) {
        results.push({ row, score });
      }
    }

    return results;
  }

  /**
   * Сортування результатів
   */
  private sortResults(results: SearchResult[], headers: string[], sortBy: string, sortOrder: string): SearchResult[] {
    const sortIndex = this.getColumnIndex(headers, sortBy);
    
    if (sortIndex === -1) {
      return results.sort((a, b) => b.score - a.score);
    }

    return results.sort((a, b) => {
      const aValue = parseFloat(a.row[sortIndex] || '0') || 0;
      const bValue = parseFloat(b.row[sortIndex] || '0') || 0;
      
      if (sortOrder === 'desc') {
        return bValue - aValue;
      } else {
        return aValue - bValue;
      }
    });
  }

  /**
   * Створення embed з результатами
   */
  private createResultsEmbed(results: SearchResult[], headers: string[], filters: SearchFilters): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setTitle('🔍 Результати розширеного пошуку')
      .setColor(0x00ff88)
      .setTimestamp();

    // Додаємо активні фільтри
    const activeFilters = this.getActiveFilters(filters);
    if (activeFilters.length > 0) {
      embed.addFields({
        name: '📋 Активні фільтри',
        value: activeFilters.join('\n'),
        inline: false
      });
    }

    // Додаємо результати (перші 10)
    const displayResults = results.slice(0, 10);
    let resultsText = '';

    for (let i = 0; i < displayResults.length; i++) {
      const result = displayResults[i];
      if (!result) continue;
      
      const nameIndex = this.getColumnIndex(headers, 'назва');
      const priceIndex = this.getColumnIndex(headers, 'ціна');
      const quantityIndex = this.getColumnIndex(headers, 'кількість');

      const name = nameIndex !== -1 ? result.row[nameIndex] || 'Н/Д' : 'Н/Д';
      const price = priceIndex !== -1 ? result.row[priceIndex] || 'Н/Д' : 'Н/Д';
      const quantity = quantityIndex !== -1 ? result.row[quantityIndex] || 'Н/Д' : 'Н/Д';

      resultsText += `${i + 1}. **${name}** - ${price} грн (${quantity} шт.)\n`;
    }

    if (resultsText) {
      embed.addFields({
        name: `📊 Знайдено результатів: ${results.length}`,
        value: resultsText,
        inline: false
      });
    }

    return embed;
  }

  /**
   * Створення компонентів навігації
   */
  private createNavigationComponents(_totalResults: number): ActionRowBuilder<ButtonBuilder>[] {
    const rows: ActionRowBuilder<ButtonBuilder>[] = [];

    // Перший ряд кнопок
    const row1 = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('export_results')
          .setLabel('📥 Експорт')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('refine_search')
          .setLabel('🔍 Уточнити пошук')
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId('clear_filters')
          .setLabel('🗑️ Очистити фільтри')
          .setStyle(ButtonStyle.Danger)
      );

    rows.push(row1);

    // Другий ряд - меню сортування
    const row2 = new ActionRowBuilder<StringSelectMenuBuilder>()
      .addComponents(
        new StringSelectMenuBuilder()
          .setCustomId('sort_results')
          .setPlaceholder('Сортування результатів')
          .addOptions([
            new StringSelectMenuOptionBuilder()
              .setLabel('За назвою (А-Я)')
              .setValue('назва_asc')
              .setDescription('Сортування за назвою за зростанням'),
            new StringSelectMenuOptionBuilder()
              .setLabel('За назвою (Я-А)')
              .setValue('назва_desc')
              .setDescription('Сортування за назвою за спаданням'),
            new StringSelectMenuOptionBuilder()
              .setLabel('За ціною (від дешевих)')
              .setValue('ціна_asc')
              .setDescription('Сортування за ціною за зростанням'),
            new StringSelectMenuOptionBuilder()
              .setLabel('За ціною (від дорогих)')
              .setValue('ціна_desc')
              .setDescription('Сортування за ціною за спаданням'),
            new StringSelectMenuOptionBuilder()
              .setLabel('За кількістю')
              .setValue('кількість_desc')
              .setDescription('Сортування за кількістю за спаданням')
          ])
      );

    rows.push(row2 as any);

    return rows;
  }

  /**
   * Отримання активних фільтрів
   */
  private getActiveFilters(filters: SearchFilters): string[] {
    const activeFilters: string[] = [];

    if (filters.name) {
      activeFilters.push(`📝 Назва: ${filters.name}`);
    }
    if (filters.client) {
      activeFilters.push(`👤 Контрагент: ${filters.client}`);
    }
    if (filters.series) {
      activeFilters.push(`🏷️ Серія: ${filters.series}`);
    }
    if (filters.priceFrom || filters.priceTo) {
      const priceRange = [];
      if (filters.priceFrom) priceRange.push(`від ${filters.priceFrom}`);
      if (filters.priceTo) priceRange.push(`до ${filters.priceTo}`);
      activeFilters.push(`💰 Ціна: ${priceRange.join(' ')} грн`);
    }
    if (filters.quantityFrom || filters.quantityTo) {
      const quantityRange = [];
      if (filters.quantityFrom) quantityRange.push(`від ${filters.quantityFrom}`);
      if (filters.quantityTo) quantityRange.push(`до ${filters.quantityTo}`);
      activeFilters.push(`📦 Кількість: ${quantityRange.join(' ')}`);
    }

    return activeFilters;
  }

  /**
   * Отримання індексу колонки
   */
  private getColumnIndex(headers: string[], field: string): number {
    return headers.findIndex(header => header.toLowerCase() === field.toLowerCase());
  }
} 