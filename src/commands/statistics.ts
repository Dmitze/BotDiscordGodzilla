/**
 * Команда для роботи зі статистикою та складними формулами Google Sheets
 * Підтримує підрахунок по парних/непарних стовпцях, агрегацію по аркушах
 * TypeScript версія 3.0.0
 */

import { SlashCommandBuilder, EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';

import { GoogleService } from '@/services/GoogleService';
import { sanitizeInput, validateCommandOptions } from '@/utils/security';
import logger from '@/utils/logger';
import { UIHelper } from '@/utils/uiHelpers';
import { DataFormatters } from '@/utils/formatters';

interface StatisticsConfig {
  sheets: string[];
  range: string;
  columnType: 'even' | 'odd' | 'all';
  operation: 'sum' | 'average' | 'count' | 'max' | 'min' | 'even_columns' | 'odd_columns' | 'complex_formula';
  groupBy?: string;
  filters?: Record<string, any>;
  customFormula?: string;
}

interface StatisticsResult {
  total: number;
  breakdown: Record<string, number>;
  summary: string;
  timestamp: Date;
  processingTime: number;
}

class StatisticsCommand {
  public readonly name = 'statistics';
  public readonly description = 'Отримання статистики з Google Sheets з підтримкою складних формул';
  public readonly usage = '/statistics <операція> <аркуші> [опції]';

  private readonly googleService: GoogleService | undefined;

  constructor(googleService?: GoogleService) {
    this.googleService = googleService;
  }

  /**
   * Створення команди
   */
  public getCommandData(): SlashCommandBuilder {
    return (new SlashCommandBuilder()
      .setName(this.name)
      .setDescription(this.description)
      .addStringOption(option =>
        option.setName('operation')
          .setDescription('Тип операції для статистики')
          .setRequired(true)
          .addChoices(
            { name: 'Сума', value: 'sum' },
            { name: 'Середнє', value: 'average' },
            { name: 'Кількість', value: 'count' },
            { name: 'Максимум', value: 'max' },
            { name: 'Мінімум', value: 'min' },
            { name: 'Парні стовпці', value: 'even_columns' },
            { name: 'Непарні стовпці', value: 'odd_columns' },
            { name: 'Складена формула', value: 'complex_formula' }
          )
      )
      .addStringOption(option =>
        option.setName('sheets')
          .setDescription('Аркуші для аналізу (через кому)')
          .setRequired(true)
      )
      .addStringOption(option =>
        option.setName('range')
          .setDescription('Діапазон даних (наприклад: H6:AB6)')
          .setRequired(false)
      )
      .addStringOption(option =>
        option.setName('column_type')
          .setDescription('Тип стовпців для аналізу')
          .setRequired(false)
          .addChoices(
            { name: 'Всі', value: 'all' },
            { name: 'Парні', value: 'even' },
            { name: 'Непарні', value: 'odd' }
          )
      )
      .addStringOption(option =>
        option.setName('group_by')
          .setDescription('Групування за стовпцем')
          .setRequired(false)
      )
      .addStringOption(option =>
        option.setName('filters')
          .setDescription('Фільтри у форматі JSON')
          .setRequired(false)
      )
      .addStringOption(option =>
        option.setName('custom_formula')
          .setDescription('Власна формула для аналізу')
          .setRequired(false)
      )) as unknown as SlashCommandBuilder;
  }

  /**
   * Виконання команди
   */
  public async execute(interaction: any): Promise<void> {
    const startTime = performance.now();

    try {
      logger.info('Початок виконання команди statistics', {
        user: interaction.user.tag,
        userId: interaction.user.id,
        guildId: interaction.guildId,
      });

      // Валідація опцій
      const options = this.extractOptions(interaction);
      const validation = validateCommandOptions(options, this.getValidationSchema());

      if (!validation.isValid) {
        await interaction.reply({
          content: `❌ Помилка валідації: ${validation.errors.join(', ')}`,
          ephemeral: true
        });
        return;
      }

      // Дефірування відповіді
      await interaction.deferReply();

      // Отримання статистики
      const result = await this.getStatistics(options);

      // Створення відповіді
      const embed = this.createStatisticsEmbed(result, options);
      const buttons = this.createActionButtons(result, options);

      const duration = performance.now() - startTime;
      logger.info(`Команда statistics виконана за ${duration.toFixed(2)}ms`, {
        user: interaction.user.tag,
        operation: options.operation,
        sheets: options.sheets.length,
        result: result.total,
      });

      const edit: any = { embeds: [embed] };
      if (buttons) {
        edit.components = [buttons];
      }
      await interaction.editReply(edit);

    } catch (error) {
      const duration = performance.now() - startTime;
      logger.error(`Помилка команди statistics після ${duration.toFixed(2)}ms:`, {
        type: 'command', component: 'StatisticsCommand', event: 'execute_failed',
        userId: interaction?.user?.id, guildId: interaction?.guildId,
        durationMs: Number(duration.toFixed(2)),
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      await this.handleError(interaction, error);
    }
  }

  /**
   * Витягування опцій з interaction
   */
  private extractOptions(interaction: any): StatisticsConfig {
    const operation = interaction.options.getString('operation', true);
    const sheetsInput = interaction.options.getString('sheets', true);
    const range = interaction.options.getString('range') || 'H6:AB6';
    const columnType = interaction.options.getString('column_type') as 'even' | 'odd' | 'all' || 'all';
    const groupBy = interaction.options.getString('group_by');
    const filtersInput = interaction.options.getString('filters');
    const customFormula = interaction.options.getString('custom_formula');

    // Санітизація вхідних даних
    const sheetsSan = sanitizeInput(sheetsInput, 'command');
    const sheets = (sheetsSan.sanitizedValue || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
    const filtersSan = filtersInput ? sanitizeInput(filtersInput, 'command') : undefined;
    const filters = filtersSan ? JSON.parse(filtersSan.sanitizedValue || '{}') : {};

    const base: StatisticsConfig = {
      sheets,
      range,
      columnType,
      operation: operation as any,
      groupBy,
      filters,
    };
    if (customFormula) {
      const cf = sanitizeInput(customFormula, 'command').sanitizedValue;
      if (cf) {
        // Only assign when non-empty to satisfy exactOptionalPropertyTypes
        (base as any).customFormula = cf;
      }
    }
    return base;
  }

  /**
   * Схема валідації
   */
  private getValidationSchema(): Record<string, any> {
    return {
      sheets: {
        required: true,
        type: 'object',
        minLength: 1,
      },
      range: {
        required: true,
        type: 'string',
        pattern: /^[A-Z]+\d+:[A-Z]+\d+$/,
      },
      operation: {
        required: true,
        type: 'string',
        enum: ['sum', 'average', 'count', 'max', 'min', 'even_columns', 'odd_columns', 'complex_formula'],
      },
    };
  }

  /**
   * Отримання статистики
   */
  private async getStatistics(config: StatisticsConfig): Promise<StatisticsResult> {
    const startTime = performance.now();

    try {
      logger.debug('Початок отримання статистики', { config });

      let total = 0;
      const breakdown: Record<string, number> = {};

      // Обробка різних типів операцій
      switch (config.operation) {
        case 'even_columns':
        case 'odd_columns':
          total = await this.calculateColumnStatistics(config, config.operation === 'even_columns');
          break;

        case 'complex_formula':
          total = await this.executeComplexFormula(config);
          break;

        default:
          total = await this.calculateBasicStatistics(config);
          break;
      }

      // Групування результатів
      if (config.groupBy) {
        breakdown[config.groupBy] = total;
      } else {
        breakdown['Загальна сума'] = total;
      }

      const processingTime = performance.now() - startTime;

      return {
        total,
        breakdown,
        summary: this.generateSummary(total, config),
        timestamp: new Date(),
        processingTime,
      };

    } catch (error) {
      logger.error('Помилка отримання статистики:', {
        type: 'command', component: 'StatisticsCommand', event: 'get_statistics_failed',
        operation: config.operation, sheets: config.sheets, range: config.range,
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Розрахунок статистики по парних/непарних стовпцях
   */
  private async calculateColumnStatistics(config: StatisticsConfig, isEven: boolean): Promise<number> {
    let total = 0;

    for (const sheetName of config.sheets) {
      try {
        if (!this.googleService) throw new Error('GoogleService не ініціалізовано');
        const data = await this.googleService.getSheetData(sheetName, config.range);

        if (!data || !data.values || data.values.length === 0) {
          logger.warn(`Немає даних в аркуші ${sheetName}`);
          continue;
        }

        const row: string[] = (data.values[0] ?? []) as string[]; // Перший рядок
        const [startRef, endRef] = config.range.includes(':')
          ? (config.range.split(':') as [string, string])
          : ([config.range, config.range] as [string, string]);
        const startCol = this.getColumnIndex(startRef);
        const endCol = this.getColumnIndex(endRef);

        for (let col = startCol; col <= endCol; col++) {
          const isEvenColumn = col % 2 === 0;

          if (isEven ? isEvenColumn : !isEvenColumn) {
            const value = parseFloat(row[col - startCol] || '0');
            if (!isNaN(value)) {
              total += value;
            }
          }
        }

        logger.debug(`Оброблено аркуш ${sheetName}`, { type: 'command', component: 'StatisticsCommand', event: 'sheet_processed', sheetName, total, isEven });

      } catch (error) {
        logger.error(`Помилка обробки аркуша ${sheetName}:`, {
          type: 'command', component: 'StatisticsCommand', event: 'sheet_process_failed', sheetName,
          errorName: error instanceof Error ? error.name : undefined,
          errorMessage: error instanceof Error ? error.message : String(error),
          stack: error instanceof Error ? error.stack : undefined,
        });
      }
    }

    return total;
  }

  /**
   * Виконання складних формул
   */
  private async executeComplexFormula(config: StatisticsConfig): Promise<number> {
    if (!config.customFormula) {
      throw new Error('Власна формула не надана');
    }
    // Функціонал складних формул наразі недоступний у сервісах. Лише валідуємо та відхиляємо.
    logger.warn('Виконання складної формули наразі не підтримується', { type: 'command', component: 'StatisticsCommand', event: 'complex_formula_unsupported' });
    throw new Error('Складні формули тимчасово недоступні');
  }

  /**
   * Розрахунок базової статистики
   */
  private async calculateBasicStatistics(config: StatisticsConfig): Promise<number> {
    let total = 0;
    let count = 0;

    for (const sheetName of config.sheets) {
      try {
        if (!this.googleService) throw new Error('GoogleService не ініціалізовано');
        const data = await this.googleService.getSheetData(sheetName, config.range);

        if (!data || !data.values) continue;

        for (const row of data.values) {
          for (const cell of row) {
            const value = parseFloat(cell || '0');
            if (!isNaN(value)) {
              switch (config.operation) {
                case 'sum':
                  total += value;
                  break;
                case 'average':
                  total += value;
                  count++;
                  break;
                case 'count':
                  if (value > 0) count++;
                  break;
                case 'max':
                  total = Math.max(total, value);
                  break;
                case 'min':
                  total = total === 0 ? value : Math.min(total, value);
                  break;
              }
            }
          }
        }

      } catch (error) {
        logger.error(`Помилка обробки аркуша ${sheetName}:`, {
          type: 'command', component: 'StatisticsCommand', event: 'sheet_process_failed', sheetName,
          errorName: error instanceof Error ? error.name : undefined,
          errorMessage: error instanceof Error ? error.message : String(error),
          stack: error instanceof Error ? error.stack : undefined,
        });
      }
    }

    return config.operation === 'average' ? (count > 0 ? total / count : 0) :
           config.operation === 'count' ? count : total;
  }

  /**
   * Отримання індексу стовпця
   */
  private getColumnIndex(column: string): number {
    const letters = column.replace(/[^A-Za-z]/g, '').toUpperCase();
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index;
  }

  /**
   * Генерація підсумку
   */
  private generateSummary(total: number, config: StatisticsConfig): string {
    const operationNames = {
      sum: 'сума',
      average: 'середнє',
      count: 'кількість',
      max: 'максимум',
      min: 'мінімум',
      even_columns: 'сума парних стовпців',
      odd_columns: 'сума непарних стовпців',
      complex_formula: 'результат формули',
    };

    return `**${operationNames[config.operation as keyof typeof operationNames]}**: ${DataFormatters.formatNumber(total)}`;
  }

  /**
   * Створення embed для відповіді
   */
  private createStatisticsEmbed(result: StatisticsResult, config: StatisticsConfig): EmbedBuilder {
    const embed = UIHelper.createBaseEmbed('📊 Статистика Google Sheets', '')
      .setColor('#00ff00')
      .setTimestamp(result.timestamp);

    // Основна інформація
    embed.addFields(
      { name: '📈 Результат', value: result.summary, inline: true },
      { name: '⏱️ Час обробки', value: `${result.processingTime.toFixed(2)}ms`, inline: true },
      { name: '📋 Аркуші', value: config.sheets.length.toString(), inline: true }
    );

    // Детальна розбивка
    if (Object.keys(result.breakdown).length > 1) {
      const breakdownText = Object.entries(result.breakdown)
        .map(([key, value]) => `**${key}**: ${DataFormatters.formatNumber(value)}`)
        .join('\n');

      embed.addFields({ name: '📊 Детальна розбивка', value: breakdownText });
    }

    // Додаткова інформація
    embed.addFields(
      { name: '🔧 Операція', value: config.operation, inline: true },
      { name: '📏 Діапазон', value: config.range, inline: true },
      { name: '📊 Тип стовпців', value: config.columnType, inline: true }
    );

    // Фільтри
    if (config.filters && Object.keys(config.filters).length > 0) {
      const filtersText = Object.entries(config.filters)
        .map(([key, value]) => `**${key}**: ${value}`)
        .join('\n');

      embed.addFields({ name: '🔍 Фільтри', value: filtersText });
    }

    return embed;
  }

  /**
   * Створення кнопок дій
   */
  private createActionButtons(_result: StatisticsResult, _config: StatisticsConfig): ActionRowBuilder<ButtonBuilder> | null {
    const row = new ActionRowBuilder<ButtonBuilder>();

    // Кнопка експорту
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(`export_stats_${Date.now()}`)
        .setLabel('📊 Експорт')
        .setStyle(ButtonStyle.Primary)
    );

    // Кнопка детального аналізу
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(`analyze_stats_${Date.now()}`)
        .setLabel('🔍 Аналіз')
        .setStyle(ButtonStyle.Secondary)
    );

    // Кнопка оновлення
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(`refresh_stats_${Date.now()}`)
        .setLabel('🔄 Оновити')
        .setStyle(ButtonStyle.Success)
    );

    return row;
  }

  /**
   * Обробка помилок
   */
  private async handleError(interaction: any, error: unknown): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

    try {
      if (interaction.deferred) {
        await interaction.editReply({
          content: `❌ Помилка отримання статистики: ${errorMessage}`,
        });
      } else {
        await interaction.reply({
          content: `❌ Помилка отримання статистики: ${errorMessage}`,
          ephemeral: true,
        });
      }
    } catch (replyError) {
      logger.error('Помилка відповіді на помилку', {
        type: 'command',
        command: this.name,
        component: 'StatisticsCommand.handleError',
        error: replyError instanceof Error ? replyError.message : String(replyError),
        stack: replyError instanceof Error ? replyError.stack : undefined,
      });
    }
  }

  /**
   * Отримання назви команди
   */
  public getName(): string {
    return this.name;
  }

  /**
   * Отримання опису команди
   */
  public getDescription(): string {
    return this.description;
  }
}

export default StatisticsCommand; 