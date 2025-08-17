/**
 * 📊 Команди аналітики та звітності
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import { EmbedBuilder } from 'discord.js';

export class AnalyticsCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('analytics', '📊 Аналітика та звітність ЗСУ', config, {}, (builder: any) => {
      return builder
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('report')
            .setDescription('📋 Генерація звітів')
            .addStringOption((option: any) =>
              option
                .setName('type')
                .setDescription('Тип звіту')
                .setRequired(true)
                .addChoices(
                  { name: 'Щоденний звіт', value: 'daily' },
                  { name: 'Тижневий звіт', value: 'weekly' },
                  { name: 'Місячний звіт', value: 'monthly' },
                  { name: 'Звіт по особовому складу', value: 'personnel' },
                  { name: 'Звіт по техніці', value: 'equipment' },
                  { name: 'Звіт по операціях', value: 'operations' },
                  { name: 'Звіт по МТЗ', value: 'materials' },
                  { name: 'Звіт по наказах', value: 'orders' }
                )
            )
            .addStringOption((option: any) =>
              option
                .setName('format')
                .setDescription('Формат звіту')
                .setRequired(false)
                .addChoices(
                  { name: 'Текстовий', value: 'text' },
                  { name: 'Excel', value: 'excel' },
                  { name: 'PDF', value: 'pdf' }
                )
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('stats')
            .setDescription('📈 Статистика та метрики')
            .addStringOption((option: any) =>
              option
                .setName('category')
                .setDescription('Категорія статистики')
                .setRequired(true)
                .addChoices(
                  { name: 'Загальна статистика', value: 'general' },
                  { name: 'Бойова готовність', value: 'combat' },
                  { name: 'Особовий склад', value: 'personnel' },
                  { name: 'Техніка', value: 'equipment' },
                  { name: 'Операції', value: 'operations' },
                  { name: 'МТЗ', value: 'materials' },
                  { name: 'Ефективність', value: 'efficiency' }
                )
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('forecast')
            .setDescription('🔮 Прогнозування та планування')
            .addStringOption((option: any) =>
              option
                .setName('type')
                .setDescription('Тип прогнозу')
                .setRequired(true)
                .addChoices(
                  { name: 'Потреби в МТЗ', value: 'materials' },
                  { name: 'Ремонт техніки', value: 'repairs' },
                  { name: 'Особовий склад', value: 'personnel' },
                  { name: 'Оперативні потреби', value: 'operations' },
                  { name: 'Бюджет', value: 'budget' }
                )
            )
            .addIntegerOption((option: any) =>
              option
                .setName('period')
                .setDescription('Період прогнозування (днів)')
                .setRequired(false)
                .setMinValue(1)
                .setMaxValue(365)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('compare')
            .setDescription('⚖️ Порівняльний аналіз')
            .addStringOption((option: any) =>
              option
                .setName('object')
                .setDescription("Об'єкт порівняння")
                .setRequired(true)
                .addChoices(
                  { name: 'Підрозділи', value: 'units' },
                  { name: 'Періоди', value: 'periods' },
                  { name: 'Показники', value: 'metrics' },
                  { name: 'Регіони', value: 'regions' }
                )
            )
            .addStringOption((option: any) =>
              option
                .setName('metric')
                .setDescription('Метрика для порівняння')
                .setRequired(true)
                .addChoices(
                  { name: 'Ефективність', value: 'efficiency' },
                  { name: 'Витрати', value: 'costs' },
                  { name: 'Результати', value: 'results' },
                  { name: 'Час виконання', value: 'time' }
                )
            )
            .addIntegerOption((option: any) =>
              option
                .setName('period')
                .setDescription('Період аналізу (днів)')
                .setRequired(false)
                .setMinValue(1)
                .setMaxValue(365)
            )
        );
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      const subcommand = interaction.options.getSubcommand();

      switch (subcommand) {
        case 'report':
          await this.handleReport(interaction);
          break;
        case 'stats':
          await this.handleStatistics(interaction);
          break;
        case 'forecast':
          await this.handleForecast(interaction);
          break;
        case 'compare':
          await this.handleComparison(interaction);
          break;
        default:
          await interaction.reply({ content: '❌ Невідома підкоманда', ephemeral: true });
      }
    } catch (error) {
      logger.error('Помилка виконання команди аналітики', {
        type: 'command',
        component: 'аналітика',
        event: 'execute_error',
        errorMessage: error instanceof Error ? error.message : String(error),
      });
      await interaction.reply({ content: '❌ Помилка аналітики', ephemeral: true });
    }
  }

  private async handleReport(interaction: any): Promise<void> {
    const type = interaction.options.getString('type', true);
    const format = interaction.options.getString('format') || 'text';

    try {
      const analyticsService = interaction.client.serviceContainer.get('AnalyticsService');
      const report = await analyticsService.generateReport(type, format);

      if (!report || !report.data || Object.keys(report.data).length === 0) {
        await interaction.reply({ content: '⚠️ Дані для звіту відсутні', ephemeral: true });
        return;
      }

      let content = `✅ Звіт згенеровано. Тип: ${this.getReportTypeName(type)}. Формат: ${format}.`;
      if (report.exportUrl) {
        content += `\nПосилання на експорт: ${report.exportUrl}`;
      }
      await interaction.reply({ content });
    } catch (error) {
      await interaction.reply({ content: '❌ Помилка генерації звіту', ephemeral: true });
    }
  }

  private async handleStatistics(interaction: any): Promise<void> {
    const category = interaction.options.getString('category', true);

    const embed = new EmbedBuilder()
      .setTitle('📈 Статистика та метрики')
      .setColor(0xff6b6b)
      .setTimestamp();

    const categoryName = this.getCategoryName(category);

    embed.setDescription(`**${categoryName}**`);

    switch (category) {
      case 'general':
        embed.addFields(
          { name: 'Загальна чисельність', value: '1,250', inline: true },
          { name: 'Бойова готовність', value: '95%', inline: true },
          { name: 'Техніка в строю', value: '87%', inline: true }
        );
        break;
      case 'combat':
        embed.addFields(
          { name: 'Бойова готовність', value: '95%', inline: true },
          { name: 'Готовність до виконання', value: '92%', inline: true },
          { name: 'Забезпеченість', value: '88%', inline: true }
        );
        break;
      case 'personnel':
        embed.addFields(
          { name: 'Особовий склад', value: '1,250', inline: true },
          { name: 'Офіцери', value: '150', inline: true },
          { name: 'Сержанти', value: '300', inline: true }
        );
        break;
      case 'equipment':
        embed.addFields(
          { name: 'Техніка в строю', value: '87%', inline: true },
          { name: 'На ремонті', value: '8%', inline: true },
          { name: 'Резерв', value: '5%', inline: true }
        );
        break;
      case 'operations':
        embed.addFields(
          { name: 'Активні операції', value: '5', inline: true },
          { name: 'Завершені операції', value: '12', inline: true },
          { name: 'Успішність', value: '94%', inline: true }
        );
        break;
      case 'materials':
        embed.addFields(
          { name: 'МТЗ в наявності', value: '85%', inline: true },
          { name: 'Потреби', value: '15%', inline: true },
          { name: 'Поставки', value: 'В процесі', inline: true }
        );
        break;
      case 'efficiency':
        embed.addFields(
          { name: 'Загальна ефективність', value: '92%', inline: true },
          { name: 'Оперативна ефективність', value: '89%', inline: true },
          { name: 'Логістична ефективність', value: '94%', inline: true }
        );
        break;
      default:
        embed.addFields({ name: 'Дані', value: 'Недоступні', inline: false });
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка прогнозування
   */
  private async handleForecast(interaction: any): Promise<void> {
    const type = interaction.options.getString('type', true);
    const period = interaction.options.getInteger('period') || 30;

    const embed = new EmbedBuilder()
      .setTitle('🔮 Прогнозування та планування')
      .setColor(0x9932cc)
      .setTimestamp();

    const forecastTypeName = this.getForecastTypeName(type);

    embed.setDescription(`**${forecastTypeName}**\n\nПеріод: ${period} днів`);
    embed.addFields(
      { name: 'Прогнозована потреба', value: '+15%', inline: true },
      { name: 'Рекомендації', value: 'Збільшити поставки', inline: true },
      { name: 'Впевненість', value: '85%', inline: true }
    );

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка порівняльного аналізу
   */
  private async handleComparison(interaction: any): Promise<void> {
    const object = interaction.options.getString('object', true);
    const metric = interaction.options.getString('metric', true);
    const period = interaction.options.getInteger('period') || 30;

    const embed = new EmbedBuilder()
      .setTitle('⚖️ Порівняльний аналіз')
      .setColor(0xff9900)
      .setTimestamp();

    const objectName = this.getObjectName(object);
    const metricName = this.getMetricName(metric);

    embed.setDescription(`**${objectName}**\n\nМетрика: ${metricName}\nПеріод: ${period} днів`);
    embed.addFields(
      { name: 'Середнє значення', value: '85%', inline: true },
      { name: 'Максимум', value: '95%', inline: true },
      { name: 'Мінімум', value: '75%', inline: true }
    );

    await interaction.reply({ embeds: [embed] });
  }

  private getReportTypeName(type: string): string {
    const map: Record<string, string> = {
      daily: 'Щоденний звіт',
      weekly: 'Тижневий звіт',
      monthly: 'Місячний звіт',
    };
    return map[type] || type;
  }

  private getCategoryName(category: string): string {
    const map: Record<string, string> = {
      general: 'Загальна статистика',
      combat: 'Бойова готовність',
    };
    return map[category] || category;
  }

  /**
   * Отримання назви типу прогнозу
   */
  private getForecastTypeName(type: string): string {
    const typeNames: Record<string, string> = {
      materials: 'Потреби в МТЗ',
      repairs: 'Ремонт техніки',
      personnel: 'Особовий склад',
      operations: 'Оперативні потреби',
      budget: 'Бюджет',
    };

    return typeNames[type] || type;
  }

  /**
   * Отримання назви об'єкта
   */
  private getObjectName(object: string): string {
    const objectNames: Record<string, string> = {
      units: 'Підрозділи',
      periods: 'Періоди',
      metrics: 'Показники',
      regions: 'Регіони',
    };

    return objectNames[object] || object;
  }

  /**
   * Отримання назви метрики
   */
  private getMetricName(metric: string): string {
    const metricNames: Record<string, string> = {
      efficiency: 'Ефективність',
      costs: 'Витрати',
      results: 'Результати',
      time: 'Час виконання',
    };

    return metricNames[metric] || metric;
  }
}
