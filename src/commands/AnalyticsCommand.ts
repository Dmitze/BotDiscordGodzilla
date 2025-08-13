/**
 * 📊 Команди аналітики та звітності
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

export class AnalyticsCommand extends BaseCommand {
  constructor(config: BotConfig) {
    // Назва та опис згідно з тестами
    super('аналітика', 'Аналітика та звітність', config, {}, (builder: any) => {
      return builder
        // підкоманда: звіт
        .addSubcommand((sub: any) =>
          sub
            .setName('звіт')
            .setDescription('Генерація звітів')
            .addStringOption((opt: any) =>
              opt
                .setName('тип')
                .setDescription('Тип звіту')
                .setRequired(true)
                .addChoices(
                  { name: 'Щоденний звіт', value: 'daily' },
                  { name: 'Тижневий звіт', value: 'weekly' },
                  { name: 'Місячний звіт', value: 'monthly' }
                )
            )
            .addStringOption((opt: any) =>
              opt
                .setName('формат')
                .setDescription('Формат звіту')
                .setRequired(false)
                .addChoices(
                  { name: 'Текстовий', value: 'text' },
                  { name: 'Excel', value: 'excel' },
                  { name: 'PDF', value: 'pdf' }
                )
            )
        )
        // підкоманда: статистика
        .addSubcommand((sub: any) =>
          sub
            .setName('статистика')
            .setDescription('Статистика та метрики')
            .addStringOption((opt: any) =>
              opt
                .setName('категорія')
                .setDescription('Категорія статистики')
                .setRequired(true)
                .addChoices(
                  { name: 'Загальна статистика', value: 'general' },
                  { name: 'Бойова готовність', value: 'combat' }
                )
            )
        )
        // підкоманда: тренди
        .addSubcommand((sub: any) =>
          sub
            .setName('тренди')
            .setDescription('Тренди використання')
            .addStringOption((opt: any) =>
              opt
                .setName('період')
                .setDescription('Період (напр. 7d, 30d)')
                .setRequired(true)
            )
        )
        // підкоманда: інсайти
        .addSubcommand((sub: any) =>
          sub.setName('інсайти').setDescription('Корисні інсайти та рекомендації')
        );
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      const sub = interaction.options.getSubcommand();

      switch (sub) {
        case 'звіт':
          await this.handleReport(interaction);
          break;
        case 'статистика':
          await this.handleStatistics(interaction);
          break;
        case 'тренди':
          await this.handleTrends(interaction);
          break;
        case 'інсайти':
          await this.handleInsights(interaction);
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
    const type = interaction.options.getString('тип', true);
    const format = interaction.options.getString('формат') || 'text';

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
    const category = interaction.options.getString('категорія', true);
    try {
      const analyticsService = interaction.client.serviceContainer.get('AnalyticsService');
      const stats = await analyticsService.getStatistics(category);
      await interaction.reply({ content: `📈 Статистика (${this.getCategoryName(category)}): ${JSON.stringify(stats)}` });
    } catch (error) {
      await interaction.reply({ content: '❌ Помилка отримання статистики', ephemeral: true });
    }
  }

  private async handleTrends(interaction: any): Promise<void> {
    const period = interaction.options.getString('період', true);
    try {
      const analyticsService = interaction.client.serviceContainer.get('AnalyticsService');
      const trends = await analyticsService.getTrends(period);
      await interaction.reply({ content: `📊 Тренди (${period}): ${JSON.stringify(trends)}` });
    } catch (error) {
      await interaction.reply({ content: '❌ Помилка отримання трендів', ephemeral: true });
    }
  }

  private async handleInsights(interaction: any): Promise<void> {
    try {
      const analyticsService = interaction.client.serviceContainer.get('AnalyticsService');
      const data = await analyticsService.getInsights();
      await interaction.reply({ content: `💡 Інсайти: ${JSON.stringify(data)}` });
    } catch (error) {
      await interaction.reply({ content: '❌ Помилка отримання інсайтів', ephemeral: true });
    }
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
}
