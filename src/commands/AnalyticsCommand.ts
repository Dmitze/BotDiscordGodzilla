/**
 * 📊 Команди аналітики та звітності
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import { EmbedBuilder, type ChatInputCommandInteraction } from 'discord.js';
import { replyWithPrivacy } from '@/ui/reply';

export class AnalyticsCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('аналітика', 'Аналітика та звітність', config, {}, (builder) => {
      builder
        .addSubcommand((subcommand) =>
          subcommand
            .setName('звіт')
            .setDescription('📋 Генерація звітів')
            .addStringOption((option) =>
              option
                .setName('тип')
                .setDescription('Тип звіту')
                .setRequired(true)
                .addChoices(
                  { name: 'Щоденний звіт', value: 'daily' },
                  { name: 'Тижневий звіт', value: 'weekly' },
                  { name: 'Місячний звіт', value: 'monthly' }
                )
            )
            .addStringOption((option) =>
              option
                .setName('формат')
                .setDescription('Формат звіту')
                .setRequired(false)
                .addChoices(
                  { name: 'Текстовий', value: 'text' },
                  { name: 'Excel', value: 'excel' },
                  { name: 'PDF', value: 'pdf' }
                )
            )
        );
      builder
        .addSubcommand((subcommand) =>
          subcommand
            .setName('статистика')
            .setDescription('📈 Статистика та метрики')
            .addStringOption((option) =>
              option
                .setName('категорія')
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
        );
      builder
        .addSubcommand((subcommand) =>
          subcommand
            .setName('тренди')
            .setDescription('📊 Тренди за період')
            .addStringOption((option) =>
              option
                .setName('період')
                .setDescription('Період трендів (напр. 7d, 30d)')
                .setRequired(true)
            )
        );
      builder
        .addSubcommand((subcommand) =>
          subcommand
            .setName('інсайти')
            .setDescription('💡 Інсайти та рекомендації')
        );
      return builder;
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      const subcommand = interaction.options.getSubcommand();

      switch (subcommand) {
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
          await replyWithPrivacy(interaction, { content: '❌ Невідома підкоманда' });
      }
    } catch (error) {
      logger.error('Помилка виконання команди аналітики', {
        type: 'command',
        component: 'аналітика',
        event: 'execute_error',
        errorMessage: error instanceof Error ? error.message : String(error),
      });
      await replyWithPrivacy(interaction, { content: '❌ Помилка аналітики' });
    }
  }

  private async handleReport(interaction: ChatInputCommandInteraction): Promise<void> {
    const type = interaction.options.getString('тип', true);
    const format = interaction.options.getString('формат') || 'text';

    try {
      const analyticsService = (interaction.client as any)?.serviceContainer?.get('AnalyticsService');
      const report = await analyticsService.generateReport(type, format);

      if (!report || !report.data || Object.keys(report.data).length === 0) {
        await replyWithPrivacy(interaction, { content: '⚠️ Дані для звіту відсутні' });
        return;
      }

      let content = `✅ Звіт згенеровано. Тип: ${this.getReportTypeName(type)}. Формат: ${format}.`;
      if (report.exportUrl) {
        content += `\nПосилання на експорт: ${report.exportUrl}`;
      }
      await interaction.reply({ content });
    } catch (error) {
      await replyWithPrivacy(interaction, { content: '❌ Помилка генерації звіту' });
    }
  }

  private async handleStatistics(interaction: ChatInputCommandInteraction): Promise<void> {
    const category = interaction.options.getString('категорія');

    const embed = new EmbedBuilder()
      .setTitle('📈 Статистика та метрики')
      .setColor(0xff6b6b)
      .setTimestamp();

    const categoryName = this.getCategoryName(category);

    embed.setDescription(`**${categoryName}**`);
    try {
      const analyticsService = (interaction.client as any)?.serviceContainer?.get('AnalyticsService');
      const stats = await analyticsService.getStatistics(category);
      // Додаємо кілька базових полів, якщо сервіс повернув дані
      if (stats) {
        if (typeof stats.totalUsers !== 'undefined') embed.addFields({ name: 'Користувачі', value: String(stats.totalUsers), inline: true });
        if (typeof stats.totalCommands !== 'undefined') embed.addFields({ name: 'Команди', value: String(stats.totalCommands), inline: true });
      }
      await interaction.reply({ embeds: [embed] });
    } catch {
      await replyWithPrivacy(interaction, { content: '❌ Помилка отримання статистики' });
    }
  }

  private async handleTrends(interaction: ChatInputCommandInteraction): Promise<void> {
    const period = interaction.options.getString('період');
    try {
      const analyticsService = (interaction.client as any)?.serviceContainer?.get('AnalyticsService');
      const trends = await analyticsService.getTrends(period);
      await interaction.reply({ content: `✅ Тренди за період ${period}: ${trends?.trends?.length ?? 0}` });
    } catch {
      await replyWithPrivacy(interaction, { content: '❌ Помилка отримання трендів' });
    }
  }

  private async handleInsights(interaction: ChatInputCommandInteraction): Promise<void> {
    try {
      const analyticsService = (interaction.client as any)?.serviceContainer?.get('AnalyticsService');
      const data = await analyticsService.getInsights();
      const msgs: string[] = [];
      if (data?.insights?.length) msgs.push('• ' + data.insights.join('\n• '));
      if (data?.recommendations?.length) msgs.push('\nРекомендації:\n• ' + data.recommendations.join('\n• '));
      await interaction.reply({ content: msgs.join('\n') || 'Немає інсайтів' });
    } catch {
      await replyWithPrivacy(interaction, { content: '❌ Помилка отримання інсайтів' });
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
      personnel: 'Особовий склад',
      equipment: 'Техніка',
      operations: 'Операції',
      materials: 'МТЗ',
      efficiency: 'Ефективність',
    };
    return map[category] || category;
  }
}
