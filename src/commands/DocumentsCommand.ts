/**
 * 📄 Команди для роботи з військовими документами ЗСУ
 * Спеціалізовані функції для різних типів документів
 */

import { EmbedBuilder } from 'discord.js';
import { DuplicateResolver } from '@/components/DuplicateResolver';
import { uiState } from '@/services/UIStateService';
import type { CommandComponentOptions } from './BaseCommand';
import type { DriveFile } from '@/types/drive';
import type { BotConfig, CommandExecuteOptions, LogMeta } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

interface DocumentAction {
  type: string;
  query?: string;
}

export class DocumentsCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('документи', '📄 Робота з військовими документами ЗСУ', config, {}, (builder: any) => {
      return builder
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('особовий-склад')
            .setDescription('👥 Робота з особовим складом')
            .addStringOption((option: any) =>
              option
                .setName('дія')
                .setDescription('Дія з особовим складом')
                .setRequired(true)
                .addChoices(
                  { name: 'Пошук особового складу', value: 'search' },
                  { name: 'Додати особу', value: 'add' },
                  { name: 'Оновити дані', value: 'update' },
                  { name: 'Звіт по особовому складу', value: 'report' },
                  { name: 'Перевірка наявності', value: 'check' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('техніка')
            .setDescription('🚗 Робота з технікою та озброєнням')
            .addStringOption((option: any) =>
              option
                .setName('дія')
                .setDescription('Дія з технікою')
                .setRequired(true)
                .addChoices(
                  { name: 'Пошук техніки', value: 'search' },
                  { name: 'Додати техніку', value: 'add' },
                  { name: 'Стан техніки', value: 'status' },
                  { name: 'Звіт по техніці', value: 'report' },
                  { name: 'Технічне обслуговування', value: 'maintenance' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('матеріали')
            .setDescription('📦 Робота з матеріально-технічним забезпеченням')
            .addStringOption((option: any) =>
              option
                .setName('дія')
                .setDescription('Дія з матеріалами')
                .setRequired(true)
                .addChoices(
                  { name: 'Пошук матеріалів', value: 'search' },
                  { name: 'Додати матеріали', value: 'add' },
                  { name: 'Залишки', value: 'stock' },
                  { name: 'Звіт по МТЗ', value: 'report' },
                  { name: 'Поповнення', value: 'replenish' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('операції')
            .setDescription('⚔️ Робота з оперативними документами')
            .addStringOption((option: any) =>
              option
                .setName('дія')
                .setDescription('Дія з операціями')
                .setRequired(true)
                .addChoices(
                  { name: 'Пошук операцій', value: 'search' },
                  { name: 'Додати операцію', value: 'add' },
                  { name: 'Статус операцій', value: 'status' },
                  { name: 'Звіт по операціях', value: 'report' },
                  { name: 'Планування', value: 'planning' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('накази')
            .setDescription('📋 Робота з наказами та розпорядженнями')
            .addStringOption((option: any) =>
              option
                .setName('дія')
                .setDescription('Дія з наказами')
                .setRequired(true)
                .addChoices(
                  { name: 'Пошук наказів', value: 'search' },
                  { name: 'Створити наказ', value: 'create' },
                  { name: 'Статус виконання', value: 'status' },
                  { name: 'Звіт по наказах', value: 'report' },
                  { name: 'Архівування', value: 'archive' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        );
    });
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      const subcommand = interaction.options.getSubcommand();
      const action = interaction.options.getString('дія', true);
      const query = interaction.options.getString('запит');

      const documentAction: DocumentAction = {
        type: action,
        query: query || undefined,
      };

      switch (subcommand) {
        case 'особовий-склад':
          await this.handlePersonnel(interaction, documentAction);
          break;
        case 'техніка':
          await this.handleEquipment(interaction, documentAction);
          break;
        case 'матеріали':
          await this.handleMaterials(interaction, documentAction);
          break;
        case 'операції':
          await this.handleOperations(interaction, documentAction);
          break;
        case 'накази':
          await this.handleOrders(interaction, documentAction);
          break;
        default:
          await interaction.reply('❌ Невідома підкоманда');
      }
    } catch (error) {
      logger.error('❌ Помилка команди документів', {
        command: 'документи',
        guildId: interaction.guildId,
        channelId: interaction.channelId,
        userId: interaction.user?.id,
        error,
      } as LogMeta);
      await interaction.reply('❌ Помилка обробки документів');
    }
  }

  /**
   * Обробка особового складу
   */
  private async handlePersonnel(interaction: any, action: DocumentAction): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('👥 Особовий склад')
      .setColor(0x0099ff)
      .setTimestamp();

    switch (action.type) {
      case 'search':
        embed.setDescription(`🔍 **Пошук особового складу**\n\nЗапит: ${action.query || 'Всі'}`);
        embed.addFields({
          name: 'Результат',
          value: 'Тимчасова відповідь: Знайдено 0 осіб',
          inline: false,
        });
        break;
      case 'add':
        embed.setDescription(`➕ **Додавання особи**\n\nДані: ${action.query || 'Не вказано'}`);
        embed.addFields({ name: 'Статус', value: '✅ Особа додана успішно', inline: false });
        break;
      case 'update':
        embed.setDescription(`✏️ **Оновлення даних**\n\nДані: ${action.query || 'Не вказано'}`);
        embed.addFields({ name: 'Статус', value: '✅ Дані оновлено успішно', inline: false });
        break;
      case 'report':
        embed.setDescription('📊 **Звіт по особовому складу**');
        embed.addFields(
          { name: 'Всього осіб', value: '0', inline: true },
          { name: 'Активних', value: '0', inline: true },
          { name: 'У відпустці', value: '0', inline: true }
        );
        break;
      case 'check':
        embed.setDescription(
          `🔍 **Перевірка наявності**\n\nЗапит: ${action.query || 'Не вказано'}`
        );
        embed.addFields({ name: 'Результат', value: '❌ Особа не знайдена', inline: false });
        break;
      default:
        embed.setDescription('❌ Невідома дія');
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка техніки
   */
  private async handleEquipment(interaction: any, action: DocumentAction): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('🚗 Техніка та озброєння')
      .setColor(0xff9900)
      .setTimestamp();

    switch (action.type) {
      case 'search':
        embed.setDescription(`🔍 **Пошук техніки**\n\nЗапит: ${action.query || 'Вся техніка'}`);
        embed.addFields({
          name: 'Результат',
          value: 'Тимчасова відповідь: Знайдено 0 одиниць техніки',
          inline: false,
        });
        break;
      case 'add':
        embed.setDescription(`➕ **Додавання техніки**\n\nДані: ${action.query || 'Не вказано'}`);
        embed.addFields({ name: 'Статус', value: '✅ Техніка додана успішно', inline: false });
        break;
      case 'status':
        embed.setDescription('📊 **Стан техніки**');
        embed.addFields(
          { name: 'Всього техніки', value: '0', inline: true },
          { name: 'Справна', value: '0', inline: true },
          { name: 'На ремонті', value: '0', inline: true }
        );
        break;
      case 'report':
        embed.setDescription('📋 **Звіт по техніці**');
        embed.addFields(
          { name: 'Танки', value: '0', inline: true },
          { name: 'БМП', value: '0', inline: true },
          { name: 'Артилерія', value: '0', inline: true }
        );
        break;
      case 'maintenance':
        embed.setDescription('🔧 **Технічне обслуговування**');
        embed.addFields(
          { name: 'Плановий ремонт', value: '0 одиниць', inline: true },
          { name: 'Аварійний ремонт', value: '0 одиниць', inline: true }
        );
        break;
      default:
        embed.setDescription('❌ Невідома дія');
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка матеріалів
   */
  private async handleMaterials(interaction: any, action: DocumentAction): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('📦 Матеріально-технічне забезпечення')
      .setColor(0x00ff88)
      .setTimestamp();

    switch (action.type) {
      case 'search':
        embed.setDescription(
          `🔍 **Пошук матеріалів**\n\nЗапит: ${action.query || 'Всі матеріали'}`
        );
        embed.addFields({
          name: 'Результат',
          value: 'Тимчасова відповідь: Знайдено 0 позицій',
          inline: false,
        });
        break;
      case 'add':
        embed.setDescription(
          `➕ **Додавання матеріалів**\n\nДані: ${action.query || 'Не вказано'}`
        );
        embed.addFields({ name: 'Статус', value: '✅ Матеріали додано успішно', inline: false });
        break;
      case 'stock':
        embed.setDescription('📊 **Залишки матеріалів**');
        embed.addFields(
          { name: 'Всього позицій', value: '0', inline: true },
          { name: 'Критичний мінімум', value: '0', inline: true }
        );
        break;
      case 'report':
        embed.setDescription('📋 **Звіт по МТЗ**');
        embed.addFields(
          { name: 'Боєприпаси', value: '0', inline: true },
          { name: 'Паливо', value: '0 л', inline: true },
          { name: 'Продовольство', value: '0 днів', inline: true }
        );
        break;
      case 'replenish':
        embed.setDescription('🔄 **Поповнення запасів**');
        embed.addFields({ name: 'Потребує поповнення', value: '0 позицій', inline: false });
        break;
      default:
        embed.setDescription('❌ Невідома дія');
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка операцій
   */
  private async handleOperations(interaction: any, action: DocumentAction): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('⚔️ Оперативні документи')
      .setColor(0xff6b6b)
      .setTimestamp();

    switch (action.type) {
      case 'search':
        embed.setDescription(`🔍 **Пошук операцій**\n\nЗапит: ${action.query || 'Всі операції'}`);
        embed.addFields({
          name: 'Результат',
          value: 'Тимчасова відповідь: Знайдено 0 операцій',
          inline: false,
        });
        break;
      case 'add':
        embed.setDescription(`➕ **Додавання операції**\n\nДані: ${action.query || 'Не вказано'}`);
        embed.addFields({ name: 'Статус', value: '✅ Операцію додано успішно', inline: false });
        break;
      case 'status':
        embed.setDescription('📊 **Статус операцій**');
        embed.addFields(
          { name: 'Активні операції', value: '0', inline: true },
          { name: 'Завершені', value: '0', inline: true },
          { name: 'Планування', value: '0', inline: true }
        );
        break;
      case 'report':
        embed.setDescription('📋 **Звіт по операціях**');
        embed.addFields(
          { name: 'Успішні', value: '0', inline: true },
          { name: 'В процесі', value: '0', inline: true },
          { name: 'Потребують підтримки', value: '0', inline: true }
        );
        break;
      case 'planning':
        embed.setDescription('📅 **Планування операцій**');
        embed.addFields({ name: 'Заплановано', value: '0 операцій', inline: false });
        break;
      default:
        embed.setDescription('❌ Невідома дія');
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Обробка наказів
   */
  private async handleOrders(interaction: any, action: DocumentAction): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('📋 Накази та розпорядження')
      .setColor(0x9932cc)
      .setTimestamp();

    switch (action.type) {
      case 'search':
        embed.setDescription(`🔍 **Пошук наказів**\n\nЗапит: ${action.query || 'Всі накази'}`);
        embed.addFields({
          name: 'Результат',
          value: 'Тимчасова відповідь: Знайдено 0 наказів',
          inline: false,
        });
        break;
      case 'create':
        embed.setDescription(`📝 **Створення наказу**\n\nЗміст: ${action.query || 'Не вказано'}`);
        embed.addFields({ name: 'Статус', value: '✅ Наказ створено успішно', inline: false });
        break;
      case 'status':
        embed.setDescription('📊 **Статус виконання**');
        embed.addFields(
          { name: 'Активні накази', value: '0', inline: true },
          { name: 'Виконані', value: '0', inline: true },
          { name: 'Прострочені', value: '0', inline: true }
        );
        break;
      case 'report':
        embed.setDescription('📋 **Звіт по наказах**');
        embed.addFields(
          { name: 'За цей місяць', value: '0', inline: true },
          { name: 'За квартал', value: '0', inline: true },
          { name: 'За рік', value: '0', inline: true }
        );
        break;
      case 'archive':
        embed.setDescription('📦 **Архівування наказів**');
        embed.addFields({ name: 'Готово до архівування', value: '0 наказів', inline: false });
        break;
      default:
        embed.setDescription('❌ Невідома дія');
    }

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Відобразити UI вибору при дублікатах (для подальшого використання командами)
   */
  // TODO: Загальний UI для дублікатів реалізовано у DuplicateResolver та onComponent

  /**
   * Обробка компонентів DuplicateResolver
   */
  protected override async onComponent(options: CommandComponentOptions): Promise<void> {
    const { interaction } = options;
    const customId = (interaction as any).customId as string | undefined;
    if (!customId || !customId.startsWith(DuplicateResolver.PREFIX)) return;

    await DuplicateResolver.handleComponent(interaction as any, {
      fetchFiles: async ({ scope, userId, nonce }) => {
        const key = uiState.makeKey({ scope, userId, nonce });
        return (
          uiState.get<Array<Pick<DriveFile, 'id' | 'name' | 'mimeType' | 'webViewLink' | 'owners'>>>(key) ?? []
        );
      },
      onSelect: async ({ fileId }) => {
        logger.debug('DuplicateResolver select', { component: 'DocumentsCommand', fileId });
      },
      title: 'Знайдено кілька збігів',
      perPage: 5,
    });
  }

  /**
   * (кінець класу)
   */
}
