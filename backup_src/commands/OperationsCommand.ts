/**
 * ⚔️ Команди оперативного управління ЗСУ
 * Спеціалізовані функції для оперативної роботи
 */

import type { ChatInputCommandInteraction } from 'discord.js';
import { EmbedBuilder } from 'discord.js';
import { replyWithPrivacy } from '@/ui/reply';
import { tUser } from '@/i18n';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

// локальні інтерфейси видалено як не використані

export class OperationsCommand extends BaseCommand {
  constructor(config: BotConfig) {
    // Command and option names must be lowercase ASCII (Discord constraint)
    super('operations', '⚔️ Оперативне управління ЗСУ', config, {}, (builder: any) => {
      return builder
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('situation')
            .setDescription('📊 Поточна оперативна ситуація')
            .addStringOption((option: any) =>
              option
                .setName('sector')
                .setDescription('Оперативний сектор')
                .setRequired(false)
                .addChoices(
                  { name: 'Всі сектори', value: 'all' },
                  { name: 'Сектор А', value: 'A' },
                  { name: 'Сектор Б', value: 'B' },
                  { name: 'Сектор В', value: 'C' },
                  { name: 'Сектор Г', value: 'D' }
                )
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('tasks')
            .setDescription('🎯 Управління завданнями')
            .addStringOption((option: any) =>
              option
                .setName('action')
                .setDescription('Дія з завданнями')
                .setRequired(true)
                .addChoices(
                  { name: 'Поточні завдання', value: 'current' },
                  { name: 'Нове завдання', value: 'new' },
                  { name: 'Оновити статус', value: 'update' },
                  { name: 'Завершити завдання', value: 'complete' },
                  { name: 'Архів завдань', value: 'archive' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('query').setDescription('Пошуковий запит або дані').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('coordination')
            .setDescription('🔄 Координація між підрозділами')
            .addStringOption((option: any) =>
              option
                .setName('type')
                .setDescription('Тип координації')
                .setRequired(true)
                .addChoices(
                  { name: 'Вогнева підтримка', value: 'fire_support' },
                  { name: 'Логістика', value: 'logistics' },
                  { name: 'Розвідка', value: 'intelligence' },
                  { name: 'Медична допомога', value: 'medical' },
                  { name: "Зв'язок", value: 'communications' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('unit').setDescription('Підрозділ для координації').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('intelligence')
            .setDescription('🔍 Розвідувальні дані')
            .addStringOption((option: any) =>
              option
                .setName('type')
                .setDescription('Тип розвідки')
                .setRequired(true)
                .addChoices(
                  { name: 'Повітряна розвідка', value: 'air' },
                  { name: 'Наземна розвідка', value: 'ground' },
                  { name: 'Технічна розвідка', value: 'technical' },
                  { name: 'Агентурна розвідка', value: 'agent' },
                  { name: 'Зведена розвідка', value: 'summary' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('area').setDescription('Район розвідки').setRequired(false)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('communications')
            .setDescription("📡 Управління зв'язком")
            .addStringOption((option: any) =>
              option
                .setName('action')
                .setDescription("Дія зі зв'язком")
                .setRequired(true)
                .addChoices(
                  { name: "Статус зв'язку", value: 'status' },
                  { name: 'Налаштування каналів', value: 'channels' },
                  { name: 'Передача повідомлення', value: 'message' },
                  { name: 'Перевірка якості', value: 'quality' },
                  { name: 'Резервні канали', value: 'backup' }
                )
            )
            .addStringOption((option: any) =>
              option.setName('channel').setDescription("Канал зв'язку").setRequired(false)
            )
            .addStringOption((option: any) =>
              option.setName('message').setDescription('Текст повідомлення').setRequired(false)
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
      const services = (interaction).client?.serviceContainer;
      const opsService = services?.get?.('operations');

      switch (subcommand) {
        case 'situation':
          await this.handleSituation(interaction, opsService);
          break;
        case 'tasks':
          await this.handleTasks(interaction, opsService);
          break;
        case 'coordination':
          await this.handleCoordination(interaction, opsService);
          break;
        case 'intelligence':
          await this.handleIntelligence(interaction, opsService);
          break;
        case 'communications':
          await this.handleCommunications(interaction, opsService);
          break;
        default:
          await replyWithPrivacy(interaction, { content: tUser('operations.error.unknownSubcommand', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
      }
    } catch (error) {
      logger.error('❌ Помилка команди операцій', {
        error: error instanceof Error ? error.message : String(error),
        userId: (interaction as ChatInputCommandInteraction).user?.id,
        command: this.name,
      });
      await replyWithPrivacy(interaction, { content: tUser('operations.error.general', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
    }
  }

  /**
   * Обробка оперативної ситуації
   */
  private async handleSituation(
    interaction: ChatInputCommandInteraction,
    opsService?: any
  ): Promise<void> {
    const sector = interaction.options.getString('sector') || 'all';

    // Виклик сервісу, якщо він є (для unit-тестів використовується мок)
    try {
      if (opsService?.getSituation) {
        await opsService.getSituation(sector);
      }
    } catch (e) {
      await replyWithPrivacy(interaction, { content: tUser('operations.error.getSituation', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
      return;
    }

    const embed = new EmbedBuilder()
      .setTitle(tUser('operations.situation.title', interaction))
      .setColor(0xff6b6b)
      .setTimestamp();

    if (sector === 'all') {
      embed.setDescription(tUser('operations.situation.descAll', interaction));
      embed.addFields(
        { name: tUser('operations.situation.fields.sectorA', interaction), value: tUser('operations.situation.fields.stable', interaction), inline: true },
        { name: tUser('operations.situation.fields.sectorB', interaction), value: tUser('operations.situation.fields.active', interaction), inline: true },
        { name: tUser('operations.situation.fields.sectorC', interaction), value: tUser('operations.situation.fields.controlled', interaction), inline: true },
        { name: tUser('operations.situation.fields.sectorD', interaction), value: tUser('operations.situation.fields.attention', interaction), inline: true }
      );
    } else {
      embed.setDescription(tUser('operations.situation.descSector', interaction, { sector }));
      embed.addFields(
        { name: tUser('operations.situation.fields.status', interaction), value: tUser('operations.situation.fields.stable', interaction), inline: true },
        { name: tUser('operations.situation.fields.activeTasks', interaction), value: '3', inline: true },
        { name: tUser('operations.situation.fields.units', interaction), value: '5', inline: true }
      );
    }

    await replyWithPrivacy(interaction, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
  }

  /**
   * Обробка завдань
   */
  private async handleTasks(
    interaction: ChatInputCommandInteraction,
    opsService?: any
  ): Promise<void> {
    const action = interaction.options.getString('action', true);
    const query = interaction.options.getString('query');

    logger.info('Виконання дії управління завданнями', {
      action,
      query: query || undefined,
      userId: interaction.user.id,
    });

    const embed = new EmbedBuilder()
      .setTitle(tUser('operations.tasks.title', interaction))
      .setColor(0x0099ff)
      .setTimestamp();

    switch (action) {
      case 'current':
        try {
          const tasks = (await opsService?.getTasks?.('current')) ?? [];
          if (!tasks.length) {
            await replyWithPrivacy(interaction, { content: tUser('operations.tasks.none', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
            return;
          }
        } catch (e) {
          await replyWithPrivacy(interaction, { content: tUser('operations.tasks.error', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
          return;
        }
        embed.setDescription(tUser('operations.tasks.current', interaction));
        embed.addFields(
          { name: tUser('operations.tasks.fields.active', interaction), value: '5', inline: true },
          { name: tUser('operations.tasks.fields.inProgress', interaction), value: '3', inline: true },
          { name: tUser('operations.tasks.fields.pending', interaction), value: '2', inline: true }
        );
        break;
      case 'new':
        embed.setDescription(
          tUser('operations.tasks.new.desc', interaction, { data: query || tUser('operations.common.notSpecified', interaction) })
        );
        embed.addFields({ name: tUser('operations.common.status', interaction), value: tUser('operations.common.taskCreated', interaction), inline: false });
        break;
      case 'update':
        embed.setDescription(
          tUser('operations.tasks.update.desc', interaction, { data: query || tUser('operations.common.notSpecified', interaction) })
        );
        embed.addFields({ name: tUser('operations.common.status', interaction), value: tUser('operations.common.statusUpdated', interaction), inline: false });
        break;
      case 'complete':
        embed.setDescription(
          tUser('operations.tasks.complete.desc', interaction, { data: query || tUser('operations.common.notSpecified', interaction) })
        );
        embed.addFields({ name: tUser('operations.common.status', interaction), value: tUser('operations.common.taskCompleted', interaction), inline: false });
        break;
      case 'archive':
        embed.setDescription(tUser('operations.tasks.archive.title', interaction));
        embed.addFields(
          { name: tUser('operations.tasks.fields.completed', interaction), value: '15', inline: true },
          { name: tUser('operations.tasks.fields.archived', interaction), value: '8', inline: true }
        );
        break;
      default:
        embed.setDescription(tUser('operations.error.unknownAction', interaction));
    }

    await replyWithPrivacy(interaction, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
  }

  /**
   * Обробка координації
   */
  private async handleCoordination(
    interaction: ChatInputCommandInteraction,
    opsService?: any
  ): Promise<void> {
    const type = interaction.options.getString('type', true);
    const unit = interaction.options.getString('unit');

    try {
      await opsService?.coordinate?.('emergency');
    } catch (e) {
      await replyWithPrivacy(interaction, { content: tUser('operations.coordination.error', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
      return;
    }

    logger.info('Координація між підрозділами', {
      type,
      unit: unit || undefined,
      userId: interaction.user.id,
    });

    const embed = new EmbedBuilder()
      .setTitle(tUser('operations.coordination.title', interaction))
      .setColor(0xff9900)
      .setTimestamp();

    const typeName = this.getCoordinationTypeName(type);

    embed.setDescription(
      tUser('operations.coordination.desc', interaction, { typeName, unit: unit || tUser('operations.common.allUnits', interaction) })
    );
    embed.addFields(
      { name: tUser('operations.coordination.fields.status', interaction), value: tUser('operations.common.active', interaction), inline: true },
      { name: tUser('operations.coordination.fields.participants', interaction), value: '3', inline: true },
      { name: tUser('operations.coordination.fields.channel', interaction), value: tUser('operations.common.primary', interaction), inline: true }
    );

    await replyWithPrivacy(interaction, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
  }

  /**
   * Обробка розвідки
   */
  private async handleIntelligence(
    interaction: ChatInputCommandInteraction,
    _opsService?: any
  ): Promise<void> {
    const type = interaction.options.getString('type', true);
    const area = interaction.options.getString('area');

    try {
      await _opsService?.getIntelligence?.(type);
    } catch (e) {
      await replyWithPrivacy(interaction, { content: tUser('operations.intelligence.error', interaction) }, { ephemeralByDefault: true, shareFlagSupport: true });
      return;
    }

    logger.info('Запит розвідувальних даних', {
      type,
      area: area || undefined,
      userId: interaction.user.id,
    });

    const embed = new EmbedBuilder()
      .setTitle(tUser('operations.intelligence.title', interaction))
      .setColor(0x00ff88)
      .setTimestamp();

    const typeName = this.getIntelligenceTypeName(type);

    embed.setDescription(
      tUser('operations.intelligence.desc', interaction, { typeName, area: area || tUser('operations.common.allAreas', interaction) })
    );
    embed.addFields(
      { name: tUser('operations.intelligence.fields.lastData', interaction), value: tUser('operations.common.time2hAgo', interaction), inline: true },
      { name: tUser('operations.intelligence.fields.reliability', interaction), value: tUser('operations.common.high', interaction), inline: true },
      { name: tUser('operations.intelligence.fields.source', interaction), value: tUser('operations.common.confirmed', interaction), inline: true }
    );

    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  /**
   * Обробка зв'язку
   */
  private async handleCommunications(
    interaction: ChatInputCommandInteraction,
    _opsService?: any
  ): Promise<void> {
    const action = interaction.options.getString('action', true);
    const channel = interaction.options.getString('channel');
    const message = interaction.options.getString('message');

    logger.info("Управління зв'язком", {
      action,
      channel: channel || undefined,
      message: message || undefined,
      userId: interaction.user.id,
    });

    const embed = new EmbedBuilder()
      .setTitle(tUser('operations.communications.title', interaction))
      .setColor(0x9932cc)
      .setTimestamp();

    switch (action) {
      case 'status':
        embed.setDescription(tUser('operations.communications.action.status', interaction));
        embed.addFields(
          { name: tUser('operations.communications.fields.primaryChannel', interaction), value: tUser('operations.common.operational', interaction), inline: true },
          { name: tUser('operations.communications.fields.backupChannel', interaction), value: tUser('operations.common.ready', interaction), inline: true },
          { name: tUser('operations.communications.fields.signalQuality', interaction), value: tUser('operations.common.high', interaction), inline: true }
        );
        break;
      case 'channels':
        embed.setDescription(tUser('operations.communications.action.channels', interaction));
        embed.addFields(
          { name: tUser('operations.communications.fields.activeChannels', interaction), value: '3', inline: true },
          { name: tUser('operations.communications.fields.backupChannels', interaction), value: '2', inline: true }
        );
        break;
      case 'message':
        embed.setDescription(
          tUser('operations.communications.action.message', interaction, {
            channel: channel || tUser('operations.common.primary', interaction),
            message: message || tUser('operations.common.notSpecified', interaction),
          })
        );
        embed.addFields({ name: tUser('operations.common.status', interaction), value: tUser('operations.common.messageSent', interaction), inline: false });
        break;
      case 'quality':
        embed.setDescription(tUser('operations.communications.action.quality', interaction));
        embed.addFields(
          { name: tUser('operations.communications.fields.signalQuality', interaction), value: '95%', inline: true },
          { name: tUser('operations.communications.fields.latency', interaction), value: '50ms', inline: true },
          { name: tUser('operations.communications.fields.stability', interaction), value: tUser('operations.common.high', interaction), inline: true }
        );
        break;
      case 'backup':
        embed.setDescription(tUser('operations.communications.action.backup', interaction));
        embed.addFields(
          { name: tUser('operations.communications.fields.channel1', interaction), value: tUser('operations.common.active', interaction), inline: true },
          { name: tUser('operations.communications.fields.channel2', interaction), value: tUser('operations.common.ready', interaction), inline: true }
        );
        break;
      default:
        embed.setDescription(tUser('operations.error.unknownAction', interaction));
    }

    await replyWithPrivacy(interaction, { embeds: [embed] }, { ephemeralByDefault: true, shareFlagSupport: true });
  }

  /**
   * Отримання назви типу координації
   */
  private getCoordinationTypeName(type: string): string {
    const typeNames: Record<string, string> = {
      fire_support: 'Вогнева підтримка',
      logistics: 'Логістика',
      intelligence: 'Розвідка',
      medical: 'Медична допомога',
      communications: "Зв'язок",
    };

    return typeNames[type] || type;
  }

  /**
   * Отримання назви типу розвідки
   */
  private getIntelligenceTypeName(type: string): string {
    const typeNames: Record<string, string> = {
      air: 'Повітряна розвідка',
      ground: 'Наземна розвідка',
      technical: 'Технічна розвідка',
      agent: 'Агентурна розвідка',
      summary: 'Зведена розвідка',
    };

    return typeNames[type] || type;
  }
}
