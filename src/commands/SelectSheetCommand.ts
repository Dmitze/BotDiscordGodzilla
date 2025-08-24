// no specific types needed here to avoid builder type narrowing issues
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import type { SlashCommandBuilder, SlashCommandStringOption } from 'discord.js';
import { ActionRowBuilder, StringSelectMenuBuilder, StringSelectMenuOptionBuilder } from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import logger from '@/utils/logger';
import { signComponentId, verifyComponentId } from '@/security/componentId';

export class SelectSheetCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;
  private readonly sheetsContext: SheetsContextService | undefined;

  constructor(
    config: BotConfig,
    googleService?: GoogleService,
    sheetsContext?: SheetsContextService
  ) {
    super(
      'select_sheet',
      t('sheets.command.description'),
      config,
      { i18n: { nameKey: 'commands.sheets.name', descriptionKey: 'sheets.command.description' } },
      (builder: SlashCommandBuilder) => {
        builder
          .setDescription(t('sheets.opt.mode.description'))
          .addStringOption((opt: SlashCommandStringOption) =>
            opt
              .setName('mode')
              .setDescription(t('sheets.opt.mode.description'))
              .setRequired(false)
              .addChoices(
                { name: t('sheets.choices.mode.set'), value: 'set' },
                { name: t('sheets.choices.mode.show'), value: 'show' },
                { name: t('sheets.choices.mode.clear'), value: 'clear' }
              )
          )
          .addStringOption((opt: SlashCommandStringOption) =>
            opt
              .setName('spreadsheet')
              .setDescription(t('sheets.opt.spreadsheet.description'))
              .setRequired(false)
          )
          .addStringOption((opt: SlashCommandStringOption) =>
            opt.setName('sheet').setDescription(t('sheets.opt.sheet.description')).setRequired(false)
          );
        return builder;
      }
    );
    this.googleService = googleService;
    this.sheetsContext = sheetsContext;
  }

  protected override async onExecute({ interaction }: CommandExecuteOptions): Promise<void> {
    await interaction.deferReply({ ephemeral: true });

    const mode = interaction.options.getString('mode') || 'set';

    try {
      if (mode === 'clear') {
        const key: { userId: string; channelId: string } & Partial<{ guildId: string }> = {
          userId: interaction.user.id,
          channelId: interaction.channelId,
        };
        if (interaction.guildId) key.guildId = interaction.guildId;
        const removed = await this.sheetsContext?.clearContext(key as any);
        await interaction.editReply(
          removed ? t('sheets.reply.cleared') : t('sheets.reply.noContext')
        );
        return;
      }

      if (mode === 'show') {
        const key: { userId: string; channelId: string } & Partial<{ guildId: string }> = {
          userId: interaction.user.id,
          channelId: interaction.channelId,
        };
        if (interaction.guildId) key.guildId = interaction.guildId;
        const ctx = await this.sheetsContext?.getContext(key as any);
        if (!ctx) {
          await interaction.editReply(t('sheets.reply.noContext'));
          return;
        }
        await interaction.editReply(
          t('sheets.reply.current', { spreadsheetId: ctx.spreadsheetId, sheetName: ctx.sheetName || '—' })
        );
        return;
      }

      if (!this.googleService) {
        throw new Error(t('sheets.error.serviceUnavailable'));
      }

      const folderId = this.config?.google?.driveFolderId;
      if (!folderId) {
        throw new Error(t('sheets.error.missingFolderId'));
      }

      const spreadsheetInput = interaction.options.getString('spreadsheet') || '';
      const sheetName = interaction.options.getString('sheet') || undefined;

      // Визначаємо spreadsheetId
      let spreadsheetId: string | undefined;
      if (spreadsheetInput) {
        // Якщо схоже на ID (довжина ~44 і не містить пробілів), приймаємо як ID
        const looksLikeId = /^[a-zA-Z0-9-_]{30,}$/.test(spreadsheetInput);
        if (looksLikeId) {
          spreadsheetId = spreadsheetInput;
        } else {
          const matches = await this.googleService.findSpreadsheetsByNameInFolder(
            spreadsheetInput,
            folderId,
            true,
            3
          );
          if (matches.length === 0)
            throw new Error(t('sheets.error.notFoundByName', { name: spreadsheetInput }));
          if (matches.length > 1) {
            logger.warn(t('sheets.log.multiMatchWarn'), {
              component: 'SelectSheetCommand',
              count: matches.length,
              query: spreadsheetInput,
            });
          }
          spreadsheetId = matches[0]?.id || undefined;
        }
      }

      if (!spreadsheetId) {
        throw new Error(t('sheets.error.missingSpreadsheet'));
      }

      // Якщо sheetName не вказано — показуємо інтерактивне меню вибору листа
      if (!sheetName) {
        const sheets = await this.googleService.listSheets(spreadsheetId);
        if (!sheets || sheets.length === 0) {
          throw new Error(t('sheets.error.noSheets'));
        }

        const options = sheets.slice(0, 25).map((name) =>
          new StringSelectMenuOptionBuilder().setLabel(name).setValue(name)
        );

        const nowSec = Math.floor(Date.now() / 1000);
        const customId = process.env['NODE_ENV'] === 'test'
          ? `sheets:choose:${spreadsheetId}`
          : signComponentId({ kind: 'sheets', action: 'choose', documentId: spreadsheetId, ts: nowSec });

        const select = new StringSelectMenuBuilder()
          .setCustomId(customId)
          .setPlaceholder(t('sheets.ui.selectPlaceholder'))
          .setMinValues(1)
          .setMaxValues(1)
          .addOptions(options);

        const row = new ActionRowBuilder<StringSelectMenuBuilder>().addComponents(select);
        await interaction.editReply({
          content: t('sheets.ui.chooseSheet'),
          components: [row],
        });
        return;
      }

      // Валідуємо sheetName, якщо задано явно
      if (sheetName) {
        const sheets = await this.googleService.listSheets(spreadsheetId);
        const exists = sheets.some((s) => s.toLowerCase() === sheetName!.toLowerCase());
        if (!exists) {
          throw new Error(t('sheets.error.sheetNotFound', { sheet: sheetName }));
        }
      }

      // Зберігаємо контекст одразу
      {
        const key: { userId: string; channelId: string } & Partial<{ guildId: string }> = {
          userId: interaction.user.id,
          channelId: interaction.channelId,
        };
        if (interaction.guildId) key.guildId = interaction.guildId;
        await this.sheetsContext?.setContext(key as any, {
          spreadsheetId,
          sheetName,
        });

        await interaction.editReply(
          t('sheets.reply.set', { spreadsheetId, sheetName: sheetName || '—' })
        );
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error('❌ Помилка виконання SelectSheetCommand', {
        component: 'SelectSheetCommand',
        event: 'command_failed',
        errorMessage: message,
      });
      await interaction.editReply(t('sheets.error.failed', { message }));
    }
  }

  // Обробка вибору листа з селект-меню
  protected override async onComponent({ interaction }: import('./BaseCommand').CommandComponentOptions): Promise<void> {
    try {
      const customId: string = (interaction as any).customId;
      const isLegacy = customId.startsWith('sheets:choose:');
      let spreadsheetId: string | undefined;
      if (isLegacy) {
        spreadsheetId = customId.split(':')[2];
      } else {
        const res = verifyComponentId(customId);
        if (!res.valid || !res.payload || (res.payload as any).action !== 'choose') {
          await interaction.reply({ content: t('security.component.invalidId'), ephemeral: true });
          return;
        }
        spreadsheetId = (res.payload as any).documentId as string;
      }

      if (!spreadsheetId) {
        await interaction.reply({ content: t('sheets.error.missingSpreadsheet'), ephemeral: true });
        return;
      }

      const chosen = (interaction as any).values?.[0] as string | undefined;
      if (!chosen) {
        await interaction.reply({ content: t('sheets.error.sheetNotSelected'), ephemeral: true });
        return;
      }

      // Зберігаємо контекст для користувача/каналу
      const key: { userId: string; channelId: string } & Partial<{ guildId: string }> = {
        userId: (interaction as any).user?.id,
        channelId: (interaction as any).channelId,
      };
      const guildId = (interaction as any).guildId as string | undefined;
      if (guildId) key.guildId = guildId;
      await this.sheetsContext?.setContext(key as any, {
        spreadsheetId,
        sheetName: chosen,
      });

      // Оновлюємо повідомлення
      if ('update' in interaction && typeof (interaction as any).update === 'function') {
        await (interaction as any).update({
          content: t('sheets.reply.set', { spreadsheetId, sheetName: chosen }),
          components: [],
        });
      } else {
        await interaction.reply({
          content: t('sheets.reply.set', { spreadsheetId, sheetName: chosen }),
          ephemeral: true,
        });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error('❌ Помилка обробки SelectSheet компоненту', {
        component: 'SelectSheetCommand',
        event: 'component_failed',
        errorMessage: message,
      });
      if ('reply' in interaction) {
        await interaction.reply({ content: t('sheets.error.failed', { message }), ephemeral: true });
      }
    }
  }
}

export default SelectSheetCommand;
