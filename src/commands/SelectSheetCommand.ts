// no specific types needed here to avoid builder type narrowing issues
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import type { SlashCommandBuilder, SlashCommandStringOption } from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import logger from '@/utils/logger';

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
      {},
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
      let sheetName = interaction.options.getString('sheet') || undefined;

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

      // Валідуємо sheetName, якщо задано
      if (sheetName) {
        const sheets = await this.googleService.listSheets(spreadsheetId);
        const exists = sheets.some(s => s.toLowerCase() === sheetName!.toLowerCase());
        if (!exists) {
          throw new Error(t('sheets.error.sheetNotFound', { sheet: sheetName }));
        }
      }

      // Зберігаємо контекст
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
}

export default SelectSheetCommand;
