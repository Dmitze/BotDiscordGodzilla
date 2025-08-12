// no specific types needed here to avoid builder type narrowing issues
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import logger from '@/utils/logger';

export class SelectSheetCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;
  private readonly sheetsContext: SheetsContextService | undefined;

  constructor(config: BotConfig, googleService?: GoogleService, sheetsContext?: SheetsContextService) {
    super(
      'вибрати_таблицю',
      '📁 Вибір Google таблиці та листа для контексту пошуку',
      config,
      {},
      (builder: any) =>
        builder
          .setDescription('Встановити/показати/очистити контекст таблиці та листа')
          .addStringOption((opt: any) =>
            opt.setName('режим')
              .setDescription('Дія: встановити, показати або очистити')
              .setRequired(false)
              .addChoices(
                { name: 'встановити', value: 'set' },
                { name: 'показати', value: 'show' },
                { name: 'очистити', value: 'clear' },
              )
          )
          .addStringOption((opt: any) =>
            opt.setName('таблиця')
              .setDescription('Назва таблиці (у папці) або ID')
              .setRequired(false)
          )
          .addStringOption((opt: any) =>
            opt.setName('лист')
              .setDescription('Назва листа в таблиці')
              .setRequired(false)
          )
    );
    this.googleService = googleService;
    this.sheetsContext = sheetsContext;
  }

  protected override async onExecute({ interaction }: CommandExecuteOptions): Promise<void> {
    await interaction.deferReply({ ephemeral: true });

    const mode = interaction.options.getString('режим') || 'set';

    try {
      if (mode === 'clear') {
        const key: { userId: string; channelId: string } & Partial<{ guildId: string }> = {
          userId: interaction.user.id,
          channelId: interaction.channelId,
        };
        if (interaction.guildId) key.guildId = interaction.guildId;
        const removed = await this.sheetsContext?.clearContext(key as any);
        await interaction.editReply(removed ? '✅ Контекст очищено' : 'ℹ️ Немає збереженого контексту');
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
          await interaction.editReply('ℹ️ Контекст не встановлено');
          return;
        }
        await interaction.editReply(`📄 Поточний контекст:\nSpreadsheet: ${ctx.spreadsheetId}\nSheet: ${ctx.sheetName || '—'}`);
        return;
      }

      if (!this.googleService) {
        throw new Error('GoogleService недоступний');
      }

      const folderId = this.config?.google?.driveFolderId;
      if (!folderId) {
        throw new Error('Не вказано GOOGLE_DRIVE_FOLDER_ID в конфігурації');
      }

      const spreadsheetInput = interaction.options.getString('таблиця') || '';
      let sheetName = interaction.options.getString('лист') || undefined;

      // Визначаємо spreadsheetId
      let spreadsheetId: string | undefined;
      if (spreadsheetInput) {
        // Якщо схоже на ID (довжина ~44 і не містить пробілів), приймаємо як ID
        const looksLikeId = /^[a-zA-Z0-9-_]{30,}$/.test(spreadsheetInput);
        if (looksLikeId) {
          spreadsheetId = spreadsheetInput;
        } else {
          const matches = await this.googleService.findSpreadsheetsByNameInFolder(spreadsheetInput, folderId, true, 3);
          if (matches.length === 0) throw new Error(`Таблицю за ім'ям "${spreadsheetInput}" не знайдено у папці`);
          if (matches.length > 1) {
            logger.warn('SelectSheet: знайдено кілька відповідників, обираємо перший', {
              component: 'SelectSheetCommand', count: matches.length, query: spreadsheetInput,
            });
          }
          spreadsheetId = matches[0]?.id || undefined;
        }
      }

      if (!spreadsheetId) {
        throw new Error('Не вказано таблицю. Задайте параметр "таблиця" (назва або ID).');
      }

      // Валідуємо sheetName, якщо задано
      if (sheetName) {
        const sheets = await this.googleService.listSheets(spreadsheetId);
        const exists = sheets.some(s => s.toLowerCase() === sheetName!.toLowerCase());
        if (!exists) {
          throw new Error(`Лист "${sheetName}" не знайдено у вибраній таблиці`);
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

      await interaction.editReply(`✅ Контекст встановлено:\nSpreadsheet: ${spreadsheetId}\nSheet: ${sheetName || '—'}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error('❌ Помилка виконання SelectSheetCommand', {
        component: 'SelectSheetCommand', event: 'command_failed', errorMessage: message,
      });
      await interaction.editReply(`❌ Помилка: ${message}`);
    }
  }
}

export default SelectSheetCommand;
