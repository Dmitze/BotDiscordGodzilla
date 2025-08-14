import { SlashCommandBuilder, EmbedBuilder } from 'discord.js';
import { BaseCommand } from '@/commands/BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';
import { extractTextFromDriveFile, summarizeText } from '@/utils/textExtractor';

export class DriveExtractCommand extends BaseCommand {
  private readonly google: GoogleService | null;

  constructor(config: BotConfig, google?: GoogleService) {
    super(
      'drive-extract',
      'Извлечь текст из последних файлов Google Drive (по папке из конфигурации)',
      config,
      {
        category: 'documents',
        usage: '/drive-extract [count] [mime] [previews]'
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addIntegerOption(opt =>
            opt
              .setName('count')
              .setDescription('Сколько файлов обработать (1-5)')
              .setRequired(false)
              .setMinValue(1)
              .setMaxValue(5)
          )
          .addStringOption(opt =>
            opt
              .setName('mime')
              .setDescription('Фильтр MIME (например, application/pdf)')
              .setRequired(false)
          )
          .addBooleanOption(opt =>
            opt
              .setName('previews')
              .setDescription('Показывать текстовые превью (true/false)')
              .setRequired(false)
          );
        return builder;
      }
    );

    this.google = google ?? null;
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const interaction = options.interaction;
    await interaction.deferReply({ ephemeral: false });

    try {
      if (!this.google) {
        await interaction.editReply('GoogleService недоступен. Обратитесь к администратору.');
        return;
      }

      const folderId = this.config.drive.folderId;
      if (!folderId) {
        await interaction.editReply('GOOGLE_DRIVE_FOLDER_ID не задан в конфигурации.');
        return;
      }

      const count = interaction.options.getInteger('count') ?? 3;
      const want = Math.max(1, Math.min(5, count));
      const mimeFilter = interaction.options.getString('mime') ?? '*';
      const showPreviews = interaction.options.getBoolean('previews') ?? true;

      // Листинг файлов
      const page = await this.google.listDriveFiles({
        folderId,
        pageSize: want,
        mimeIncludes: mimeFilter === '*' ? [] : [mimeFilter],
      });
      const files: DriveFile[] = page.files.slice(0, want);

      if (files.length === 0) {
        await interaction.editReply('В папке нет файлов по заданным условиям.');
        return;
      }

      // Извлечение текста
      const results: { file: DriveFile; text: string; warnings: string[]; mimeType: string }[] = [];

      for (const f of files) {
        try {
          const meta = await this.google.getDriveFile(f.id);
          const res = await extractTextFromDriveFile(this.google, meta);
          results.push({ file: meta, text: res.text, warnings: res.warnings, mimeType: res.mimeType });
        } catch (e) {
          logger.error('Ошибка при обработке файла в /drive-extract', {
            type: 'command',
            command: 'drive-extract',
            fileId: f.id,
            error: e instanceof Error ? e.message : String(e),
          });
          results.push({ file: f, text: '', warnings: [String(e)], mimeType: f.mimeType ?? '' });
        }
      }

      // Формирование ответа
      const embed = new EmbedBuilder()
        .setTitle('Извлечение текста из файлов Google Drive')
        .setDescription(`Папка: ${folderId}\nФайлов: ${results.length}${mimeFilter && mimeFilter !== '*' ? `\nФильтр MIME: ${mimeFilter}` : ''}`)
        .setColor(0x2f855a);

      for (const r of results) {
        const preview = showPreviews ? summarizeText(r.text, 300).replace(/\n/g, ' ⏎ ') : 'превью скрыто';
        const warn = r.warnings?.length ? `\n⚠️ ${r.warnings.join('; ')}` : '';
        embed.addFields({
          name: `${r.file.name ?? 'без имени'} (${r.mimeType || r.file.mimeType || 'unknown'})`,
          value: `${preview || '[пусто]'}${warn}`.slice(0, 1024),
        });
      }

      await interaction.editReply({ embeds: [embed] });
    } catch (error) {
      logger.error('Ошибка выполнения команды /drive-extract', {
        type: 'command',
        command: 'drive-extract',
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Ошибка при извлечении текста. Попробуйте позже.');
    }
  }
}
