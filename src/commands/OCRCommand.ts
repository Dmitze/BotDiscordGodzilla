import { SlashCommandBuilder } from 'discord.js';
import type { Attachment } from 'discord.js';
import { BaseCommand } from '@/commands/BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import { t } from '@/i18n';

export class OCRCommand extends BaseCommand {
  private readonly google: GoogleService | null;

  constructor(config: BotConfig, google?: GoogleService) {
    super(
      'ocr',
      t('ocr.command.description'),
      config,
      {
        category: 'documents',
        usage: '/ocr [image] | [drive_id]'
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addAttachmentOption(opt =>
            opt
              .setName('image')
              .setDescription(t('ocr.opt.image.description'))
              .setRequired(false)
          )
          .addStringOption(opt =>
            opt
              .setName('drive_id')
              .setDescription(t('ocr.opt.drive_id.description'))
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
      const attachment = interaction.options.getAttachment('image') as Attachment | null;
      const driveId = interaction.options.getString('drive_id');

      if (!attachment && !driveId) {
        await interaction.editReply(t('ocr.error.needInput'));
        return;
      }

      if (!this.google) {
        await interaction.editReply(t('ocr.error.serviceUnavailable'));
        return;
      }

      let text = '';

      if (attachment) {
        const url = attachment.url;
        const res = await fetch(url);
        if (!res.ok) throw new Error(t('ocr.error.downloadFail', { status: String(res.status), statusText: res.statusText }));
        const buf = Buffer.from(await res.arrayBuffer());
        text = await this.google.extractTextFromBuffer(buf);
      } else if (driveId) {
        const file = await this.google.getDriveFileMetadata(driveId);
        text = await this.google.extractTextFromImage(file);
      }

      const result = text?.trim() || t('ocr.result.empty');
      const provider = this.config.google.ocrProvider ?? 'vision';
      await interaction.editReply(`${t('ocr.result.provider', { provider })}\n\n${result.slice(0, 1900)}`);
    } catch (error) {
      logger.error('Ошибка выполнения OCR команды', {
        type: 'command',
        command: 'ocr',
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply(t('ocr.error.process'));
    }
  }
}
