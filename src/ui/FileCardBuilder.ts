import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle, type MessageActionRowComponentBuilder } from 'discord.js';
import { UI_COLORS, UI_EMOJI, i18nTitleFor } from './constants';
import type { DriveFile } from '@/types/drive';
import { t } from '@/i18n';

export interface FileCardOptions {
  locale?: string;
  showOwner?: boolean;
  showDates?: boolean;
  showOpen?: boolean;
  showDownload?: boolean;
  showSummary?: boolean;
  showQuestion?: boolean;
  hideWebLink?: boolean;
}

export function buildFileEmbed(file: DriveFile, opts: FileCardOptions = {}) {
  const name = String(file.name || file.id);
  const isFolder = String(file.mimeType || '').includes('folder');
  const icon = isFolder ? UI_EMOJI.folder : UI_EMOJI.file;
  const embed = new EmbedBuilder()
    .setTitle(`${icon} ${i18nTitleFor('file')}: ${name}`)
    .setColor(UI_COLORS.success)
    .setTimestamp();

  const parts: string[] = [];
  if (opts.showOwner && Array.isArray(file.owners) && file.owners.length) {
    parts.push(`${t('files.card.owner') || 'Власник'}: ${file.owners.join(', ')}`);
  }
  if (opts.showDates && file.modifiedTime) {
    parts.push(`${t('files.card.modified') || 'Оновлено'}: ${file.modifiedTime}`);
  }
  if (parts.length) embed.setDescription(parts.join('\n'));

  return embed;
}

export function buildFileActions(file: DriveFile, opts: FileCardOptions = {}) {
  const rows: ActionRowBuilder<MessageActionRowComponentBuilder>[] = [];
  const payload = Buffer.from(JSON.stringify({ id: file.id })).toString('base64');

  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>();
  if (opts.showOpen !== false) {
    row.addComponents(
      new ButtonBuilder().setCustomId(`drive:open:${payload}`).setLabel(t('files.buttons.open') || 'Відкрити').setStyle(ButtonStyle.Primary)
    );
  }
  if (opts.showDownload !== false) {
    row.addComponents(
      new ButtonBuilder().setCustomId(`drive:download:${payload}`).setLabel(t('files.buttons.download') || 'Завантажити').setStyle(ButtonStyle.Secondary)
    );
  }
  if (opts.showSummary !== false) {
    row.addComponents(
      new ButtonBuilder().setCustomId(`drive:summary:${payload}`).setLabel(t('files.buttons.summary') || 'Резюме').setStyle(ButtonStyle.Secondary)
    );
  }
  if (opts.showQuestion !== false) {
    row.addComponents(
      new ButtonBuilder().setCustomId(`drive:question:${payload}`).setLabel(t('files.buttons.question') || 'Питання').setStyle(ButtonStyle.Secondary)
    );
  }

  if (row.components.length) rows.push(row);

  if (!opts.hideWebLink && (file as any).webViewLink) {
    const linkRow = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(
      new ButtonBuilder().setLabel(t('files.buttons.source') || 'Джерело').setStyle(ButtonStyle.Link).setURL(String((file as any).webViewLink))
    );
    rows.push(linkRow);
  }

  return rows;
}
