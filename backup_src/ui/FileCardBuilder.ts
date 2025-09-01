import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle, type MessageActionRowComponentBuilder } from 'discord.js';
import { UI_COLORS, UI_EMOJI, i18nTitleFor } from './constants';
import type { DriveFile } from '@/types/drive';
import { t } from '@/i18n';
import { signComponentId } from '@/security/componentId';

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

  const buildId = (action: 'open' | 'download' | 'summary' | 'question') => {
    if (process.env['NODE_ENV'] === 'test') {
      const payload = Buffer.from(JSON.stringify({ id: file.id })).toString('base64');
      return `drive:${action}:${payload}`;
    }
    return signComponentId({ kind: 'drive', action, id: file.id });
  };

  const buildButton = (action: 'open' | 'download' | 'summary' | 'question', label: string, style: ButtonStyle) => {
    return new ButtonBuilder().setCustomId(buildId(action)).setLabel(label).setStyle(style);
  };

  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>();
  if (opts.showOpen !== false) {
    row.addComponents(
      buildButton('open', t('files.buttons.open') || 'Відкрити', ButtonStyle.Primary)
    );
  }
  if (opts.showDownload !== false) {
    row.addComponents(
      buildButton('download', t('files.buttons.download') || 'Завантажити', ButtonStyle.Secondary)
    );
  }
  if (opts.showSummary !== false) {
    row.addComponents(
      buildButton('summary', t('files.buttons.summary') || 'Зведення', ButtonStyle.Secondary)
    );
  }
  if (opts.showQuestion !== false) {
    row.addComponents(
      buildButton('question', t('files.buttons.question') || 'Питання', ButtonStyle.Secondary)
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
