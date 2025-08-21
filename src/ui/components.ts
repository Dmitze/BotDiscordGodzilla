import {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  type MessageActionRowComponentBuilder,
} from 'discord.js';
import { t } from '@/i18n';

export type BuildIdFn = (args: { sid: string; page: number; ts: number; action?: 'toggle' | 'reset' | 'close' }) => string;

export function buildSearchPaginationRows(params: {
  sid: string;
  safePage: number;
  totalPages: number;
  changesOnly: boolean;
  allowLink: boolean;
  folderId?: string;
  buildId: BuildIdFn;
}): ActionRowBuilder<MessageActionRowComponentBuilder>[] {
  const { sid, safePage, totalPages, changesOnly, allowLink, folderId, buildId } = params;
  const ts = Math.floor(Date.now() / 1000);

  const row1 = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(
    new ButtonBuilder().setCustomId(buildId({ sid, page: 1, ts }))
      .setLabel(t('files.search.buttons.first'))
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === 1),
    new ButtonBuilder().setCustomId(buildId({ sid, page: Math.max(1, safePage - 1), ts }))
      .setLabel(t('files.search.buttons.prev'))
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === 1),
    new ButtonBuilder().setCustomId(buildId({ sid, page: Math.min(totalPages, safePage + 1), ts }))
      .setLabel(t('files.search.buttons.next'))
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === totalPages),
    new ButtonBuilder().setCustomId(buildId({ sid, page: totalPages, ts }))
      .setLabel(t('files.search.buttons.last'))
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === totalPages),
  );

  const row2 = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(
    new ButtonBuilder().setCustomId(buildId({ sid, page: safePage, ts, action: 'toggle' }))
      .setLabel(t('files.search.buttons.showChangesOnly'))
      .setStyle(changesOnly ? ButtonStyle.Primary : ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(buildId({ sid, page: 1, ts, action: 'reset' }))
      .setLabel(t('files.search.buttons.resetBaseline'))
      .setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(buildId({ sid, page: safePage, ts, action: 'close' }))
      .setLabel(t('files.search.buttons.close'))
      .setStyle(ButtonStyle.Danger),
  );

  const rows: ActionRowBuilder<MessageActionRowComponentBuilder>[] = [row1, row2];

  if (allowLink && folderId && folderId !== 'root') {
    const folderUrl = `https://drive.google.com/drive/folders/${encodeURIComponent(folderId)}`;
    const linkBtn = new ButtonBuilder().setLabel(t('files.buttons.source')).setStyle(ButtonStyle.Link).setURL(folderUrl);
    rows.push(new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(linkBtn));
  }

  return rows;
}
