import { EmbedBuilder, ButtonBuilder, ButtonStyle, ActionRowBuilder, type ChatInputCommandInteraction, type MessageActionRowComponentBuilder } from 'discord.js';
import { t } from '@/i18n';

export interface ReadDeps {
  config: any;
  getGoogleService: (interaction: ChatInputCommandInteraction) => any | undefined;
  isMimeAllowed: (mime: string, allowed: string[]) => boolean;
  isOwnerAllowed: (owners: string[], allowlist: string[]) => boolean;
  isTooLarge: (bytes: number, limitMb: number) => boolean;
  getSubcommandTitle: (name: 'пошук' | 'читати' | 'аналіз' | string) => string;
  sanitizeTextForChat: (text: string, maxLen: number) => string;
  buildPaginatedChunks: (text: string, opts: { maxChunkLen: number }) => string[];
  summarizeTlDr: (text: string, opts: { budget: number; minSentLen: number }) => string;
  generateSessionId: (prefix: string) => string;
  buildTextCustomId: (args: { sid: string; page: number; action?: 'close' }) => string;
  textSessions: {
    set: (sid: string, v: { fileId: string; fileName: string; chunks: string[]; createdAt: number; link?: string }) => void;
  };
  mapGoogleApiErrorToMessage: (e: any) => string | null;
}

// --- Helpers extracted to reduce complexity ---
function isAllowedByPolicy(meta: any, driveCfg: any, isMimeAllowed: (mime: string, allowed: string[]) => boolean, isOwnerAllowed: (owners: string[], allowlist: string[]) => boolean): { ok: boolean; reason?: 'mime' | 'owner' } {
  if (driveCfg?.allowedMime && !isMimeAllowed(meta.mimeType, driveCfg.allowedMime)) {
    return { ok: false, reason: 'mime' };
  }
  if (driveCfg?.ownerAllowlist?.length) {
    const owners = (meta.owners as any[])?.map((o: any) => o?.emailAddress || o?.displayName).filter(Boolean) || [];
    if (!isOwnerAllowed(owners, driveCfg.ownerAllowlist)) {
      return { ok: false, reason: 'owner' };
    }
  }
  return { ok: true };
}

async function tryExportSheetAsXlsx(interaction: ChatInputCommandInteraction, svc: any, fileId: string, meta: any): Promise<boolean> {
  try {
    const xlsxBuf = await svc.exportDriveFile(
      fileId,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    const baseName = String(meta.name || fileId);
    const fileName = baseName.endsWith('.xlsx') ? baseName : `${baseName}.xlsx`;
    await interaction.editReply({
      content: t('files.read.downloadedSheet') || 'Завантажено таблицю як .xlsx',
      files: [{ attachment: xlsxBuf, name: fileName }],
    });
    return true;
  } catch {
    return false;
  }
}

function buildSourceLink(meta: any, driveCfg: any): string {
  const linkAllowed = !(driveCfg?.hideWebLink);
  return linkAllowed ? String(meta.webViewLink || '') : '';
}

async function replyQuickEmbed(interaction: ChatInputCommandInteraction, title: string, description: string, link?: string): Promise<void> {
  const embed = new EmbedBuilder().setTitle(title).setDescription(description).setColor(0x22c55e).setTimestamp();
  if (link) {
    const linkBtn = new ButtonBuilder().setLabel('Джерело').setStyle(ButtonStyle.Link).setURL(link);
    const rowLink = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(linkBtn);
    await interaction.editReply({ embeds: [embed], components: [rowLink] });
  } else {
    await interaction.editReply({ embeds: [embed] });
  }
}

function createTextSession(
  deps: ReadDeps,
  args: { fileId: string; fileName: string; chunks: string[]; link?: string }
): { sid: string; rows: ActionRowBuilder<MessageActionRowComponentBuilder>[] } {
  const sid = deps.generateSessionId('txt');
  const sessionObj: { fileId: string; fileName: string; chunks: string[]; createdAt: number; link?: string } = {
    fileId: args.fileId,
    fileName: args.fileName,
    chunks: args.chunks,
    createdAt: Math.floor(Date.now() / 1000),
  };
  if (args.link) sessionObj.link = args.link;
  deps.textSessions.set(sid, sessionObj);

  const openBtn = new ButtonBuilder()
    .setCustomId(deps.buildTextCustomId({ sid, page: 1 }))
    .setLabel('Показати ще')
    .setStyle(ButtonStyle.Primary);
  const closeBtn = new ButtonBuilder()
    .setCustomId(deps.buildTextCustomId({ sid, page: 1, action: 'close' }))
    .setLabel(t('files.search.buttons.close'))
    .setStyle(ButtonStyle.Danger);
  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(openBtn, closeBtn);
  const rows: ActionRowBuilder<MessageActionRowComponentBuilder>[] = [row];
  if (args.link) {
    const linkBtn = new ButtonBuilder().setLabel('Джерело').setStyle(ButtonStyle.Link).setURL(args.link);
    const rowLink = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(linkBtn);
    rows.push(rowLink);
  }
  return { sid, rows };
}

export async function handleReadTextFlow(
  interaction: ChatInputCommandInteraction,
  options: { fileId: string },
  deps: ReadDeps
): Promise<void> {
  const {
    config,
    getGoogleService,
    isMimeAllowed,
    isOwnerAllowed,
    isTooLarge,
    sanitizeTextForChat,
    buildPaginatedChunks,
    summarizeTlDr,
    getSubcommandTitle,
    mapGoogleApiErrorToMessage,
  } = deps;

  const svc = getGoogleService(interaction);
  if (!svc) {
    await interaction.editReply({ content: t('files.error.serviceUnavailable') });
    return;
  }

  try {
    const meta = await svc.getDriveFileMetadata(options.fileId);
    if (!meta || !meta.mimeType) {
      await interaction.editReply({ content: t('files.error.metadata') });
      return;
    }

    const driveCfg = config.drive;
    const policy = isAllowedByPolicy(meta, driveCfg, isMimeAllowed, isOwnerAllowed);
    if (!policy.ok) {
      await interaction.editReply({ content: policy.reason === 'mime' ? t('files.policy.disallowedMime') : t('files.policy.deniedOwner') });
      return;
    }

    const sizeBytes = Number(meta.size || 0) || 0;
    const tooLarge = isTooLarge(sizeBytes, (driveCfg?.fileMaxSizeMb ?? 0));

    // Якщо це Google Sheets — віддаємо як .xlsx вкладення
    if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
      const done = await tryExportSheetAsXlsx(interaction, svc, options.fileId, meta);
      if (done) return;
    }

    const extracted = await (svc).extractTextForChat(options.fileId);
    const safeText = String(extracted?.text || '').trim();

    if (!safeText) {
      if (tooLarge) {
        const link = buildSourceLink(meta, driveCfg);
        const sizeMb = (sizeBytes / (1024 * 1024)).toFixed(1);
        const summary = t('files.summary.largeFile', {
          name: String(meta.name || ''),
          mimeType: String(meta.mimeType || ''),
          size: sizeMb,
        });
        const linkText = link ? `\n${t('files.summary.link')}: ${link}` : '';
        await interaction.editReply({ content: `${summary}${linkText}` });
        return;
      }
      await interaction.editReply({ content: t('files.error.noText') });
      return;
    }

    const fileName = String(meta.name || options.fileId);
    const quick = sanitizeTextForChat(safeText, 1800);
    if (quick.length >= safeText.length) {
      const link = buildSourceLink(meta, driveCfg);
      const title = `📄 ${getSubcommandTitle('читати')}: ${fileName}`;
      await replyQuickEmbed(interaction, title, quick, link || undefined);
      return;
    }

    const tldr = summarizeTlDr(safeText, { budget: 800, minSentLen: 40 });
    const chunks = buildPaginatedChunks(safeText, { maxChunkLen: 1800 });
    const link = buildSourceLink(meta, config.drive);
    const sessionArgs = link ? { fileId: options.fileId, fileName, chunks, link } : { fileId: options.fileId, fileName, chunks };
    const { rows } = createTextSession(deps, sessionArgs as { fileId: string; fileName: string; chunks: string[]; link?: string });

    const embed = new EmbedBuilder()
      .setTitle(`📄 ${getSubcommandTitle('читати')}: ${fileName}`)
      .setDescription(tldr)
      .setColor(0x22c55e)
      .setTimestamp();
    await interaction.editReply({ embeds: [embed], components: rows });
  } catch (error) {
    const msg = mapGoogleApiErrorToMessage(error) || t('files.error.process');
    await interaction.editReply({ content: msg });
  }
}
