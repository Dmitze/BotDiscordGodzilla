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
    getSubcommandTitle,
    sanitizeTextForChat,
    buildPaginatedChunks,
    summarizeTlDr,
    generateSessionId,
    buildTextCustomId,
    textSessions,
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
    if (driveCfg?.allowedMime && !isMimeAllowed(meta.mimeType, driveCfg.allowedMime)) {
      await interaction.editReply({ content: t('files.policy.disallowedMime') });
      return;
    }
    if (driveCfg?.ownerAllowlist?.length) {
      const owners = (meta.owners as any[])?.map((o: any) => o?.emailAddress || o?.displayName).filter(Boolean) || [];
      if (!isOwnerAllowed(owners, driveCfg.ownerAllowlist)) {
        await interaction.editReply({ content: t('files.policy.deniedOwner') });
        return;
      }
    }

    const sizeBytes = Number((meta as any).size || 0) || 0;
    const tooLarge = isTooLarge(sizeBytes, (driveCfg?.fileMaxSizeMb ?? 0));

    // Якщо це Google Sheets — віддаємо як .xlsx вкладення
    if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
      try {
        const xlsxBuf = await svc.exportDriveFile(
          options.fileId,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        const baseName = String(meta.name || options.fileId);
        const fileName = baseName.endsWith('.xlsx') ? baseName : `${baseName}.xlsx`;
        await interaction.editReply({
          content: t('files.read.downloadedSheet') || 'Завантажено таблицю як .xlsx',
          files: [{ attachment: xlsxBuf, name: fileName }],
        });
        return;
      } catch (e) {
        // Якщо експорт не вдався — переходимо до текстового фолу-бека нижче
      }
    }

    const extracted = await (svc).extractTextForChat(options.fileId);
    const safeText = String(extracted?.text || '').trim();

    if (!safeText) {
      if (tooLarge) {
        const linkAllowed = !(driveCfg?.hideWebLink);
        const link = linkAllowed ? String((meta as any).webViewLink || '') : '';
        const sizeMb = (sizeBytes / (1024 * 1024)).toFixed(1);
        const summary = t('files.summary.largeFile', {
          name: String(meta.name || ''),
          mimeType: String(meta.mimeType || ''),
          size: sizeMb,
        });
        const linkText = linkAllowed && link ? `\n${t('files.summary.link')}: ${link}` : '';
        await interaction.editReply({ content: `${summary}${linkText}` });
        return;
      }
      await interaction.editReply({ content: t('files.error.noText') });
      return;
    }

    const fileName = String(meta.name || options.fileId);
    const quick = sanitizeTextForChat(safeText, 1800);
    if (quick.length >= safeText.length) {
      const linkAllowed = !(driveCfg?.hideWebLink);
      const link = linkAllowed ? String((meta as any).webViewLink || '') : '';
      const embed = new EmbedBuilder()
        .setTitle(`📄 ${getSubcommandTitle('читати')}: ${fileName}`)
        .setDescription(quick)
        .setColor(0x22c55e)
        .setTimestamp();
      if (link) {
        const linkBtn = new ButtonBuilder().setLabel('Джерело').setStyle(ButtonStyle.Link).setURL(link);
        const rowLink = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(linkBtn);
        await interaction.editReply({ embeds: [embed], components: [rowLink] });
      } else {
        await interaction.editReply({ embeds: [embed] });
      }
      return;
    }

    const tldr = summarizeTlDr(safeText, { budget: 800, minSentLen: 40 });
    const chunks = buildPaginatedChunks(safeText, { maxChunkLen: 1800 });
    const sid = generateSessionId('txt');
    const linkAllowed = !(config.drive?.hideWebLink);
    const link = linkAllowed ? String((meta as any).webViewLink || '') : '';
    const sessionObj: { fileId: string; fileName: string; chunks: string[]; createdAt: number; link?: string } = {
      fileId: options.fileId,
      fileName,
      chunks,
      createdAt: Math.floor(Date.now() / 1000),
    };
    if (link) sessionObj.link = link;
    textSessions.set(sid, sessionObj);

    const openBtn = new ButtonBuilder()
      .setCustomId(buildTextCustomId({ sid, page: 1 }))
      .setLabel('Показати ще')
      .setStyle(ButtonStyle.Primary);
    const closeBtn = new ButtonBuilder()
      .setCustomId(buildTextCustomId({ sid, page: 1, action: 'close' }))
      .setLabel(t('files.search.buttons.close'))
      .setStyle(ButtonStyle.Danger);
    const row = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(openBtn, closeBtn);
    const rows: ActionRowBuilder<MessageActionRowComponentBuilder>[] = [row];
    if (link) {
      const linkBtn = new ButtonBuilder().setLabel('Джерело').setStyle(ButtonStyle.Link).setURL(link);
      const rowLink = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(linkBtn);
      rows.push(rowLink);
    }

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
