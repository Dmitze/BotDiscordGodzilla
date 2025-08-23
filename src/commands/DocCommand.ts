import {
  EmbedBuilder,
  type SlashCommandBuilder,
  ButtonBuilder,
  ButtonStyle,
  ActionRowBuilder,
  type MessageActionRowComponentBuilder,
} from 'discord.js';
import { BaseCommand, type CommandExecuteOptions, type CommandAutocompleteOptions } from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import logger from '@/utils/logger';
import type { DocBlock } from '@/types/docs';
import { t } from '@/i18n';
import { replyWithPrivacy } from '@/ui/reply';

export class DocCommand extends BaseCommand {
  private readonly google: GoogleService | null;

  constructor(config: BotConfig, google?: GoogleService) {
    super(
      'doc',
      t('doc.command.description'),
      config,
      {
        category: 'documents',
        usage: '/doc blocks <documentId|url> [limit]',
        i18n: { nameKey: 'commands.doc.name', descriptionKey: 'doc.command.description' }
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('blocks')
              .setDescription(t('doc.sub.blocks.description'))
              .addStringOption(opt =>
                opt
                  .setName('documentid')
                  .setDescription(t('doc.opt.documentid.description'))
                  .setRequired(true)
                  .setAutocomplete(true)
              )
              .addIntegerOption(opt =>
                opt
                  .setName('limit')
                  .setDescription(t('doc.opt.limit.description'))
                  .setRequired(false)
                  .setMinValue(1)
                  .setMaxValue(25)
              )
              .addStringOption(opt =>
                opt
                  .setName('format')
                  .setDescription(t('doc.opt.format.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: 'short', value: 'short' },
                    { name: 'full', value: 'full' },
                    { name: 'headings', value: 'headings' },
                  )
              )
          );
        return builder;
      }
    );

    this.google = google ?? null;
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const interaction = options.interaction;

    const sub = interaction.options.getSubcommand();
    if (sub !== 'blocks') {
      await replyWithPrivacy(interaction as any, { content: t('doc.unknownSub') });
      return;
    }

    const documentInput = interaction.options.getString('documentid', true);
    const documentId = extractDocId(documentInput);
    if (!documentId) {
      await replyWithPrivacy(interaction as any, { content: t('doc.id.invalid') });
      return;
    }
    const limit = interaction.options.getInteger('limit') ?? 10;
    const format = (interaction.options.getString('format') ?? 'short') as FormatMode;

    await interaction.deferReply({ ephemeral: false });

    try {
      if (!this.google) {
        await interaction.editReply(t('doc.google.unavailable'));
        return;
      }

      const blocks: DocBlock[] = await this.google.getDocumentBlocks(documentId);

      const pageSize = Math.max(1, Math.min(25, limit));
      const page = 1;
      const { embed, components } = buildBlocksPage({
        blocks,
        documentId,
        page,
        pageSize,
        format,
      });

      await interaction.editReply({ embeds: [embed], components });
    } catch (error) {
      logger.error('Ошибка выполнения /doc blocks', {
        type: 'command',
        command: 'doc blocks',
        documentId,
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply(t('doc.error.fetchBlocks'));
    }
  }

  protected override async onAutocomplete(options: CommandAutocompleteOptions): Promise<void> {
    const interaction = options.interaction;
    try {
      if (!this.google) {
        await interaction.respond([]);
        return;
      }

      const folderId = this.config.drive?.folderId;
      if (!folderId) {
        await interaction.respond([]);
        return;
      }

      const query = (options.query || '').trim();
      // Ограничиваемся только Google Docs
      const { files } = await this.google.listDriveFiles({
        folderId,
        query,
        mimeIncludes: ['application/vnd.google-apps.document'],
        pageSize: 10,
      });

      const choices = files.slice(0, 10).map(f => ({
        name: f.name.length > 90 ? f.name.slice(0, 87) + '…' : f.name,
        value: f.id,
      }));

      await interaction.respond(choices);
    } catch (error) {
      logger.warn('Автодополнение /doc blocks не удалось', {
        type: 'command',
        command: 'doc blocks',
        event: 'autocomplete_failed',
        error: error instanceof Error ? error.message : String(error),
      });
      try {
        await interaction.respond([]);
      } catch { /* noop */ }
    }
  }

  protected override async onComponent(options: import('@/commands/BaseCommand').CommandComponentOptions): Promise<void> {
    const interaction = options.interaction;
    if (!interaction.isButton()) return;

    try {
      const payload = (options as any)?.context?.componentPayload as
        | { kind?: string; documentId?: string; page?: number; pageSize?: number; format?: string; ts?: number }
        | undefined;
      const isSigned = !!payload && payload.kind === 'docblk';
      const parsed = isSigned ? payload : parseCustomId(interaction.customId);
      if (!parsed || parsed.kind !== 'docblk') return;

      const { documentId, page, pageSize, format, ts } = parsed as {
        documentId?: string;
        page?: number;
        pageSize?: number;
        format?: FormatMode | string;
        ts?: number;
      };

      // Narrow required fields
      if (!documentId || typeof page !== 'number' || typeof pageSize !== 'number' || !format) {
        await replyWithPrivacy(interaction as any, { content: t('doc.sessionExpired') });
        return;
      }

      const docId: string = documentId;
      const pg: number = page;
      const pgSize: number = pageSize;
      const fmt: FormatMode = format as FormatMode;

      // Ограничение времени жизни: 10 минут
      const nowSec = Math.floor(Date.now() / 1000);
      if (ts && nowSec - ts > 10 * 60) {
        await replyWithPrivacy(interaction as any, { content: t('doc.sessionExpired') });
        return;
      }

      if (!this.google) {
        await replyWithPrivacy(interaction as any, { content: t('doc.google.unavailable') });
        return;
      }

      const blocks: DocBlock[] = await this.google.getDocumentBlocks(docId);
      const { embed, components } = buildBlocksPage({ blocks, documentId: docId, page: pg, pageSize: pgSize, format: fmt });

      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ embeds: [embed], components });
      } else {
        await interaction.update({ embeds: [embed], components });
      }
    } catch (error) {
      logger.error('Ошибка обработки кнопки DocCommand', {
        type: 'command',
        command: 'doc',
        event: 'component_error',
        error: error instanceof Error ? error.message : String(error),
      });
      if (!interaction.deferred && !interaction.replied) {
        await replyWithPrivacy(interaction as any, { content: t('doc.error.updatePage') });
      }
    }
  }
}

// Извлекает ID документа из полной ссылки или возвращает строку, если она похожа на ID
function extractDocId(input: string): string | null {
  const trimmed = input.trim();
  // Прямой ID (обычно base64-like: буквы, цифры, -, _)
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) return trimmed;
  try {
    const url = new URL(trimmed);
    // Формат: https://docs.google.com/document/d/<ID>/...
    const m = url.pathname.match(/\/document\/d\/([a-zA-Z0-9-_]+)/);
    if (m && m[1]) return m[1];
    // Альтернатива: https://docs.google.com/document/u/0/d/<ID>/...
    const m2 = url.pathname.match(/\/document\/u\/\d+\/d\/([a-zA-Z0-9-_]+)/);
    if (m2 && m2[1]) return m2[1];
    // Редкий формат: /open?id=<ID>
    const openId = url.searchParams.get('id');
    if (openId && /^[a-zA-Z0-9-_]{20,}$/.test(openId)) return openId;
  } catch {
    // not a URL
  }
  return null;
}

// ===== Вспомогательные типы и функции форматирования/пагинации =====
type FormatMode = 'short' | 'full' | 'headings';

function toPreview(text: string | undefined, max = 300): string {
  const str = (text || '').trim();
  if (!str) return t('doc.kinds.empty');
  const s = str.replace(/\s+/g, ' ');
  return s.length > max ? s.slice(0, max - 1) + '…' : s;
}

function formatBlock(b: DocBlock, indexBase: number, mode: FormatMode): { name: string; value: string } | null {
  const idx = indexBase;
  switch (mode) {
    case 'headings':
      if (b.kind === 'heading') {
        return { name: `${idx}. H${b.level}`, value: toPreview(b.text, 200) };
      }
      return null;
    case 'full': {
      switch (b.kind) {
        case 'heading':
          return { name: `${idx}. ${t('doc.kinds.heading', { level: b.level })}`, value: toPreview(b.text, 900) };
        case 'listItem':
          return { name: `${idx}. ${t('doc.kinds.listItem')}${b.listId ? ` (${b.listId})` : ''}`, value: `• ${toPreview(b.text, 900)}` };
        case 'paragraph':
          return { name: `${idx}. ${t('doc.kinds.paragraph')}`, value: toPreview(b.text, 900) };
        case 'table': {
          const rowsCount = b.rows.length;
          const firstRow = b.rows[0];
          const colsCount = firstRow && Array.isArray(firstRow.cells) ? firstRow.cells.length : 0;
          const header = t('doc.kinds.table', { rows: rowsCount, cols: colsCount });
          const firstRowText = rowsCount > 0 && colsCount > 0 && firstRow
            ? firstRow.cells.map(c => c.text).join(' | ')
            : t('doc.kinds.emptyTable');
          return { name: `${idx}. table ${header}`, value: toPreview(firstRowText, 900) };
        }
        case 'footnote':
          return { name: `${idx}. ${t('doc.kinds.footnote', { id: b.id })}`, value: toPreview(b.text, 600) };
      }
      // exhaustive
    }
    case 'short':
    default: {
      switch (b.kind) {
        case 'heading':
          return { name: `${idx}. ${t('doc.kinds.heading', { level: b.level })}`, value: `📝 ${toPreview(b.text, 300)}` };
        case 'listItem':
          return { name: `${idx}. ${t('doc.kinds.listItem')}${b.listId ? ` (${b.listId})` : ''}`, value: `• ${toPreview(b.text, 300)}` };
        case 'paragraph':
          return { name: `${idx}. ${t('doc.kinds.paragraph')}`, value: toPreview(b.text, 300) };
        case 'table': {
          const rowsCount = b.rows.length;
          const firstRow = b.rows[0];
          const colsCount = firstRow && Array.isArray(firstRow.cells) ? firstRow.cells.length : 0;
          const header = t('doc.kinds.table', { rows: rowsCount, cols: colsCount });
          const firstRowText = rowsCount > 0 && colsCount > 0 && firstRow
            ? firstRow.cells.map(c => c.text).join(' | ')
            : t('doc.kinds.emptyTable');
          return { name: `${idx}. table ${header}`, value: toPreview(firstRowText, 300) };
        }
        case 'footnote':
          return { name: `${idx}. ${t('doc.kinds.footnote', { id: b.id })}`, value: toPreview(b.text, 200) };
      }
      // exhaustive
    }
  }
}

function buildBlocksPage(args: {
  blocks: DocBlock[];
  documentId: string;
  page: number;
  pageSize: number; // 1..25
  format: FormatMode;
}): { embed: EmbedBuilder; components: ActionRowBuilder<MessageActionRowComponentBuilder>[] } {
  const { blocks, documentId, page, pageSize, format } = args;
  const total = blocks.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const start = (safePage - 1) * pageSize;
  const slice = blocks.slice(start, start + pageSize);

  const embed = new EmbedBuilder()
    .setTitle(t('doc.embed.title'))
    .setDescription(t('doc.embed.desc', { documentId, total, page: safePage, totalPages, pageSize, format }))
    .setColor(0x3b82f6);

  let idx = start + 1;
  const fields: { name: string; value: string }[] = [];
  for (const b of slice) {
    const f = formatBlock(b, idx, format);
    if (f) fields.push(f);
    idx++;
  }
  for (const f of fields.slice(0, 25)) embed.addFields(f);

  const nowSec = Math.floor(Date.now() / 1000);
  const row = new ActionRowBuilder<MessageActionRowComponentBuilder>().addComponents(
    new ButtonBuilder()
      .setCustomId(buildCustomId({ documentId, page: 1, pageSize, format, ts: nowSec }))
      .setLabel('⏮')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === 1),
    new ButtonBuilder()
      .setCustomId(buildCustomId({ documentId, page: Math.max(1, safePage - 1), pageSize, format, ts: nowSec }))
      .setLabel('◀')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === 1),
    new ButtonBuilder()
      .setCustomId(buildCustomId({ documentId, page: Math.min(totalPages, safePage + 1), pageSize, format, ts: nowSec }))
      .setLabel('▶')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === totalPages),
    new ButtonBuilder()
      .setCustomId(buildCustomId({ documentId, page: totalPages, pageSize, format, ts: nowSec }))
      .setLabel('⏭')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage === totalPages),
  );

  return { embed, components: [row] };
}

function buildCustomId(args: { documentId: string; page: number; pageSize: number; format: FormatMode; ts?: number }): string {
  const { documentId, page, pageSize, format, ts } = args;
  const t = ts ?? Math.floor(Date.now() / 1000);
  // компактный custom_id, укладываемся в лимит Discord (<=100 символов)
  return `docblk|d=${documentId}|p=${page}|s=${pageSize}|f=${format}|t=${t}`;
}

function parseCustomId(customId: string):
  | { kind: 'docblk'; documentId: string; page: number; pageSize: number; format: FormatMode; ts?: number }
  | null {
  if (!customId.startsWith('docblk|')) return null;
  const parts = customId.split('|').slice(1);
  const map = new Map(parts.map(kv => {
    const i = kv.indexOf('=');
    return i > 0 ? [kv.slice(0, i), kv.slice(i + 1)] as const : [kv, ''];
  }));
  const d = map.get('d');
  const p = Number(map.get('p') || '1');
  const s = Number(map.get('s') || '10');
  const f = (map.get('f') || 'short') as FormatMode;
  const t = Number(map.get('t') || '');
  if (!d) return null;
  const result: { kind: 'docblk'; documentId: string; page: number; pageSize: number; format: FormatMode; ts?: number } = {
    kind: 'docblk',
    documentId: d,
    page: Number.isFinite(p) ? p : 1,
    pageSize: Number.isFinite(s) ? Math.max(1, Math.min(25, s)) : 10,
    format: f,
  };
  if (Number.isFinite(t)) {
    result.ts = t;
  }
  return result;
}
