import { ActionRowBuilder, ButtonBuilder, ButtonStyle, EmbedBuilder, StringSelectMenuBuilder, type StringSelectMenuInteraction, type Interaction, type ButtonInteraction } from 'discord.js';
import { replyWithPrivacy } from '@/ui/reply';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

/**
 * DuplicateResolver — UI для дизамбигуации, когда найдено несколько одинаковых имен.
 * customId формат: dup|{scope}|{userId}|{nonce}|{action}|{page}
 *  - scope: произвольная строка (напр. "doc", "files")
 *  - action: prev|next|select|cancel
 */
export class DuplicateResolver {
  public static readonly PREFIX = 'dup|';

  static buildPage({
    scope,
    userId,
    nonce,
    files,
    page = 0,
    perPage = 5,
    title = 'Знайдено кілька збігів',
  }: {
    scope: string;
    userId: string;
    nonce: string;
    files: Pick<DriveFile, 'id' | 'name' | 'mimeType' | 'webViewLink' | 'owners'>[];
    page?: number;
    perPage?: number;
    title?: string;
  }) {
    const total = files.length;
    const pages = Math.max(1, Math.ceil(total / perPage));
    const current = Math.min(Math.max(0, page), pages - 1);
    const start = current * perPage;
    const pageItems = files.slice(start, start + perPage);

    const embed = new EmbedBuilder()
      .setTitle(title)
      .setColor(0x2f3136)
      .setDescription(
        pageItems
          .map((f, idx) => `• ${start + idx + 1}. ${f.name} (${f.mimeType})`)
          .join('\n') || '—'
      )
      .setFooter({ text: `Сторінка ${current + 1}/${pages} | Всього: ${total}` });

    const select = new StringSelectMenuBuilder()
      .setCustomId(this.customId(scope, userId, nonce, 'select', current))
      .setPlaceholder('Оберіть один варіант')
      .addOptions(
        ...pageItems.map((f) => ({
          label: f.name ?? '—',
          value: f.id,
          description: f.mimeType?.slice(0, 95) ?? undefined,
        }))
      );

    const prevBtn = new ButtonBuilder()
      .setCustomId(this.customId(scope, userId, nonce, 'prev', current))
      .setStyle(ButtonStyle.Secondary)
      .setLabel('◀ Prev')
      .setDisabled(current === 0);

    const nextBtn = new ButtonBuilder()
      .setCustomId(this.customId(scope, userId, nonce, 'next', current))
      .setStyle(ButtonStyle.Secondary)
      .setLabel('Next ▶')
      .setDisabled(current >= pages - 1);

    const cancelBtn = new ButtonBuilder()
      .setCustomId(this.customId(scope, userId, nonce, 'cancel', current))
      .setStyle(ButtonStyle.Danger)
      .setLabel('✖ Скасувати');

    const rows = [
      new ActionRowBuilder<StringSelectMenuBuilder>().addComponents(select),
      new ActionRowBuilder<ButtonBuilder>().addComponents(prevBtn, nextBtn, cancelBtn),
    ];

    return { embed, rows, current, pages };
  }

  static async handleComponent(interaction: Interaction, resolver: {
    fetchFiles: (ctx: { scope: string; userId: string; nonce: string }) => Promise<Pick<DriveFile, 'id' | 'name' | 'mimeType' | 'webViewLink' | 'owners'>[]>;
    onSelect: (ctx: { scope: string; userId: string; fileId: string; nonce: string }) => Promise<void>;
    title?: string;
    perPage?: number;
  }) {
    try {
      if (!('customId' in interaction)) return;
      const customId = (interaction as any).customId as string;
      if (!customId?.startsWith(this.PREFIX)) return;

      const parts = customId.split('|');
      // dup|scope|userId|nonce|action|page
      const scope = parts[1] ?? 'default';
      const userId = parts[2] ?? '0';
      const nonce = parts[3] ?? '0';
      const action = parts[4] ?? 'select';
      const page = Number(parts[5] ?? '0') || 0;

      // Только владелец может интерактить
      if ((interaction as any).user?.id && (interaction as any).user.id !== userId) {
        await replyWithPrivacy(interaction as any, { content: '⛔ Це меню не для вас' });
        return;
      }

      const files = await resolver.fetchFiles({ scope, userId, nonce });
      const perPage = resolver.perPage ?? 5;

      if (action === 'cancel') {
        await (interaction as any).update?.({ content: '❌ Скасовано', components: [], embeds: [] });
        return;
      }

      if (action === 'select' && (interaction as any).isStringSelectMenu?.()) {
        const select = interaction as StringSelectMenuInteraction;
        const [fileId] = select.values ?? [];
        if (!fileId) {
          await replyWithPrivacy(select as any, { content: '❌ Не обрано елемент' });
          return;
        }
        await resolver.onSelect({ scope, userId, fileId, nonce });
        await select.update({ content: '✅ Обрано', components: [], embeds: [] });
        return;
      }

      // Пагинация
      let nextPage = page;
      if (action === 'next') nextPage = page + 1;
      if (action === 'prev') nextPage = Math.max(0, page - 1);

      const { embed, rows } = this.buildPage({
        scope,
        userId,
        nonce,
        files,
        page: nextPage,
        perPage,
        title: resolver.title ?? 'Знайдено кілька збігів',
      });

      const button = interaction as ButtonInteraction;
      await button.update({ embeds: [embed], components: rows });
    } catch (error) {
      logger.error('❌ DuplicateResolver.handleComponent error', { error: String(error) });
      try {
        await replyWithPrivacy(interaction as any, { content: '❌ Помилка' });
      } catch {
        // ignore
      }
    }
  }

  static customId(scope: string, userId: string, nonce: string, action: 'prev' | 'next' | 'select' | 'cancel', page: number) {
    return `${this.PREFIX}${scope}|${userId}|${nonce}|${action}|${page}`;
  }
}
