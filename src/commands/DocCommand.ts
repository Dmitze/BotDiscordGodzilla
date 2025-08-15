import type { SlashCommandBuilder } from 'discord.js';
import { EmbedBuilder } from 'discord.js';
import { BaseCommand } from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';
import type { CommandExecuteOptions } from '@/commands/BaseCommand';
import type { GoogleService } from '@/services/GoogleService';
import logger from '@/utils/logger';
import type { DocBlock } from '@/types/docs';

export class DocCommand extends BaseCommand {
  private readonly google: GoogleService | null;

  constructor(config: BotConfig, google?: GoogleService) {
    super(
      'doc',
      'Работа с Google Docs',
      config,
      {
        category: 'documents',
        usage: '/doc blocks <documentId> [limit]'
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('blocks')
              .setDescription('Показать структурированные блоки Google Docs')
              .addStringOption(opt =>
                opt
                  .setName('documentid')
                  .setDescription('ID документа Google Docs')
                  .setRequired(true)
              )
              .addIntegerOption(opt =>
                opt
                  .setName('limit')
                  .setDescription('Сколько блоков отобразить (1-25)')
                  .setRequired(false)
                  .setMinValue(1)
                  .setMaxValue(25)
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
      await interaction.reply({ content: 'Неизвестная подкоманда', ephemeral: true });
      return;
    }

    const documentId = interaction.options.getString('documentid', true);
    const limit = interaction.options.getInteger('limit') ?? 10;

    await interaction.deferReply({ ephemeral: false });

    try {
      if (!this.google) {
        await interaction.editReply('GoogleService недоступен. Обратитесь к администратору.');
        return;
      }

      const blocks: DocBlock[] = await this.google.getDocumentBlocks(documentId);
      const shown = Math.max(1, Math.min(25, limit));

      const embed = new EmbedBuilder()
        .setTitle('Структура документа Google Docs')
        .setDescription(`documentId: ${documentId}\nВсего блоков: ${blocks.length}\nПоказываю: ${Math.min(shown, blocks.length)}`)
        .setColor(0x3b82f6);

      const toPreview = (text?: string, max = 300): string => {
        const t = (text || '').trim();
        if (!t) return '[пусто]';
        const s = t.replace(/\s+/g, ' ');
        return s.length > max ? s.slice(0, max - 1) + '…' : s;
      };

      const fields = blocks.slice(0, shown).map((b, i) => {
        let name = `${i + 1}. ${b.kind}`;
        let value = '';
        switch (b.kind) {
          case 'heading':
            name = `${i + 1}. heading h${b.level}`;
            value = `📝 ${toPreview(b.text, 300)}`;
            break;
          case 'listItem':
            name = `${i + 1}. listItem${b.listId ? ` (${b.listId})` : ''}`;
            value = `• ${toPreview(b.text, 300)}`;
            break;
          case 'paragraph':
            name = `${i + 1}. paragraph`;
            value = toPreview(b.text, 300);
            break;
          case 'table': {
            const rowsCount = b.rows.length;
            const firstRow = b.rows[0];
            const colsCount = firstRow?.cells?.length ?? 0;
            name = `${i + 1}. table ${rowsCount}x${colsCount}`;
            if (rowsCount > 0 && colsCount > 0) {
              const firstRowText = firstRow!.cells.map(c => c.text).join(' | ');
              value = toPreview(firstRowText, 300);
            } else {
              value = '[empty table]';
            }
            break;
          }
          case 'footnote':
            name = `${i + 1}. footnote ${b.id}`;
            value = toPreview(b.text, 200);
            break;
        }
        if (!value) value = '[нет данных]';
        return { name, value } as const;
      });

      // Discord ограничение: max 25 полей в embed
      for (const f of fields.slice(0, 25)) {
        embed.addFields(f);
      }

      await interaction.editReply({ embeds: [embed] });
    } catch (error) {
      logger.error('Ошибка выполнения /doc blocks', {
        type: 'command',
        command: 'doc blocks',
        documentId,
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Ошибка при получении блоков документа. Проверьте ID и доступ.');
    }
  }
}
