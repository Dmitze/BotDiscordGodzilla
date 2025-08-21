import { Client, Message, EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import logger from '@/utils/logger';
import type { IntentDetector } from './IntentDetector';
import type { MemoryService } from './MemoryService';
import type { DriveIndexerService, DriveSearchResult } from '@/services/DriveIndexerService';
import { tokenizeQuery, buildSnippet, highlightSnippet } from '@/utils/highlight';

export class ChatRouter {
  constructor(
    private readonly client: Client,
    private readonly memory: MemoryService,
    private readonly intents: IntentDetector,
    private readonly getService?: (name: string) => unknown
  ) {}

  bind(): void {
    this.client.on('messageCreate', (m: Message) => {
      void this.handleMessage(m);
    });
    logger.info('💬 ChatRouter bound to messageCreate');
  }

  private handleMessage = async (msg: Message): Promise<void> => {
    try {
      if (!msg || !msg.author || msg.author.bot) return;
      const content = (msg.content || '').trim();
      if (!content) return;

      const meta = {
        type: 'chat',
        event: 'message_in',
        userId: msg.author.id,
        channelId: msg.channelId,
        messageId: msg.id,
      } as const;
      logger.info('chat_message_in', meta);

      const intent = await this.intents.detectWithAI(content);

      switch (intent.type) {
        case 'SEARCH':
          await this.replySearch(msg, intent.params?.['query'] || content);
          break;
        case 'HELP':
          await this.replyHelp(msg);
          break;
        case 'ANALYZE_SHEET':
          await this.replyAnalyzeSheet(msg);
          break;
        case 'ANALYZE_FILE':
          await this.replyAnalyzeFile(msg);
          break;
        case 'QNA_GENERAL':
          await this.replyQna(msg, content);
          break;
        default:
          await this.replyUnknown(msg);
      }
    } catch (e) {
      logger.error('chat_handle_error', {
        type: 'chat',
        component: 'ChatRouter',
        error: e instanceof Error ? e.message : String(e),
      });
      try {
        await msg.reply('❌ Ошибка обработки сообщения. Попробуйте уточнить запрос.');
      } catch (replyErr) {
        logger.debug('reply_failed_suppressed', {
          type: 'chat',
          component: 'ChatRouter',
          error: replyErr instanceof Error ? replyErr.message : String(replyErr),
        });
      }
    }
  };

  private async replySearch(msg: Message, queryRaw: string): Promise<void> {
    const query = (queryRaw || '').trim();
    if (!query) {
      await msg.reply('Вкажіть запит для пошуку. Приклад: "пошук договор поставки"');
      return;
    }
    const svc = (this.getService?.('driveIndexer') ?? undefined) as DriveIndexerService | undefined;
    if (!svc) {
      await msg.reply('Пошук наразі недоступний. Сервіс індексації не активний.');
      return;
    }
    try {
      const results: DriveSearchResult[] = await svc.search(query, 5);
      if (!results.length) {
        await msg.reply('Нічого не знайдено за вашим запитом. Спробуйте інші ключові слова.');
        return;
      }
      const terms = tokenizeQuery(query);
      const embeds = results.slice(0, 3).map(r => {
        const e = new EmbedBuilder()
          .setColor('#2b6cb0')
          .setTitle(this.decorateTitle(r.file.name, r.file.mimeType))
          .setDescription(highlightSnippet(buildSnippet(r.file.snippet || '', terms, 240), terms))
          .addFields(
            ...(r.file.modifiedTime ? [{ name: 'Оновлено', value: new Date(r.file.modifiedTime).toLocaleString('uk-UA') }] : []),
            ...(Array.isArray(r.file.owners) && r.file.owners.length
              ? [{ name: 'Власники', value: r.file.owners.join(', ') }]
              : []),
            ...(typeof r.file.size === 'number' ? [{ name: 'Розмір', value: `${r.file.size} B` }] : [])
          );
        return e;
      });

      const buttons = new ActionRowBuilder<ButtonBuilder>().addComponents(
        ...results.slice(0, 3).map(r =>
          new ButtonBuilder()
            .setCustomId(`search|expand|${r.file.id}`)
            .setLabel('Розгорнути')
            .setStyle(ButtonStyle.Primary)
        )
      );

      await msg.reply({ embeds, components: [buttons] });
    } catch (e) {
      logger.error('search_reply_failed', { error: e instanceof Error ? e.message : String(e) });
      await msg.reply('❌ Сталася помилка під час пошуку. Спробуйте пізніше.');
    }
  }

  private decorateTitle(name: string, mime?: string): string {
    const icon = this.mimeIcon(mime || '');
    return `${icon} ${name}`;
  }

  private mimeIcon(mime: string): string {
    if (/google-apps.document/.test(mime)) return '📄';
    if (/pdf/.test(mime)) return '📑';
    if (/image\//.test(mime)) return '🖼️';
    if (/sheet|excel|spreadsheet/.test(mime)) return '📊';
    return '📁';
  }

  private async replyHelp(msg: Message): Promise<void> {
    await msg.reply(
      'Я ассистент. Могу: 1) анализировать таблицы Google Sheets; 2) анализировать файлы Google Drive; 3) отвечать на вопросы. Примеры: "проанализируй таблицу "Личный состав"", "сколько записей в "Отчёт_июль"", "что в файле с id ..."'
    );
  }

  private async replyAnalyzeSheet(msg: Message): Promise<void> {
    await msg.reply(
      'Принято: анализ таблицы. Уточните имя листа/таблицы в кавычках или ID. Функция анализа будет активирована на следующем шаге рефакторинга.'
    );
  }

  private async replyAnalyzeFile(msg: Message): Promise<void> {
    await msg.reply(
      'Принято: анализ файла/документа. Уточните ID файла. Функция анализа будет активирована на следующем шаге рефакторинга.'
    );
  }

  private async replyQna(msg: Message, content: string): Promise<void> {
    this.memory.addTurn(msg.channelId, msg.author.id, {
      role: 'user',
      content,
      ts: Date.now(),
    });
    await msg.reply('Принято. Чат-режим активен. Генерация ответа ИИ будет подключена на следующем шаге.');
  }

  private async replyUnknown(msg: Message): Promise<void> {
    await msg.reply(
      'Не понял запрос. Примеры: "проанализируй таблицу "Сводка"", "сколько записей в "Отчёт"", "что в файле с id ..."'
    );
  }
}
