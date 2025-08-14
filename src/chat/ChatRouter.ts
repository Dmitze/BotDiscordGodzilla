import type { Client, Message } from 'discord.js';
import logger from '@/utils/logger';
import type { IntentDetector } from './IntentDetector';
import type { MemoryService } from './MemoryService';

export class ChatRouter {
  constructor(
    private readonly client: Client,
    private readonly memory: MemoryService,
    private readonly intents: IntentDetector
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

      const intent = this.intents.detect(content);

      switch (intent.type) {
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
