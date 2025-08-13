import type { Client, Message } from 'discord.js';
import logger from '@/utils/logger';
import { IntentDetector } from './IntentDetector';
import { MemoryService } from './MemoryService';

export class ChatRouter {
  constructor(
    private readonly client: Client,
    private readonly memory: MemoryService,
    private readonly intents: IntentDetector
  ) {}

  bind(): void {
    this.client.on('messageCreate', this.handleMessage);
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
        case 'HELP': {
          await msg.reply(
            'Я ассистент. Могу: 1) анализировать таблицы Google Sheets; 2) анализировать файлы Google Drive; 3) отвечать на вопросы. Примеры: "проанализируй таблицу \"Личный состав\"", "сколько записей в \"Отчёт_июль\"", "что в файле с id ..."'
          );
          break;
        }
        case 'ANALYZE_SHEET': {
          // Заглушка — подключим SheetsAnalyzer на следующем шаге
          await msg.reply(
            'Принято: анализ таблицы. Уточните имя листа/таблицы в кавычках или ID. Функция анализа будет активирована на следующем шаге рефакторинга.'
          );
          break;
        }
        case 'ANALYZE_FILE': {
          // Заглушка — подключим DriveAnalyzer на следующем шаге
          await msg.reply(
            'Принято: анализ файла/документа. Уточните ID файла. Функция анализа будет активирована на следующем шаге рефакторинга.'
          );
          break;
        }
        case 'QNA_GENERAL': {
          // Заглушка — подключим AIService.generate на следующем шаге
          this.memory.addTurn(msg.channelId, msg.author.id, {
            role: 'user',
            content,
            ts: Date.now(),
          });
          await msg.reply('Принято. Чат-режим активен. Генерация ответа ИИ будет подключена на следующем шаге.');
          break;
        }
        default: {
          await msg.reply(
            'Не понял запрос. Примеры: "проанализируй таблицу \"Сводка\"", "сколько записей в \"Отчёт\"", "что в файле с id ..."'
          );
        }
      }
    } catch (e) {
      logger.error('chat_handle_error', {
        type: 'chat',
        component: 'ChatRouter',
        error: e instanceof Error ? e.message : String(e),
      });
      try {
        await msg.reply('❌ Ошибка обработки сообщения. Попробуйте уточнить запрос.');
      } catch {}
    }
  };
}
