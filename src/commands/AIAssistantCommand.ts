/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import * as xlsx from 'xlsx';
import { AnalyticsService } from '@/services/AnalyticsService';
import { t } from '@/i18n';

interface AIQueryResult {
  response: string;
  confidence: number;
  action?: string;
  actionData?: {
    type: string;
    format?: string;
  };
}

interface ValidationResult {
  isValid: boolean;
  errors: string[];
}

interface ValidationSchema {
  [key: string]: {
    required: boolean;
    type: string;
    maxLength?: number;
    sanitize?: string;
  };
}

export class AIAssistantCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;

  constructor(config: BotConfig, googleService?: GoogleService) {
    super('ai', t('ai.command.description'), config, {}, (builder: any) => {
      return builder
        .addStringOption((option: any) =>
          option
            .setName('запит')
            .setDescription(t('ai.opt.query.description'))
            .setRequired(true)
            .setMaxLength(1000)
        )
        .addStringOption((option: any) =>
          option
            .setName('контекст')
            .setDescription(t('ai.opt.context.description'))
            .setRequired(false)
            .setMaxLength(500)
        );
    });

    this.googleService = googleService;
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      // Перевірка прав доступу
      const hasAccess = await this.checkPermission(interaction);
      if (!hasAccess) {
        return;
      }

      // Миттєво відправляємо defer, щоб уникнути таймауту Discord ("Приложение не отвечает")
      await interaction.deferReply();

      // Валідація параметрів
      const commandOptions = {
        query: interaction.options.getString('запит'),
        context: interaction.options.getString('контекст'),
      };

      const validation = this.validateCommandOptions(commandOptions);
      if (!validation.isValid) {
        await interaction.editReply({
          content: t('ai.validation.failed', { errors: validation.errors.join('\n') }),
        });
        return;
      }

      // Логування події
      this.logSecurityEvent('ai_command_executed', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        command: this.name,
        query: commandOptions.query,
        context: commandOptions.context,
        guildId: interaction.guildId,
        channelId: interaction.channelId,
      });

      // Обробка запиту через AI
      const result = await this.processAIQuery(
        interaction.user.id,
        commandOptions.query || '',
        commandOptions.context
      );

      // Формування відповіді
      let response = `🤖 **${t('ai.reply.title')}**\n\n`;

      if (result.confidence < 0.7) {
        response += t('ai.reply.lowConfidence', { pct: Math.round(result.confidence * 100) });
        response += '\n';
      }

      response += t('ai.reply.query', { query: String(commandOptions.query) });
      response += '\n\n';
      response += t('ai.reply.answer', { answer: result.response });

      // Додавання контексту якщо є
      if (commandOptions.context) {
        response += '\n\n' + t('ai.reply.context', { context: String(commandOptions.context) });
      }

      // Додавання додаткової інформації
      if (result.actionData) {
        response += `\n\n` + t('ai.reply.action', { action: result.actionData.type });
        if (result.actionData.format) {
          response += ` ` + t('ai.reply.format', { format: result.actionData.format });
        }
      }

      // Відправка відповіді
      await interaction.editReply({
        content: response,
        ephemeral: false,
      });

      // Логування успішного виконання
      logger.info('AI command executed successfully', {
        type: 'command',
        command: this.name,
        component: 'AIAssistantCommand.onExecute',
        userTag: interaction.user.tag,
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        action: result.action,
        confidence: result.confidence,
        hasActionData: !!result.actionData,
      });
    } catch (error) {
      logger.error('AI Assistant command error', {
        type: 'command',
        command: this.name,
        component: 'AIAssistantCommand.onExecute',
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      const errorMessage = t('ai.error.process');

      if (interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermission(interaction: any): Promise<boolean> {
    try {
      const { PermissionManager } = await import('../core/PermissionManager');
      const permissionManager = new PermissionManager(this.config);
      const result = await permissionManager.checkPermission(
        interaction.user,
        interaction.member ?? null,
        interaction.commandName,
        interaction.channelId
      );

      if (!result.allowed) {
        this.logSecurityEvent('command_access_denied', {
          severity: 'medium',
          reason: result.reason,
          userLevel: result.userLevel,
          userId: interaction.user.id,
          ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
          channelId: interaction.channelId,
          command: interaction.commandName,
        });

        await interaction.reply({ content: t('ai.error.accessDenied', { reason: result.reason || '' }), ephemeral: true });
        return false;
      }

      logger.info('✅ Доступ до AI-команди дозволено', {
        type: 'security',
        event: 'permission_granted',
        component: 'AIAssistantCommand.checkPermission',
        userId: interaction.user.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        userLevel: result.userLevel,
        remainingUses: result.remainingUses,
      });
      return true;
    } catch (error) {
      logger.error('❌ Помилка перевірки прав у AI-команді', {
        type: 'security',
        component: 'AIAssistantCommand.checkPermission',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        userId: interaction.user?.id,
        ...(interaction.guildId ? { guildId: interaction.guildId } : {}),
        channelId: interaction.channelId,
        severity: 'high',
      });
      // За замовчуванням не пускати у разі помилки
      await interaction.reply({ content: t('ai.error.permissionCheck'), ephemeral: true });
      return false;
    }
  }

  /**
   * Валідація параметрів команди
   */
  private validateCommandOptions(options: Record<string, any>): ValidationResult {
    const validationSchema: ValidationSchema = {
      query: {
        required: true,
        type: 'string',
        maxLength: 1000,
        sanitize: 'ai_prompt',
      },
      context: {
        required: false,
        type: 'string',
        maxLength: 500,
        sanitize: 'ai_prompt',
      },
    };

    const errors: string[] = [];

    for (const [key, schema] of Object.entries(validationSchema)) {
      const value = options[key];

      if (schema.required && !value) {
        errors.push(`${key} є обов'язковим`);
        continue;
      }

      if (value && typeof value !== schema.type) {
        errors.push(`${key} має бути типу ${schema.type}`);
        continue;
      }

      if (value && schema.maxLength && value.length > schema.maxLength) {
        errors.push(`${key} не може бути довшим за ${schema.maxLength} символів`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
    };
  }

  /**
   * Логування події безпеки
   */
  private logSecurityEvent(eventType: string, data: Record<string, any>): void {
    logger.security(eventType, data['userId'] || 'unknown', {
      type: 'security',
      component: 'AIAssistantCommand',
      command: this.name,
      ...data,
    });
  }

  /**
   * Обробка AI запиту: визначення наміру та виконання дій
   */
  private async processAIQuery(
    _userId: string,
    query: string,
    _context?: string
  ): Promise<AIQueryResult> {
    const q = (query || '').toLowerCase();

    // 1) OCR зображень
    const intentOcrImage = /(картин|изображен|image|photo|png|jpg|jpeg)/i.test(q) && /(ocr|текст|прочитай|извле(ки|чи))/i.test(q);
    if (intentOcrImage) {
      if (!this.googleService) {
        return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
      }
      const folderId = this.config.google.driveFolderId;
      if (!folderId) {
        return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
      }
      const nameTokens = (query.match(/[\p{L}\p{N}\-_.]{2,}/giu) || []).filter(w => w.length >= 2).slice(0, 5);
      const nameQuery = nameTokens.join(' ').trim();
      let index = await this.googleService.getDriveIndex(folderId);
      if (!index) index = await this.googleService.buildDriveIndex(folderId, { ttlSeconds: 1800, recursive: true, maxDepth: -1 });
      const qlc = (s: string) => s.toLowerCase();
      const matchesName = (name?: string) => !nameQuery || qlc(name || '').includes(qlc(nameQuery));
      const isImage = (mt?: string) => !!(mt && /^image\//i.test(mt));
      const candidates = (index || []).filter(f => isImage((f as any).mimeType) && matchesName((f as any).name));
      if (!candidates.length) {
        return { response: t('ai.ocr.noImages'), confidence: 0.85, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
      }
      for (const f of candidates.slice(0, 5)) {
        try {
          const text = await this.googleService.extractTextFromImage(f as any);
          if (!text.trim()) continue;
          const preview = text.length > 1500 ? text.slice(0, 1500) + '…' : text;
          return { response: t('ai.ocr.result', { name: String((f as any).name ?? ''), id: String((f as any).id ?? ''), preview }), confidence: 0.9, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
        } catch (e) {
          logger.warn('OCR error', { type: 'command', component: 'AIAssistantCommand.processAIQuery', fileId: (f as any).id, err: e instanceof Error ? e.message : String(e) });
        }
      }
      return { response: t('ai.ocr.cannotRead'), confidence: 0.7, action: 'ocr_image', actionData: { type: 'analyze', format: 'text' } };
    }

    // 2) Аналітика таблиць за статусом (опціонально за місяць)
    const intentAnalytics = /(группируй|сгруппируй|групу(ва|пу)й|посчитай|підрахуй)/i.test(q) && /(статус|status)/i.test(q);
    if (intentAnalytics) {
      if (!this.googleService) {
        return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
      }
      const folderId = this.config.google.driveFolderId;
      if (!folderId) {
        return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
      }

      const nameTokens = (query.match(/[\p{L}\p{N}\-_.]{2,}/giu) || []).filter(w => w.length >= 2).slice(0, 5);
      const nameQuery = nameTokens.join(' ').trim();
      const monthMap: Record<string, number> = { 'январ':1, 'лют':2, 'фев':2, 'берез':3, 'март':3, 'квіт':4, 'апрел':4, 'май':5, 'трав':5, 'июн':6, 'черв':6, 'июл':7, 'лип':7, 'авг':8, 'серп':8, 'сен':9, 'верес':9, 'окт':10, 'жовт':10, 'нояб':11, 'листоп':11, 'дек':12, 'груд':12 };
      const monthKey = Object.keys(monthMap).find(k => q.includes(k));
      const monthNum = monthKey ? monthMap[monthKey] : undefined;

      const files = await this.googleService.listDriveFilesInFolder(folderId, { recursive: true, type: 'any', limit: 100, maxDepth: -1, ...(nameQuery ? { query: nameQuery } : {}) });
      const tableLike = files.filter(f => {
        const mt = (f.mimeType || '');
        return mt === 'application/vnd.google-apps.spreadsheet' || mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mt === 'application/vnd.ms-excel';
      });
      if (!tableLike.length) {
        return { response: t('ai.analytics.noTables'), confidence: 0.85, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
      }

      const readGoogleSheet = async (spreadsheetId: string): Promise<Array<Record<string, any>>> => {
        try {
          const sheetTitles = await this.googleService!.listSheets(spreadsheetId);
          const first = sheetTitles[0] || 'Лист1';
          const data = await this.googleService!.getSheetData(spreadsheetId, `${first}!A1:Z2000`);
          const rows = (data.values || []) as any[];
          if (!rows.length) return [];
          const headerRow: string[] = rows[0] ?? [];
          const rest = rows.slice(1);
          const headers = headerRow.map(h => String(h ?? '').trim());
          return rest.map(row => {
            const obj: Record<string, any> = {};
            headers.forEach((h, i) => { if (!h) return; obj[h] = row[i]; });
            return obj;
          });
        } catch { return []; }
      };
      const readExcelBuffer = (buf: Buffer): Array<Record<string, any>> => {
        try {
          const wb = xlsx.read(buf, { type: 'buffer' });
          const firstName = wb.SheetNames[0];
          if (!firstName) return [];
          const sheet = wb.Sheets[firstName];
          if (!sheet) return [];
          return xlsx.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '' });
        } catch { return []; }
      };

      const analytics = new AnalyticsService();
      for (const f of tableLike.slice(0, 5)) {
        try {
          const mt = f.mimeType || '';
          let rows: Array<Record<string, any>> = [];
          if (mt === 'application/vnd.google-apps.spreadsheet') rows = await readGoogleSheet(f.id!);
          else rows = readExcelBuffer(await this.googleService.downloadDriveFile(f.id!));
          if (!rows.length) continue;
          const schema = analytics.inferSchema(rows);
          const statusKey = schema.find(k => /статус|status/i.test(k)) || schema[0];
          const dateKey = schema.find(k => /дата|date/i.test(k));
          let filtered = rows;
          if (monthNum && dateKey) {
            filtered = rows.filter(r => {
              const v = r[dateKey!];
              const d = v ? new Date(v) : null;
              return d instanceof Date && !isNaN(+d) && d.getMonth() + 1 === monthNum;
            });
          }
          if (!statusKey) continue;
          const groups = analytics.groupBy(filtered, [statusKey]);
          const lines: string[] = [];
          for (const [gk, arr] of Object.entries(groups)) {
            const cnt = analytics.aggregate(arr as any[], null as any, 'count');
            lines.push(`${gk || '—'}: ${cnt}`);
          }
          const head = `Файл: ${f.name} (id: ${f.id})`;
          return { response: head + '\n' + lines.join('\n'), confidence: 0.9, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
        } catch (e) {
          logger.warn('Analytics failed for file', { type: 'command', component: 'AIAssistantCommand.processAIQuery', fileId: (f as any).id, err: e instanceof Error ? e.message : String(e) });
        }
      }
      return { response: 'Не вдалося виконати аналітику: дані порожні або структура невідома.', confidence: 0.7, action: 'table_analytics', actionData: { type: 'analyze', format: 'text' } };
    }

    // 3) Витягти текст з Docs/Word/PDF
    const intentExtractText = /(pdf|word|docx|docs?|документ|файл)/i.test(q) && /(покажи|выведи|витягни|извле(ки|чи)|текст)/i.test(q);
    if (intentExtractText) {
      if (!this.googleService) {
        return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
      }
      const folderId = this.config.google.driveFolderId;
      if (!folderId) {
        return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
      }
      const nameTokens = (query.match(/[\p{L}\p{N}\-_.]{2,}/giu) || []).filter(w => w.length >= 2).slice(0, 5);
      const nameQuery = nameTokens.join(' ').trim();
      let index = await this.googleService.getDriveIndex(folderId);
      if (!index) index = await this.googleService.buildDriveIndex(folderId, { ttlSeconds: 1800, recursive: true, maxDepth: -1 });
      const qlc = (s: string) => s.toLowerCase();
      const matchesName = (name?: string) => !nameQuery || qlc(name || '').includes(qlc(nameQuery));
      const isDocLike = (mt?: string) => mt === 'application/vnd.google-apps.document' || mt === 'application/pdf' || mt === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || mt === 'application/msword';
      const candidates = (index || []).filter(f => isDocLike((f as any).mimeType) && matchesName((f as any).name));
      if (!candidates.length) {
        return { response: 'Не знайдено відповідних документів (Docs/Word/PDF) за вашим описом.', confidence: 0.85, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
      }
      for (const f of candidates.slice(0, 5)) {
        try {
          const text = await this.googleService.extractTextFromFile(f as any);
          if (!text.trim()) continue;
          const preview = text.length > 1500 ? text.slice(0, 1500) + '…' : text;
          return { response: `Файл: ${(f as any).name} (id: ${(f as any).id})\n\n${preview}`, confidence: 0.9, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
        } catch (e) {
          logger.warn('Не вдалося витягти текст з документу', { type: 'command', component: 'AIAssistantCommand.processAIQuery', fileId: (f as any).id, fileName: (f as any).name, error: e instanceof Error ? e.message : String(e) });
        }
      }
      return { response: 'Не вдалося витягти текст: документи порожні або формат не підтримується. Уточніть назву файла або надішліть приклад.', confidence: 0.7, action: 'extract_text', actionData: { type: 'analyze', format: 'text' } };
    }

    // 4) Аналіз автобусів у таблицях
    const intentAnalyzeBuses = /автобус|bus/.test(q) && /(сколько|скiльки|скільки|осталось|залишил(о|ось)|бг|остат)/.test(q);
    if (intentAnalyzeBuses) {
      if (!this.googleService) {
        return { response: t('ai.error.googleUnavailable'), confidence: 0.6, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
      }
      const nameTokens = (query.match(/[\p{L}\p{N}\-_.]{2,}/giu) || []).filter(w => w.length >= 2).slice(0, 4);
      const folderId = this.config.google.driveFolderId;
      if (!folderId) {
        return { response: t('ai.error.missingDriveFolderId'), confidence: 0.7, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
      }
      const nameQuery = nameTokens.join(' ').trim();
      const baseOpts: { recursive?: boolean; type?: 'sheet' | 'folder' | 'any'; query?: string; limit?: number; pageToken?: string; maxDepth?: number } = { recursive: true, type: 'any', limit: 100, maxDepth: -1 };
      if (nameQuery) baseOpts.query = nameQuery;
      const files = await this.googleService.listDriveFilesInFolder(folderId, baseOpts);
      const candidates = files.filter(f => {
        const mt = f.mimeType || '';
        return mt === 'application/vnd.google-apps.spreadsheet' || mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mt === 'application/vnd.ms-excel';
      });
      if (!candidates.length) {
        return { response: 'Не знайдено придатних таблиць (Google Sheets/Excel) за вашим описом.', confidence: 0.85, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
      }
      const countBuses = (rows: Array<Record<string, any>>): number => {
        if (!rows.length) return 0;
        const norm = (s: unknown) => String(s ?? '').toLowerCase();
        const keys = Object.keys(rows[0] || {});
        const typeKey = keys.find(k => /(тип|вид|категор|vehicle|type|category)/i.test(k)) || (keys[0] as string | undefined);
        const statusKey = keys.find(k => /(статус|state|status)/i.test(k)) || ((keys[1] as string | undefined) ?? (keys[0] as string | undefined));
        if (!typeKey || !statusKey) return 0;
        let count = 0;
        for (const r of rows) {
          if (/автобус|bus/i.test(norm(r[typeKey])) && /(бг|остат|остал|залиш|в наличии|на складе)/i.test(norm(r[statusKey]))) count++;
        }
        return count;
      };
      const readGoogleSheet = async (spreadsheetId: string): Promise<Array<Record<string, any>>> => {
        try {
          const sheetTitles = await this.googleService!.listSheets(spreadsheetId);
          const first = sheetTitles[0] || 'Лист1';
          const data = await this.googleService!.getSheetData(spreadsheetId, `${first}!A1:Z1000`);
          const rows = (data.values || []) as any[];
          if (!rows.length) return [];
          const headerRow: string[] = rows[0] ?? [];
          const rest = rows.slice(1);
          const headers = headerRow.map(h => String(h ?? '').trim());
          return rest.map(row => { const obj: Record<string, any> = {}; headers.forEach((h, i) => { if (!h) return; obj[h] = row[i]; }); return obj; });
        } catch { return []; }
      };
      const readExcelBuffer = (buf: Buffer): Array<Record<string, any>> => {
        try {
          const wb = xlsx.read(buf, { type: 'buffer' });
          const firstName = wb.SheetNames[0];
          if (!firstName) return [];
          const sheet = wb.Sheets[firstName];
          if (!sheet) return [];
          return xlsx.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '' });
        } catch { return []; }
      };
      for (const f of candidates.slice(0, 5)) {
        try {
          const mt = f.mimeType || '';
          let rows: Array<Record<string, any>> = [];
          if (mt === 'application/vnd.google-apps.spreadsheet') rows = await readGoogleSheet(f.id!);
          else rows = readExcelBuffer(await this.googleService.downloadDriveFile(f.id!));
          if (!rows.length) continue;
          const total = countBuses(rows);
          return { response: `Файл: ${f.name} (id: ${f.id})\nРезультат: автобусів зі статусом БГ/залишок — ${total} шт.`, confidence: 0.92, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
        } catch (e) {
          logger.warn('Не вдалося обробити файл-кандидат', { type: 'command', component: 'AIAssistantCommand.processAIQuery', fileId: (f as any).id, fileName: (f as any).name, error: e instanceof Error ? e.message : String(e) });
        }
      }
      return { response: 'Не вдалося виконати аналіз: таблиці порожні або структура невідома.', confidence: 0.7, action: 'analyze_buses', actionData: { type: 'analyze', format: 'text' } };
    }

    // 5) Список таблиць/Excel у папці
    const intentListSheets = /таблиц|таблицы|лист(ы|и)?|sheets?|список.*таблиц|какие.*таблиц|google\s*диск|google\s*sheets/.test(q) && /какие|покажи|список|list|что|найд/i.test(q);
    if (intentListSheets) {
      try {
        if (!this.googleService) {
          return { response: 'GoogleService не доступний для цієї команди. Перевірте ініціалізацію сервісів або конфігурацію.', confidence: 0.6, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
        }
        const folderId = this.config.google.driveFolderId;
        if (!folderId) {
          return { response: 'Не налаштовано google.driveFolderId у конфігурації. Додайте ID каталогу з таблицями.', confidence: 0.6, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
        }
        const files = await this.googleService.listDriveFilesInFolder(folderId, { recursive: true, type: 'sheet', limit: 50, maxDepth: -1 });
        if (!files.length) {
          const spreadsheetId = this.config.google.spreadsheetId;
          if (spreadsheetId) {
            try {
              const sheetTitles = await this.googleService.listSheets(spreadsheetId);
              if (sheetTitles && sheetTitles.length >= 0) {
                return { response: 'Таблиці не знайдені у вказаній папці Google Drive. Проте доступ до таблиці з конфігурації працює. Переконайтесь, що потрібні файли знаходяться у цій папці або змініть GOOGLE_DRIVE_FOLDER_ID на правильну папку.', confidence: 0.92, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
              }
            } catch {
              return { response: 'Не вдалось отримати доступ до таблиць: папка порожня або недоступна, а також немає доступу до таблиці з конфігурації. Перевірте, що ви надали доступ сервісному акаунту та що файли знаходяться у вказаній папці.', confidence: 0.7, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
            }
          }
          return { response: 'Таблиці не знайдені у вказаній папці Google Drive.', confidence: 0.9, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
        }
        const lines = files.slice(0, 20).map((f, idx) => {
          const mime = f.mimeType || '';
          const label = mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ? 'Excel (.xlsx)' : mime === 'application/vnd.ms-excel' ? 'Excel (.xls)' : 'Google Sheets';
          return `${idx + 1}. ${f.name} [${label}] (id: ${f.id})`;
        });
        const more = files.length > 20 ? `\n… та ще ${files.length - 20}` : '';
        return { response: `Знайдено ${files.length} таблиць/Excel-файлів:\n\n${lines.join('\n')}${more}`, confidence: 0.95, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
      } catch (error) {
        logger.error('Помилка отримання списку таблиць', { type: 'command', component: 'AIAssistantCommand.processAIQuery', error: error instanceof Error ? error.message : String(error) });
        return { response: 'Сталася помилка при отриманні списку таблиць Google. Спробуйте пізніше або перевірте доступи.', confidence: 0.4, action: 'list_sheets', actionData: { type: 'list', format: 'text' } };
      }
    }

    // 6) Fallback
    const response = `Це тимчасова відповідь AI на запит: "${query}"`;
    return { response, confidence: 0.8, action: 'search', actionData: { type: 'search', format: 'text' } };
  }
}
