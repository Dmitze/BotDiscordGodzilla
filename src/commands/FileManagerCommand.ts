/**
 * Команда для роботи з Google Drive та різними форматами файлів
 * Включає пошук, читання та аналіз файлів
 */

import { AttachmentBuilder, EmbedBuilder, ChatInputCommandInteraction } from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import type { drive_v3 } from 'googleapis';
import type { GoogleService } from '@/services/GoogleService';
import type { AIService } from '@/services/AIService';
import type { SheetsContextService } from '@/services/SheetsContextService';

interface FileSearchOptions {
  query: string;
  folder?: string;
}

interface FileReadOptions {
  fileId: string;
}

interface FileAnalysisOptions {
  fileId: string;
  analysisType: 'summary' | 'detailed' | 'key_points';
}

interface FileReportOptions {
  fileId: string;
  format: 'txt' | 'pdf' | 'docx';
}

interface FileResult {
  success: boolean;
  message: string;
  data?: any;
  file?: Buffer;
  fileName?: string;
}

interface ValidationResult {
  isValid: boolean;
  errors: string[];
  data?: any;
}

export class FileManagerCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('файли', '📁 Робота з файлами в Google Drive', config, {}, (builder: any) => {
      return builder
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('пошук')
            .setDescription('Пошук файлів у Google Drive')
            .addStringOption((option: any) =>
              option
                .setName('запит')
                .setDescription('Назва файлу для пошуку')
                .setRequired(true)
                .setMaxLength(200)
            )
            .addStringOption((option: any) =>
              option
                .setName('папка')
                .setDescription('ID папки для пошуку (опціонально)')
                .setRequired(false)
                .setMaxLength(50)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('читати')
            .setDescription('Читати вміст файлу')
            .addStringOption((option: any) =>
              option
                .setName('id')
                .setDescription('ID файлу в Google Drive')
                .setRequired(true)
                .setMaxLength(50)
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('аналіз')
            .setDescription('AI-аналіз вмісту файлу')
            .addStringOption((option: any) =>
              option
                .setName('id')
                .setDescription('ID файлу в Google Drive')
                .setRequired(true)
                .setMaxLength(50)
            )
            .addStringOption((option: any) =>
              option
                .setName('тип')
                .setDescription('Тип аналізу')
                .setRequired(false)
                .addChoices(
                  { name: 'Короткий зміст', value: 'summary' },
                  { name: 'Детальний аналіз', value: 'detailed' },
                  { name: 'Ключові моменти', value: 'key_points' }
                )
            )
        )
        .addSubcommand((subcommand: any) =>
          subcommand
            .setName('звіт')
            .setDescription('Створити звіт на основі файлу')
            .addStringOption((option: any) =>
              option
                .setName('id')
                .setDescription('ID файлу в Google Drive')
                .setRequired(true)
                .setMaxLength(50)
            )
            .addStringOption((option: any) =>
              option
                .setName('формат')
                .setDescription('Формат звіту')
                .setRequired(false)
                .addChoices(
                  { name: 'Текст', value: 'txt' },
                  { name: 'PDF', value: 'pdf' },
                  { name: 'Word', value: 'docx' }
                )
            )
        );
    });
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

      const subcommand = interaction.options.getSubcommand();

      // Валідація параметрів
      const commandOptions = this.extractOptions(interaction, subcommand);
      const validation = this.validateOptions(commandOptions, subcommand);

      if (!validation.isValid) {
        await interaction.reply({
          content: `❌ Помилка валідації:\n${validation.errors.join('\n')}`,
          ephemeral: true,
        });
        return;
      }

      // Логування події
      this.logSecurityEvent('file_manager_command_executed', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        subcommand,
        options: validation.data,
      });

      // Відповідь про обробку
      await interaction.deferReply();

      // Виконання підкоманди
      let result: FileResult;
      switch (subcommand) {
        case 'пошук':
          result = await this.handleSearch(interaction, validation.data as FileSearchOptions);
          break;
        case 'читати':
          result = await this.handleRead(interaction, validation.data as FileReadOptions);
          break;
        case 'аналіз':
          result = await this.handleAnalyze(interaction, validation.data as FileAnalysisOptions);
          break;
        case 'звіт':
          result = await this.handleReport(interaction, validation.data as FileReportOptions);
          break;
        default:
          throw new Error(`Невідома підкоманда: ${subcommand}`);
      }

      // Відправка результату
      await this.sendResult(interaction, result, subcommand);

      // Логування успішного виконання
      logger.info('File manager command executed successfully', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        subcommand,
        success: true,
      });
    } catch (error) {
      logger.error('File Manager command error', {
        error: error instanceof Error ? error.message : String(error),
        userId: interaction.user?.id,
        subcommand: (() => {
          try {
            return interaction.options.getSubcommand();
          } catch {
            return undefined;
          }
        })(),
      });

      const errorMessage =
        '❌ Помилка при роботі з файлами. Спробуйте ще раз або зверніться до адміністратора.';

      if (interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else if (interaction.replied) {
        await interaction.followUp({ content: errorMessage, ephemeral: true });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermission(_interaction: ChatInputCommandInteraction): Promise<boolean> {
    // TODO: Реалізувати перевірку прав доступу
    // Тимчасова реалізація - дозволяємо всім
    return true;
  }

  /**
   * Витяг параметрів з interaction
   */
  private extractOptions(interaction: ChatInputCommandInteraction, subcommand: string): any {
    const options: any = {};

    switch (subcommand) {
      case 'пошук':
        options.query = interaction.options.getString('запит');
        options.folder = interaction.options.getString('папка');
        break;
      case 'читати':
      case 'аналіз':
      case 'звіт':
        options.fileId = interaction.options.getString('id');
        break;
    }

    if (subcommand === 'аналіз') {
      options.analysisType = interaction.options.getString('тип') || 'summary';
    }

    if (subcommand === 'звіт') {
      options.format = interaction.options.getString('формат') || 'txt';
    }

    return options;
  }

  /**
   * Валідація параметрів
   */
  private validateOptions(options: any, subcommand: string): ValidationResult {
    const errors: string[] = [];

    switch (subcommand) {
      case 'пошук':
        if (!options.query || options.query.length < 2) {
          errors.push('Запит повинен містити мінімум 2 символи');
        }
        break;
      case 'читати':
      case 'аналіз':
      case 'звіт':
        if (!options.fileId || options.fileId.length < 10) {
          errors.push('ID файлу повинен містити мінімум 10 символів');
        }
        break;
    }

    return {
      isValid: errors.length === 0,
      errors,
      data: options,
    };
  }

  /**
   * Обробка пошуку файлів
   */
  private async handleSearch(
    interaction: ChatInputCommandInteraction,
    options: FileSearchOptions
  ): Promise<FileResult> {
    const svc = this.getGoogleService(interaction);
    if (!svc) {
      return { success: false, message: '❌ GoogleService недоступний' };
    }

    const folderId = options.folder || this.config.google?.driveFolderId;
    if (!folderId) {
      return {
        success: false,
        message:
          '❌ Не вказано ID папки. Додайте параметр "папка" або налаштуйте driveFolderId у конфігурації.',
      };
    }

    const files: drive_v3.Schema$File[] = await svc.listDriveFilesInFolder(folderId, {
      recursive: true,
      type: 'any',
      query: options.query,
      limit: 50,
      maxDepth: 4,
    });

    if (!files.length) {
      return {
        success: true,
        message: `🔍 Пошук: "${options.query}"\nПапка: ${folderId}\n\nНічого не знайдено.`,
      };
    }

    const lines = files.slice(0, 20).map((f, idx) => {
      const icon =
        f.mimeType === 'application/vnd.google-apps.folder'
          ? '📁'
          : f.mimeType === 'application/vnd.google-apps.spreadsheet'
            ? '📊'
            : f.mimeType === 'application/vnd.google-apps.document'
              ? '📄'
              : '📦';
      return `${idx + 1}. ${icon} ${f.name} — ${f.id}`;
    });

    const more = files.length > 20 ? `\n…та ще ${files.length - 20}` : '';
    const msg = `🔍 **Результати пошуку**\nЗапит: ${options.query}\nПапка: ${folderId}\nЗнайдено: ${files.length}\n\n${lines.join('\n')}${more}`;
    return { success: true, message: msg };
  }

  /**
   * Обробка читання файлу
   */
  private async handleRead(
    interaction: ChatInputCommandInteraction,
    options: FileReadOptions
  ): Promise<FileResult> {
    const svc = this.getGoogleService(interaction);
    if (!svc) return { success: false, message: '❌ GoogleService недоступний' };

    const meta = await svc.getDriveFileMetadata(options.fileId);
    if (!meta || !meta.mimeType)
      return { success: false, message: '❌ Неможливо отримати метадані файлу' };

    const isGApp = meta.mimeType.startsWith('application/vnd.google-apps');
    let fileBuf: Buffer;
    let fileName: string;

    if (isGApp) {
      // Експорт для Google типів
      if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
        fileBuf = await svc.exportDriveFile(
          options.fileId,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        fileName = `${meta.name || 'spreadsheet'}.xlsx`;
      } else if (meta.mimeType === 'application/vnd.google-apps.document') {
        fileBuf = await svc.exportDriveFile(options.fileId, 'application/pdf');
        fileName = `${meta.name || 'document'}.pdf`;
      } else if (meta.mimeType === 'application/vnd.google-apps.presentation') {
        fileBuf = await svc.exportDriveFile(options.fileId, 'application/pdf');
        fileName = `${meta.name || 'presentation'}.pdf`;
      } else {
        // За замовчуванням пробуємо PDF
        fileBuf = await svc.exportDriveFile(options.fileId, 'application/pdf');
        fileName = `${meta.name || 'file'}.pdf`;
      }
    } else {
      // Звичайний файл
      fileBuf = await svc.downloadDriveFile(options.fileId);
      fileName = meta.name || `${options.fileId}`;
    }

    return {
      success: true,
      message: `📄 **Файл завантажено**\n\nНазва: ${meta.name}\nТип: ${meta.mimeType}`,
      file: fileBuf,
      fileName,
    };
  }

  /**
   * Обробка аналізу файлу
   */
  private async handleAnalyze(
    interaction: ChatInputCommandInteraction,
    options: FileAnalysisOptions
  ): Promise<FileResult> {
    const analysisTypeName = this.getAnalysisTypeName(options.analysisType);

    const anyClient = interaction.client as any;
    const ai = anyClient?.serviceContainer?.get?.('ai') as AIService | undefined;
    const google = this.getGoogleService(interaction);
    const sheetsContext = anyClient?.serviceContainer?.get?.('sheetsContext') as
      | SheetsContextService
      | undefined;

    // Отримуємо текстову витримку з файлу для аналізу (offline)
    let contextText = '';
    try {
      if (google) {
        const meta = await google.getDriveFileMetadata(options.fileId);
        if (meta.mimeType === 'application/vnd.google-apps.document') {
          const buf = await google.exportDriveFile(options.fileId, 'text/plain');
          contextText = buf.toString('utf8').slice(0, 4000);
        } else if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
          const buf = await google.exportDriveFile(options.fileId, 'text/csv');
          contextText = buf.toString('utf8').slice(0, 4000);
        } else {
          contextText = `File: ${meta.name} (${meta.mimeType})`;
        }
      }
    } catch {}

    // Додатковий контекст із SheetsContextService (якщо є)
    let sheetCtxNote = '';
    try {
      if (sheetsContext) {
        const ctx = await (sheetsContext as any).get?.('current');
        if (ctx) sheetCtxNote = `\nContext: ${JSON.stringify(ctx).slice(0, 500)}`;
      }
    } catch {}

    let analysis = `Тип аналізу: ${analysisTypeName}\n${sheetCtxNote}`;
    if (ai) {
      try {
        const res = await (ai as any).generate?.(
          `Проаналізуй наступний вміст та надай ${analysisTypeName}:\n\n${contextText}`,
          { maxTokens: 512 }
        );
        if (res && typeof res.content === 'string') {
          analysis = res.content;
        }
      } catch {
        // Фолбек на локальне резюме без мережі
        analysis = `${analysis}\n\nЗведення (локальне): ${contextText.slice(0, 800)}`;
      }
    } else {
      analysis = `${analysis}\n\nЗведення (локальне): ${contextText.slice(0, 800)}`;
    }

    return {
      success: true,
      message: `🤖 **AI-аналіз файлу**\n\n${analysis}`,
    };
  }

  /**
   * Обробка створення звіту
   */
  private async handleReport(
    interaction: ChatInputCommandInteraction,
    options: FileReportOptions
  ): Promise<FileResult> {
    const google = this.getGoogleService(interaction);
    if (!google) return { success: false, message: '❌ GoogleService недоступний' };

    const meta = await google.getDriveFileMetadata(options.fileId);
    const baseName = (meta.name || 'report').replace(/\.[^/.]+$/, '');

    // Збираємо коротку інформацію про файл як вміст звіту
    let content = `Звіт по файлу\nНазва: ${meta.name}\nТип: ${meta.mimeType}\nОновлено: ${meta.modifiedTime}`;
    try {
      if (meta.mimeType === 'application/vnd.google-apps.document') {
        const buf = await google.exportDriveFile(options.fileId, 'text/plain');
        content += `\n\nФрагмент вмісту:\n${buf.toString('utf8').slice(0, 1000)}`;
      } else if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
        const buf = await google.exportDriveFile(options.fileId, 'text/csv');
        content += `\n\nПерші рядки (CSV):\n${buf.toString('utf8').slice(0, 1000)}`;
      }
    } catch {}

    if (options.format === 'txt') {
      return {
        success: true,
        message: '📋 Звіт сформовано (TXT)',
        file: Buffer.from(content, 'utf8'),
        fileName: `${baseName}-report.txt`,
      };
    }

    if (options.format === 'pdf') {
      const pdf = await this.renderPdf(content, baseName);
      return {
        success: true,
        message: '📋 Звіт сформовано (PDF)',
        file: pdf,
        fileName: `${baseName}-report.pdf`,
      };
    }

    // docx
    const docx = await this.renderDocx(content, baseName);
    return {
      success: true,
      message: '📋 Звіт сформовано (DOCX)',
      file: docx,
      fileName: `${baseName}-report.docx`,
    };
  }

  // Helpers
  private async renderPdf(text: string, title: string): Promise<Buffer> {
    return await new Promise<Buffer>((resolve, reject) => {
      try {
        // Dynamic require to avoid dependency at module load time
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const PDFDocument = require('pdfkit');
        const doc = new PDFDocument({ margin: 50 });
        const chunks: Buffer[] = [];
        doc.on('data', (d: any) => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(d)));
        doc.on('end', () => resolve(Buffer.concat(chunks)));
        doc.on('error', reject);

        doc.fontSize(18).text(title, { underline: true });
        doc.moveDown();
        doc.fontSize(12).text(text);
        doc.end();
      } catch (e) {
        reject(e as any);
      }
    });
  }

  private async renderDocx(text: string, title: string): Promise<Buffer> {
    // Dynamic require for docx
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const docx = require('docx');
    const { Document, Packer, Paragraph, HeadingLevel, TextRun } = docx;
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ text: title, heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ children: [new TextRun(text)] }),
          ],
        },
      ],
    });
    const buf = await Packer.toBuffer(doc);
    return Buffer.from(buf);
  }

  /**
   * Відправка результату
   */
  private async sendResult(
    interaction: ChatInputCommandInteraction,
    result: FileResult,
    subcommand: string
  ): Promise<void> {
    if (!result.success) {
      await interaction.editReply({ content: result.message });
      return;
    }

    const embed = new EmbedBuilder()
      .setTitle(`📁 ${this.getSubcommandTitle(subcommand)}`)
      .setDescription(result.message)
      .setColor(0x00ff00)
      .setTimestamp();

    if (result.file && result.fileName) {
      const attachment = new AttachmentBuilder(result.file, { name: result.fileName });
      await interaction.editReply({ embeds: [embed], files: [attachment] });
    } else {
      await interaction.editReply({ embeds: [embed] });
    }
  }

  /**
   * Отримання назви типу аналізу
   */
  private getAnalysisTypeName(type: FileAnalysisOptions['analysisType']): string {
    const typeNames: Record<string, string> = {
      summary: 'Короткий зміст',
      detailed: 'Детальний аналіз',
      key_points: 'Ключові моменти',
    };
    return typeNames[type] || type;
  }

  /**
   * Отримання заголовку підкоманди
   */
  private getSubcommandTitle(subcommand: string): string {
    const titles: Record<string, string> = {
      пошук: 'Пошук файлів',
      читати: 'Читання файлу',
      аналіз: 'AI-аналіз',
      звіт: 'Створення звіту',
    };

    return titles[subcommand] || 'Робота з файлами';
  }

  /**
   * Логування події безпеки
   */
  private logSecurityEvent(eventType: string, data: Record<string, any>): void {
    logger.info('security_event', {
      eventType,
      ...data,
    });
  }

  /**
   * Отримання GoogleService через ServiceContainer
   */
  private getGoogleService(interaction: ChatInputCommandInteraction): GoogleService | undefined {
    try {
      const anyClient = interaction.client as any;
      const svc = anyClient?.serviceContainer?.get?.('google');
      return svc as GoogleService | undefined;
    } catch (e) {
      logger.warn('FileManager: не вдалося отримати GoogleService', {
        component: 'FileManagerCommand',
        event: 'service_resolve_failed',
        error: String(e),
      });
      return undefined;
    }
  }
}
