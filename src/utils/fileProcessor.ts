/**
 * Модуль для роботи з Google Drive та різними форматами файлів
 * Включає читання PDF, Word, Google Docs та створення звітів
 * TypeScript версія
 */

import logger from './logger';
import fs from 'fs/promises';
import path from 'path';
import { google } from 'googleapis';

// Конфігурація
const FILE_CONFIG = {
  SUPPORTED_FORMATS: ['pdf', 'docx', 'doc', 'txt', 'gdoc'] as const,
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  TEMP_DIR: './data/tmp',
  DOWNLOAD_TIMEOUT: 30000, // 30 секунд
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000, // 1 секунда
};

interface FileStats {
  filesProcessed: number;
  filesDownloaded: number;
  filesAnalyzed: number;
  errors: number;
}

interface FileMetadata {
  id: string;
  name: string;
  mimeType: string;
  size: string;
  modifiedTime: string;
  webViewLink: string;
  createdTime?: string;
  description?: string;
}

interface FileContent {
  metadata: FileMetadata;
  content: string;
  fileType: string;
}

interface ReportData {
  title?: string;
  summary?: string;
  data?: any;
  conclusions?: string;
}

/**
 * Клас для роботи з файлами
 */
class FileProcessor {
  private drive: any;
  private docs: any;
  private stats: FileStats;

  constructor() {
    this.drive = null;
    this.docs = null;
    this.stats = {
      filesProcessed: 0,
      filesDownloaded: 0,
      filesAnalyzed: 0,
      errors: 0,
    };
    this.initializeGoogleAPIs();
  }

  /**
   * Ініціалізація Google APIs
   */
  private async initializeGoogleAPIs(): Promise<void> {
    try {
      if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
        logger.warn('Google Application Credentials not found');
        return;
      }

      const auth = new google.auth.GoogleAuth({
        keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
        scopes: [
          'https://www.googleapis.com/auth/drive.readonly',
          'https://www.googleapis.com/auth/documents.readonly',
        ],
      });

      this.drive = google.drive({ version: 'v3', auth });
      this.docs = google.docs({ version: 'v1', auth });

      logger.info('Google APIs initialized successfully');
    } catch (error) {
      logger.error('Failed to initialize Google APIs:', error);
    }
  }

  /**
   * Пошук файлів у Google Drive
   */
  async searchFiles(query: string, folderId?: string): Promise<FileMetadata[]> {
    try {
      if (!this.drive) {
        throw new Error('Google Drive API not initialized');
      }

      let searchQuery = `name contains '${this.sanitizeQuery(query)}'`;

      if (folderId) {
        searchQuery += ` and '${folderId}' in parents`;
      }

      const response = await this.drive.files.list({
        q: searchQuery,
        fields: 'files(id,name,mimeType,size,modifiedTime,webViewLink)',
        pageSize: 20,
      });

      const files = response.data.files || [];
      this.stats.filesProcessed += files.length;

      logger.info(`Found ${files.length} files for query: ${query}`);
      return files;
    } catch (error) {
      this.stats.errors++;
      logger.error('File search error:', error);
      throw new Error('Помилка пошуку файлів');
    }
  }

  /**
   * Отримання метаданих файлу
   */
  async getFileMetadata(fileId: string): Promise<FileMetadata> {
    try {
      if (!this.drive) {
        throw new Error('Google Drive API not initialized');
      }

      const response = await this.drive.files.get({
        fileId,
        fields: 'id,name,mimeType,size,modifiedTime,createdTime,webViewLink,description',
      });

      return response.data;
    } catch (error) {
      this.stats.errors++;
      logger.error('File metadata error:', error);
      throw new Error('Помилка отримання метаданих файлу');
    }
  }

  /**
   * Завантаження файлу
   */
  async downloadFile(fileId: string, fileName: string): Promise<string> {
    try {
      if (!this.drive) {
        throw new Error('Google Drive API not initialized');
      }

      await this.ensureTempDir();

      const filePath = path.join(FILE_CONFIG.TEMP_DIR, fileName);

      const response = await this.drive.files.get(
        {
          fileId,
          alt: 'media',
        },
        {
          responseType: 'stream',
        }
      );

      const writer = require('fs').createWriteStream(filePath);
      response.data.pipe(writer);

      return new Promise((resolve, reject) => {
        writer.on('finish', () => {
          this.stats.filesDownloaded++;
          logger.info(`File downloaded: ${fileName}`);
          resolve(filePath);
        });
        writer.on('error', reject);
      });
    } catch (error) {
      this.stats.errors++;
      logger.error('File download error:', error);
      throw new Error('Помилка завантаження файлу');
    }
  }

  /**
   * Читання Google Doc
   */
  async readGoogleDoc(documentId: string): Promise<string> {
    try {
      if (!this.docs) {
        throw new Error('Google Docs API not initialized');
      }

      const response = await this.docs.documents.get({
        documentId,
      });

      return this.extractTextFromGoogleDoc(response.data);
    } catch (error) {
      this.stats.errors++;
      logger.error('Google Doc reading error:', error);
      throw new Error('Помилка читання Google Doc');
    }
  }

  /**
   * Витяг тексту з Google Doc
   */
  private extractTextFromGoogleDoc(document: any): string {
    try {
      let text = '';

      if (document.body && document.body.content) {
        for (const element of document.body.content) {
          if (element.paragraph) {
            for (const paragraphElement of element.paragraph.elements) {
              if (paragraphElement.textRun) {
                text += paragraphElement.textRun.content;
              }
            }
            text += '\n';
          }
        }
      }

      return text.trim();
    } catch (error) {
      logger.error('Text extraction error:', error);
      return '';
    }
  }

  /**
   * Читання PDF файлу
   */
  async readPDF(filePath: string): Promise<string> {
    try {
      // Тут можна додати бібліотеку для читання PDF
      // Наприклад, pdf-parse або pdf2pic
      logger.warn('PDF reading not implemented yet');
      return 'PDF reading not implemented';
    } catch (error) {
      logger.error('PDF reading error:', error);
      throw new Error('Помилка читання PDF файлу');
    }
  }

  /**
   * Читання Word файлу
   */
  async readWord(filePath: string): Promise<string> {
    try {
      // Тут можна додати бібліотеку для читання Word
      // Наприклад, mammoth або docx
      logger.warn('Word reading not implemented yet');
      return 'Word reading not implemented';
    } catch (error) {
      logger.error('Word reading error:', error);
      throw new Error('Помилка читання Word файлу');
    }
  }

  /**
   * Читання текстового файлу
   */
  async readTextFile(filePath: string): Promise<string> {
    try {
      const content = await fs.readFile(filePath, 'utf8');
      return content;
    } catch (error) {
      logger.error('Text file reading error:', error);
      throw new Error('Помилка читання текстового файлу');
    }
  }

  /**
   * Читання змісту файлу
   */
  async readFileContent(fileId: string): Promise<FileContent> {
    try {
      const metadata = await this.getFileMetadata(fileId);
      let content = '';

      const fileType = this.getFileType(metadata.mimeType, metadata.name);

      switch (fileType) {
        case 'gdoc':
          content = await this.readGoogleDoc(fileId);
          break;

        case 'pdf':
          const pdfPath = await this.downloadFile(fileId, `${fileId}.pdf`);
          content = await this.readPDF(pdfPath);
          await this.cleanupTempFile(pdfPath);
          break;

        case 'docx':
        case 'doc':
          const wordPath = await this.downloadFile(fileId, `${fileId}.${fileType}`);
          content = await this.readWord(wordPath);
          await this.cleanupTempFile(wordPath);
          break;

        case 'txt':
          const txtPath = await this.downloadFile(fileId, `${fileId}.txt`);
          content = await this.readTextFile(txtPath);
          await this.cleanupTempFile(txtPath);
          break;

        default:
          throw new Error(`Непідтримуваний тип файлу: ${fileType}`);
      }

      this.stats.filesAnalyzed++;

      return {
        metadata,
        content,
        fileType,
      };
    } catch (error: any) {
      this.stats.errors++;
      logger.error('File content reading error:', error);
      throw new Error(`Помилка читання файлу: ${error.message}`);
    }
  }

  /**
   * Визначення типу файлу
   */
  private getFileType(mimeType: string, fileName: string): string {
    try {
      const mimeTypes: Record<string, string> = {
        'application/vnd.google-apps.document': 'gdoc',
        'application/pdf': 'pdf',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/msword': 'doc',
        'text/plain': 'txt',
      };

      if (mimeType && mimeTypes[mimeType]) {
        return mimeTypes[mimeType];
      }

      // Визначення за розширенням файлу
      const extension = path.extname(fileName).toLowerCase();
      const extensionMap: Record<string, string> = {
        '.pdf': 'pdf',
        '.docx': 'docx',
        '.doc': 'doc',
        '.txt': 'txt',
      };

      return extensionMap[extension] || 'unknown';
    } catch (error) {
      logger.error('File type detection error:', error);
      return 'unknown';
    }
  }

  /**
   * Створення звіту
   */
  async createReport(reportData: ReportData, format = 'txt'): Promise<string> {
    try {
      await this.ensureTempDir();

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const fileName = `report_${timestamp}.${format}`;
      const filePath = path.join(FILE_CONFIG.TEMP_DIR, fileName);

      switch (format) {
        case 'txt':
          await this.createTextReport(reportData, filePath);
          break;
        case 'pdf':
          await this.createPDFReport(reportData, filePath);
          break;
        case 'docx':
          await this.createWordReport(reportData, filePath);
          break;
        default:
          throw new Error(`Непідтримуваний формат звіту: ${format}`);
      }

      logger.info(`Report created: ${fileName}`);
      return filePath;
    } catch (error: any) {
      this.stats.errors++;
      logger.error('Report creation error:', error);
      throw new Error(`Помилка створення звіту: ${error.message}`);
    }
  }

  /**
   * Створення текстового звіту
   */
  private async createTextReport(reportData: ReportData, filePath: string): Promise<void> {
    try {
      let content = '';

      if (reportData.title) {
        content += `ЗВІТ: ${reportData.title}\n`;
        content += '='.repeat(50) + '\n\n';
      }

      if (reportData.summary) {
        content += `Короткий зміст:\n${reportData.summary}\n\n`;
      }

      if (reportData.data) {
        content += `Дані:\n${JSON.stringify(reportData.data, null, 2)}\n\n`;
      }

      if (reportData.conclusions) {
        content += `Висновки:\n${reportData.conclusions}\n\n`;
      }

      content += `Створено: ${new Date().toLocaleString('uk-UA')}\n`;

      await fs.writeFile(filePath, content, 'utf8');
    } catch (error) {
      logger.error('Text report creation error:', error);
      throw error;
    }
  }

  /**
   * Створення PDF звіту
   */
  private async createPDFReport(reportData: ReportData, filePath: string): Promise<void> {
    try {
      // Тут можна додати бібліотеку для створення PDF
      // Наприклад, puppeteer або jsPDF
      logger.warn('PDF report creation not implemented yet');

      // Тимчасова реалізація - створюємо текстовий файл
      await this.createTextReport(reportData, filePath.replace('.pdf', '.txt'));
    } catch (error) {
      logger.error('PDF report creation error:', error);
      throw error;
    }
  }

  /**
   * Створення Word звіту
   */
  private async createWordReport(reportData: ReportData, filePath: string): Promise<void> {
    try {
      // Тут можна додати бібліотеку для створення Word
      // Наприклад, docx
      logger.warn('Word report creation not implemented yet');

      // Тимчасова реалізація - створюємо текстовий файл
      await this.createTextReport(reportData, filePath.replace('.docx', '.txt'));
    } catch (error) {
      logger.error('Word report creation error:', error);
      throw error;
    }
  }

  /**
   * Створення тимчасової директорії
   */
  private async ensureTempDir(): Promise<void> {
    try {
      await fs.mkdir(FILE_CONFIG.TEMP_DIR, { recursive: true });
    } catch (error) {
      logger.error('Temp directory creation error:', error);
      throw error;
    }
  }

  /**
   * Очищення тимчасового файлу
   */
  private async cleanupTempFile(filePath: string): Promise<void> {
    try {
      await fs.unlink(filePath);
      logger.info(`Temp file cleaned up: ${filePath}`);
    } catch (error) {
      logger.warn(`Failed to cleanup temp file ${filePath}:`, error);
    }
  }

  /**
   * Очищення всіх тимчасових файлів
   */
  async cleanupAllTempFiles(): Promise<void> {
    try {
      const files = await fs.readdir(FILE_CONFIG.TEMP_DIR);

      for (const file of files) {
        const filePath = path.join(FILE_CONFIG.TEMP_DIR, file);
        const stats = await fs.stat(filePath);

        // Видалення файлів старіше 1 години
        if (Date.now() - stats.mtime.getTime() > 60 * 60 * 1000) {
          await this.cleanupTempFile(filePath);
        }
      }

      logger.info('Temp files cleanup completed');
    } catch (error) {
      logger.error('Temp files cleanup error:', error);
    }
  }

  /**
   * Санітизація пошукового запиту
   */
  private sanitizeQuery(query: string): string {
    try {
      if (typeof query !== 'string') {
        return '';
      }

      return query
        .trim()
        .slice(0, 100)
        .replace(/[<>\"'&]/g, '');
    } catch (error) {
      logger.error('Query sanitization error:', error);
      return '';
    }
  }

  /**
   * Отримання статистики
   */
  getStats(): any {
    return {
      ...this.stats,
      tempDir: FILE_CONFIG.TEMP_DIR,
      supportedFormats: FILE_CONFIG.SUPPORTED_FORMATS,
    };
  }
}

// Експорт екземпляру класу
export const fileProcessor = new FileProcessor();
export default fileProcessor; 