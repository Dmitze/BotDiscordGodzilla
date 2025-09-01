/**
 * Google Sheets Service using node-google-spreadsheet library
 * Implements the same interface as GoogleService for compatibility
 */

import { google } from 'googleapis';
import type { drive_v3 } from 'googleapis';
import type { DocBlock } from '@/types/docs';
import type { MetricsService } from './MetricsService';
import { createHash } from 'crypto';
import type { BotConfig, SheetData, BatchSheetData } from '@/types';
import type { DriveListQuery, DriveListResult, DriveFile } from '@/types/drive';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';
import { sanitizeTextForChat, normalizeText } from '@/utils/fileProcessor';
import { validateInput } from '@/utils/security';

interface GoogleServiceStats {
  service: string;
  uptime: number;
  requests: number;
  errors: number;
  averageResponseTime: number;
  connectionPoolUsage: number;
  cacheHits: number;
  cacheMisses: number;
}

interface GoogleServiceOptions {
  useCache?: boolean;
  cacheTTL?: number;
  forceRefresh?: boolean;
  batchSize?: number;
  retryFailed?: boolean;
  cacheResults?: boolean;
  maxRetries?: number;
  valueInputOption?: 'RAW' | 'USER_ENTERED';
  clearCache?: boolean;
}

export class GoogleSheetsService extends BaseServiceClass {
  private auth: InstanceType<typeof google.auth.JWT> | null = null;
  private drive: drive_v3.Drive | null = null;
  private stats: GoogleServiceStats;
  private cacheService: CacheService;
  private metrics?: MetricsService;

  constructor(config: BotConfig) {
    super('GoogleSheetsService', config);
    this.cacheService = new CacheService(config);
    this.stats = {
      service: 'GoogleSheetsService',
      uptime: 0,
      requests: 0,
      errors: 0,
      averageResponseTime: 0,
      connectionPoolUsage: 0,
      cacheHits: 0,
      cacheMisses: 0,
    };
  }

  /** Подключение MetricsService (вызывается из ServiceManager) */
  public setMetricsService(ms: MetricsService): void {
    this.metrics = ms;
  }

  /**
   * Извлечение текста для чата с валидацией/санитизацией, хэш-контролем
   * Возвращает текст, источник (parser/export/ocr), контрольную сумму и modifiedTime
   */
  public async extractTextForChat(fileId: string): Promise<{
    text: string;
    checksum: string;
    modifiedTime?: string;
    source: 'export' | 'parser' | 'ocr' | 'raw';
    warnings: string[];
  }> {
    const overallStart = Date.now();
    const meta = await this.getDriveFileMetadata(fileId);
    const mime = String(meta.mimeType || '');
    const modified = meta.modifiedTime ?? '';

    // Используем модификацию и потенциальный контент-хэш для кэш-ключа
    const baseCacheKey = `extract:text:${fileId}:${modified}:${mime}`;
    try {
      const cached = await this.cacheService.get<{ text: string; checksum: string; modifiedTime?: string; source: 'export' | 'parser' | 'ocr' | 'raw'; warnings: string[] }>(baseCacheKey);
      if (cached?.text) {
        this.stats.cacheHits++;
        // Cache hit: lightweight metrics
        try { this.metrics?.recordFileOperation({ operation: 'extract_text_cache', status: 'success', mime, fileId }); } catch {}
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }

    const warnings: string[] = [];
    let buffer: Buffer | null = null;
    let text = '';
    let source: 'export' | 'parser' | 'ocr' | 'raw' = 'raw';

    try {
      // For Google Sheets, we export as CSV
      if (mime === 'application/vnd.google-apps.spreadsheet') {
        buffer = await this.exportDriveFile(fileId, 'text/csv');
        text = buffer.toString('utf8');
        source = 'export';
      } 
      // For Google Docs, we export as plain text
      else if (mime === 'application/vnd.google-apps.document') {
        buffer = await this.exportDriveFile(fileId, 'text/plain');
        text = buffer.toString('utf8');
        source = 'export';
      }
      // For other file types, we download directly
      else {
        buffer = await this.downloadDriveFile(fileId);
        text = buffer.toString('utf8');
        source = 'raw';
      }
    } catch (error) {
      logger.error('❌ Помилка витягання тексту', {
        type: 'processing_error',
        event: 'extract_text_failed',
        component: 'GoogleSheetsService',
        fileId,
        mime,
        error: error instanceof Error ? error.message : String(error),
      });
      text = '';
      try { this.metrics?.recordFileOperation({ operation: 'extract_text', status: 'error', mime, fileId }); } catch {}
    }

    // Нормализация и очистка
    text = normalizeText(text);
    const val = validateInput(text, { inputType: 'message' });
    if (!val.isValid) warnings.push(`sanitization: ${val.errors.join(', ')}`);
    const safe = sanitizeTextForChat(text);

    // Контрольная сумма
    const checksum = buffer ? createHash('sha256').update(buffer).digest('hex') : createHash('sha256').update(safe).digest('hex');

    const result = { text: safe, checksum, modifiedTime: modified, source, warnings } as const;
    try {
      const ttl = this.config.drive?.ttlTextSec ?? 300;
      await this.cacheService.set(baseCacheKey, result, ttl);
    } catch { /* istanbul ignore next */ }
    // Final overall metrics for extractTextForChat
    try {
      this.metrics?.recordFileOperation({ operation: 'extract_text', status: 'success', mime, fileId });
      this.metrics?.observeFileOperationLatency('extract_text', mime, Date.now() - overallStart);
      this.metrics?.observeTextSizeBytes(source, Buffer.byteLength(safe, 'utf8'));
      if (mime) this.metrics?.incrementMimeType(mime);
    } catch {}
    return result;
  }

  /**
   * Отримати метадані файлу Google Drive
   */
  public async getDriveFileMetadata(fileId: string): Promise<drive_v3.Schema$File> {
    try {
      if (!this.drive) throw new Error('Drive API не ініціалізовано');
      
      const res = await this.drive.files.get({
        fileId,
        fields: 'id,name,mimeType,size,modifiedTime,owners(displayName),parents',
      });
      
      return res.data;
    } catch (error) {
      logger.error('❌ Помилка отримання метаданих файлу Drive', {
        type: 'api_error',
        event: 'drive_get_metadata_failed',
        component: 'GoogleSheetsService',
        fileId,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Завантаження бінарного файлу з Drive (не Google Docs/Sheets типи)
   */
  public async downloadDriveFile(fileId: string): Promise<Buffer> {
    try {
      if (!this.drive) throw new Error('Drive API не ініціалізовано');
      
      const res = await this.drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'arraybuffer' }
      );
      
      return Buffer.from(res.data as any);
    } catch (error) {
      logger.error('❌ Помилка завантаження файла Drive', {
        type: 'api_error',
        event: 'drive_download_failed',
        component: 'GoogleSheetsService',
        fileId,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Експорт файлу Google Docs/Sheets/Slides у вказаний MIME тип
   */
  public async exportDriveFile(fileId: string, mimeType: string): Promise<Buffer> {
    try {
      // For Google Sheets, use node-google-spreadsheet library
      if (mimeType === 'text/csv' || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        // Dynamically import the module to avoid ES module issues
        const { GoogleSpreadsheet } = await import('google-spreadsheet');
        const doc = new GoogleSpreadsheet(fileId, this.auth!);
        await doc.loadInfo();
        
        // Export as XLSX or CSV
        if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
          const buffer = await doc.downloadAsXLSX();
          return Buffer.from(buffer);
        } else if (mimeType === 'text/csv') {
          // For CSV, we use the Drive API export method which is more reliable
          if (!this.drive) throw new Error('Drive API не ініціалізовано');
          
          const res = await this.drive.files.export(
            { fileId, mimeType: 'text/csv' },
            { responseType: 'arraybuffer' }
          );
          
          return Buffer.from(res.data as any);
        }
      }
      
      // For other file types or as fallback, use the Drive API export
      if (!this.drive) throw new Error('Drive API не ініціалізовано');
      
      const res = await this.drive.files.export(
        { fileId, mimeType },
        { responseType: 'arraybuffer' }
      );
      
      return Buffer.from(res.data as any);
    } catch (error) {
      logger.error('❌ Помилка експорту файла Drive', {
        type: 'api_error',
        event: 'drive_export_failed',
        component: 'GoogleSheetsService',
        fileId,
        mimeType,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Ініціалізація Google сервісів
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація Google Sheets Service...', {
        type: 'system',
        event: 'google_sheets_service_init',
        component: 'GoogleSheetsService',
      });

      // Ініціалізація кешу
      await this.cacheService.initialize();

      // У тестовому режимі або коли сервіс вимкнено/немає credentials — пропускаємо зовнішню ініціалізацію
      const disabled =
        process.env['NODE_ENV'] === 'test' ||
        process.env['DISABLE_GOOGLE_SERVICE'] === 'true' ||
        !this.config.google?.credentials;

      if (disabled) {
        logger.warn('🧪 Режим тесту/відключено/немає credentials: пропущено auth/API для GoogleSheetsService', {
          type: 'system',
          event: 'google_sheets_service_init_skipped_external',
          component: 'GoogleSheetsService',
        });
      } else {
        // Створення автентифікації
        await this.initializeAuth();

        // Ініціалізація Drive API клієнта
        this.initializeDriveAPI();
      }

      logger.info('✅ Google Sheets Service ініціалізовано', {
        type: 'system',
        event: 'google_sheets_service_init_success',
        component: 'GoogleSheetsService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google Sheets Service', {
          type: 'system',
          event: 'google_sheets_service_init_failed',
          component: 'GoogleSheetsService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          severity: 'critical',
        });
      } else {
        logger.error('❌ Помилка ініціалізації Google Sheets Service', {
          type: 'system',
          event: 'google_sheets_service_init_failed',
          component: 'GoogleSheetsService',
          errorMessage: String(error),
          severity: 'critical',
        });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація автентифікації
   */
  private async initializeAuth(): Promise<void> {
    try {
      // Перевірка наявності credentials
      if (!this.config.google.credentials) {
        throw new Error('Google credentials не налаштовано');
      }

      // Створення JWT автентифікації
      const jwt = new google.auth.JWT(
        this.config.google.credentials.client_email,
        undefined,
        this.config.google.credentials.private_key,
        [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
        ]
      );

      // Авторизація
      await jwt.authorize();
      this.auth = jwt;
      logger.info('✅ Google автентифікація успішна', {
        type: 'system',
        event: 'google_auth_success',
        component: 'GoogleSheetsService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка Google автентифікації', {
          type: 'api_error',
          event: 'google_auth_failed',
          component: 'GoogleSheetsService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'google',
        });
      } else {
        logger.error('❌ Помилка Google автентифікації', {
          type: 'api_error',
          event: 'google_auth_failed',
          component: 'GoogleSheetsService',
          service: 'google',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація Drive API клієнта
   */
  private initializeDriveAPI(): void {
    try {
      if (!this.auth) throw new Error('Auth client is not initialized');
      
      // Google Drive API
      this.drive = google.drive({ version: 'v3', auth: this.auth });

      logger.info('✅ Google Drive API клієнт ініціалізовано', {
        type: 'system',
        event: 'google_drive_api_init_success',
        component: 'GoogleSheetsService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google Drive API', {
          type: 'api_error',
          event: 'google_drive_api_init_failed',
          component: 'GoogleSheetsService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'google',
        });
      } else {
        logger.error('❌ Помилка ініціалізації Google Drive API', {
          type: 'api_error',
          event: 'google_drive_api_init_failed',
          component: 'GoogleSheetsService',
          service: 'google',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<any> {
    try {
      // Перевірка автентифікації
      if (!this.auth) {
        return {
          healthy: false,
          service: this.name,
          error: 'Auth не ініціалізовано',
        };
      }

      // Перевірка API клієнтів
      if (!this.drive) {
        return {
          healthy: false,
          service: this.name,
          error: 'Drive API клієнт не ініціалізовано',
        };
      }

      // Тестовий запит до Google Drive API
      try {
        await this.drive.about.get({
          fields: 'user'
        });
      } catch (error) {
        return {
          healthy: false,
          service: this.name,
          error: `Помилка тестового запиту: ${String(error)}`,
        };
      }

      return {
        healthy: true,
        service: this.name,
        details: {
          requests: this.stats.requests,
          errors: this.stats.errors,
          averageResponseTime: this.stats.averageResponseTime,
        },
      };
    } catch (error) {
      return {
        healthy: false,
        service: this.name,
        error: `Health check failed: ${String(error)}`,
      };
    }
  }

  /**
   * Завершення роботи
   */
  protected async onShutdown(): Promise<void> {
    try {
      // Зупинка кеш сервісу
      await this.cacheService.shutdown();

      // Скидання API клієнтів
      this.drive = null;
      this.auth = null;

      logger.info('✅ Google Sheets Service зупинено', {
        type: 'system',
        event: 'google_sheets_service_shutdown_success',
        component: 'GoogleSheetsService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка зупинки Google Sheets Service', {
          type: 'system',
          event: 'google_sheets_service_shutdown_failed',
          component: 'GoogleSheetsService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка зупинки Google Sheets Service', {
          type: 'system',
          event: 'google_sheets_service_shutdown_failed',
          component: 'GoogleSheetsService',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Отримання статистики
   */
  protected onGetStats(): Partial<GoogleServiceStats> {
    return this.stats;
  }
}