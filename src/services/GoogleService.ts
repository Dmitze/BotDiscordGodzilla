/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */

import { google, sheets_v4, drive_v3, docs_v1 } from 'googleapis';
import type { 
  BotConfig, 
  HealthStatus, 
  ServiceStats,
  SheetData,
  BatchSheetData
} from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';

interface GoogleServiceStats extends ServiceStats {
  requests: number;
  errors: number;
  averageResponseTime: number;
  connectionPoolUsage: number;
  cacheHits: number;
  cacheMisses: number;
}

interface ConnectionInfo {
  inUse: boolean;
  lastUsed: number;
  requestCount: number;
}

interface GoogleServiceOptions {
  useCache?: boolean;
  cacheTTL?: number;
  forceRefresh?: boolean;
  batchSize?: number;
  retryFailed?: boolean;
  cacheResults?: boolean;
  maxRetries?: number;
  valueInputOption?: string;
  clearCache?: boolean;
}

export class GoogleService extends BaseServiceClass {
  private auth: any = null;
  private sheets: sheets_v4.Sheets | null = null;
  private drive: drive_v3.Drive | null = null;
  private docs: docs_v1.Docs | null = null;
  private connectionPool = new Map<string, ConnectionInfo>();
  private readonly retryAttempts = 3;
  private readonly retryDelay = 1000;
  private stats: GoogleServiceStats;
  private cacheService: CacheService;

  constructor(config: BotConfig) {
    super('GoogleService', config);
    this.cacheService = new CacheService(config);
    this.stats = {
      service: 'GoogleService',
      uptime: 0,
      requests: 0,
      errors: 0,
      averageResponseTime: 0,
      connectionPoolUsage: 0,
      cacheHits: 0,
      cacheMisses: 0,
    };
  }

  /**
   * Ініціалізація Google сервісів
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація Google Service...', { type: 'system', event: 'google_service_init' });

      // Ініціалізація кешу
      await this.cacheService.initialize();

      // Створення автентифікації
      await this.initializeAuth();

      // Ініціалізація API клієнтів
      await this.initializeAPIs();

      // Створення connection pool
      await this.initializeConnectionPool();

      logger.info('✅ Google Service ініціалізовано', { type: 'system', event: 'google_service_init_success' });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google Service', {
          type: 'system', event: 'google_service_init_failed',
          errorName: error.name, errorMessage: error.message, stack: error.stack,
          severity: 'critical',
        });
      } else {
        logger.error('❌ Помилка ініціалізації Google Service', { type: 'system', event: 'google_service_init_failed', errorMessage: String(error), severity: 'critical' });
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
      this.auth = new google.auth.JWT(
        this.config.google.credentials.client_email,
        undefined,
        this.config.google.credentials.private_key,
        [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/documents',
        ]
      );

      // Авторизація
      await this.auth.authorize();
      logger.info('✅ Google автентифікація успішна', { type: 'system', event: 'google_auth_success' });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка Google автентифікації', { type: 'api_error', event: 'google_auth_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'google' });
      } else {
        logger.error('❌ Помилка Google автентифікації', { type: 'api_error', event: 'google_auth_failed', service: 'google', errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація API клієнтів
   */
  private async initializeAPIs(): Promise<void> {
    try {
      // Google Sheets API
      this.sheets = google.sheets({ version: 'v4', auth: this.auth });

      // Google Drive API
      this.drive = google.drive({ version: 'v3', auth: this.auth });

      // Google Docs API
      this.docs = google.docs({ version: 'v1', auth: this.auth });

      logger.info('✅ Google API клієнти ініціалізовано', { type: 'system', event: 'google_api_init_success' });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google API', { type: 'api_error', event: 'google_api_init_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'google' });
      } else {
        logger.error('❌ Помилка ініціалізації Google API', { type: 'api_error', event: 'google_api_init_failed', service: 'google', errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація Connection Pool
   */
  private async initializeConnectionPool(): Promise<void> {
    try {
      const apiTypes = ['sheets', 'drive', 'docs'];
      
      for (const apiType of apiTypes) {
        this.connectionPool.set(apiType, {
          inUse: false,
          lastUsed: Date.now(),
          requestCount: 0,
        });
      }

      logger.info('✅ Connection Pool ініціалізовано', { type: 'system', event: 'connection_pool_init_success' });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Connection Pool', { type: 'system', event: 'connection_pool_init_failed', errorName: error.name, errorMessage: error.message, stack: error.stack });
      } else {
        logger.error('❌ Помилка ініціалізації Connection Pool', { type: 'system', event: 'connection_pool_init_failed', errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Отримання з'єднання з пулу
   */
  private getConnection(apiType: string): boolean {
    const connection = this.connectionPool.get(apiType);
    if (!connection) {
      return false;
    }

    if (connection.inUse) {
      return false;
    }

    connection.inUse = true;
    connection.lastUsed = Date.now();
    connection.requestCount++;
    return true;
  }

  /**
   * Звільнення з'єднання
   */
  private releaseConnection(apiType: string): void {
    const connection = this.connectionPool.get(apiType);
    if (connection) {
      connection.inUse = false;
    }
  }

  /**
   * Виконання операції з retry
   */
  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    apiType: string,
    maxRetries: number = this.retryAttempts
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const connection = this.getConnection(apiType);
        if (!connection) {
          throw new Error(`Немає доступних з'єднань для ${apiType}`);
        }

        const startTime = Date.now();
        const result = await operation();
        const duration = Date.now() - startTime;

        this.releaseConnection(apiType);
        this.updateStats(true, duration);

        return result;
      } catch (error) {
        lastError = error as Error;
        this.releaseConnection(apiType);
        this.updateStats(false, 0);

        if (attempt < maxRetries) {
          const delay = this.retryDelay * Math.pow(2, attempt);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }

    throw lastError || new Error('Всі спроби виконання невдалі');
  }

  /**
   * Отримання даних з Google Sheets
   */
  public async getSheetData(
    spreadsheetId: string,
    range: string,
    options: GoogleServiceOptions = {}
  ): Promise<SheetData> {
    const { useCache = true, cacheTTL = 300, forceRefresh = false } = options;

    try {
      // Перевірка кешу
      if (useCache && !forceRefresh) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          const cached = await this.cacheService.get<SheetData>(cacheKey);
          if (cached) {
            this.stats.cacheHits++;
            logger.debug('✅ Використано кешовані дані Sheets', {
              type: 'system', event: 'cache_hit',
              spreadsheetId: spreadsheetId.substring(0, 10) + '...',
              range,
              rowsCount: cached.values.length
            });
            return cached;
          } else {
            this.stats.cacheMisses++;
          }
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка читання з кешу Sheets', { type: 'system', event: 'cache_read_failed', errorName: cacheError.name, errorMessage: cacheError.message, stack: cacheError.stack, component: 'CacheService' });
          } else {
            logger.warn('⚠️ Помилка читання з кешу Sheets', { type: 'system', event: 'cache_read_failed', component: 'CacheService', errorMessage: String(cacheError) });
          }
          this.stats.cacheMisses++;
        }
      }

      const result = await this.executeWithRetry(
        async () => {
          if (!this.sheets) throw new Error('Sheets API не ініціалізовано');
          
          const response = await this.sheets.spreadsheets.values.get({
            spreadsheetId,
            range,
          });

          return {
            range: response.data.range || range,
            majorDimension: response.data.majorDimension || 'ROWS',
            values: response.data.values || [],
          };
        },
        'sheets'
      );

      // Збереження в кеш
      if (useCache) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          await this.cacheService.set(cacheKey, result, cacheTTL * 1000);
          logger.debug('💾 Дані Sheets збережено в кеш', {
            type: 'system', event: 'cache_write',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range,
            rowsCount: result.values.length,
            ttl: `${cacheTTL}s`
          });
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка збереження в кеш Sheets', { type: 'system', event: 'cache_write_failed', errorName: cacheError.name, errorMessage: cacheError.message, stack: cacheError.stack, component: 'CacheService' });
          } else {
            logger.warn('⚠️ Помилка збереження в кеш Sheets', { type: 'system', event: 'cache_write_failed', component: 'CacheService', errorMessage: String(cacheError) });
          }
        }
      }

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання даних з Sheets', { type: 'api_error', event: 'sheets_get_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId, range });
      } else {
        logger.error('❌ Помилка отримання даних з Sheets', { type: 'api_error', event: 'sheets_get_failed', service: 'sheets', spreadsheetId, range, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Запис даних в Google Sheets
   */
  public async writeSheetData(
    spreadsheetId: string,
    range: string,
    values: string[][],
    options: GoogleServiceOptions = {}
  ): Promise<void> {
    const { valueInputOption = 'RAW', clearCache = true } = options;

    try {
      await this.executeWithRetry(
        async () => {
          if (!this.sheets) throw new Error('Sheets API не ініціалізовано');
          
          await this.sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption,
            requestBody: {
              values,
            },
          });
        },
        'sheets'
      );

      // Очищення кешу
      if (clearCache) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          await this.cacheService.delete(cacheKey);
          logger.debug('🗑️ Кеш Sheets очищено', {
            type: 'system', event: 'cache_delete',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range
          });
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка очищення кешу Sheets', { type: 'system', event: 'cache_delete_failed', errorName: cacheError.name, errorMessage: cacheError.message, stack: cacheError.stack, component: 'CacheService' });
          } else {
            logger.warn('⚠️ Помилка очищення кешу Sheets', { type: 'system', event: 'cache_delete_failed', component: 'CacheService', errorMessage: String(cacheError) });
          }
        }
      }
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка запису в Sheets', { type: 'api_error', event: 'sheets_write_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId, range });
      } else {
        logger.error('❌ Помилка запису в Sheets', { type: 'api_error', event: 'sheets_write_failed', service: 'sheets', spreadsheetId, range, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Batch отримання даних з Google Sheets
   */
  public async batchGetSheetData(
    spreadsheetId: string,
    ranges: string[],
    options: GoogleServiceOptions = {}
  ): Promise<BatchSheetData> {
    const { 
      batchSize = 10, 
      retryFailed = true,
      maxRetries = 3
    } = options;

    try {
      const chunks = this.chunkArray(ranges, batchSize);
      const results: SheetData[] = [];
      const failedRanges: string[] = [];

      for (const chunk of chunks) {
        try {
          const result = await this.executeWithRetry(
            async () => {
              if (!this.sheets) throw new Error('Sheets API не ініціалізовано');
              
              const response = await this.sheets.spreadsheets.values.batchGet({
                spreadsheetId,
                ranges: chunk,
              });

              return response.data.valueRanges || [];
            },
            'sheets',
            maxRetries
          );

          results.push(
            ...result.map(vr => ({
              range: vr.range ?? '',
              majorDimension: vr.majorDimension ?? 'ROWS',
              values: vr.values ?? [],
            }))
          );
        } catch (error) {
          if (error instanceof Error) {
            logger.error('❌ Помилка batch запиту', { type: 'api_error', event: 'sheets_batch_get_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId, ranges: chunk });
          } else {
            logger.error('❌ Помилка batch запиту', { type: 'api_error', event: 'sheets_batch_get_failed', service: 'sheets', spreadsheetId, ranges: chunk, errorMessage: String(error) });
          }
          if (retryFailed) {
            failedRanges.push(...chunk);
          }
        }
      }

      // Повторна спроба для невдалих ranges
      if (retryFailed && failedRanges.length > 0) {
        for (const range of failedRanges) {
          try {
            const result = await this.getSheetData(spreadsheetId, range, { useCache: false });
            results.push(result);
          } catch (error) {
            if (error instanceof Error) {
              logger.error(`❌ Повторна спроба невдала для range: ${range}`, { type: 'api_error', event: 'sheets_retry_get_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId, range });
            } else {
              logger.error(`❌ Повторна спроба невдала для range: ${range}`, { type: 'api_error', event: 'sheets_retry_get_failed', service: 'sheets', spreadsheetId, range, errorMessage: String(error) });
            }
          }
        }
      }

      return {
        valueRanges: results,
        spreadsheetId,
      };
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка batch отримання даних', { type: 'api_error', event: 'sheets_batch_get_failed_final', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId });
      } else {
        logger.error('❌ Помилка batch отримання даних', { type: 'api_error', event: 'sheets_batch_get_failed_final', service: 'sheets', spreadsheetId, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Batch запис даних в Google Sheets
   */
  public async batchWriteSheetData(
    spreadsheetId: string,
    data: Array<{ range: string; values: string[][] }>,
    options: GoogleServiceOptions = {}
  ): Promise<void> {
    const { 
      batchSize = 10, 
      retryFailed = true,
      maxRetries = 3,
      clearCache = true
    } = options;

    try {
      const chunks = this.chunkArray(data, batchSize);
      const failedBatches: Array<{ range: string; values: string[][] }> = [];

      for (const chunk of chunks) {
        try {
          await this.executeWithRetry(
            async () => {
              if (!this.sheets) throw new Error('Sheets API не ініціалізовано');
              
              const requests = chunk.map(item => ({
                updateCells: {
                  range: {
                    sheetId: 0, // TODO: Отримати sheetId
                    startRowIndex: 0,
                    endRowIndex: item.values.length,
                    startColumnIndex: 0,
                    endColumnIndex: item.values[0]?.length || 0,
                  },
                  rows: item.values.map(row => ({
                    values: row.map(cell => ({ userEnteredValue: { stringValue: cell } })),
                  })),
                  fields: 'userEnteredValue',
                },
              }));

              await this.sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                requestBody: { requests },
              });
            },
            'sheets',
            maxRetries
          );

          // Очищення кешу
          if (clearCache) {
            for (const item of chunk) {
              const cacheKey = `sheets:${spreadsheetId}:${item.range}`;
              try {
                await this.cacheService.delete(cacheKey);
                logger.debug('🗑️ Кеш Sheets очищено', {
                  type: 'system', event: 'cache_delete',
                  spreadsheetId: spreadsheetId.substring(0, 10) + '...',
                  range: item.range
                });
              } catch (cacheError) {
                if (cacheError instanceof Error) {
                  logger.warn('⚠️ Помилка очищення кешу Sheets', { type: 'system', event: 'cache_delete_failed', errorName: cacheError.name, errorMessage: cacheError.message, stack: cacheError.stack, component: 'CacheService' });
                } else {
                  logger.warn('⚠️ Помилка очищення кешу Sheets', { type: 'system', event: 'cache_delete_failed', component: 'CacheService', errorMessage: String(cacheError) });
                }
              }
            }
          }
        } catch (error) {
          if (error instanceof Error) {
            logger.error('❌ Помилка batch запису', { type: 'api_error', event: 'sheets_batch_write_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId });
          } else {
            logger.error('❌ Помилка batch запису', { type: 'api_error', event: 'sheets_batch_write_failed', service: 'sheets', spreadsheetId, errorMessage: String(error) });
          }
          if (retryFailed) {
            failedBatches.push(...chunk);
          }
        }
      }

      // Повторна спроба для невдалих batch
      if (retryFailed && failedBatches.length > 0) {
        for (const item of failedBatches) {
          try {
            await this.writeSheetData(spreadsheetId, item.range, item.values, { useCache: false });
          } catch (error) {
            if (error instanceof Error) {
              logger.error(`❌ Повторна спроба невдала для range: ${item.range}`, { type: 'api_error', event: 'sheets_retry_write_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId, range: item.range });
            } else {
              logger.error(`❌ Повторна спроба невдала для range: ${item.range}`, { type: 'api_error', event: 'sheets_retry_write_failed', service: 'sheets', spreadsheetId, range: item.range, errorMessage: String(error) });
            }
          }
        }
      }
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка batch запису даних', { type: 'api_error', event: 'sheets_batch_write_failed_final', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'sheets', spreadsheetId });
      } else {
        logger.error('❌ Помилка batch запису даних', { type: 'api_error', event: 'sheets_batch_write_failed_final', service: 'sheets', spreadsheetId, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Пошук файлів в Google Drive
   */
  public async searchFiles(
    query: string,
    _options: GoogleServiceOptions = {}
  ): Promise<drive_v3.Schema$File[]> {
    try {
      const result = await this.executeWithRetry(
        async () => {
          if (!this.drive) throw new Error('Drive API не ініціалізовано');
          
          const response = await this.drive.files.list({
            q: query,
            fields: 'files(id,name,mimeType,size,modifiedTime)',
            pageSize: 100,
          });

          return response.data.files || [];
        },
        'drive'
      );

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка пошуку файлів', { type: 'api_error', event: 'drive_search_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'drive', query });
      } else {
        logger.error('❌ Помилка пошуку файлів', { type: 'api_error', event: 'drive_search_failed', service: 'drive', query, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Отримання метаданих файлу
   */
  public async getFileMetadata(
    fileId: string,
    fields: string = '*'
  ): Promise<drive_v3.Schema$File> {
    try {
      const result = await this.executeWithRetry(
        async () => {
          if (!this.drive) throw new Error('Drive API не ініціалізовано');
          
          const response = await this.drive.files.get({
            fileId,
            fields,
          });

          return response.data;
        },
        'drive'
      );

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання метаданих файлу', { type: 'api_error', event: 'drive_metadata_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'drive', fileId });
      } else {
        logger.error('❌ Помилка отримання метаданих файлу', { type: 'api_error', event: 'drive_metadata_failed', service: 'drive', fileId, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Отримання контенту документа
   */
  public async getDocumentContent(documentId: string): Promise<string> {
    try {
      const result = await this.executeWithRetry(
        async () => {
          if (!this.docs) throw new Error('Docs API не ініціалізовано');
          
          const response = await this.docs.documents.get({
            documentId,
          });

          // Парсинг контенту документа
          const content = this.parseDocumentContent(response.data);
          return content;
        },
        'docs'
      );

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання контенту документа', { type: 'api_error', event: 'docs_content_failed', errorName: error.name, errorMessage: error.message, stack: error.stack, service: 'docs', documentId });
      } else {
        logger.error('❌ Помилка отримання контенту документа', { type: 'api_error', event: 'docs_content_failed', service: 'docs', documentId, errorMessage: String(error) });
      }
      throw error;
    }
  }

  /**
   * Парсинг контенту документа
   */
  private parseDocumentContent(document: docs_v1.Schema$Document): string {
    if (!document.body?.content) {
      return '';
    }

    let content = '';
    for (const element of document.body.content) {
      if (element.paragraph) {
        for (const element2 of element.paragraph.elements || []) {
          if (element2.textRun?.content) {
            content += element2.textRun.content;
          }
        }
        content += '\n';
      }
    }

    return content.trim();
  }

  /**
   * Отримання статистики з'єднань
   */
  public getConnectionStats(): Record<string, ConnectionInfo> {
    const stats: Record<string, ConnectionInfo> = {};
    for (const [apiType, connection] of this.connectionPool.entries()) {
      stats[apiType] = { ...connection };
    }
    return stats;
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
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
      if (!this.sheets || !this.drive || !this.docs) {
        return {
          healthy: false,
          service: this.name,
          error: 'API клієнти не ініціалізовано',
        };
      }

      // Тестовий запит до Sheets API
      try {
        await this.sheets.spreadsheets.get({
          spreadsheetId: this.config.google.spreadsheetId,
          ranges: ['A1:A1'],
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
          connectionPoolSize: this.connectionPool.size,
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

      // Очищення connection pool
      this.connectionPool.clear();

      // Скидання API клієнтів
      this.sheets = null;
      this.drive = null;
      this.docs = null;
      this.auth = null;

      logger.info('✅ Google Service зупинено', { type: 'system', event: 'google_service_shutdown_success' });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка зупинки Google Service', { type: 'system', event: 'google_service_shutdown_failed', errorName: error.name, errorMessage: error.message, stack: error.stack });
      } else {
        logger.error('❌ Помилка зупинки Google Service', { type: 'system', event: 'google_service_shutdown_failed', errorMessage: String(error) });
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

  /**
   * Розбивка масиву на чанки
   */
  private chunkArray<T>(array: T[], chunkSize: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Оновлення статистики
   */
  private updateStats(success: boolean, duration: number): void {
    this.stats.requests++;
    if (!success) {
      this.stats.errors++;
    }

    // Оновлення середнього часу відповіді
    const totalTime = this.stats.averageResponseTime * (this.stats.requests - 1) + duration;
    this.stats.averageResponseTime = totalTime / this.stats.requests;

    // Оновлення використання connection pool
    let inUseConnections = 0;
    for (const connection of this.connectionPool.values()) {
      if (connection.inUse) {
        inUseConnections++;
      }
    }
    this.stats.connectionPoolUsage = (inUseConnections / this.connectionPool.size) * 100;
  }
} 