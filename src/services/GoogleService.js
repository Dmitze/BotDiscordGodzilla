/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */

const BaseService = require('../core/BaseService');
const logger = require('../utils/logger');
const { google } = require('googleapis');

class GoogleService extends BaseService {
  constructor(config) {
    super('GoogleService', config);
    this.auth = null;
    this.sheets = null;
    this.drive = null;
    this.docs = null;
    this.connectionPool = new Map();
    this.maxConnections = 10;
    this.connectionTimeout = 30000; // 30 секунд
    this.retryAttempts = 3;
    this.retryDelay = 1000;
  }

  /**
   * Ініціалізація Google сервісів
   */
  async onInitialize() {
    try {
      logger.info('🔧 Ініціалізація Google Service...');

      // Створення автентифікації
      await this.initializeAuth();

      // Ініціалізація API клієнтів
      await this.initializeAPIs();

      // Створення connection pool
      await this.initializeConnectionPool();

      logger.info('✅ Google Service ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Google Service:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація автентифікації
   */
  async initializeAuth() {
    try {
      // Перевірка наявності credentials
      if (!this.config.google.credentials) {
        throw new Error('Google credentials не налаштовано');
      }

      // Створення JWT автентифікації
      this.auth = new google.auth.JWT(
        this.config.google.credentials.client_email,
        null,
        this.config.google.credentials.private_key,
        [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/documents',
        ]
      );

      // Авторизація
      await this.auth.authorize();
      logger.info('✅ Google автентифікація успішна');
    } catch (error) {
      logger.error('❌ Помилка Google автентифікації:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація API клієнтів
   */
  async initializeAPIs() {
    try {
      // Google Sheets API
      this.sheets = google.sheets({ version: 'v4', auth: this.auth });

      // Google Drive API
      this.drive = google.drive({ version: 'v3', auth: this.auth });

      // Google Docs API
      this.docs = google.docs({ version: 'v1', auth: this.auth });

      logger.info('✅ Google API клієнти ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Google API:', error);
      throw error;
    }
  }

  /**
   * Ініціалізація Connection Pool
   */
  async initializeConnectionPool() {
    try {
      // Створення пулу з'єднань для різних API
      this.connectionPool.set('sheets', {
        client: this.sheets,
        inUse: false,
        lastUsed: Date.now(),
      });

      this.connectionPool.set('drive', {
        client: this.drive,
        inUse: false,
        lastUsed: Date.now(),
      });

      this.connectionPool.set('docs', {
        client: this.docs,
        inUse: false,
        lastUsed: Date.now(),
      });

      logger.info('✅ Connection Pool ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації Connection Pool:', error);
      throw error;
    }
  }

  /**
   * Отримання з'єднання з пулу
   */
  async getConnection(apiType) {
    const connection = this.connectionPool.get(apiType);
    if (!connection) {
      throw new Error(`Невідомий тип API: ${apiType}`);
    }

    // Перевірка чи з'єднання вільне
    if (connection.inUse) {
      throw new Error(`З'єднання ${apiType} зайняте`);
    }

    connection.inUse = true;
    connection.lastUsed = Date.now();

    return connection.client;
  }

  /**
   * Повернення з'єднання в пул
   */
  releaseConnection(apiType) {
    const connection = this.connectionPool.get(apiType);
    if (connection) {
      connection.inUse = false;
      connection.lastUsed = Date.now();
    }
  }

  /**
   * Виконання запиту з retry логікою
   */
  async executeWithRetry(operation, apiType, maxRetries = this.retryAttempts) {
    let lastError;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const client = await this.getConnection(apiType);
        
        const result = await Promise.race([
          operation(client),
          new Promise((_, reject) => 
            setTimeout(() => reject(new Error('Timeout')), this.connectionTimeout)
          )
        ]);

        this.releaseConnection(apiType);
        this.updateStats(true);
        
        return result;
      } catch (error) {
        lastError = error;
        this.releaseConnection(apiType);
        this.updateStats(false);

        if (attempt === maxRetries) {
          logger.error(`❌ Помилка після ${maxRetries} спроб:`, error);
          break;
        }

        // Експоненціальна затримка
        const delay = this.retryDelay * Math.pow(2, attempt - 1);
        logger.warn(`⚠️ Спроба ${attempt}/${maxRetries} невдала, повтор через ${delay}ms`);
        
        await new Promise(resolve => setTimeout(resolve, delay));
      }
    }

    throw lastError;
  }

  /**
   * Отримання даних з Google Sheets
   */
  async getSheetData(spreadsheetId, range, options = {}) {
    const startTime = Date.now();
    const {
      useCache = true,
      cacheTTL = 300000, // 5 хвилин
      forceRefresh = false
    } = options;

    try {
      // Перевірка кешу
      if (useCache && this.serviceContainer && !forceRefresh) {
        const cacheService = this.serviceContainer.get('cache');
        if (cacheService) {
          const cacheKey = `sheets:${spreadsheetId}:${range}`;
          const cachedData = await cacheService.get(cacheKey);
          if (cachedData) {
            logger.info(`📋 Дані отримано з кешу: ${range}`);
            this.updateStats(true, Date.now() - startTime);
            return cachedData;
          }
        }
      }

      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.spreadsheets.values.get({
            spreadsheetId: spreadsheetId || this.config.google.spreadsheetId,
            range: range || 'Sheet1',
            valueRenderOption: options.valueRenderOption || 'UNFORMATTED_VALUE',
            dateTimeRenderOption: options.dateTimeRenderOption || 'SERIAL_NUMBER',
          });

          return response.data;
        },
        'sheets'
      );

      // Кешування результатів
      if (useCache && this.serviceContainer) {
        const cacheService = this.serviceContainer.get('cache');
        if (cacheService) {
          const cacheKey = `sheets:${spreadsheetId}:${range}`;
          await cacheService.set(cacheKey, result, cacheTTL);
        }
      }

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка отримання даних з Sheets:', error);
      throw error;
    }
  }

  /**
   * Запис даних в Google Sheets
   */
  async writeSheetData(spreadsheetId, range, values, options = {}) {
    const startTime = Date.now();

    try {
      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.spreadsheets.values.update({
            spreadsheetId: spreadsheetId || this.config.google.spreadsheetId,
            range: range,
            valueInputOption: options.valueInputOption || 'RAW',
            resource: {
              values: values,
            },
          });

          return response.data;
        },
        'sheets'
      );

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка запису в Sheets:', error);
      throw error;
    }
  }

  /**
   * Пошук файлів в Google Drive
   */
  async searchFiles(query, options = {}) {
    const startTime = Date.now();

    try {
      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.files.list({
            q: query,
            pageSize: options.pageSize || 10,
            fields: options.fields || 'files(id,name,mimeType,size,modifiedTime,webViewLink)',
            orderBy: options.orderBy || 'modifiedTime desc',
          });

          return response.data;
        },
        'drive'
      );

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка пошуку файлів:', error);
      throw error;
    }
  }

  /**
   * Отримання вмісту файлу
   */
  async getFileContent(fileId, options = {}) {
    const startTime = Date.now();

    try {
      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.files.get({
            fileId: fileId,
            alt: 'media',
          });

          return response.data;
        },
        'drive'
      );

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка отримання вмісту файлу:', error);
      throw error;
    }
  }

  /**
   * Отримання метаданих файлу
   */
  async getFileMetadata(fileId, fields = '*') {
    const startTime = Date.now();

    try {
      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.files.get({
            fileId: fileId,
            fields: fields,
          });

          return response.data;
        },
        'drive'
      );

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка отримання метаданих файлу:', error);
      throw error;
    }
  }

  /**
   * Отримання вмісту Google Docs
   */
  async getDocumentContent(documentId) {
    const startTime = Date.now();

    try {
      const result = await this.executeWithRetry(
        async (client) => {
          const response = await client.documents.get({
            documentId: documentId,
          });

          return response.data;
        },
        'docs'
      );

      this.updateStats(true, Date.now() - startTime);
      return result;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка отримання вмісту документа:', error);
      throw error;
    }
  }

  /**
   * Batch операції для Sheets з оптимізацією
   */
  async batchGetSheetData(spreadsheetId, ranges, options = {}) {
    const startTime = Date.now();
    const {
      batchSize = 10,
      cacheResults = true,
      cacheTTL = 300000, // 5 хвилин
      retryFailed = true,
      maxRetries = 3
    } = options;

    try {
      // Розбиття ranges на батчі для оптимізації
      const batches = this.chunkArray(ranges, batchSize);
      const results = [];
      const failedRanges = [];

      for (let i = 0; i < batches.length; i++) {
        const batch = batches[i];
        logger.info(`📊 Обробка batch ${i + 1}/${batches.length} (${batch.length} ranges)`);

        try {
          const batchResult = await this.executeWithRetry(
            async (client) => {
              const response = await client.spreadsheets.values.batchGet({
                spreadsheetId: spreadsheetId || this.config.google.spreadsheetId,
                ranges: batch,
                majorDimension: options.majorDimension || 'ROWS',
                valueRenderOption: options.valueRenderOption || 'UNFORMATTED_VALUE',
                dateTimeRenderOption: options.dateTimeRenderOption || 'SERIAL_NUMBER',
              });

              return response.data;
            },
            'sheets',
            maxRetries
          );

          results.push(batchResult);

          // Кешування результатів якщо увімкнено
          if (cacheResults && this.serviceContainer) {
            const cacheService = this.serviceContainer.get('cache');
            if (cacheService) {
              for (let j = 0; j < batch.length; j++) {
                const range = batch[j];
                const valueRange = batchResult.valueRanges[j];
                const cacheKey = `sheets:${spreadsheetId}:${range}`;
                await cacheService.set(cacheKey, valueRange, cacheTTL);
              }
            }
          }

        } catch (error) {
          logger.error(`❌ Помилка batch ${i + 1}:`, error);
          failedRanges.push(...batch);
          
          if (!retryFailed) {
            throw error;
          }
        }
      }

      // Повторна спроба для невдалих ranges
      if (retryFailed && failedRanges.length > 0) {
        logger.info(`🔄 Повторна спроба для ${failedRanges.length} ranges`);
        const retryResults = await this.batchGetSheetData(
          spreadsheetId, 
          failedRanges, 
          { ...options, retryFailed: false, maxRetries: 1 }
        );
        results.push(retryResults);
      }

      // Об'єднання результатів
      const combinedResult = {
        valueRanges: results.flatMap(r => r.valueRanges || []),
        spreadsheetId: spreadsheetId || this.config.google.spreadsheetId,
      };

      this.updateStats(true, Date.now() - startTime);
      logger.info(`✅ Batch запит завершено: ${ranges.length} ranges за ${Date.now() - startTime}ms`);
      
      return combinedResult;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка batch отримання даних:', error);
      throw error;
    }
  }

  /**
   * Batch запис в Sheets з оптимізацією
   */
  async batchWriteSheetData(spreadsheetId, data, options = {}) {
    const startTime = Date.now();
    const {
      batchSize = 10,
      valueInputOption = 'RAW',
      retryFailed = true,
      maxRetries = 3,
      clearCache = true
    } = options;

    try {
      // Розбиття data на батчі для оптимізації
      const batches = this.chunkArray(data, batchSize);
      const results = [];
      const failedBatches = [];

      for (let i = 0; i < batches.length; i++) {
        const batch = batches[i];
        logger.info(`📝 Обробка batch запису ${i + 1}/${batches.length} (${batch.length} ranges)`);

        try {
          const batchResult = await this.executeWithRetry(
            async (client) => {
              const response = await client.spreadsheets.values.batchUpdate({
                spreadsheetId: spreadsheetId || this.config.google.spreadsheetId,
                resource: {
                  valueInputOption: valueInputOption,
                  data: batch,
                },
              });

              return response.data;
            },
            'sheets',
            maxRetries
          );

          results.push(batchResult);

          // Очищення кешу для змінених ranges
          if (clearCache && this.serviceContainer) {
            const cacheService = this.serviceContainer.get('cache');
            if (cacheService) {
              for (const item of batch) {
                const cacheKey = `sheets:${spreadsheetId}:${item.range}`;
                await cacheService.delete(cacheKey);
              }
            }
          }

        } catch (error) {
          logger.error(`❌ Помилка batch запису ${i + 1}:`, error);
          failedBatches.push(batch);
          
          if (!retryFailed) {
            throw error;
          }
        }
      }

      // Повторна спроба для невдалих batches
      if (retryFailed && failedBatches.length > 0) {
        logger.info(`🔄 Повторна спроба для ${failedBatches.length} batches`);
        const retryResults = await this.batchWriteSheetData(
          spreadsheetId, 
          failedBatches.flat(), 
          { ...options, retryFailed: false, maxRetries: 1 }
        );
        results.push(retryResults);
      }

      // Об'єднання результатів
      const combinedResult = {
        totalUpdatedCells: results.reduce((sum, r) => sum + (r.totalUpdatedCells || 0), 0),
        totalUpdatedRows: results.reduce((sum, r) => sum + (r.totalUpdatedRows || 0), 0),
        totalUpdatedColumns: results.reduce((sum, r) => sum + (r.totalUpdatedColumns || 0), 0),
        totalUpdatedSheets: results.reduce((sum, r) => sum + (r.totalUpdatedSheets || 0), 0),
        responses: results,
      };

      this.updateStats(true, Date.now() - startTime);
      logger.info(`✅ Batch запис завершено: ${data.length} ranges за ${Date.now() - startTime}ms`);
      
      return combinedResult;
    } catch (error) {
      this.updateStats(false, Date.now() - startTime);
      logger.error('❌ Помилка batch запису:', error);
      throw error;
    }
  }

  /**
   * Отримання статистики з'єднань
   */
  getConnectionStats() {
    const stats = {};
    
    for (const [apiType, connection] of this.connectionPool.entries()) {
      stats[apiType] = {
        inUse: connection.inUse,
        lastUsed: connection.lastUsed,
        idleTime: Date.now() - connection.lastUsed,
      };
    }

    return stats;
  }

  /**
   * Health check
   */
  async onHealthCheck() {
    try {
      // Перевірка автентифікації
      if (!this.auth) {
        return {
          healthy: false,
          error: 'Google автентифікація не ініціалізована',
          service: this.name,
        };
      }

      // Перевірка підключення до Sheets API
      const testData = await this.getSheetData(
        this.config.google.spreadsheetId,
        'A1:A1'
      );

      return {
        healthy: true,
        service: this.name,
        connectionStats: this.getConnectionStats(),
      };
    } catch (error) {
      return {
        healthy: false,
        error: error.message,
        service: this.name,
      };
    }
  }

  /**
   * Завершення роботи
   */
  async onShutdown() {
    try {
      // Очищення connection pool
      this.connectionPool.clear();
      
      logger.info('✅ Google Service завершено');
    } catch (error) {
      logger.error('❌ Помилка завершення Google Service:', error);
    }
  }

  /**
   * Отримання розширеної статистики
   */
  getStats() {
    return {
      ...super.getStats(),
      connections: this.getConnectionStats(),
      maxConnections: this.maxConnections,
      connectionTimeout: this.connectionTimeout,
      retryAttempts: this.retryAttempts,
    };
  }

  /**
   * Розбиття масиву на батчі
   */
  chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Оновлення статистики
   */
  updateStats(success, duration) {
    if (!this.requestStats) {
      this.requestStats = { success: 0, totalDuration: 0, averageDuration: 0 };
    }
    if (!this.errorStats) {
      this.errorStats = { count: 0, lastError: null };
    }

    if (success) {
      this.requestStats.success++;
      this.requestStats.totalDuration += duration;
      this.requestStats.averageDuration = this.requestStats.totalDuration / this.requestStats.success;
    } else {
      this.errorStats.count++;
      this.errorStats.lastError = new Date();
    }
  }
}

module.exports = GoogleService;
