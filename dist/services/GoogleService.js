"use strict";
/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.GoogleService = void 0;
const googleapis_1 = require("googleapis");
const BaseService_1 = require("@/core/BaseService");
const CacheService_1 = require("./CacheService");
const logger_1 = require("@/utils/logger");
class GoogleService extends BaseService_1.BaseService {
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
        this.cacheService = new CacheService_1.CacheService(config);
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
    async onInitialize() {
        try {
            logger_1.logger.info('🔧 Ініціалізація Google Service...');
            // Ініціалізація кешу
            await this.cacheService.initialize();
            // Створення автентифікації
            await this.initializeAuth();
            // Ініціалізація API клієнтів
            await this.initializeAPIs();
            // Створення connection pool
            await this.initializeConnectionPool();
            logger_1.logger.info('✅ Google Service ініціалізовано');
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка ініціалізації Google Service:', error);
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
            this.auth = new googleapis_1.google.auth.JWT(this.config.google.credentials.client_email, null, this.config.google.credentials.private_key, [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/documents',
            ]);
            // Авторизація
            await this.auth.authorize();
            logger_1.logger.info('✅ Google автентифікація успішна');
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка Google автентифікації:', error);
            throw error;
        }
    }
    /**
     * Ініціалізація API клієнтів
     */
    async initializeAPIs() {
        try {
            // Google Sheets API
            this.sheets = googleapis_1.google.sheets({ version: 'v4', auth: this.auth });
            // Google Drive API
            this.drive = googleapis_1.google.drive({ version: 'v3', auth: this.auth });
            // Google Docs API
            this.docs = googleapis_1.google.docs({ version: 'v1', auth: this.auth });
            logger_1.logger.info('✅ Google API клієнти ініціалізовано');
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка ініціалізації Google API:', error);
            throw error;
        }
    }
    /**
     * Ініціалізація Connection Pool
     */
    async initializeConnectionPool() {
        try {
            const apiTypes = ['sheets', 'drive', 'docs'];
            for (const apiType of apiTypes) {
                this.connectionPool.set(apiType, {
                    inUse: false,
                    lastUsed: Date.now(),
                    requestCount: 0,
                });
            }
            logger_1.logger.info('✅ Connection Pool ініціалізовано');
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка ініціалізації Connection Pool:', error);
            throw error;
        }
    }
    /**
     * Отримання з'єднання з пулу
     */
    getConnection(apiType) {
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
    releaseConnection(apiType) {
        const connection = this.connectionPool.get(apiType);
        if (connection) {
            connection.inUse = false;
        }
    }
    /**
     * Виконання операції з retry
     */
    async executeWithRetry(operation, apiType, maxRetries = this.retryAttempts) {
        let lastError = null;
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
            }
            catch (error) {
                lastError = error;
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
    async getSheetData(spreadsheetId, range, options = {}) {
        const { useCache = true, cacheTTL = 300, forceRefresh = false } = options;
        try {
            // Перевірка кешу
            if (useCache && !forceRefresh) {
                const cacheKey = `sheets:${spreadsheetId}:${range}`;
                try {
                    const cached = await this.cacheService.get(cacheKey);
                    if (cached) {
                        this.stats.cacheHits++;
                        logger_1.logger.debug('✅ Використано кешовані дані Sheets', {
                            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
                            range,
                            rowsCount: cached.values.length
                        });
                        return cached;
                    }
                    else {
                        this.stats.cacheMisses++;
                    }
                }
                catch (cacheError) {
                    logger_1.logger.warn('⚠️ Помилка читання з кешу Sheets:', cacheError);
                    this.stats.cacheMisses++;
                }
            }
            const result = await this.executeWithRetry(async () => {
                if (!this.sheets)
                    throw new Error('Sheets API не ініціалізовано');
                const response = await this.sheets.spreadsheets.values.get({
                    spreadsheetId,
                    range,
                });
                return {
                    range: response.data.range || range,
                    majorDimension: response.data.majorDimension || 'ROWS',
                    values: response.data.values || [],
                };
            }, 'sheets');
            // Збереження в кеш
            if (useCache) {
                const cacheKey = `sheets:${spreadsheetId}:${range}`;
                try {
                    await this.cacheService.set(cacheKey, result, { ttl: cacheTTL * 1000 });
                    logger_1.logger.debug('💾 Дані Sheets збережено в кеш', {
                        spreadsheetId: spreadsheetId.substring(0, 10) + '...',
                        range,
                        rowsCount: result.values.length,
                        ttl: `${cacheTTL}s`
                    });
                }
                catch (cacheError) {
                    logger_1.logger.warn('⚠️ Помилка збереження в кеш Sheets:', cacheError);
                }
            }
            return result;
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка отримання даних з Sheets:', error);
            throw error;
        }
    }
    /**
     * Запис даних в Google Sheets
     */
    async writeSheetData(spreadsheetId, range, values, options = {}) {
        const { valueInputOption = 'RAW', clearCache = true } = options;
        try {
            await this.executeWithRetry(async () => {
                if (!this.sheets)
                    throw new Error('Sheets API не ініціалізовано');
                await this.sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range,
                    valueInputOption,
                    requestBody: {
                        values,
                    },
                });
            }, 'sheets');
            // Очищення кешу
            if (clearCache) {
                const cacheKey = `sheets:${spreadsheetId}:${range}`;
                try {
                    await this.cacheService.delete(cacheKey);
                    logger_1.logger.debug('🗑️ Кеш Sheets очищено', {
                        spreadsheetId: spreadsheetId.substring(0, 10) + '...',
                        range
                    });
                }
                catch (cacheError) {
                    logger_1.logger.warn('⚠️ Помилка очищення кешу Sheets:', cacheError);
                }
            }
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка запису в Sheets:', error);
            throw error;
        }
    }
    /**
     * Batch отримання даних з Google Sheets
     */
    async batchGetSheetData(spreadsheetId, ranges, options = {}) {
        const { batchSize = 10, cacheResults = true, cacheTTL = 300, retryFailed = true, maxRetries = 3 } = options;
        try {
            const chunks = this.chunkArray(ranges, batchSize);
            const results = [];
            const failedRanges = [];
            for (const chunk of chunks) {
                try {
                    const result = await this.executeWithRetry(async () => {
                        if (!this.sheets)
                            throw new Error('Sheets API не ініціалізовано');
                        const response = await this.sheets.spreadsheets.values.batchGet({
                            spreadsheetId,
                            ranges: chunk,
                        });
                        return response.data.valueRanges || [];
                    }, 'sheets', maxRetries);
                    results.push(...result);
                }
                catch (error) {
                    logger_1.logger.error('❌ Помилка batch запиту:', error);
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
                    }
                    catch (error) {
                        logger_1.logger.error(`❌ Повторна спроба невдала для range: ${range}`, error);
                    }
                }
            }
            return {
                valueRanges: results,
                spreadsheetId,
            };
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка batch отримання даних:', error);
            throw error;
        }
    }
    /**
     * Batch запис даних в Google Sheets
     */
    async batchWriteSheetData(spreadsheetId, data, options = {}) {
        const { batchSize = 10, valueInputOption = 'RAW', retryFailed = true, maxRetries = 3, clearCache = true } = options;
        try {
            const chunks = this.chunkArray(data, batchSize);
            const failedBatches = [];
            for (const chunk of chunks) {
                try {
                    await this.executeWithRetry(async () => {
                        if (!this.sheets)
                            throw new Error('Sheets API не ініціалізовано');
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
                    }, 'sheets', maxRetries);
                    // Очищення кешу
                    if (clearCache) {
                        for (const item of chunk) {
                            const cacheKey = `sheets:${spreadsheetId}:${item.range}`;
                            // TODO: Очистити кеш
                        }
                    }
                }
                catch (error) {
                    logger_1.logger.error('❌ Помилка batch запису:', error);
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
                    }
                    catch (error) {
                        logger_1.logger.error(`❌ Повторна спроба невдала для range: ${item.range}`, error);
                    }
                }
            }
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка batch запису даних:', error);
            throw error;
        }
    }
    /**
     * Пошук файлів в Google Drive
     */
    async searchFiles(query, options = {}) {
        try {
            const result = await this.executeWithRetry(async () => {
                if (!this.drive)
                    throw new Error('Drive API не ініціалізовано');
                const response = await this.drive.files.list({
                    q: query,
                    fields: 'files(id,name,mimeType,size,modifiedTime)',
                    pageSize: 100,
                });
                return response.data.files || [];
            }, 'drive');
            return result;
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка пошуку файлів:', error);
            throw error;
        }
    }
    /**
     * Отримання метаданих файлу
     */
    async getFileMetadata(fileId, fields = '*') {
        try {
            const result = await this.executeWithRetry(async () => {
                if (!this.drive)
                    throw new Error('Drive API не ініціалізовано');
                const response = await this.drive.files.get({
                    fileId,
                    fields,
                });
                return response.data;
            }, 'drive');
            return result;
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка отримання метаданих файлу:', error);
            throw error;
        }
    }
    /**
     * Отримання контенту документа
     */
    async getDocumentContent(documentId) {
        try {
            const result = await this.executeWithRetry(async () => {
                if (!this.docs)
                    throw new Error('Docs API не ініціалізовано');
                const response = await this.docs.documents.get({
                    documentId,
                });
                // Парсинг контенту документа
                const content = this.parseDocumentContent(response.data);
                return content;
            }, 'docs');
            return result;
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка отримання контенту документа:', error);
            throw error;
        }
    }
    /**
     * Парсинг контенту документа
     */
    parseDocumentContent(document) {
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
    getConnectionStats() {
        const stats = {};
        for (const [apiType, connection] of this.connectionPool.entries()) {
            stats[apiType] = { ...connection };
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
            }
            catch (error) {
                return {
                    healthy: false,
                    service: this.name,
                    error: `Помилка тестового запиту: ${error}`,
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
        }
        catch (error) {
            return {
                healthy: false,
                service: this.name,
                error: `Health check failed: ${error}`,
            };
        }
    }
    /**
     * Завершення роботи
     */
    async onShutdown() {
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
            logger_1.logger.info('✅ Google Service зупинено');
        }
        catch (error) {
            logger_1.logger.error('❌ Помилка зупинки Google Service:', error);
            throw error;
        }
    }
    /**
     * Отримання статистики
     */
    onGetStats() {
        return this.stats;
    }
    /**
     * Розбивка масиву на чанки
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
exports.GoogleService = GoogleService;
//# sourceMappingURL=GoogleService.js.map