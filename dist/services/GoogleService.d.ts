/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */
import { drive_v3 } from 'googleapis';
import type { BotConfig, HealthStatus, ServiceStats, SheetData, BatchSheetData } from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
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
    maxRetries?: number;
    valueInputOption?: string;
    clearCache?: boolean;
}
export declare class GoogleService extends BaseServiceClass {
    private auth;
    private sheets;
    private drive;
    private docs;
    private connectionPool;
    private readonly maxConnections;
    private readonly connectionTimeout;
    private readonly retryAttempts;
    private readonly retryDelay;
    private stats;
    private cacheService;
    constructor(config: BotConfig);
    /**
     * Ініціалізація Google сервісів
     */
    protected onInitialize(): Promise<void>;
    /**
     * Ініціалізація автентифікації
     */
    private initializeAuth;
    /**
     * Ініціалізація API клієнтів
     */
    private initializeAPIs;
    /**
     * Ініціалізація Connection Pool
     */
    private initializeConnectionPool;
    /**
     * Отримання з'єднання з пулу
     */
    private getConnection;
    /**
     * Звільнення з'єднання
     */
    private releaseConnection;
    /**
     * Виконання операції з retry
     */
    private executeWithRetry;
    /**
     * Отримання даних з Google Sheets
     */
    getSheetData(spreadsheetId: string, range: string, options?: GoogleServiceOptions): Promise<SheetData>;
    /**
     * Запис даних в Google Sheets
     */
    writeSheetData(spreadsheetId: string, range: string, values: string[][], options?: GoogleServiceOptions): Promise<void>;
    /**
     * Batch отримання даних з Google Sheets
     */
    batchGetSheetData(spreadsheetId: string, ranges: string[], options?: GoogleServiceOptions): Promise<BatchSheetData>;
    /**
     * Batch запис даних в Google Sheets
     */
    batchWriteSheetData(spreadsheetId: string, data: Array<{
        range: string;
        values: string[][];
    }>, options?: GoogleServiceOptions): Promise<void>;
    /**
     * Пошук файлів в Google Drive
     */
    searchFiles(query: string, options?: GoogleServiceOptions): Promise<drive_v3.Schema$File[]>;
    /**
     * Отримання метаданих файлу
     */
    getFileMetadata(fileId: string, fields?: string): Promise<drive_v3.Schema$File>;
    /**
     * Отримання контенту документа
     */
    getDocumentContent(documentId: string): Promise<string>;
    /**
     * Парсинг контенту документа
     */
    private parseDocumentContent;
    /**
     * Отримання статистики з'єднань
     */
    getConnectionStats(): Record<string, ConnectionInfo>;
    /**
     * Health check
     */
    protected onHealthCheck(): Promise<HealthStatus>;
    /**
     * Завершення роботи
     */
    protected onShutdown(): Promise<void>;
    /**
     * Отримання статистики
     */
    protected onGetStats(): Partial<GoogleServiceStats>;
    /**
     * Розбивка масиву на чанки
     */
    private chunkArray;
    /**
     * Оновлення статистики
     */
    private updateStats;
}
export {};
//# sourceMappingURL=GoogleService.d.ts.map