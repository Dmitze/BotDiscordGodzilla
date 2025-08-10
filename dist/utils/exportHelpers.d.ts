interface ExportOptions {
    filename?: string;
    sheetName?: string;
    includeMetadata?: boolean;
    metadata?: Record<string, any>;
    format?: 'xlsx' | 'csv';
    userId?: string;
    guildId?: string;
}
interface ExportResult {
    filePath: string;
    fileSize: number;
    format: string;
    rows: number;
    columns: number;
}
interface AnalysisData {
    type?: string;
    results?: Record<string, any>;
}
declare class ExportHelpers {
    private tmpDir;
    constructor();
    /**
     * Створення тимчасової папки
     */
    private ensureTmpDir;
    /**
     * Експорт в Excel з метаданими
     */
    exportToExcel(data: any[][], headers: string[], options?: ExportOptions): Promise<ExportResult>;
    /**
     * Експорт в CSV з метаданими
     */
    exportToCSV(data: any[][], headers: string[], options?: ExportOptions): Promise<ExportResult>;
    /**
     * Створення аркушу з метаданими
     */
    private createMetadataSheet;
    /**
     * Створення метаданих для CSV
     */
    private createMetadataCSV;
    /**
     * Експорт результатів пошуку
     */
    exportSearchResults(results: any[][], headers: string[], searchFilters: any, options?: ExportOptions): Promise<ExportResult>;
    /**
     * Експорт всієї таблиці
     */
    exportFullTable(data: any[][], headers: string[], options?: ExportOptions): Promise<ExportResult>;
    /**
     * Створення звіту з аналізом даних
     */
    exportAnalysisReport(data: any[][], _headers: string[], analysis: AnalysisData, options?: ExportOptions): Promise<ExportResult>;
    /**
     * Очищення старих файлів
     */
    cleanupOldFiles(): void;
    /**
     * Запис метрик експорту
     */
    private recordExportMetrics;
    /**
     * Отримання статистики експорту
     */
    getExportStats(): any;
    /**
     * Валідація розміру файлу
     */
    validateFileSize(fileSize: number): boolean;
}
export default ExportHelpers;
//# sourceMappingURL=exportHelpers.d.ts.map