/**
 * Утиліти для форматування даних
 * TypeScript версія
 */
interface Metrics {
    [key: string]: number | string;
}
interface Stats {
    total: number;
    success: number;
    errors: number;
    avgTime: number;
    [key: string]: any;
}
declare class DataFormatters {
    /**
     * Форматування числа з роздільниками
     */
    static formatNumber(num: number | null | undefined, locale?: string): string;
    /**
     * Форматування валюти
     */
    static formatCurrency(amount: number | null | undefined, currency?: string, locale?: string): string;
    /**
     * Форматування дати
     */
    static formatDate(date: Date | string | null | undefined, locale?: string): string;
    /**
     * Форматування часу роботи
     */
    static formatUptime(ms: number): string;
    /**
     * Форматування розміру файлу
     */
    static formatFileSize(bytes: number): string;
    /**
     * Форматування таблиці для Discord
     */
    static formatTable(data: any[][], headers: string[], maxRows?: number): string;
    /**
     * Форматування прогрес-бару
     */
    static formatProgress(current: number, total: number, width?: number): string;
    /**
     * Форматування статусу
     */
    static formatStatus(status: string, showIcon?: boolean): string;
    /**
     * Форматування метрик
     */
    static formatMetrics(metrics: Metrics): string;
    /**
     * Форматування помилки
     */
    static formatError(error: Error | string, includeDetails?: boolean): string;
    /**
     * Форматування часу виконання
     */
    static formatExecutionTime(startTime: number): string;
    /**
     * Форматування списку
     */
    static formatList(items: string[], title?: string | null, maxItems?: number): string;
    /**
     * Форматування статистики
     */
    static formatStats(stats: Stats): string;
    /**
     * Форматування дати та часу
     */
    static formatDateTime(date: Date | string, locale?: string): string;
    /**
     * Форматування відсотків
     */
    static formatPercentage(value: number, total: number, decimals?: number): string;
    /**
     * Обрізання тексту
     */
    static truncateText(text: string, maxLength: number, suffix?: string): string;
    /**
     * Капіталізація першої літери
     */
    static capitalizeFirst(text: string): string;
}
export default DataFormatters;
export { DataFormatters };
//# sourceMappingURL=formatters.d.ts.map