/**
 * Розширений логер для Discord AI Assistant Bot
 * Рефакторована версія з покращеними можливостями
 * TypeScript версія 3.0.0 - Повністю рефакторовано
 */
interface LogMeta {
    [key: string]: any;
    timestamp?: string;
    level?: string;
    service?: string;
    userId?: string;
    guildId?: string;
    channelId?: string;
    requestId?: string;
    correlationId?: string;
    type?: string;
    severity?: string;
    category?: string;
    component?: string;
    logLevel?: string;
    processId?: number;
    memory?: NodeJS.MemoryUsage;
}
interface LoggerStats {
    totalLogs: number;
    errors: number;
    commands: number;
    apiRequests: number;
    performance: number;
    security: number;
    system: number;
    debug: number;
    warnings: number;
    lastLogTime: Date;
    averageLogSize: number;
    logBufferSize: number;
}
interface LogEntry {
    timestamp: Date;
    level: string;
    message: string;
    meta: LogMeta;
    size: number;
}
declare class Logger {
    private logger;
    private stats;
    private logBuffer;
    private cleanupInterval;
    private flushInterval;
    private isInitialized;
    private readonly logsDir;
    constructor();
    /**
     * Санітізація метаданих логів: маскує секрети, обрізає великі значення, прибирає цикли
     */
    private sanitizeMeta;
    /**
     * Ініціалізація логера з детальним логуванням
     */
    private initialize;
    /**
     * Створення папки для логів
     */
    private ensureLogsDirectory;
    /**
     * Створення форматів логування
     */
    private createFormats;
    /**
     * Створення транспортів
     */
    private createTransports;
    /**
     * Отримання рівня логування
     */
    private getLogLevel;
    /**
     * Налаштування обробки необроблених помилок
     */
    private setupExceptionHandling;
    /**
     * Запуск періодичних завдань
     */
    private startPeriodicTasks;
    /**
     * Створення резервного логера
     */
    private createFallbackLogger;
    /**
     * Логування з детальною інформацією
     */
    private log;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Додавання до буфера логів
     */
    private addToBuffer;
    /**
     * Скидання буфера логів
     */
    private flushLogBuffer;
    /**
     * Очищення старих логів
     */
    private cleanupOldLogs;
    /**
     * Логування інформації
     */
    info(message: string, meta?: LogMeta): void;
    /**
     * Логування помилок
     */
    error(message: string, meta?: LogMeta): void;
    /**
     * Логування попереджень
     */
    warn(message: string, meta?: LogMeta): void;
    /**
     * Логування дебагу
     */
    debug(message: string, meta?: LogMeta): void;
    /**
     * Логування команд з детальною інформацією
     */
    command(command: string, user: string, duration: number, success?: boolean, meta?: LogMeta): void;
    /**
     * Логування помилок команд
     */
    commandError(command: string, user: string, error: Error, duration: number, meta?: LogMeta): void;
    /**
     * Логування API запитів
     */
    apiRequest(service: string, endpoint: string, duration: number, success?: boolean, meta?: LogMeta): void;
    /**
     * Логування помилок API
     */
    apiError(service: string, endpoint: string, error: Error, duration: number, meta?: LogMeta): void;
    /**
     * Логування подій безпеки
     */
    security(event: string, user: string, details?: LogMeta): void;
    /**
     * Логування продуктивності
     */
    performance(operation: string, duration: number, details?: LogMeta): void;
    /**
     * Логування системних подій
     */
    system(event: string, details?: LogMeta): void;
    /**
     * Отримання детальної статистики логера
     */
    getStats(): LoggerStats;
    /**
     * Отримання буфера логів
     */
    getLogBuffer(): LogEntry[];
    /**
     * Очищення ресурсів
     */
    cleanup(): Promise<void>;
    /**
     * Перевірка стану логера
     */
    isHealthy(): boolean;
}
declare const logger: Logger;
export default logger;
export { Logger, type LogMeta, type LoggerStats, type LogEntry };
//# sourceMappingURL=logger.d.ts.map