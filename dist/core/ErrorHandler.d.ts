/**
 * Error Handler для Discord бота
 * Централізована обробка помилок з покращеною архітектурою
 * TypeScript версія
 */
interface ErrorType {
    severity: 'critical' | 'error' | 'warn' | 'info';
    category: string;
    retryable: boolean;
    maxRetries: number;
    notificationThreshold: number;
}
interface ErrorInfo extends ErrorType {
    type: string;
}
interface ErrorContext {
    type?: string;
    promise?: Promise<any>;
    [key: string]: any;
}
interface ErrorStats {
    totalErrors: number;
    errorCounts: Record<string, number>;
    notificationQueueSize: number;
    isActive: boolean;
}
interface ServiceContainer {
    get(service: string): any;
    shutdown(): Promise<void>;
}
declare class ErrorHandler {
    private serviceContainer;
    private errorTypes;
    private errorCounts;
    private isActive;
    private notificationQueue;
    private maxQueueSize;
    constructor(serviceContainer: ServiceContainer);
    /**
     * Ініціалізація обробника помилок
     */
    initialize(): Promise<void>;
    /**
     * Реєстрація типів помилок
     */
    private registerErrorTypes;
    /**
     * Обробка помилки
     */
    handle(error: Error, context?: ErrorContext): Promise<{
        handled: boolean;
        errorInfo?: ErrorInfo;
        retryable?: boolean;
        maxRetries?: number;
        error?: Error;
    }>;
    /**
     * Обробка необроблених помилок
     */
    handleUncaughtException(error: Error): void;
    /**
     * Обробка необроблених rejections
     */
    handleUnhandledRejection(reason: any, promise: Promise<any>): void;
    /**
     * Спроба graceful shutdown
     */
    private attemptGracefulShutdown;
    /**
     * Додавання сповіщення до черги
     */
    private queueNotification;
    /**
     * Запуск обробника сповіщень
     */
    private startNotificationProcessor;
    /**
     * Генерація ID для сповіщення
     */
    private generateNotificationId;
    /**
     * Класифікація помилки
     */
    private classifyError;
    /**
     * Логування помилки
     */
    private logError;
    /**
     * Підрахунок помилок
     */
    private incrementErrorCount;
    /**
     * Відправка сповіщення
     */
    private sendNotification;
    /**
     * Перевірка чи потрібно відправити сповіщення
     */
    private shouldSendNotification;
    /**
     * Відправка Discord сповіщення
     */
    private sendDiscordNotification;
    /**
     * Отримання кольору для помилки
     */
    private getErrorColor;
    /**
     * Пошук каналу для сповіщень
     */
    private findNotificationChannel;
    /**
     * Отримання зрозумілого повідомлення про помилку
     */
    private getUserFriendlyMessage;
    /**
     * Отримання статистики помилок
     */
    getStats(): ErrorStats;
    /**
     * Очищення статистики помилок
     */
    clearErrorStats(): void;
    /**
     * Перевірка активності
     */
    isActive(): boolean;
    /**
     * Завершення роботи
     */
    shutdown(): Promise<void>;
}
export { ErrorHandler };
//# sourceMappingURL=ErrorHandler.d.ts.map