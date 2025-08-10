/**
 * Утиліта для повторних спроб операцій
 * TypeScript версія
 */
interface RetryOptions {
    maxAttempts?: number;
    delay?: number;
    backoff?: 'fixed' | 'exponential' | 'linear';
    factor?: number;
    maxDelay?: number;
    timeout?: number;
    onRetry?: (attempt: number, error: Error) => void;
    shouldRetry?: (error: Error) => boolean;
}
interface RetryResult<T> {
    success: boolean;
    data?: T;
    error?: Error;
    attempts: number;
    totalTime: number;
}
declare class RetryManager {
    private static defaultOptions;
    /**
     * Виконання операції з повторними спробами
     */
    static execute<T>(operation: () => Promise<T>, options?: RetryOptions): Promise<RetryResult<T>>;
    /**
     * Розрахунок затримки між спробами
     */
    private static calculateDelay;
    /**
     * Затримка виконання
     */
    private static sleep;
    /**
     * Створення функції з повторними спробами
     */
    static createRetryFunction<T extends (...args: any[]) => Promise<any>>(fn: T, options?: RetryOptions): (...args: Parameters<T>) => Promise<RetryResult<Awaited<ReturnType<T>>>>;
    /**
     * Retry для HTTP запитів
     */
    static httpRequest<T>(requestFn: () => Promise<T>, options?: RetryOptions): Promise<RetryResult<T>>;
    /**
     * Retry для операцій з базою даних
     */
    static databaseOperation<T>(operation: () => Promise<T>, options?: RetryOptions): Promise<RetryResult<T>>;
    /**
     * Retry для файлових операцій
     */
    static fileOperation<T>(operation: () => Promise<T>, options?: RetryOptions): Promise<RetryResult<T>>;
    /**
     * Retry для Discord API операцій
     */
    static discordOperation<T>(operation: () => Promise<T>, options?: RetryOptions): Promise<RetryResult<T>>;
}
export default RetryManager;
export { RetryManager };
//# sourceMappingURL=retry.d.ts.map