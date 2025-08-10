/**
 * AI Service для Discord бота
 * Централізоване управління AI функціоналом
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { BotConfig, HealthStatus, ServiceStats, AIResponse, AIRequestOptions } from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
interface AIServiceStats extends ServiceStats {
    totalRequests: number;
    successfulRequests: number;
    failedRequests: number;
    averageResponseTime: number;
    totalResponseTime: number;
    cacheHits: number;
    cacheMisses: number;
    providerSwitches: number;
    contextCleanups: number;
}
interface ConversationContext {
    messages: Array<{
        role: 'user' | 'assistant' | 'system';
        content: string;
    }>;
    timestamp: number;
    requestCount: number;
}
export declare class AIService extends BaseServiceClass {
    private providers;
    private currentProvider;
    private conversationMemory;
    private stats;
    private memoryCleanupInterval;
    private healthCheckInterval;
    private cacheService;
    constructor(config: BotConfig);
    /**
     * Ініціалізація AI сервісу з детальним логуванням
     */
    protected onInitialize(): Promise<void>;
    /**
     * Створення AI провайдерів з детальним логуванням
     */
    private createProviders;
    /**
     * Створення OpenAI провайдера з покращеною обробкою помилок
     */
    private createOpenAIProvider;
    /**
     * Створення Ollama провайдера з покращеною обробкою помилок
     */
    private createOllamaProvider;
    /**
     * Валідація конфігурації з детальним логуванням
     */
    private validateConfiguration;
    /**
     * Генерація відповіді з покращеною обробкою помилок
     */
    generateResponse(prompt: string, options?: AIRequestOptions): Promise<AIResponse>;
    /**
     * Валідація та санітизація промпту
     */
    private validateAndSanitizePrompt;
    /**
     * Аналіз даних з детальним логуванням
     */
    analyzeData(data: string, analysisType?: 'summary' | 'sentiment' | 'keywords'): Promise<AIResponse>;
    /**
     * Генерація звіту з детальним логуванням
     */
    generateReport(data: string, options?: {
        format?: string;
        length?: string;
    }): Promise<AIResponse>;
    /**
     * Обробка природномовного запиту з детальним логуванням
     */
    processNaturalLanguageQuery(userId: string, userInput: string, context?: Record<string, unknown>): Promise<AIResponse>;
    /**
     * Отримання контексту розмови
     */
    getConversationContext(userId: string): ConversationContext;
    /**
     * Збереження в контекст з валідацією
     */
    saveToContext(userId: string, role: 'user' | 'assistant' | 'system', content: string): void;
    /**
     * Очищення контексту
     */
    clearContext(userId: string): void;
    /**
     * Створення промпту для аналізу
     */
    private buildAnalysisPrompt;
    /**
     * Створення промпту для звіту
     */
    private buildReportPrompt;
    /**
     * Створення промпту для розмови
     */
    private buildConversationPrompt;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Створення ключа кешу
     */
    private buildCacheKey;
    /**
     * Запуск очищення пам'яті
     */
    private startMemoryCleanup;
    /**
     * Запуск health check
     */
    private startHealthCheck;
    /**
     * Очищення пам'яті з детальним логуванням
     */
    private cleanupMemory;
    /**
     * Health check з детальним логуванням
     */
    protected onHealthCheck(): Promise<HealthStatus>;
    /**
     * Завершення роботи з детальним логуванням
     */
    protected onShutdown(): Promise<void>;
    /**
     * Отримання статистики з детальним логуванням
     */
    protected onGetStats(): Partial<AIServiceStats>;
}
export {};
//# sourceMappingURL=AIService.d.ts.map