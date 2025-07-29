/**
 * AI Service для Discord бота
 * Централізоване управління AI функціоналом
 */

import OpenAI from 'openai';
import type { 
  BaseService, 
  BotConfig, 
  HealthStatus, 
  ServiceStats,
  AIResponse,
  AIRequest,
  AIRequestOptions
} from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
// TODO: Створити типизовані утиліти
const logger = {
  info: (message: string, ...args: unknown[]) => console.log(message, ...args),
  error: (message: string, ...args: unknown[]) => console.error(message, ...args),
  warn: (message: string, ...args: unknown[]) => console.warn(message, ...args),
  debug: (message: string, ...args: unknown[]) => console.debug(message, ...args),
};

const sanitizeInput = (input: string): string => {
  return input.trim().replace(/[<>]/g, '');
};

interface AIServiceStats extends ServiceStats {
  totalRequests: number;
  successfulRequests: number;
  failedRequests: number;
  averageResponseTime: number;
  totalResponseTime: number;
  cacheHits: number;
  cacheMisses: number;
}

interface ConversationContext {
  messages: Array<{ role: 'user' | 'assistant' | 'system'; content: string }>;
  timestamp: number;
  requestCount: number;
}

interface AIProvider {
  generate(prompt: string, options?: AIRequestOptions): Promise<AIResponse>;
}

interface OllamaProvider {
  host: string;
  model: string;
}

export class AIService extends BaseServiceClass {
  private providers: Record<string, AIProvider> = {};
  private currentProvider: string;
  private conversationMemory = new Map<string, ConversationContext>();
  private stats: AIServiceStats;
  private memoryCleanupInterval: NodeJS.Timeout | null = null;

  constructor(config: BotConfig) {
    super('AIService', config);
    this.currentProvider = config.ai.provider;
    this.stats = {
      service: 'AIService',
      uptime: 0,
      requests: 0,
      errors: 0,
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      averageResponseTime: 0,
      totalResponseTime: 0,
      cacheHits: 0,
      cacheMisses: 0,
    };
  }

  /**
   * Ініціалізація AI сервісу
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('🤖 Ініціалізація AI сервісу...');

      // Створення провайдерів
      await this.createProviders();

      // Валідація конфігурації
      this.validateConfiguration();

      // Запуск очищення пам'яті
      this.startMemoryCleanup();

      logger.info('✅ AI сервіс ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації AI сервісу:', error);
      throw error;
    }
  }

  /**
   * Створення AI провайдерів
   */
  private async createProviders(): Promise<void> {
    // OpenAI провайдер
    if (this.config.ai['openai'].apiKey) {
      this.providers['openai'] = this.createOpenAIProvider();
      logger.debug('✅ OpenAI провайдер створено');
    }

    // Ollama провайдер
    if (this.config.ai['ollama'].host) {
      this.providers['ollama'] = this.createOllamaProvider();
      logger.debug('✅ Ollama провайдер створено');
    }

    if (Object.keys(this.providers).length === 0) {
      throw new Error('Жоден AI провайдер не налаштовано');
    }
  }

  /**
   * Створення OpenAI провайдера
   */
  private createOpenAIProvider(): AIProvider {
    try {
      const openai = new OpenAI({
        apiKey: this.config.ai.openai.apiKey,
        maxRetries: 3,
        timeout: 30000,
      });

      return {
        async generate(prompt: string, options: AIRequestOptions = {}): Promise<AIResponse> {
          const startTime = Date.now();
          
          try {
            const response = await openai.chat.completions.create({
              model: options.model || this.config.ai.openai.model,
              messages: [{ role: 'user', content: prompt }],
              max_tokens: options.maxTokens || this.config.ai.openai.maxTokens,
              temperature: options.temperature || this.config.ai.openai.temperature,
            });

            const duration = Date.now() - startTime;
            
            return {
              content: response.choices[0]?.message?.content || '',
              provider: 'openai',
              model: response.model,
              tokens: response.usage?.total_tokens || 0,
              duration,
            };
          } catch (error) {
            throw new Error(`OpenAI error: ${error}`);
          }
        },
      };
    } catch (error) {
      logger.error('Помилка створення OpenAI провайдера:', error);
      throw error;
    }
  }

  /**
   * Створення Ollama провайдера
   */
  private createOllamaProvider(): AIProvider {
    const ollamaConfig = this.config.ai.ollama;
    
    return {
      async generate(prompt: string, options: AIRequestOptions = {}): Promise<AIResponse> {
        const startTime = Date.now();
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000);

        try {
          const response = await fetch(`${ollamaConfig.host}/api/generate`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              model: options.model || ollamaConfig.model,
              prompt,
              stream: false,
              options: {
                temperature: options.temperature || 0.7,
                num_predict: options.maxTokens || 1000,
              },
            }),
            signal: controller.signal,
          });

          clearTimeout(timeoutId);

          if (!response.ok) {
            throw new Error(`Ollama API error: ${response.statusText}`);
          }

          const data = await response.json();
          const duration = Date.now() - startTime;

          return {
            content: data.response || '',
            provider: 'ollama',
            model: data.model || ollamaConfig.model,
            tokens: data.eval_count || 0,
            duration,
          };
        } catch (error) {
          clearTimeout(timeoutId);
          throw new Error(`Ollama error: ${error}`);
        }
      },
    };
  }

  /**
   * Валідація конфігурації
   */
  private validateConfiguration(): void {
    if (!this.providers[this.currentProvider]) {
      throw new Error(`Поточний провайдер ${this.currentProvider} не налаштовано`);
    }

    logger.info(`✅ AI конфігурація валідна, активний провайдер: ${this.currentProvider}`);
  }

  /**
   * Генерація відповіді
   */
  public async generateResponse(
    prompt: string,
    options: AIRequestOptions = {}
  ): Promise<AIResponse> {
    const {
      useCache = true,
      cacheTTL = 3600,
      forceRefresh = false,
      retryAttempts = 3,
      timeout = 30000,
      provider = this.currentProvider,
    } = options;

    try {
      // Перевірка кешу
      if (useCache && !forceRefresh) {
        const cacheKey = this.hashPrompt(prompt);
        // TODO: Реалізувати кешування через CacheService
        // const cached = await this.cacheService.get(cacheKey);
        // if (cached) {
        //   this.stats.cacheHits++;
        //   return cached;
        // }
      }

      // Sanitize input
      const sanitizedPrompt = sanitizeInput(prompt);
      if (!sanitizedPrompt) {
        throw new Error('Порожній або невалідний промпт');
      }

      // Retry logic
      let lastError: Error | null = null;
      for (let attempt = 0; attempt <= retryAttempts; attempt++) {
        try {
          const startTime = Date.now();
          
          // Спробувати основний провайдер
          let response: AIResponse;
          if (this.providers[provider]) {
            response = await this.providers[provider].generate(sanitizedPrompt, options);
          } else {
            // Fallback до Ollama
            response = await this.providers.ollama.generate(sanitizedPrompt, options);
          }

          const duration = Date.now() - startTime;
          this.updateStats(true, duration);

          // Збереження в кеш
          if (useCache) {
            const cacheKey = this.hashPrompt(prompt);
            // TODO: Зберегти в кеш
            // await this.cacheService.set(cacheKey, response, cacheTTL);
          }

          return response;
        } catch (error) {
          lastError = error as Error;
          
          if (attempt < retryAttempts) {
            const delay = 1000 * Math.pow(2, attempt);
            await new Promise(resolve => setTimeout(resolve, delay));
            logger.warn(`Спроба ${attempt + 1} невдала, повтор через ${delay}ms`);
          }
        }
      }

      throw lastError || new Error('Всі спроби генерації невдалі');
    } catch (error) {
      this.updateStats(false, 0);
      logger.error('❌ Помилка генерації відповіді:', error);
      throw error;
    }
  }

  /**
   * Аналіз даних
   */
  public async analyzeData(
    data: string,
    analysisType: 'summary' | 'sentiment' | 'keywords' = 'summary'
  ): Promise<AIResponse> {
    const prompt = this.buildAnalysisPrompt(data, analysisType);
    return this.generateResponse(prompt, { provider: this.currentProvider });
  }

  /**
   * Генерація звіту
   */
  public async generateReport(
    data: string,
    options: { format?: string; length?: string } = {}
  ): Promise<AIResponse> {
    const prompt = this.buildReportPrompt(data, options);
    return this.generateResponse(prompt, { provider: this.currentProvider });
  }

  /**
   * Обробка природномовного запиту
   */
  public async processNaturalLanguageQuery(
    userId: string,
    userInput: string,
    context: Record<string, unknown> = {}
  ): Promise<AIResponse> {
    const conversationContext = this.getConversationContext(userId);
    const prompt = this.buildConversationPrompt(userInput, conversationContext, context);
    
    const response = await this.generateResponse(prompt, { provider: this.currentProvider });
    
    // Збереження в контекст
    this.saveToContext(userId, 'user', userInput);
    this.saveToContext(userId, 'assistant', response.content);
    
    return response;
  }

  /**
   * Отримання контексту розмови
   */
  public getConversationContext(userId: string): ConversationContext {
    const context = this.conversationMemory.get(userId);
    if (!context) {
      return {
        messages: [],
        timestamp: Date.now(),
        requestCount: 0,
      };
    }
    return context;
  }

  /**
   * Збереження в контекст
   */
  public saveToContext(userId: string, role: 'user' | 'assistant' | 'system', content: string): void {
    let context = this.conversationMemory.get(userId);
    
    if (!context) {
      context = {
        messages: [],
        timestamp: Date.now(),
        requestCount: 0,
      };
    }

    context.messages.push({ role, content });
    context.timestamp = Date.now();
    context.requestCount++;

    // Обмеження розміру контексту
    if (context.messages.length > 20) {
      context.messages = context.messages.slice(-10);
    }

    this.conversationMemory.set(userId, context);
  }

  /**
   * Очищення контексту
   */
  public clearContext(userId: string): void {
    this.conversationMemory.delete(userId);
  }

  /**
   * Створення промпту для аналізу
   */
  private buildAnalysisPrompt(data: string, analysisType: string): string {
    const prompts = {
      summary: `Надай короткий зміст наступного тексту:\n\n${data}`,
      sentiment: `Проаналізуй емоційний тон наступного тексту:\n\n${data}`,
      keywords: `Виділи ключові слова з наступного тексту:\n\n${data}`,
    };

    return prompts[analysisType as keyof typeof prompts] || prompts.summary;
  }

  /**
   * Створення промпту для звіту
   */
  private buildReportPrompt(data: string, options: { format?: string; length?: string }): string {
    const format = options.format || 'text';
    const length = options.length || 'medium';
    
    return `Створи ${length} звіт у форматі ${format} на основі наступних даних:\n\n${data}`;
  }

  /**
   * Створення промпту для розмови
   */
  private buildConversationPrompt(
    userInput: string,
    context: ConversationContext,
    additionalContext: Record<string, unknown> = {}
  ): string {
    let prompt = 'Ти - корисний AI асистент. Відповідай українською мовою.\n\n';

    // Додавання контексту
    if (context.messages.length > 0) {
      prompt += 'Контекст розмови:\n';
      for (const message of context.messages.slice(-5)) {
        prompt += `${message.role}: ${message.content}\n`;
      }
      prompt += '\n';
    }

    // Додатковий контекст
    if (Object.keys(additionalContext).length > 0) {
      prompt += 'Додатковий контекст:\n';
      for (const [key, value] of Object.entries(additionalContext)) {
        prompt += `${key}: ${value}\n`;
      }
      prompt += '\n';
    }

    prompt += `Користувач: ${userInput}\nАсистент:`;
    
    return prompt;
  }

  /**
   * Оновлення статистики
   */
  private updateStats(success: boolean, responseTime: number): void {
    this.stats.totalRequests++;
    this.stats.totalResponseTime += responseTime;
    
    if (success) {
      this.stats.successfulRequests++;
    } else {
      this.stats.failedRequests++;
    }

    this.stats.averageResponseTime = this.stats.totalResponseTime / this.stats.totalRequests;
  }

  /**
   * Запуск очищення пам'яті
   */
  private startMemoryCleanup(): void {
    this.memoryCleanupInterval = setInterval(() => {
      this.cleanupMemory();
    }, 300000); // Кожні 5 хвилин
  }

  /**
   * Очищення пам'яті
   */
  private cleanupMemory(): void {
    const now = Date.now();
    const maxAge = 3600000; // 1 година

    for (const [userId, context] of this.conversationMemory.entries()) {
      if (now - context.timestamp > maxAge) {
        this.conversationMemory.delete(userId);
      }
    }

    logger.debug(`🧹 Очищено ${this.conversationMemory.size} контекстів розмов`);
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    try {
      if (!this.providers[this.currentProvider]) {
        return {
          healthy: false,
          service: this.name,
          error: 'Активний провайдер не налаштовано',
        };
      }

      // Тестовий запит
      try {
        await this.generateResponse('Тест', { useCache: false });
      } catch (error) {
        return {
          healthy: false,
          service: this.name,
          error: `Тестовий запит невдалий: ${error}`,
        };
      }

      return {
        healthy: true,
        service: this.name,
        details: {
          activeProvider: this.currentProvider,
          availableProviders: Object.keys(this.providers),
          conversationContexts: this.conversationMemory.size,
          totalRequests: this.stats.totalRequests,
          successRate: this.stats.successfulRequests / this.stats.totalRequests,
        },
      };
    } catch (error) {
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
  protected async onShutdown(): Promise<void> {
    try {
      if (this.memoryCleanupInterval) {
        clearInterval(this.memoryCleanupInterval);
        this.memoryCleanupInterval = null;
      }

      this.conversationMemory.clear();
      this.providers = {};

      logger.info('✅ AI Service зупинено');
    } catch (error) {
      logger.error('❌ Помилка зупинки AI Service:', error);
      throw error;
    }
  }

  /**
   * Отримання статистики
   */
  protected onGetStats(): Partial<AIServiceStats> {
    return this.stats;
  }

  /**
   * Хешування промпту для кешування
   */
  private hashPrompt(prompt: string): string {
    const crypto = require('crypto');
    return crypto.createHash('md5').update(prompt).digest('hex');
  }
} 