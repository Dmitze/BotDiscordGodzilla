/**
 * AI Service для Discord бота
 * Централізоване управління AI функціоналом
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import OpenAI from 'openai';
import type { BotConfig, HealthStatus, ServiceStats, AIResponse, AIRequestOptions } from '@/types';

import { BaseService as BaseServiceClass } from '@/core/BaseService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';

// Константи для AI сервісу
const AI_SERVICE_CONSTANTS = {
  MAX_RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000, // 1 секунда
  REQUEST_TIMEOUT: 30000, // 30 секунд
  MEMORY_CLEANUP_INTERVAL: 300000, // 5 хвилин
  MAX_CONTEXT_AGE: 3600000, // 1 година
  MAX_CONTEXT_MESSAGES: 20,
  MAX_PROMPT_LENGTH: 4000,
  MIN_PROMPT_LENGTH: 1,
} as const;

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
  providerSwitches: number;
  contextCleanups: number;
}

interface ConversationContext {
  messages: Array<{ role: 'user' | 'assistant' | 'system'; content: string }>;
  timestamp: number;
  requestCount: number;
}

interface AIProvider {
  generate(prompt: string, options?: AIRequestOptions): Promise<AIResponse>;
  isHealthy(): Promise<boolean>;
}

// (видалено невикористаний інтерфейс OllamaProvider)

export class AIService extends BaseServiceClass {
  private providers: Record<string, AIProvider> = {};
  private currentProvider: string;
  private conversationMemory = new Map<string, ConversationContext>();
  private stats: AIServiceStats;
  private memoryCleanupInterval: NodeJS.Timeout | null = null;
  private healthCheckInterval: NodeJS.Timeout | null = null;
  private cacheService: CacheService;

  constructor(config: BotConfig) {
    super('AIService', config);
    this.currentProvider = config.ai.provider;
    this.cacheService = new CacheService(config);
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
      providerSwitches: 0,
      contextCleanups: 0,
    };
  }

  /**
   * Ініціалізація AI сервісу з детальним логуванням
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('🤖 Ініціалізація AI сервісу...');

      // Ініціалізація кешу
      await this.cacheService.initialize();

      // Створення провайдерів
      await this.createProviders();

      // Валідація конфігурації
      this.validateConfiguration();

      // Запуск очищення пам'яті
      this.startMemoryCleanup();

      // Запуск health check
      this.startHealthCheck();

      logger.info('✅ AI сервіс ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації AI сервісу:', { error });
      throw error;
    }
  }

  /**
   * Створення AI провайдерів з детальним логуванням
   */
  private async createProviders(): Promise<void> {
    try {
      logger.info('🔧 Створення AI провайдерів...');

      // OpenAI провайдер
      if (this.config.ai['openai'].apiKey) {
        this.providers['openai'] = this.createOpenAIProvider();
        logger.debug('✅ OpenAI провайдер створено');
      } else {
        logger.warn('⚠️ OpenAI API ключ не налаштовано');
      }

      // Ollama провайдер
      if (this.config.ai['ollama'].host) {
        this.providers['ollama'] = this.createOllamaProvider();
        logger.debug('✅ Ollama провайдер створено');
      } else {
        logger.warn('⚠️ Ollama хост не налаштовано');
      }

      if (Object.keys(this.providers).length === 0) {
        throw new Error('Жоден AI провайдер не налаштовано');
      }

      logger.info(`✅ Створено ${Object.keys(this.providers).length} AI провайдерів`);
    } catch (error) {
      logger.error('❌ Помилка створення AI провайдерів:', { error });
      throw error;
    }
  }

  /**
   * Створення OpenAI провайдера з покращеною обробкою помилок
   */
  private createOpenAIProvider(): AIProvider {
    try {
      const openaiCfg = this.config.ai.openai;
      const openai = new OpenAI({
        apiKey: openaiCfg.apiKey,
        maxRetries: AI_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS,
        timeout: AI_SERVICE_CONSTANTS.REQUEST_TIMEOUT,
      });

      return {
        async generate(prompt: string, options: AIRequestOptions = {}): Promise<AIResponse> {
          const startTime = Date.now();

          try {
            logger.debug('🔄 OpenAI запит...', {
              model: options.model || openaiCfg.model,
              maxTokens: options.maxTokens || openaiCfg.maxTokens,
              temperature: options.temperature || openaiCfg.temperature,
            });

            const response = await openai.chat.completions.create({
              model: options.model || openaiCfg.model,
              messages: [{ role: 'user', content: prompt }],
              max_tokens: options.maxTokens || openaiCfg.maxTokens,
              temperature: options.temperature || openaiCfg.temperature,
            });

            const duration = Date.now() - startTime;

            logger.debug('✅ OpenAI відповідь отримана', {
              duration: `${duration}ms`,
              model: response.model,
              tokens: response.usage?.total_tokens || 0,
            });

            return {
              content: response.choices[0]?.message?.content || '',
              provider: 'openai',
              model: response.model,
              tokens: response.usage?.total_tokens || 0,
              duration,
            };
          } catch (error) {
            const duration = Date.now() - startTime;
            logger.error('❌ Помилка OpenAI запиту:', {
              error: error instanceof Error ? error.message : String(error),
              duration: `${duration}ms`,
            });
            throw new Error(
              `OpenAI error: ${error instanceof Error ? error.message : String(error)}`
            );
          }
        },
        async isHealthy(): Promise<boolean> {
          try {
            await openai.models.list();
            return true;
          } catch (error) {
            logger.error('❌ OpenAI health check невдалий:', { error });
            return false;
          }
        },
      };
    } catch (error) {
      logger.error('❌ Помилка створення OpenAI провайдера:', { error });
      throw error;
    }
  }

  /**
   * Створення Ollama провайдера з покращеною обробкою помилок
   */
  private createOllamaProvider(): AIProvider {
    const ollamaConfig = this.config.ai.ollama;

    return {
      async generate(prompt: string, options: AIRequestOptions = {}): Promise<AIResponse> {
        const startTime = Date.now();
        const controller = new AbortController();
        const timeoutId = setTimeout(
          () => controller.abort(),
          AI_SERVICE_CONSTANTS.REQUEST_TIMEOUT
        );

        try {
          logger.debug('🔄 Ollama запит...', {
            host: ollamaConfig.host,
            model: options.model || ollamaConfig.model,
            temperature: options.temperature || 0.7,
          });

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

          const data: any = await response.json();

          const duration = Date.now() - startTime;

          logger.debug('✅ Ollama відповідь отримана', {
            duration: `${duration}ms`,
            model: data.model || ollamaConfig.model,
          });

          return {
            content: data.response || '',
            provider: 'ollama',
            model: data.model || ollamaConfig.model,
            tokens: data.eval_count || 0,
            duration,
          };
        } catch (error) {
          clearTimeout(timeoutId);
          const duration = Date.now() - startTime;
          logger.error('❌ Помилка Ollama запиту:', {
            error: error instanceof Error ? error.message : String(error),
            duration: `${duration}ms`,
          });
          throw new Error(
            `Ollama error: ${error instanceof Error ? error.message : String(error)}`
          );
        }
      },
      async isHealthy(): Promise<boolean> {
        try {
          const response = await fetch(`${ollamaConfig.host}/api/tags`);
          return response.ok;
        } catch (error) {
          logger.error('❌ Ollama health check невдалий:', { error });
          return false;
        }
      },
    };
  }

  /**
   * Валідація конфігурації з детальним логуванням
   */
  private validateConfiguration(): void {
    try {
      if (!this.providers[this.currentProvider]) {
        throw new Error(`Поточний провайдер ${this.currentProvider} не налаштовано`);
      }

      logger.info(`✅ AI конфігурація валідна, активний провайдер: ${this.currentProvider}`);
      logger.info(`📊 Доступні провайдери: ${Object.keys(this.providers).join(', ')}`);
    } catch (error) {
      logger.error('❌ Помилка валідації AI конфігурації:', { error });
      throw error;
    }
  }

  /**
   * Генерація відповіді з покращеною обробкою помилок
   */
  public async generateResponse(
    prompt: string,
    options: AIRequestOptions = {}
  ): Promise<AIResponse> {
    const {
      useCache = true,
      cacheTTL = 3600,
      forceRefresh = false,
      retryAttempts = AI_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS,
      provider = this.currentProvider,
    } = options;

    try {
      // Валідація промпту
      const sanitizedPrompt = this.validateAndSanitizePrompt(prompt);

      // Перевірка кешу
      if (useCache && !forceRefresh) {
        const cacheKey = this.buildCacheKey(sanitizedPrompt, options);
        try {
          const cached = await this.cacheService.get<AIResponse>(cacheKey);
          if (cached) {
            this.stats.cacheHits++;
            logger.debug('✅ Використано кешовану відповідь', {
              cacheKey: cacheKey.substring(0, 20) + '...',
              provider: cached.provider,
              tokens: cached.tokens,
            });
            return cached;
          } else {
            this.stats.cacheMisses++;
          }
        } catch (cacheError) {
          logger.warn('⚠️ Помилка читання з кешу:', { error: cacheError });
          this.stats.cacheMisses++;
        }
      }

      // Retry logic з fallback
      let lastError: Error | null = null;
      let usedProvider: string = provider;

      for (let attempt = 0; attempt <= retryAttempts; attempt++) {
        try {
          const startTime = Date.now();

          // Спробувати основний провайдер
          let response: AIResponse;
          const primary = this.providers[usedProvider];
          if (primary) {
            response = await primary.generate(sanitizedPrompt, options);
          } else {
            // Fallback до іншого провайдера
            const fallbackProvider = Object.keys(this.providers).find(p => p !== usedProvider);
            if (fallbackProvider) {
              usedProvider = fallbackProvider;
              this.stats.providerSwitches++;
              logger.warn(`🔄 Переключення на провайдер ${usedProvider}`);
              const fallbackImpl = this.providers[usedProvider];
              if (!fallbackImpl) {
                throw new Error('Немає доступних провайдерів');
              }
              response = await fallbackImpl.generate(sanitizedPrompt, options);
            } else {
              throw new Error('Немає доступних провайдерів');
            }
          }

          const duration = Date.now() - startTime;

          this.updateStats(true, duration);

          // Збереження в кеш
          if (useCache) {
            const cacheKey = this.buildCacheKey(sanitizedPrompt, options);
            try {
              await this.cacheService.set(cacheKey, response, cacheTTL);
              logger.debug('💾 Відповідь збережена в кеш', {
                cacheKey: cacheKey.substring(0, 20) + '...',
                ttl: `${cacheTTL}s`,
                provider: response.provider,
              });
            } catch (cacheError) {
              logger.warn('⚠️ Помилка збереження в кеш:', { error: cacheError });
            }
          }

          logger.info(`✅ AI відповідь згенерована за ${duration}ms`, {
            provider: usedProvider,
            tokens: response.tokens,
            duration: `${duration}ms`,
          });

          return response;
        } catch (error) {
          lastError = error as Error;

          if (attempt < retryAttempts) {
            const delay = AI_SERVICE_CONSTANTS.RETRY_DELAY * Math.pow(2, attempt);
            logger.warn(`🔄 Спроба ${attempt + 1} невдала, повтор через ${delay}ms`, {
              error: lastError.message,
              provider: usedProvider,
            });
            await new Promise(resolve => setTimeout(resolve, delay));
          }
        }
      }

      throw lastError || new Error('Всі спроби генерації невдалі');
    } catch (error) {
      this.updateStats(false, 0);
      logger.error('❌ Помилка генерації відповіді:', {
        error: error instanceof Error ? error.message : String(error),
        prompt: prompt.substring(0, 100) + '...',
        provider: provider,
      });
      throw error;
    }
  }

  /**
   * Валідація та санітизація промпту
   */
  private validateAndSanitizePrompt(prompt: string): string {
    const sanitized = sanitizeInput(prompt);

    if (!sanitized || sanitized.length < AI_SERVICE_CONSTANTS.MIN_PROMPT_LENGTH) {
      throw new Error('Порожній або занадто короткий промпт');
    }

    if (sanitized.length > AI_SERVICE_CONSTANTS.MAX_PROMPT_LENGTH) {
      logger.warn('⚠️ Промпт занадто довгий, обрізаю...');
      return sanitized.substring(0, AI_SERVICE_CONSTANTS.MAX_PROMPT_LENGTH);
    }

    return sanitized;
  }

  /**
   * Аналіз даних з детальним логуванням
   */
  public async analyzeData(
    data: string,
    analysisType: 'summary' | 'sentiment' | 'keywords' = 'summary'
  ): Promise<AIResponse> {
    try {
      logger.info(`📊 Аналіз даних: ${analysisType}`, {
        dataLength: data.length,
        analysisType,
      });

      const prompt = this.buildAnalysisPrompt(data, analysisType);
      const response = await this.generateResponse(prompt, { provider: this.currentProvider });

      logger.info(`✅ Аналіз даних завершено`, {
        analysisType,
        responseLength: response.content.length,
        duration: `${response.duration}ms`,
      });

      return response;
    } catch (error) {
      logger.error('❌ Помилка аналізу даних:', { error });
      throw error;
    }
  }

  /**
   * Генерація звіту з детальним логуванням
   */
  public async generateReport(
    data: string,
    options: { format?: string; length?: string } = {}
  ): Promise<AIResponse> {
    try {
      logger.info('📋 Генерація звіту', {
        dataLength: data.length,
        format: options.format || 'text',
        length: options.length || 'medium',
      });

      const prompt = this.buildReportPrompt(data, options);
      const response = await this.generateResponse(prompt, { provider: this.currentProvider });

      logger.info('✅ Звіт згенеровано', {
        responseLength: response.content.length,
        duration: `${response.duration}ms`,
      });

      return response;
    } catch (error) {
      logger.error('❌ Помилка генерації звіту:', { error });
      throw error;
    }
  }

  /**
   * Обробка природномовного запиту з детальним логуванням
   */
  public async processNaturalLanguageQuery(
    userId: string,
    userInput: string,
    context: Record<string, unknown> = {}
  ): Promise<AIResponse> {
    try {
      logger.info('💬 Обробка природномовного запиту', {
        userId,
        inputLength: userInput.length,
        contextKeys: Object.keys(context),
      });

      const conversationContext = this.getConversationContext(userId);
      const prompt = this.buildConversationPrompt(userInput, conversationContext, context);

      const response = await this.generateResponse(prompt, { provider: this.currentProvider });

      // Збереження в контекст
      this.saveToContext(userId, 'user', userInput);
      this.saveToContext(userId, 'assistant', response.content);

      logger.info('✅ Природномовний запит оброблено', {
        userId,
        responseLength: response.content.length,
        duration: `${response.duration}ms`,
      });

      return response;
    } catch (error) {
      logger.error('❌ Помилка обробки природномовного запиту:', { error });
      throw error;
    }
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
   * Збереження в контекст з валідацією
   */
  public saveToContext(
    userId: string,
    role: 'user' | 'assistant' | 'system',
    content: string
  ): void {
    try {
      let context = this.conversationMemory.get(userId);

      if (!context) {
        context = {
          messages: [],
          timestamp: Date.now(),
          requestCount: 0,
        };
      }

      // Валідація контенту
      const sanitizedContent = sanitizeInput(content);
      if (!sanitizedContent) {
        logger.warn('⚠️ Спроба зберегти порожній контент в контекст');
        return;
      }

      context.messages.push({ role, content: sanitizedContent });
      context.timestamp = Date.now();
      context.requestCount++;

      // Обмеження розміру контексту
      if (context.messages.length > AI_SERVICE_CONSTANTS.MAX_CONTEXT_MESSAGES) {
        context.messages = context.messages.slice(-10);
        logger.debug('🧹 Контекст обрізано до останніх 10 повідомлень');
      }

      this.conversationMemory.set(userId, context);

      logger.debug('💾 Контекст збережено', {
        userId,
        role,
        messageCount: context.messages.length,
      });
    } catch (error) {
      logger.error('❌ Помилка збереження контексту:', { error });
    }
  }

  /**
   * Очищення контексту
   */
  public clearContext(userId: string): void {
    try {
      this.conversationMemory.delete(userId);
      logger.info('🧹 Контекст очищено', { userId });
    } catch (error) {
      logger.error('❌ Помилка очищення контексту:', { error });
    }
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
   * Створення ключа кешу
   */
  private buildCacheKey(prompt: string, options: AIRequestOptions): string {
    const keyData = {
      prompt: prompt.substring(0, 500),
      provider: options.provider || this.currentProvider,
      model: options.model,
      temperature: options.temperature,
      maxTokens: options.maxTokens,
    };
    // Використання Node.js crypto для стабільного ключа
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const crypto = require('crypto');
    const keyString = JSON.stringify(keyData);
    return `ai:${crypto.createHash('sha256').update(keyString).digest('hex').substring(0, 32)}`;
  }

  /**
   * Запуск очищення пам'яті
   */
  private startMemoryCleanup(): void {
    this.memoryCleanupInterval = setInterval(() => {
      this.cleanupMemory();
    }, AI_SERVICE_CONSTANTS.MEMORY_CLEANUP_INTERVAL);

    logger.info("🧹 Запущено очищення пам'яті AI сервісу");
  }

  /**
   * Запуск health check
   */
  private startHealthCheck(): void {
    this.healthCheckInterval = setInterval(async () => {
      try {
        const health = await this.onHealthCheck();
        if (!health.healthy) {
          logger.warn('⚠️ AI сервіс health check виявив проблеми:', health);
        }
      } catch (error) {
        logger.error('❌ Помилка AI сервіс health check:', { error });
      }
    }, 60000); // Кожну хвилину

    logger.info('🏥 Запущено health check AI сервісу');
  }

  /**
   * Очищення пам'яті з детальним логуванням
   */
  private cleanupMemory(): void {
    try {
      const now = Date.now();
      let cleanedCount = 0;

      for (const [userId, context] of this.conversationMemory.entries()) {
        if (now - context.timestamp > AI_SERVICE_CONSTANTS.MAX_CONTEXT_AGE) {
          this.conversationMemory.delete(userId);
          cleanedCount++;
        }
      }

      if (cleanedCount > 0) {
        this.stats.contextCleanups++;
        logger.info(`🧹 Очищено ${cleanedCount} застарілих контекстів розмов`);
      }
    } catch (error) {
      logger.error("❌ Помилка очищення пам'яті AI сервісу: ", { error });
    }
  }

  /**
   * Health check з детальним логуванням
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

      // Перевірка здоров'я активного провайдера
      const active = this.providers[this.currentProvider];
      if (!active) {
        return {
          healthy: false,
          service: this.name,
          error: 'Активний провайдер не налаштовано',
        };
      }
      const isHealthy = await active.isHealthy();
      if (!isHealthy) {
        return {
          healthy: false,
          service: this.name,
          error: 'Активний провайдер нездоровий',
        };
      }

      // Тестовий запит
      try {
        await this.generateResponse('Тест', { useCache: false });
      } catch (error) {
        return {
          healthy: false,
          service: this.name,
          error: `Тестовий запит невдалий: ${error instanceof Error ? error.message : String(error)}`,
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
          successRate:
            this.stats.totalRequests > 0
              ? (this.stats.successfulRequests / this.stats.totalRequests) * 100
              : 0,
          averageResponseTime: this.stats.averageResponseTime,
        },
      };
    } catch (error) {
      return {
        healthy: false,
        service: this.name,
        error: `Health check failed: ${error instanceof Error ? error.message : String(error)}`,
      };
    }
  }

  /**
   * Завершення роботи з детальним логуванням
   */
  protected async onShutdown(): Promise<void> {
    try {
      logger.info('🛑 Завершення роботи AI сервісу...');

      // Зупинка інтервалів
      if (this.memoryCleanupInterval) {
        clearInterval(this.memoryCleanupInterval);
        this.memoryCleanupInterval = null;
      }

      if (this.healthCheckInterval) {
        clearInterval(this.healthCheckInterval);
        this.healthCheckInterval = null;
      }

      // Зупинка кеш сервісу
      await this.cacheService.shutdown();

      // Очищення пам'яті
      this.conversationMemory.clear();
      this.providers = {};

      logger.info('✅ AI Service зупинено');
    } catch (error) {
      logger.error('❌ Помилка зупинки AI Service:', { error });
      throw error;
    }
  }

  /**
   * Отримання статистики з детальним логуванням
   */
  protected onGetStats(): Partial<AIServiceStats> {
    return {
      ...this.stats,
      successRate:
        this.stats.totalRequests > 0
          ? (this.stats.successfulRequests / this.stats.totalRequests) * 100
          : 0,
    };
  }
}
