/**
 * AI Service для Discord бота
 * Централізоване управління AI функціоналом
 */

const logger = require('../utils/logger');
const { sanitizeInput } = require('../utils/security');

class AIService {
  constructor(bot) {
    this.bot = bot;
    this.config = bot.config.ai;
    this.providers = {};
    this.currentProvider = this.config.provider;
    this.conversationMemory = new Map();
    this.stats = {
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      averageResponseTime: 0,
      totalResponseTime: 0,
    };
    this.isActive = false;
  }

  /**
   * Ініціалізація AI сервісу
   */
  async initialize() {
    try {
      logger.info('🤖 Ініціалізація AI сервісу...');

      // Створення провайдерів
      await this.createProviders();

      // Валідація конфігурації
      this.validateConfiguration();

      // Запуск очищення пам'яті
      this.startMemoryCleanup();

      this.isActive = true;
      logger.info('✅ AI сервіс ініціалізовано');
    } catch (error) {
      logger.error('❌ Помилка ініціалізації AI сервісу:', error);
      throw error;
    }
  }

  /**
   * Створення AI провайдерів
   */
  async createProviders() {
    // OpenAI провайдер
    if (this.config.openai.apiKey) {
      this.providers.openai = this.createOpenAIProvider();
      logger.debug('✅ OpenAI провайдер створено');
    }

    // Ollama провайдер
    if (this.config.ollama.host) {
      this.providers.ollama = this.createOllamaProvider();
      logger.debug('✅ Ollama провайдер створено');
    }

    if (Object.keys(this.providers).length === 0) {
      throw new Error('Жоден AI провайдер не налаштовано');
    }
  }

  /**
   * Створення OpenAI провайдера
   */
  createOpenAIProvider() {
    try {
      const { OpenAI } = require('openai');
      return new OpenAI({
        apiKey: this.config.openai.apiKey,
        maxRetries: 3,
        timeout: 30000,
      });
    } catch (error) {
      logger.error('Помилка створення OpenAI провайдера:', error);
      return null;
    }
  }

  /**
   * Створення Ollama провайдера
   */
  createOllamaProvider() {
    return {
      async generate(prompt, options = {}) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000);

        try {
          const response = await fetch(`${this.config.ollama.host}/api/generate`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              model: options.model || this.config.ollama.model,
              prompt: sanitizeInput(prompt, 'ai_prompt'),
              stream: false,
              options: {
                temperature: options.temperature || 0.7,
                num_predict: options.maxTokens || 2000,
              },
            }),
            signal: controller.signal,
          });

          clearTimeout(timeoutId);

          if (!response.ok) {
            throw new Error(`Ollama API error: ${response.status}`);
          }

          const result = await response.json();
          return result.response || 'Порожня відповідь від Ollama';
        } catch (error) {
          clearTimeout(timeoutId);
          if (error.name === 'AbortError') {
            throw new Error('Таймаут запиту до Ollama');
          }
          throw error;
        }
      },
    };
  }

  /**
   * Валідація конфігурації
   */
  validateConfiguration() {
    if (!this.providers[this.currentProvider]) {
      logger.warn(
        `Поточний провайдер ${this.currentProvider} недоступний, використовуємо fallback`
      );
      this.currentProvider = Object.keys(this.providers)[0];
    }
  }

  /**
   * Генерація відповіді
   */
  async generateResponse(prompt, options = {}) {
    const startTime = Date.now();
    this.stats.totalRequests++;

    const {
      useCache = true,
      cacheTTL = 600000, // 10 хвилин
      forceRefresh = false,
      retryAttempts = 3,
      timeout = 30000,
      provider = this.currentProvider
    } = options;

    try {
      const sanitizedPrompt = sanitizeInput(prompt, 'ai_prompt');

      // Перевірка кешу
      if (useCache && !forceRefresh && this.bot.serviceContainer) {
        const cacheService = this.bot.serviceContainer.get('cache');
        if (cacheService) {
          const cacheKey = `ai:${provider}:${this.hashPrompt(sanitizedPrompt)}`;
          const cachedResponse = await cacheService.get(cacheKey);
          if (cachedResponse) {
            logger.info(`📋 AI відповідь отримано з кешу`);
            this.updateStats(true, Date.now() - startTime);
            return cachedResponse;
          }
        }
      }

      // Генерація відповіді з retry механізмом
      let response;
      let lastError;

      for (let attempt = 1; attempt <= retryAttempts; attempt++) {
        try {
          if (provider === 'openai' && this.providers.openai) {
            response = await this.generateOpenAIResponse(sanitizedPrompt, { ...options, timeout });
          } else if (provider === 'ollama' && this.providers.ollama) {
            response = await this.providers.ollama.generate(sanitizedPrompt, { ...options, timeout });
          } else {
            throw new Error(`Невідомий провайдер: ${provider}`);
          }

          // Кешування відповіді
          if (useCache && this.bot.serviceContainer) {
            const cacheService = this.bot.serviceContainer.get('cache');
            if (cacheService) {
              const cacheKey = `ai:${provider}:${this.hashPrompt(sanitizedPrompt)}`;
              await cacheService.set(cacheKey, response, cacheTTL);
            }
          }

          const responseTime = Date.now() - startTime;
          this.updateStats(true, responseTime);
          return response;

        } catch (error) {
          lastError = error;
          logger.warn(`⚠️ Спроба ${attempt}/${retryAttempts} невдала:`, error.message);
          
          if (attempt < retryAttempts) {
            // Експоненціальна затримка
            const delay = Math.min(1000 * Math.pow(2, attempt - 1), 10000);
            await new Promise(resolve => setTimeout(resolve, delay));
          }
        }
      }

      // Fallback до іншого провайдера
      if (provider !== 'ollama' && this.providers.ollama) {
        logger.warn('Fallback до Ollama провайдера');
        return this.generateResponse(prompt, { ...options, provider: 'ollama' });
      }

      throw lastError;
    } catch (error) {
      const responseTime = Date.now() - startTime;
      this.updateStats(false, responseTime);
      logger.error('❌ Помилка генерації AI відповіді:', error);
      throw error;
    }
  }

  /**
   * Генерація відповіді через OpenAI
   */
  async generateOpenAIResponse(prompt, options = {}) {
    const response = await this.providers.openai.chat.completions.create({
      model: options.model || this.config.openai.model,
      messages: [
        {
          role: 'system',
          content:
            'Ти - корисний AI асистент для Discord бота, що працює з військовими документами ЗСУ.',
        },
        {
          role: 'user',
          content: prompt,
        },
      ],
      max_tokens: options.maxTokens || this.config.openai.maxTokens,
      temperature: options.temperature || this.config.openai.temperature,
    });

    return response.choices[0]?.message?.content || 'Порожня відповідь';
  }

  /**
   * Аналіз даних
   */
  async analyzeData(data, analysisType = 'summary') {
    try {
      const prompt = this.buildAnalysisPrompt(data, analysisType);
      return await this.generateResponse(prompt);
    } catch (error) {
      logger.error('Помилка аналізу даних:', error);
      throw new Error('Не вдалося проаналізувати дані');
    }
  }

  /**
   * Створення звіту
   */
  async generateReport(data, options = {}) {
    try {
      const prompt = this.buildReportPrompt(data, options);
      return await this.generateResponse(prompt);
    } catch (error) {
      logger.error('Помилка генерації звіту:', error);
      throw new Error('Не вдалося створити звіт');
    }
  }

  /**
   * Обробка природномовного запиту
   */
  async processNaturalLanguageQuery(userId, userInput, context = {}) {
    try {
      const conversationContext = this.getConversationContext(userId);
      const prompt = this.buildConversationPrompt(userInput, conversationContext, context);

      const response = await this.generateResponse(prompt);

      // Збереження контексту
      this.saveToContext(userId, 'user', userInput);
      this.saveToContext(userId, 'assistant', response);

      return response;
    } catch (error) {
      logger.error('Помилка обробки запиту:', error);
      throw new Error('Не вдалося обробити запит');
    }
  }

  /**
   * Отримання контексту розмови
   */
  getConversationContext(userId) {
    const context = this.conversationMemory.get(userId);
    if (!context) return [];

    // Повертаємо останні 10 повідомлень
    return context.slice(-10);
  }

  /**
   * Збереження в контекст
   */
  saveToContext(userId, role, content) {
    if (!this.conversationMemory.has(userId)) {
      this.conversationMemory.set(userId, []);
    }

    const context = this.conversationMemory.get(userId);
    context.push({
      role,
      content: sanitizeInput(content, 'ai_prompt'),
      timestamp: new Date(),
    });

    // Обмежуємо розмір контексту
    if (context.length > 20) {
      context.splice(0, context.length - 20);
    }
  }

  /**
   * Очищення контексту
   */
  clearContext(userId) {
    this.conversationMemory.delete(userId);
  }

  /**
   * Створення промпту для аналізу
   */
  buildAnalysisPrompt(data, analysisType) {
    const prompts = {
      summary: `Проаналізуй наступні дані та створи короткий зміст:\n${JSON.stringify(data, null, 2)}`,
      detailed: `Проведи детальний аналіз наступних даних:\n${JSON.stringify(data, null, 2)}`,
      key_points: `Виділи ключові моменти з наступних даних:\n${JSON.stringify(data, null, 2)}`,
    };

    return prompts[analysisType] || prompts.summary;
  }

  /**
   * Створення промпту для звіту
   */
  buildReportPrompt(data, options) {
    return `Створи звіт на основі наступних даних:
Тип звіту: ${options.type || 'загальний'}
Формат: ${options.format || 'текстовий'}
Дані: ${JSON.stringify(data, null, 2)}

Звіт повинен бути структурованим та інформативним.`;
  }

  /**
   * Створення промпту для розмови
   */
  buildConversationPrompt(userInput, context, additionalContext) {
    let prompt =
      'Ти - корисний AI асистент для Discord бота, що працює з військовими документами ЗСУ.\n\n';

    if (context.length > 0) {
      prompt += 'Контекст попередньої розмови:\n';
      context.forEach(msg => {
        prompt += `${msg.role}: ${msg.content}\n`;
      });
      prompt += '\n';
    }

    if (additionalContext.sheetData) {
      prompt += `Дані з таблиці: ${JSON.stringify(additionalContext.sheetData)}\n\n`;
    }

    prompt += `Поточний запит користувача: ${userInput}`;
    return prompt;
  }

  /**
   * Оновлення статистики
   */
  updateStats(success, responseTime) {
    if (success) {
      this.stats.successfulRequests++;
    } else {
      this.stats.failedRequests++;
    }

    this.stats.totalResponseTime += responseTime;
    this.stats.averageResponseTime = this.stats.totalResponseTime / this.stats.totalRequests;
  }

  /**
   * Запуск очищення пам'яті
   */
  startMemoryCleanup() {
    setInterval(() => {
      this.cleanupMemory();
    }, 3600000); // Кожну годину
  }

  /**
   * Очищення пам'яті
   */
  cleanupMemory() {
    const now = Date.now();
    const maxAge = 3600000; // 1 година

    for (const [userId, context] of this.conversationMemory.entries()) {
      const filteredContext = context.filter(msg => now - msg.timestamp.getTime() < maxAge);

      if (filteredContext.length === 0) {
        this.conversationMemory.delete(userId);
      } else {
        this.conversationMemory.set(userId, filteredContext);
      }
    }

    logger.debug(`Очищено пам'ять, залишилось ${this.conversationMemory.size} користувачів`);
  }

  /**
   * Перевірка активності
   */
  isActive() {
    return this.isActive;
  }

  /**
   * Отримання статистики
   */
  getStats() {
    return {
      ...this.stats,
      activeProvider: this.currentProvider,
      availableProviders: Object.keys(this.providers),
      memorySize: this.conversationMemory.size,
      isActive: this.isActive,
    };
  }

  /**
   * Завершення роботи
   */
  async shutdown() {
    logger.info('🛑 Завершення роботи AI сервісу...');
    this.isActive = false;
    this.conversationMemory.clear();
    logger.info('✅ AI сервіс завершено');
  }

  /**
   * Хешування промпту для кешування
   */
  hashPrompt(prompt) {
    const crypto = require('crypto');
    return crypto.createHash('md5').update(prompt).digest('hex');
  }
}

module.exports = AIService;
