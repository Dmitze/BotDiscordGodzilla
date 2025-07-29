/**
 * Розширений AI-модуль для Discord Bot
 * Включає природномовний інтерфейс, контекстну пам'ять та аналіз даних
 */

const logger = require('./logger');
const { sanitizeInput } = require('./security');

// Конфігурація AI
const AI_CONFIG = {
  OPENAI_MODEL: process.env.OPENAI_MODEL || 'gpt-3.5-turbo',
  OLLAMA_MODEL: process.env.OLLAMA_MODEL || 'llama2',
  MAX_TOKENS: parseInt(process.env.OPENAI_MAX_TOKENS) || 2000,
  TEMPERATURE: parseFloat(process.env.OPENAI_TEMPERATURE) || 0.7,
  MAX_CONTEXT_LENGTH: 4000,
  MEMORY_TTL: 3600, // 1 година
  REQUEST_TIMEOUT: 30000, // 30 секунд
};

// In-memory кеш для контекстної пам'яті (в продакшені використовуйте Redis)
const conversationMemory = new Map();

/**
 * Клас для роботи з AI
 */
class AIEnhanced {
  constructor() {
    this.providers = {
      openai: this.createOpenAIProvider(),
      ollama: this.createOllamaProvider(),
    };
    this.currentProvider = process.env.AI_PROVIDER || 'openai';
    this.stats = {
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      averageResponseTime: 0,
    };
  }

  /**
   * Створення OpenAI провайдера
   */
  createOpenAIProvider() {
    try {
      const { OpenAI } = require('openai');
      if (!process.env.OPENAI_API_KEY) {
        logger.warn('OpenAI API key not found');
        return null;
      }
      return new OpenAI({
        apiKey: process.env.OPENAI_API_KEY,
        maxRetries: 3,
        timeout: AI_CONFIG.REQUEST_TIMEOUT,
      });
    } catch (error) {
      logger.error('Failed to create OpenAI provider:', error);
      return null;
    }
  }

  /**
   * Створення Ollama провайдера
   */
  createOllamaProvider() {
    try {
      const ollamaUrl = process.env.OLLAMA_URL || 'http://localhost:11434';

      return {
        async generate(prompt, options = {}) {
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), AI_CONFIG.REQUEST_TIMEOUT);

          try {
            const response = await fetch(`${ollamaUrl}/api/generate`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                model: options.model || AI_CONFIG.OLLAMA_MODEL,
                prompt: sanitizeInput(prompt, 'ai_prompt'),
                stream: false,
                options: {
                  temperature: options.temperature || AI_CONFIG.TEMPERATURE,
                  num_predict: options.maxTokens || AI_CONFIG.MAX_TOKENS,
                },
              }),
              signal: controller.signal,
            });

            clearTimeout(timeoutId);

            if (!response.ok) {
              throw new Error(`Ollama API error: ${response.status} ${response.statusText}`);
            }

            const result = await response.json();
            return result.response || 'Порожня відповідь від Ollama';
          } catch (error) {
            clearTimeout(timeoutId);
            if (error.name === 'AbortError') {
              throw new Error('Ollama request timeout');
            }
            throw error;
          }
        },
      };
    } catch (error) {
      logger.error('Failed to create Ollama provider:', error);
      return null;
    }
  }

  /**
   * Отримання контексту розмови
   * @param {string} userId - ID користувача
   * @returns {Array} - Історія розмови
   */
  getConversationContext(userId) {
    try {
      const memory = conversationMemory.get(userId);
      if (!memory) return [];

      // Очищення застарілих повідомлень
      const now = Date.now();
      const validMessages = memory.messages.filter(
        msg => now - msg.timestamp < AI_CONFIG.MEMORY_TTL * 1000
      );

      // Оновлення пам'яті
      if (validMessages.length !== memory.messages.length) {
        memory.messages = validMessages;
        conversationMemory.set(userId, memory);
      }

      return validMessages.slice(-10); // Останні 10 повідомлень
    } catch (error) {
      logger.error('Error getting conversation context:', error);
      return [];
    }
  }

  /**
   * Збереження повідомлення в контекст
   * @param {string} userId - ID користувача
   * @param {string} role - Роль (user/assistant)
   * @param {string} content - Зміст повідомлення
   */
  saveToContext(userId, role, content) {
    try {
      if (!conversationMemory.has(userId)) {
        conversationMemory.set(userId, { messages: [] });
      }

      const memory = conversationMemory.get(userId);
      memory.messages.push({
        role,
        content: sanitizeInput(content, 'ai_prompt'),
        timestamp: Date.now(),
      });

      // Обмеження розміру пам'яті
      if (memory.messages.length > 20) {
        memory.messages = memory.messages.slice(-20);
      }
    } catch (error) {
      logger.error('Error saving to context:', error);
    }
  }

  /**
   * Аналіз природномовного запиту
   * @param {string} userInput - Користувацький запит
   * @returns {Object} - Результат аналізу
   */
  async analyzeNaturalLanguage(userInput) {
    try {
      const sanitizedInput = sanitizeInput(userInput, 'ai_prompt');

      if (!sanitizedInput) {
        throw new Error('Invalid input');
      }

      const analysisPrompt = `
        Проаналізуй наступний запит користувача та визнач:
        1. Тип запиту (пошук, аналіз, допомога, інше)
        2. Ключові слова для пошуку
        3. Поля для фільтрації
        4. Додаткові параметри
        
        Запит: "${sanitizedInput}"
        
        Відповідь у форматі JSON:
        {
          "type": "search|analysis|help|other",
          "keywords": ["слово1", "слово2"],
          "fields": ["поле1", "поле2"],
          "parameters": {"параметр1": "значення1"}
        }
      `;

      const response = await this.generateResponse(analysisPrompt, { maxTokens: 500 });

      try {
        return JSON.parse(response);
      } catch (parseError) {
        logger.warn('Failed to parse AI analysis response:', parseError);
        return {
          type: 'other',
          keywords: [sanitizedInput],
          fields: [],
          parameters: {},
        };
      }
    } catch (error) {
      logger.error('Natural language analysis error:', error);
      return {
        type: 'other',
        keywords: [userInput],
        fields: [],
        parameters: {},
      };
    }
  }

  /**
   * Генерація відповіді від AI
   * @param {string} prompt - Запит до AI
   * @param {Object} options - Додаткові опції
   * @returns {Promise<string>} - Відповідь AI
   */
  async generateResponse(prompt, options = {}) {
    const startTime = Date.now();
    this.stats.totalRequests++;

    try {
      const provider = this.providers[this.currentProvider];
      if (!provider) {
        throw new Error(`AI provider '${this.currentProvider}' not available`);
      }

      const sanitizedPrompt = sanitizeInput(prompt, 'ai_prompt');
      if (!sanitizedPrompt) {
        throw new Error('Invalid prompt');
      }

      const response = await provider.generate(sanitizedPrompt, {
        model: options.model || AI_CONFIG.OPENAI_MODEL,
        maxTokens: options.maxTokens || AI_CONFIG.MAX_TOKENS,
        temperature: options.temperature || AI_CONFIG.TEMPERATURE,
        ...options,
      });

      const responseTime = Date.now() - startTime;
      this.updateStats(true, responseTime);

      return response;
    } catch (error) {
      const responseTime = Date.now() - startTime;
      this.updateStats(false, responseTime);

      logger.error('AI response generation error:', error);

      // Спробувати інший провайдер якщо поточний не працює
      if (this.currentProvider === 'openai' && this.providers.ollama) {
        logger.info('Trying Ollama as fallback...');
        this.currentProvider = 'ollama';
        return this.generateResponse(prompt, options);
      }

      throw new Error(`AI service error: ${error.message}`);
    }
  }

  /**
   * Аналіз даних
   * @param {Array} data - Дані для аналізу
   * @param {string} analysisType - Тип аналізу
   * @returns {Promise<string>} - Результат аналізу
   */
  async analyzeData(data, analysisType = 'summary') {
    try {
      if (!Array.isArray(data) || data.length === 0) {
        throw new Error('No data provided for analysis');
      }

      const dataSummary = data
        .slice(0, 10)
        .map(item => (typeof item === 'string' ? item : JSON.stringify(item)))
        .join('\n');

      const analysisPrompt = `
        Проаналізуй наступні дані та надай ${analysisType}:
        
        ${dataSummary}
        
        ${data.length > 10 ? `\n... та ще ${data.length - 10} записів` : ''}
        
        Будь ласка, надай структуровану відповідь українською мовою.
      `;

      return await this.generateResponse(analysisPrompt, { maxTokens: 1000 });
    } catch (error) {
      logger.error('Data analysis error:', error);
      throw new Error(`Помилка аналізу даних: ${error.message}`);
    }
  }

  /**
   * Генерація звіту
   * @param {Array} data - Дані для звіту
   * @param {Object} options - Опції звіту
   * @returns {Promise<string>} - Згенерований звіт
   */
  async generateReport(data, options = {}) {
    try {
      if (!Array.isArray(data) || data.length === 0) {
        throw new Error('No data provided for report');
      }

      const reportPrompt = `
        Створи детальний звіт на основі наступних даних:
        
        ${JSON.stringify(data.slice(0, 20), null, 2)}
        
        ${data.length > 20 ? `\n... та ще ${data.length - 20} записів` : ''}
        
        Тип звіту: ${options.type || 'загальний'}
        Формат: ${options.format || 'текстовий'}
        
        Включи:
        - Загальну статистику
        - Ключові висновки
        - Рекомендації
      `;

      return await this.generateResponse(reportPrompt, { maxTokens: 1500 });
    } catch (error) {
      logger.error('Report generation error:', error);
      throw new Error(`Помилка генерації звіту: ${error.message}`);
    }
  }

  /**
   * Обробка природномовного запиту
   * @param {string} userId - ID користувача
   * @param {string} userInput - Користувацький запит
   * @param {Array} sheetData - Дані з таблиці (опціонально)
   * @returns {Promise<string>} - Відповідь AI
   */
  async processNaturalLanguageQuery(userId, userInput, sheetData = null) {
    try {
      const sanitizedInput = sanitizeInput(userInput, 'ai_prompt');
      if (!sanitizedInput) {
        throw new Error('Invalid input');
      }

      // Отримання контексту
      const context = this.getConversationContext(userId);

      // Аналіз запиту
      const analysis = await this.analyzeNaturalLanguage(sanitizedInput);

      // Формування промпту з контекстом
      let prompt = `Ти - AI асистент для роботи з Google Sheets. Користувач запитує: "${sanitizedInput}"\n\n`;

      if (context.length > 0) {
        prompt += 'Контекст попередньої розмови:\n';
        context.forEach(msg => {
          prompt += `${msg.role}: ${msg.content}\n`;
        });
        prompt += '\n';
      }

      if (sheetData && sheetData.length > 0) {
        prompt += `Дані з таблиці (перші 5 записів):\n${JSON.stringify(sheetData.slice(0, 5), null, 2)}\n\n`;
      }

      prompt += `Аналіз запиту: ${JSON.stringify(analysis)}\n\n`;
      prompt += 'Надай корисну та інформативну відповідь українською мовою.';

      // Генерація відповіді
      const response = await this.generateResponse(prompt, { maxTokens: 1000 });

      // Збереження в контекст
      this.saveToContext(userId, 'user', sanitizedInput);
      this.saveToContext(userId, 'assistant', response);

      return response;
    } catch (error) {
      logger.error('Natural language query processing error:', error);
      return `Вибачте, сталася помилка при обробці вашого запиту: ${error.message}`;
    }
  }

  /**
   * Отримання довідкового повідомлення
   * @returns {string} - Довідка
   */
  getHelpMessage() {
    return `
🤖 **AI Асистент - Довідка**

**Доступні функції:**
• **Природномовний пошук** - "знайди товари iPhone"
• **Аналіз даних** - "проаналізуй залишки"
• **Генерація звітів** - "створи звіт по продажах"
• **Контекстна пам'ять** - бот пам'ятає попередню розмову

**Приклади запитів:**
• "Покажи товари з ціною вище 1000"
• "Які товари найпопулярніші?"
• "Створи звіт по контрагентах"
• "Проаналізуй тренди продажів"

**Підтримувані провайдери:**
• OpenAI GPT (основний)
• Ollama (локальний, резервний)
    `;
  }

  /**
   * Очищення контексту користувача
   * @param {string} userId - ID користувача
   */
  clearContext(userId) {
    try {
      conversationMemory.delete(userId);
      logger.info(`Context cleared for user ${userId}`);
    } catch (error) {
      logger.error('Error clearing context:', error);
    }
  }

  /**
   * Отримання статистики
   * @returns {Object} - Статистика
   */
  getStats() {
    return {
      ...this.stats,
      currentProvider: this.currentProvider,
      availableProviders: Object.keys(this.providers).filter(key => this.providers[key] !== null),
      memorySize: conversationMemory.size,
    };
  }

  /**
   * Оновлення статистики
   * @param {boolean} success - Чи успішний запит
   * @param {number} responseTime - Час відповіді
   */
  updateStats(success, responseTime) {
    if (success) {
      this.stats.successfulRequests++;
    } else {
      this.stats.failedRequests++;
    }

    // Оновлення середнього часу відповіді
    const totalRequests = this.stats.successfulRequests + this.stats.failedRequests;
    this.stats.averageResponseTime =
      (this.stats.averageResponseTime * (totalRequests - 1) + responseTime) / totalRequests;
  }
}

// Експорт екземпляру класу
module.exports = {
  aiEnhanced: new AIEnhanced(),
};
