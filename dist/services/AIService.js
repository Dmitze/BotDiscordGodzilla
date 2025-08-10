"use strict";
/**
 * AI Service для Discord бота
 * Централізоване управління AI функціоналом
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AIService = void 0;
const openai_1 = __importDefault(require("openai"));
const BaseService_1 = require("@/core/BaseService");
const CacheService_1 = require("./CacheService");
const logger_1 = __importDefault(require("@/utils/logger"));
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
};
const sanitizeInput = (input) => {
    return input.trim().replace(/[<>]/g, '');
};
// (видалено невикористаний інтерфейс OllamaProvider)
class AIService extends BaseService_1.BaseService {
    constructor(config) {
        super('AIService', config);
        this.providers = {};
        this.conversationMemory = new Map();
        this.memoryCleanupInterval = null;
        this.healthCheckInterval = null;
        this.currentProvider = config.ai.provider;
        this.cacheService = new CacheService_1.CacheService(config);
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
    async onInitialize() {
        try {
            logger_1.default.info('🤖 Ініціалізація AI сервісу...');
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
            logger_1.default.info('✅ AI сервіс ініціалізовано');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка ініціалізації AI сервісу:', { error });
            throw error;
        }
    }
    /**
     * Створення AI провайдерів з детальним логуванням
     */
    async createProviders() {
        try {
            logger_1.default.info('🔧 Створення AI провайдерів...');
            // OpenAI провайдер
            if (this.config.ai['openai'].apiKey) {
                this.providers['openai'] = this.createOpenAIProvider();
                logger_1.default.debug('✅ OpenAI провайдер створено');
            }
            else {
                logger_1.default.warn('⚠️ OpenAI API ключ не налаштовано');
            }
            // Ollama провайдер
            if (this.config.ai['ollama'].host) {
                this.providers['ollama'] = this.createOllamaProvider();
                logger_1.default.debug('✅ Ollama провайдер створено');
            }
            else {
                logger_1.default.warn('⚠️ Ollama хост не налаштовано');
            }
            if (Object.keys(this.providers).length === 0) {
                throw new Error('Жоден AI провайдер не налаштовано');
            }
            logger_1.default.info(`✅ Створено ${Object.keys(this.providers).length} AI провайдерів`);
        }
        catch (error) {
            logger_1.default.error('❌ Помилка створення AI провайдерів:', { error });
            throw error;
        }
    }
    /**
     * Створення OpenAI провайдера з покращеною обробкою помилок
     */
    createOpenAIProvider() {
        try {
            const openaiCfg = this.config.ai.openai;
            const openai = new openai_1.default({
                apiKey: openaiCfg.apiKey,
                maxRetries: AI_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS,
                timeout: AI_SERVICE_CONSTANTS.REQUEST_TIMEOUT,
            });
            return {
                async generate(prompt, options = {}) {
                    const startTime = Date.now();
                    try {
                        logger_1.default.debug('🔄 OpenAI запит...', {
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
                        logger_1.default.debug('✅ OpenAI відповідь отримана', {
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
                    }
                    catch (error) {
                        const duration = Date.now() - startTime;
                        logger_1.default.error('❌ Помилка OpenAI запиту:', {
                            error: error instanceof Error ? error.message : String(error),
                            duration: `${duration}ms`,
                        });
                        throw new Error(`OpenAI error: ${error instanceof Error ? error.message : String(error)}`);
                    }
                },
                async isHealthy() {
                    try {
                        await openai.models.list();
                        return true;
                    }
                    catch (error) {
                        logger_1.default.error('❌ OpenAI health check невдалий:', { error });
                        return false;
                    }
                },
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка створення OpenAI провайдера:', { error });
            throw error;
        }
    }
    /**
     * Створення Ollama провайдера з покращеною обробкою помилок
     */
    createOllamaProvider() {
        const ollamaConfig = this.config.ai.ollama;
        return {
            async generate(prompt, options = {}) {
                const startTime = Date.now();
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), AI_SERVICE_CONSTANTS.REQUEST_TIMEOUT);
                try {
                    logger_1.default.debug('🔄 Ollama запит...', {
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
                    const data = await response.json();
                    const duration = Date.now() - startTime;
                    logger_1.default.debug('✅ Ollama відповідь отримана', {
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
                }
                catch (error) {
                    clearTimeout(timeoutId);
                    const duration = Date.now() - startTime;
                    logger_1.default.error('❌ Помилка Ollama запиту:', {
                        error: error instanceof Error ? error.message : String(error),
                        duration: `${duration}ms`,
                    });
                    throw new Error(`Ollama error: ${error instanceof Error ? error.message : String(error)}`);
                }
            },
            async isHealthy() {
                try {
                    const response = await fetch(`${ollamaConfig.host}/api/tags`);
                    return response.ok;
                }
                catch (error) {
                    logger_1.default.error('❌ Ollama health check невдалий:', { error });
                    return false;
                }
            },
        };
    }
    /**
     * Валідація конфігурації з детальним логуванням
     */
    validateConfiguration() {
        try {
            if (!this.providers[this.currentProvider]) {
                throw new Error(`Поточний провайдер ${this.currentProvider} не налаштовано`);
            }
            logger_1.default.info(`✅ AI конфігурація валідна, активний провайдер: ${this.currentProvider}`);
            logger_1.default.info(`📊 Доступні провайдери: ${Object.keys(this.providers).join(', ')}`);
        }
        catch (error) {
            logger_1.default.error('❌ Помилка валідації AI конфігурації:', { error });
            throw error;
        }
    }
    /**
     * Генерація відповіді з покращеною обробкою помилок
     */
    async generateResponse(prompt, options = {}) {
        const { useCache = true, cacheTTL = 3600, forceRefresh = false, retryAttempts = AI_SERVICE_CONSTANTS.MAX_RETRY_ATTEMPTS, provider = this.currentProvider, } = options;
        try {
            // Валідація промпту
            const sanitizedPrompt = this.validateAndSanitizePrompt(prompt);
            // Перевірка кешу
            if (useCache && !forceRefresh) {
                const cacheKey = this.buildCacheKey(sanitizedPrompt, options);
                try {
                    const cached = await this.cacheService.get(cacheKey);
                    if (cached) {
                        this.stats.cacheHits++;
                        logger_1.default.debug('✅ Використано кешовану відповідь', {
                            cacheKey: cacheKey.substring(0, 20) + '...',
                            provider: cached.provider,
                            tokens: cached.tokens,
                        });
                        return cached;
                    }
                    else {
                        this.stats.cacheMisses++;
                    }
                }
                catch (cacheError) {
                    logger_1.default.warn('⚠️ Помилка читання з кешу:', { error: cacheError });
                    this.stats.cacheMisses++;
                }
            }
            // Retry logic з fallback
            let lastError = null;
            let usedProvider = provider;
            for (let attempt = 0; attempt <= retryAttempts; attempt++) {
                try {
                    const startTime = Date.now();
                    // Спробувати основний провайдер
                    let response;
                    const primary = this.providers[usedProvider];
                    if (primary) {
                        response = await primary.generate(sanitizedPrompt, options);
                    }
                    else {
                        // Fallback до іншого провайдера
                        const fallbackProvider = Object.keys(this.providers).find(p => p !== usedProvider);
                        if (fallbackProvider) {
                            usedProvider = fallbackProvider;
                            this.stats.providerSwitches++;
                            logger_1.default.warn(`🔄 Переключення на провайдер ${usedProvider}`);
                            const fallbackImpl = this.providers[usedProvider];
                            if (!fallbackImpl) {
                                throw new Error('Немає доступних провайдерів');
                            }
                            response = await fallbackImpl.generate(sanitizedPrompt, options);
                        }
                        else {
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
                            logger_1.default.debug('💾 Відповідь збережена в кеш', {
                                cacheKey: cacheKey.substring(0, 20) + '...',
                                ttl: `${cacheTTL}s`,
                                provider: response.provider,
                            });
                        }
                        catch (cacheError) {
                            logger_1.default.warn('⚠️ Помилка збереження в кеш:', { error: cacheError });
                        }
                    }
                    logger_1.default.info(`✅ AI відповідь згенерована за ${duration}ms`, {
                        provider: usedProvider,
                        tokens: response.tokens,
                        duration: `${duration}ms`,
                    });
                    return response;
                }
                catch (error) {
                    lastError = error;
                    if (attempt < retryAttempts) {
                        const delay = AI_SERVICE_CONSTANTS.RETRY_DELAY * Math.pow(2, attempt);
                        logger_1.default.warn(`🔄 Спроба ${attempt + 1} невдала, повтор через ${delay}ms`, {
                            error: lastError.message,
                            provider: usedProvider,
                        });
                        await new Promise(resolve => setTimeout(resolve, delay));
                    }
                }
            }
            throw lastError || new Error('Всі спроби генерації невдалі');
        }
        catch (error) {
            this.updateStats(false, 0);
            logger_1.default.error('❌ Помилка генерації відповіді:', {
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
    validateAndSanitizePrompt(prompt) {
        const sanitized = sanitizeInput(prompt);
        if (!sanitized || sanitized.length < AI_SERVICE_CONSTANTS.MIN_PROMPT_LENGTH) {
            throw new Error('Порожній або занадто короткий промпт');
        }
        if (sanitized.length > AI_SERVICE_CONSTANTS.MAX_PROMPT_LENGTH) {
            logger_1.default.warn('⚠️ Промпт занадто довгий, обрізаю...');
            return sanitized.substring(0, AI_SERVICE_CONSTANTS.MAX_PROMPT_LENGTH);
        }
        return sanitized;
    }
    /**
     * Аналіз даних з детальним логуванням
     */
    async analyzeData(data, analysisType = 'summary') {
        try {
            logger_1.default.info(`📊 Аналіз даних: ${analysisType}`, {
                dataLength: data.length,
                analysisType,
            });
            const prompt = this.buildAnalysisPrompt(data, analysisType);
            const response = await this.generateResponse(prompt, { provider: this.currentProvider });
            logger_1.default.info(`✅ Аналіз даних завершено`, {
                analysisType,
                responseLength: response.content.length,
                duration: `${response.duration}ms`,
            });
            return response;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка аналізу даних:', { error });
            throw error;
        }
    }
    /**
     * Генерація звіту з детальним логуванням
     */
    async generateReport(data, options = {}) {
        try {
            logger_1.default.info('📋 Генерація звіту', {
                dataLength: data.length,
                format: options.format || 'text',
                length: options.length || 'medium',
            });
            const prompt = this.buildReportPrompt(data, options);
            const response = await this.generateResponse(prompt, { provider: this.currentProvider });
            logger_1.default.info('✅ Звіт згенеровано', {
                responseLength: response.content.length,
                duration: `${response.duration}ms`,
            });
            return response;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка генерації звіту:', { error });
            throw error;
        }
    }
    /**
     * Обробка природномовного запиту з детальним логуванням
     */
    async processNaturalLanguageQuery(userId, userInput, context = {}) {
        try {
            logger_1.default.info('💬 Обробка природномовного запиту', {
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
            logger_1.default.info('✅ Природномовний запит оброблено', {
                userId,
                responseLength: response.content.length,
                duration: `${response.duration}ms`,
            });
            return response;
        }
        catch (error) {
            logger_1.default.error('❌ Помилка обробки природномовного запиту:', { error });
            throw error;
        }
    }
    /**
     * Отримання контексту розмови
     */
    getConversationContext(userId) {
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
    saveToContext(userId, role, content) {
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
                logger_1.default.warn('⚠️ Спроба зберегти порожній контент в контекст');
                return;
            }
            context.messages.push({ role, content: sanitizedContent });
            context.timestamp = Date.now();
            context.requestCount++;
            // Обмеження розміру контексту
            if (context.messages.length > AI_SERVICE_CONSTANTS.MAX_CONTEXT_MESSAGES) {
                context.messages = context.messages.slice(-10);
                logger_1.default.debug('🧹 Контекст обрізано до останніх 10 повідомлень');
            }
            this.conversationMemory.set(userId, context);
            logger_1.default.debug('💾 Контекст збережено', {
                userId,
                role,
                messageCount: context.messages.length,
            });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка збереження контексту:', { error });
        }
    }
    /**
     * Очищення контексту
     */
    clearContext(userId) {
        try {
            this.conversationMemory.delete(userId);
            logger_1.default.info('🧹 Контекст очищено', { userId });
        }
        catch (error) {
            logger_1.default.error('❌ Помилка очищення контексту:', { error });
        }
    }
    /**
     * Створення промпту для аналізу
     */
    buildAnalysisPrompt(data, analysisType) {
        const prompts = {
            summary: `Надай короткий зміст наступного тексту:\n\n${data}`,
            sentiment: `Проаналізуй емоційний тон наступного тексту:\n\n${data}`,
            keywords: `Виділи ключові слова з наступного тексту:\n\n${data}`,
        };
        return prompts[analysisType] || prompts.summary;
    }
    /**
     * Створення промпту для звіту
     */
    buildReportPrompt(data, options) {
        const format = options.format || 'text';
        const length = options.length || 'medium';
        return `Створи ${length} звіт у форматі ${format} на основі наступних даних:\n\n${data}`;
    }
    /**
     * Створення промпту для розмови
     */
    buildConversationPrompt(userInput, context, additionalContext = {}) {
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
    updateStats(success, responseTime) {
        this.stats.totalRequests++;
        this.stats.totalResponseTime += responseTime;
        if (success) {
            this.stats.successfulRequests++;
        }
        else {
            this.stats.failedRequests++;
        }
        this.stats.averageResponseTime = this.stats.totalResponseTime / this.stats.totalRequests;
    }
    /**
     * Створення ключа кешу
     */
    buildCacheKey(prompt, options) {
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
    startMemoryCleanup() {
        this.memoryCleanupInterval = setInterval(() => {
            this.cleanupMemory();
        }, AI_SERVICE_CONSTANTS.MEMORY_CLEANUP_INTERVAL);
        logger_1.default.info("🧹 Запущено очищення пам'яті AI сервісу");
    }
    /**
     * Запуск health check
     */
    startHealthCheck() {
        this.healthCheckInterval = setInterval(async () => {
            try {
                const health = await this.onHealthCheck();
                if (!health.healthy) {
                    logger_1.default.warn('⚠️ AI сервіс health check виявив проблеми:', health);
                }
            }
            catch (error) {
                logger_1.default.error('❌ Помилка AI сервіс health check:', { error });
            }
        }, 60000); // Кожну хвилину
        logger_1.default.info('🏥 Запущено health check AI сервісу');
    }
    /**
     * Очищення пам'яті з детальним логуванням
     */
    cleanupMemory() {
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
                logger_1.default.info(`🧹 Очищено ${cleanedCount} застарілих контекстів розмов`);
            }
        }
        catch (error) {
            logger_1.default.error("❌ Помилка очищення пам'яті AI сервісу: ", { error });
        }
    }
    /**
     * Health check з детальним логуванням
     */
    async onHealthCheck() {
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
            }
            catch (error) {
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
                    successRate: this.stats.totalRequests > 0
                        ? (this.stats.successfulRequests / this.stats.totalRequests) * 100
                        : 0,
                    averageResponseTime: this.stats.averageResponseTime,
                },
            };
        }
        catch (error) {
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
    async onShutdown() {
        try {
            logger_1.default.info('🛑 Завершення роботи AI сервісу...');
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
            logger_1.default.info('✅ AI Service зупинено');
        }
        catch (error) {
            logger_1.default.error('❌ Помилка зупинки AI Service:', { error });
            throw error;
        }
    }
    /**
     * Отримання статистики з детальним логуванням
     */
    onGetStats() {
        return {
            ...this.stats,
            successRate: this.stats.totalRequests > 0
                ? (this.stats.successfulRequests / this.stats.totalRequests) * 100
                : 0,
        };
    }
}
exports.AIService = AIService;
//# sourceMappingURL=AIService.js.map