import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import type { AIService } from '@/services/AIService';
import logger from '@/utils/logger';

export interface LanguageDetectionResult {
  language: string;
  confidence: number;
  supported: boolean;
}

export interface TranslationResult {
  originalText: string;
  translatedText: string;
  sourceLanguage: string;
  targetLanguage: string;
  confidence: number;
}

// New interface for cross-language search
export interface CrossLanguageSearchResult {
  fileId: string;
  fileName: string;
  snippet: string;
  sourceLanguage: string;
  targetLanguage: string;
  relevanceScore: number;
}

// New interface for interface localization
export interface LocalizationConfig {
  userLanguage: string;
  preferredLanguages: string[];
  autoDetect: boolean;
}

export class MultilingualDocumentProcessor extends BaseService {
  private ai: AIService | null = null;
  private supportedLanguages: string[] = ['uk', 'en', 'ru', 'pl', 'de', 'fr', 'es'];
  private languageNames: Record<string, string> = {
    'uk': 'Українська',
    'en': 'English',
    'ru': 'Русский',
    'pl': 'Polski',
    'de': 'Deutsch',
    'fr': 'Français',
    'es': 'Español'
  };
  
  // New properties for enhanced functionality
  private localizationConfigs: Map<string, LocalizationConfig> = new Map();
  private translationCache: Map<string, TranslationResult> = new Map();
  private readonly CACHE_TTL = 30 * 60 * 1000; // 30 minutes

  constructor(config: BotConfig) {
    super('MultilingualDocumentProcessor', config);
  }

  /**
   * Ініціалізує сервіс з необхідними залежностями
   */
  initializeServices(ai: AIService): void {
    this.ai = ai;
  }

  /**
   * Визначає мову документа
   */
  async detectLanguage(text: string): Promise<LanguageDetectionResult> {
    try {
      if (!this.ai) {
        throw new Error('AIService не ініціалізовано');
      }

      if (!text || text.trim().length === 0) {
        return {
          language: 'unknown',
          confidence: 0,
          supported: false
        };
      }

      // Для коротких текстів використовуємо простий підхід
      if (text.length < 100) {
        return this.detectLanguageSimple(text);
      }

      // Для довгих текстів використовуємо AI
      const prompt = `
Визнач мову наступного тексту та надай результат у форматі JSON:

Текст: "${text.substring(0, 500)}"

Відповідай тільки JSON у форматі:
{
  "language": "код мови (uk, en, ru, тощо)",
  "confidence": число від 0 до 1
}
`;

      const response = await this.ai.generateResponse(prompt, {
        temperature: 0.1,
        maxTokens: 100
      });

      const result = this.parseLanguageDetectionResponse(response.content);
      
      return {
        language: result.language,
        confidence: result.confidence,
        supported: this.supportedLanguages.includes(result.language)
      };
    } catch (error) {
      logger.error('Помилка визначення мови документа', {
        component: 'MultilingualDocumentProcessor',
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Повертаємо результат за замовчуванням
      return {
        language: 'unknown',
        confidence: 0,
        supported: false
      };
    }
  }

  /**
   * Просте визначення мови для коротких текстів
   */
  private detectLanguageSimple(text: string): LanguageDetectionResult {
    const ukrainianChars = /[іїєґ]/gi;
    const russianChars = /[ыэъ]/gi;
    const englishChars = /[a-z]/gi;
    
    const ukrainianMatches = (text.match(ukrainianChars) || []).length;
    const russianMatches = (text.match(russianChars) || []).length;
    const englishMatches = (text.match(englishChars) || []).length;
    
    if (ukrainianMatches > russianMatches && ukrainianMatches > englishMatches) {
      return {
        language: 'uk',
        confidence: Math.min(0.9, ukrainianMatches / text.length * 10),
        supported: true
      };
    }
    
    if (russianMatches > ukrainianMatches && russianMatches > englishMatches) {
      return {
        language: 'ru',
        confidence: Math.min(0.8, russianMatches / text.length * 10),
        supported: true
      };
    }
    
    if (englishMatches > ukrainianMatches && englishMatches > russianMatches) {
      return {
        language: 'en',
        confidence: Math.min(0.8, englishMatches / text.length * 10),
        supported: true
      };
    }
    
    return {
      language: 'unknown',
      confidence: 0.1,
      supported: false
    };
  }

  /**
   * Перекладає документ
   */
  async translateDocument(
    text: string, 
    targetLanguage: string,
    sourceLanguage?: string
  ): Promise<TranslationResult> {
    try {
      // Check cache first
      const cacheKey = `${text.substring(0, 100)}_${sourceLanguage || 'auto'}_${targetLanguage}`;
      const cachedResult = this.translationCache.get(cacheKey);
      
      if (cachedResult) {
        // Check if cache is still valid
        const now = Date.now();
        if (now - cachedResult.confidence < this.CACHE_TTL) {
          return cachedResult;
        } else {
          // Remove expired cache entry
          this.translationCache.delete(cacheKey);
        }
      }
      
      if (!this.ai) {
        throw new Error('AIService не ініціалізовано');
      }

      if (!text || text.trim().length === 0) {
        return {
          originalText: '',
          translatedText: '',
          sourceLanguage: sourceLanguage || 'unknown',
          targetLanguage,
          confidence: 0
        };
      }

      // Визначаємо мову джерела якщо не вказана
      let srcLang = sourceLanguage;
      if (!srcLang || srcLang === 'unknown') {
        const detection = await this.detectLanguage(text);
        srcLang = detection.language;
      }

      // Якщо мова джерела співпадає з цільовою, повертаємо оригінал
      if (srcLang === targetLanguage) {
        return {
          originalText: text,
          translatedText: text,
          sourceLanguage: srcLang,
          targetLanguage,
          confidence: 1
        };
      }

      // Виконуємо переклад
      const prompt = `
Переклади наступний текст з ${this.getLanguageName(srcLang)} на ${this.getLanguageName(targetLanguage)}.
Зберігай форматування та структуру тексту.

Текст для перекладу:
"${text}"

Відповідай тільки перекладеним текстом без додаткових пояснень.
`;

      const response = await this.ai.generateResponse(prompt, {
        temperature: 0.3,
        maxTokens: Math.min(4000, text.length * 2)
      });

      const result: TranslationResult = {
        originalText: text,
        translatedText: response.content.trim(),
        sourceLanguage: srcLang || 'unknown',
        targetLanguage,
        confidence: Date.now() // Using timestamp as cache expiration marker
      };
      
      // Cache the result
      this.translationCache.set(cacheKey, result);
      
      return result;
    } catch (error) {
      logger.error('Помилка перекладу документа', {
        component: 'MultilingualDocumentProcessor',
        targetLanguage,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Повертаємо оригінал у разі помилки
      return {
        originalText: text,
        translatedText: text,
        sourceLanguage: sourceLanguage || 'unknown',
        targetLanguage,
        confidence: 0
      };
    }
  }

  /**
   * Перекладає документ Google Drive
   */
  async translateDriveFile(
    file: DriveFile,
    content: string,
    targetLanguage: string
  ): Promise<TranslationResult> {
    try {
      // Визначаємо мову документа
      const detection = await this.detectLanguage(content);
      
      // Перекладаємо документ
      const translation = await this.translateDocument(
        content,
        targetLanguage,
        detection.language
      );
      
      return translation;
    } catch (error) {
      logger.error('Помилка перекладу файлу Google Drive', {
        component: 'MultilingualDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        targetLanguage,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Встановлює конфігурацію локалізації для користувача
   */
  setUserLocalizationConfig(userId: string, config: LocalizationConfig): void {
    this.localizationConfigs.set(userId, config);
  }

  /**
   * Отримує конфігурацію локалізації для користувача
   */
  getUserLocalizationConfig(userId: string): LocalizationConfig | undefined {
    return this.localizationConfigs.get(userId);
  }

  /**
   * Локалізує інтерфейс для користувача
   */
  async localizeInterface(userId: string, interfaceText: string): Promise<string> {
    try {
      const config = this.localizationConfigs.get(userId);
      if (!config) {
        // Return original text if no config
        return interfaceText;
      }
      
      // Use user's preferred language or auto-detect
      const targetLanguage = config.userLanguage;
      
      // Translate interface text
      const translation = await this.translateDocument(
        interfaceText,
        targetLanguage
      );
      
      return translation.translatedText;
    } catch (error) {
      logger.error('Помилка локалізації інтерфейсу', {
        component: 'MultilingualDocumentProcessor',
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Return original text in case of error
      return interfaceText;
    }
  }

  /**
   * Виконує крос-мовний пошук
   */
  async crossLanguageSearch(
    query: string,
    documents: Array<{ id: string; name: string; content: string }>,
    targetLanguage: string
  ): Promise<CrossLanguageSearchResult[]> {
    try {
      // Detect query language
      const queryDetection = await this.detectLanguage(query);
      const queryLanguage = queryDetection.language;
      
      // If query is already in target language, do regular search
      if (queryLanguage === targetLanguage) {
        return this.performRegularSearch(query, documents);
      }
      
      // Translate query to target language
      const translatedQuery = await this.translateDocument(
        query,
        targetLanguage,
        queryLanguage
      );
      
      // Search in translated documents
      const results: CrossLanguageSearchResult[] = [];
      
      for (const doc of documents) {
        // Detect document language
        const docDetection = await this.detectLanguage(doc.content);
        const docLanguage = docDetection.language;
        
        // If document is already in target language, search directly
        if (docLanguage === targetLanguage) {
          const relevance = this.calculateRelevance(translatedQuery.translatedText, doc.content);
          if (relevance > 0.1) {
            results.push({
              fileId: doc.id,
              fileName: doc.name,
              snippet: this.extractSnippet(doc.content, translatedQuery.translatedText),
              sourceLanguage: docLanguage,
              targetLanguage,
              relevanceScore: relevance
            });
          }
        } else {
          // Translate document to target language
          const translatedDoc = await this.translateDocument(
            doc.content,
            targetLanguage,
            docLanguage
          );
          
          const relevance = this.calculateRelevance(translatedQuery.translatedText, translatedDoc.translatedText);
          if (relevance > 0.1) {
            results.push({
              fileId: doc.id,
              fileName: doc.name,
              snippet: this.extractSnippet(translatedDoc.translatedText, translatedQuery.translatedText),
              sourceLanguage: docLanguage,
              targetLanguage,
              relevanceScore: relevance
            });
          }
        }
      }
      
      // Sort by relevance score
      return results.sort((a, b) => b.relevanceScore - a.relevanceScore);
    } catch (error) {
      logger.error('Помилка крос-мовного пошуку', {
        component: 'MultilingualDocumentProcessor',
        targetLanguage,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Fall back to regular search
      return this.performRegularSearch(query, documents);
    }
  }

  /**
   * Виконує звичайний пошук
   */
  private performRegularSearch(
    query: string,
    documents: Array<{ id: string; name: string; content: string }>
  ): CrossLanguageSearchResult[] {
    const results: CrossLanguageSearchResult[] = [];
    
    for (const doc of documents) {
      const relevance = this.calculateRelevance(query, doc.content);
      if (relevance > 0.1) {
        results.push({
          fileId: doc.id,
          fileName: doc.name,
          snippet: this.extractSnippet(doc.content, query),
          sourceLanguage: 'unknown',
          targetLanguage: 'unknown',
          relevanceScore: relevance
        });
      }
    }
    
    return results.sort((a, b) => b.relevanceScore - a.relevanceScore);
  }

  /**
   * Обчислює релевантність між запитом та текстом
   */
  private calculateRelevance(query: string, text: string): number {
    // Simple keyword matching approach
    const queryWords = query.toLowerCase().split(/\s+/);
    const textWords = text.toLowerCase().split(/\s+/);
    
    let matches = 0;
    for (const queryWord of queryWords) {
      if (textWords.some(textWord => textWord.includes(queryWord))) {
        matches++;
      }
    }
    
    return matches / queryWords.length;
  }

  /**
   * Витягує фрагмент тексту навколо ключових слів
   */
  private extractSnippet(text: string, query: string, snippetLength: number = 200): string {
    const queryWords = query.toLowerCase().split(/\s+/);
    
    // Find the first occurrence of any query word
    let bestPosition = -1;
    for (const queryWord of queryWords) {
      const position = text.toLowerCase().indexOf(queryWord);
      if (position !== -1 && (bestPosition === -1 || position < bestPosition)) {
        bestPosition = position;
      }
    }
    
    if (bestPosition === -1) {
      // If no query word found, return beginning of text
      return text.substring(0, snippetLength) + (text.length > snippetLength ? '...' : '');
    }
    
    // Extract snippet centered around the found position
    const start = Math.max(0, bestPosition - Math.floor(snippetLength / 2));
    const end = Math.min(text.length, start + snippetLength);
    
    let snippet = text.substring(start, end);
    if (start > 0) snippet = '...' + snippet;
    if (end < text.length) snippet = snippet + '...';
    
    return snippet;
  }

  /**
   * Отримує назву мови
   */
  private getLanguageName(languageCode: string): string {
    return this.languageNames[languageCode] || languageCode;
  }

  /**
   * Парсить відповідь визначення мови
   */
  private parseLanguageDetectionResponse(response: string): { language: string; confidence: number } {
    try {
      // Шукаємо JSON у відповіді
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        return {
          language: parsed.language || 'unknown',
          confidence: typeof parsed.confidence === 'number' ? parsed.confidence : 0.5
        };
      }
    } catch (error) {
      logger.warn('Помилка парсингу відповіді визначення мови', {
        component: 'MultilingualDocumentProcessor',
        error: error instanceof Error ? error.message : String(error)
      });
    }
    
    // Повертаємо значення за замовчуванням
    return {
      language: 'unknown',
      confidence: 0.5
    };
  }

  /**
   * Отримує список підтримуваних мов
   */
  getSupportedLanguages(): { code: string; name: string }[] {
    return this.supportedLanguages.map(code => ({
      code,
      name: this.languageNames[code] || code
    }));
  }

  /**
   * Перевіряє чи підтримується мова
   */
  isLanguageSupported(languageCode: string): boolean {
    return this.supportedLanguages.includes(languageCode);
  }

  // === BaseService required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('MultilingualDocumentProcessor ініціалізовано', {
      component: 'MultilingualDocumentProcessor'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('MultilingualDocumentProcessor зупинено', {
      component: 'MultilingualDocumentProcessor'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      supportedLanguages: this.supportedLanguages.length,
      cachedTranslations: this.translationCache.size,
      userConfigs: this.localizationConfigs.size
    };
  }

  protected onGetStats(): any {
    return {
      supportedLanguages: this.supportedLanguages.length,
      cachedTranslations: this.translationCache.size,
      userConfigs: this.localizationConfigs.size
    };
  }
}