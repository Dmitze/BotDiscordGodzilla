import { BaseServiceClass } from '@/core/BaseService';
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

export class MultilingualDocumentProcessor extends BaseServiceClass {
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

      return {
        originalText: text,
        translatedText: response.content.trim(),
        sourceLanguage: srcLang || 'unknown',
        targetLanguage,
        confidence: 0.9 // Припускаємо високу впевненість для AI перекладу
      };
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

  // === BaseServiceClass required methods ===
  
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
      supportedLanguages: this.supportedLanguages.length
    };
  }

  protected onGetStats(): any {
    return {
      supportedLanguages: this.supportedLanguages.length
    };
  }
}