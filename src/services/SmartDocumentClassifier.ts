/**
 * Розумна класифікація та автоматичне тегування документів
 * Smart Document Classification & Auto-Tagging Service
 */

import type { AIService } from './AIService';
import type { GoogleService } from './GoogleService';
import { AIPromptTemplateService, type PromptContext } from './AIPromptTemplateService';
import logger from '@/utils/logger';

export interface DocumentTag {
  category: string;
  subcategory?: string;
  keywords: string[];
  confidence: number;
  automated: boolean;
  createdAt: Date;
}

export interface DocumentClassification {
  primaryCategory: 'military' | 'administrative' | 'legal' | 'financial' | 'technical' | 'personal';
  secondaryCategories: string[];
  tags: DocumentTag[];
  sensitivity: 'public' | 'internal' | 'confidential' | 'secret' | 'top_secret';
  retentionPeriod?: number; // in days
  requiredApprovals?: string[];
  relatedDocuments?: string[];
  extractedMetadata: DocumentMetadata;
  confidence: number;
}

export interface DocumentMetadata {
  author?: string;
  organization?: string;
  documentNumber?: string;
  dateCreated?: Date;
  dateModified?: Date;
  recipients?: string[];
  subjects?: string[];
  amounts?: Array<{ value: number; currency: string; context: string }>;
  dates?: Array<{ date: Date; context: string }>;
  locations?: Array<{ name: string; type: 'city' | 'region' | 'country' | 'address' }>;
  references?: Array<{ type: 'law' | 'regulation' | 'order' | 'document'; reference: string }>;
}

export class SmartDocumentClassifier {
  private classificationCache = new Map<string, { classification: DocumentClassification; timestamp: number }>();
  private readonly CACHE_TTL = 24 * 60 * 60 * 1000; // 24 години

  constructor(
    private aiService: AIService,
    private googleService: GoogleService
  ) {}

  /**
   * Класифікація документа з повним аналізом
   */
  async classifyDocument(fileId: string, content?: string): Promise<DocumentClassification> {
    try {
      // Перевірка кешу
      const cached = this.classificationCache.get(fileId);
      if (cached && (Date.now() - cached.timestamp) < this.CACHE_TTL) {
        return cached.classification;
      }

      // Отримання контенту якщо не надано
      if (!content) {
        content = await this.extractDocumentContent(fileId);
      }

      // Отримання метаданих файлу
      const fileMetadata = await this.googleService.getFileInfo(fileId);

      // Класифікація документа
      const classification = await this.performClassification(content, fileMetadata);

      // Збереження в кеш
      this.classificationCache.set(fileId, {
        classification,
        timestamp: Date.now()
      });

      logger.info('Документ класифіковано', {
        component: 'SmartDocumentClassifier',
        fileId,
        primaryCategory: classification.primaryCategory,
        tagsCount: classification.tags.length,
        confidence: classification.confidence
      });

      return classification;

    } catch (error) {
      logger.error('Помилка класифікації документа', {
        component: 'SmartDocumentClassifier',
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Виконання класифікації з AI
   */
  private async performClassification(content: string, fileMetadata: any): Promise<DocumentClassification> {
    const classificationPrompt = this.buildClassificationPrompt(content, fileMetadata);
    
    const aiResponse = await this.aiService.generateResponse(classificationPrompt, {
      model: 'gpt-4',
      temperature: 0.1,
      maxTokens: 1500,
      useCache: true
    });

    return this.parseClassificationResponse(aiResponse.content);
  }

  /**
   * Побудова промпта для класифікації
   */
  private buildClassificationPrompt(content: string, fileMetadata: any): string {
    return `
Ти — експерт-архіваріус з досвідом класифікації документів українських державних та військових установ.
Проаналізуй документ та надай детальну класифікацію.

📄 ІНФОРМАЦІЯ ПРО ФАЙЛ:
- Назва: ${fileMetadata.name || 'Невідома'}
- Тип: ${fileMetadata.mimeType || 'Невідомий'}
- Розмір: ${fileMetadata.size ? `${Math.round(fileMetadata.size / 1024)} KB` : 'Невідомий'}
- Створено: ${fileMetadata.createdTime || 'Невідомо'}

📝 ЗМІСТ ДОКУМЕНТА:
${content.substring(0, 3000)}

🎯 ЗАВДАННЯ КЛАСИФІКАЦІЇ:

1. **ОСНОВНА КАТЕГОРІЯ** (military/administrative/legal/financial/technical/personal)
2. **ДОДАТКОВІ КАТЕГОРІЇ** (підкатегорії, специфічні типи)
3. **ТЕГИ** (ключові слова, теми, сфери)
4. **РІВЕНЬ СЕКРЕТНОСТІ** (public/internal/confidential/secret/top_secret)
5. **МЕТАДАНІ** (автор, організація, номери, дати, суми, посилання)
6. **ЗБЕРІГАННЯ** (термін зберігання, необхідні затвердження)

📋 КРИТЕРІЇ КЛАСИФІКАЦІЇ:

**ВІЙСЬКОВІ ДОКУМЕНТИ:**
- Оперативні плани, накази, звіти
- Бойові донесення, розвіддані
- Матеріально-технічне забезпечення
- Кадрові питання ЗСУ

**АДМІНІСТРАТИВНІ:**
- Розпорядження, постанови
- Протоколи засідань
- Кадрові наказі
- Господарські документи

**ЮРИДИЧНІ:**
- Договори, угоди
- Нормативні акти
- Судові рішення
- Правові висновки

**ФІНАНСОВІ:**
- Бюджетні документи
- Фінансові звіти
- Договори постачання
- Аудиторські висновки

🔍 АЛГОРИТМ АНАЛІЗУ:
1. Визнач тип документа за структурою та змістом
2. Виділи ключові терміни та концепції
3. Знайди метадані (дати, номери, особи, суми)
4. Оціни рівень секретності
5. Визнач терміни зберігання
6. Створи систему тегів

📤 ФОРМАТ ВІДПОВІДІ (JSON):
{
  "primaryCategory": "категорія",
  "secondaryCategories": ["підкатегорія1", "підкатегорія2"],
  "tags": [
    {
      "category": "основна_тема",
      "subcategory": "підтема",
      "keywords": ["ключове_слово1", "ключове_слово2"],
      "confidence": 0.95,
      "automated": true,
      "createdAt": "2024-01-15T10:00:00Z"
    }
  ],
  "sensitivity": "рівень_секретності",
  "retentionPeriod": 1825,
  "requiredApprovals": ["роль1", "роль2"],
  "extractedMetadata": {
    "author": "автор",
    "organization": "організація",
    "documentNumber": "номер_документа",
    "dateCreated": "2024-01-15T10:00:00Z",
    "recipients": ["адресат1", "адресат2"],
    "subjects": ["тема1", "тема2"],
    "amounts": [
      {
        "value": 10000,
        "currency": "UAH",
        "context": "бюджетне_асигнування"
      }
    ],
    "dates": [
      {
        "date": "2024-02-01T00:00:00Z",
        "context": "дедлайн_виконання"
      }
    ],
    "locations": [
      {
        "name": "Київ",
        "type": "city"
      }
    ],
    "references": [
      {
        "type": "law",
        "reference": "Закон України №123"
      }
    ]
  },
  "confidence": 0.92
}

🇺🇦 ВІДПОВІДЬ СУВОРО У JSON ФОРМАТІ:`;
  }

  /**
   * Парсинг відповіді AI
   */
  private parseClassificationResponse(response: string): DocumentClassification {
    try {
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('Не знайдено JSON у відповіді');
      }

      const parsed = JSON.parse(jsonMatch[0]);

      // Конвертація дат
      if (parsed.extractedMetadata) {
        if (parsed.extractedMetadata.dateCreated) {
          parsed.extractedMetadata.dateCreated = new Date(parsed.extractedMetadata.dateCreated);
        }
        if (parsed.extractedMetadata.dateModified) {
          parsed.extractedMetadata.dateModified = new Date(parsed.extractedMetadata.dateModified);
        }
        if (parsed.extractedMetadata.dates) {
          parsed.extractedMetadata.dates = parsed.extractedMetadata.dates.map((d: any) => ({
            ...d,
            date: new Date(d.date)
          }));
        }
      }

      // Конвертація дат в тегах
      if (parsed.tags) {
        parsed.tags = parsed.tags.map((tag: any) => ({
          ...tag,
          createdAt: tag.createdAt ? new Date(tag.createdAt) : new Date(),
          automated: tag.automated !== false
        }));
      }

      return {
        primaryCategory: parsed.primaryCategory || 'administrative',
        secondaryCategories: parsed.secondaryCategories || [],
        tags: parsed.tags || [],
        sensitivity: parsed.sensitivity || 'internal',
        retentionPeriod: parsed.retentionPeriod,
        requiredApprovals: parsed.requiredApprovals || [],
        relatedDocuments: parsed.relatedDocuments || [],
        extractedMetadata: parsed.extractedMetadata || {},
        confidence: parsed.confidence || 0.5
      };

    } catch (error) {
      logger.warn('Помилка парсингу класифікації', {
        component: 'SmartDocumentClassifier',
        error
      });

      // Fallback класифікація
      return {
        primaryCategory: 'administrative',
        secondaryCategories: [],
        tags: [{
          category: 'general',
          keywords: [],
          confidence: 0.1,
          automated: true,
          createdAt: new Date()
        }],
        sensitivity: 'internal',
        extractedMetadata: {},
        confidence: 0.1
      };
    }
  }

  /**
   * Витяг контенту документа
   */
  private async extractDocumentContent(fileId: string): Promise<string> {
    try {
      const fileInfo = await this.googleService.getFileInfo(fileId);
      
      if (fileInfo.mimeType?.includes('document')) {
        return await this.googleService.exportDocument(fileId, 'text/plain');
      } else if (fileInfo.mimeType?.includes('spreadsheet')) {
        // Для таблиць витягуємо структуровані дані
        const sheets = await this.googleService.getSpreadsheetData(fileId);
        return this.formatSpreadsheetsContent(sheets);
      } else if (fileInfo.mimeType?.includes('pdf')) {
        // TODO: Інтеграція з PDF парсером
        return 'PDF документ (обробка буде додана)';
      }
      
      return 'Не вдалося витягти текст з документа';
    } catch (error) {
      logger.warn('Помилка витягу контенту для класифікації', {
        component: 'SmartDocumentClassifier',
        fileId,
        error
      });
      return 'Помилка читання документа';
    }
  }

  /**
   * Форматування контенту таблиць
   */
  private formatSpreadsheetsContent(sheets: any): string {
    try {
      let content = '';
      for (const sheet of sheets) {
        content += `Аркуш: ${sheet.title}\n`;
        if (sheet.data && sheet.data.length > 0) {
          // Перші кілька рядків для аналізу
          content += sheet.data.slice(0, 10)
            .map((row: any[]) => row.join(' | '))
            .join('\n');
          content += '\n\n';
        }
      }
      return content;
    } catch (error) {
      return 'Помилка обробки табличних даних';
    }
  }

  /**
   * Автоматичне додавання тегів до Google Drive
   */
  async applyTagsToFile(fileId: string, classification: DocumentClassification): Promise<void> {
    try {
      // Створюємо описи на основі класифікації
      const tags = classification.tags.map(tag => 
        tag.keywords.join(', ')
      ).join('; ');

      const description = `
Категорія: ${classification.primaryCategory}
Теги: ${tags}
Секретність: ${classification.sensitivity}
Класифіковано: ${new Date().toISOString()}
Впевненість: ${Math.round(classification.confidence * 100)}%
`.trim();

      // Оновлюємо опис файлу
      await this.googleService.updateFileMetadata(fileId, {
        description
      });

      logger.info('Теги застосовано до файлу', {
        component: 'SmartDocumentClassifier',
        fileId,
        tagsCount: classification.tags.length
      });

    } catch (error) {
      logger.warn('Помилка застосування тегів', {
        component: 'SmartDocumentClassifier',
        fileId,
        error
      });
    }
  }

  /**
   * Пошук документів за тегами та категоріями
   */
  async searchByClassification(
    category?: string,
    tags?: string[],
    sensitivity?: string
  ): Promise<Array<{ fileId: string; classification: DocumentClassification }>> {
    const results: Array<{ fileId: string; classification: DocumentClassification }> = [];

    // Проходимо по кешу класифікацій
    for (const [fileId, cached] of this.classificationCache.entries()) {
      const classification = cached.classification;
      let matches = true;

      if (category && classification.primaryCategory !== category) {
        matches = false;
      }

      if (sensitivity && classification.sensitivity !== sensitivity) {
        matches = false;
      }

      if (tags && tags.length > 0) {
        const documentKeywords = classification.tags.flatMap(tag => tag.keywords);
        const hasMatchingTags = tags.some(tag => 
          documentKeywords.some(keyword => 
            keyword.toLowerCase().includes(tag.toLowerCase())
          )
        );
        if (!hasMatchingTags) {
          matches = false;
        }
      }

      if (matches) {
        results.push({ fileId, classification });
      }
    }

    return results.sort((a, b) => b.classification.confidence - a.classification.confidence);
  }

  /**
   * Генерація звіту по класифікації документів
   */
  async generateClassificationReport(): Promise<string> {
    const allClassifications = Array.from(this.classificationCache.values())
      .map(cached => cached.classification);

    if (allClassifications.length === 0) {
      return '📊 Немає даних для звіту по класифікації документів';
    }

    // Підрахунок статистики
    const categoryStats = new Map<string, number>();
    const sensitivityStats = new Map<string, number>();
    const tagStats = new Map<string, number>();

    for (const classification of allClassifications) {
      // Категорії
      const current = categoryStats.get(classification.primaryCategory) || 0;
      categoryStats.set(classification.primaryCategory, current + 1);

      // Рівні секретності
      const sensitivityCurrent = sensitivityStats.get(classification.sensitivity) || 0;
      sensitivityStats.set(classification.sensitivity, sensitivityCurrent + 1);

      // Теги
      for (const tag of classification.tags) {
        for (const keyword of tag.keywords) {
          const tagCurrent = tagStats.get(keyword) || 0;
          tagStats.set(keyword, tagCurrent + 1);
        }
      }
    }

    // Формування звіту
    let report = '📊 **ЗВІТ ПО КЛАСИФІКАЦІЇ ДОКУМЕНТІВ**\n\n';
    
    report += `📈 **Загальна статистика:**\n`;
    report += `- Всього документів: ${allClassifications.length}\n`;
    report += `- Середня впевненість: ${Math.round(allClassifications.reduce((sum, c) => sum + c.confidence, 0) / allClassifications.length * 100)}%\n\n`;

    report += `📂 **По категоріях:**\n`;
    for (const [category, count] of categoryStats.entries()) {
      const percentage = Math.round(count / allClassifications.length * 100);
      report += `- ${this.translateCategory(category)}: ${count} (${percentage}%)\n`;
    }

    report += `\n🔐 **По рівнях секретності:**\n`;
    for (const [sensitivity, count] of sensitivityStats.entries()) {
      const percentage = Math.round(count / allClassifications.length * 100);
      report += `- ${this.translateSensitivity(sensitivity)}: ${count} (${percentage}%)\n`;
    }

    // Топ тегів
    const topTags = Array.from(tagStats.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);

    report += `\n🏷️ **Найпопулярніші теги:**\n`;
    for (const [tag, count] of topTags) {
      report += `- ${tag}: ${count}\n`;
    }

    return report;
  }

  /**
   * Допоміжні методи перекладу
   */
  private translateCategory(category: string): string {
    const translations = {
      military: 'Військові',
      administrative: 'Адміністративні',
      legal: 'Юридичні',
      financial: 'Фінансові',
      technical: 'Технічні',
      personal: 'Персональні'
    };
    return translations[category as keyof typeof translations] || category;
  }

  private translateSensitivity(sensitivity: string): string {
    const translations = {
      public: 'Відкриті',
      internal: 'Внутрішні',
      confidential: 'Конфіденційні',
      secret: 'Таємні',
      top_secret: 'Цілком таємні'
    };
    return translations[sensitivity as keyof typeof translations] || sensitivity;
  }
}