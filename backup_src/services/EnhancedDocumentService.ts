/**
 * Розширений сервіс документів для глибокого аналізу та обробки
 * Enhanced Document Service для AI Bot
 */

import type { GoogleService } from './GoogleService';
import type { AIService } from './AIService';
import logger from '@/utils/logger';

interface DocumentInsight {
  summary: string;
  keyPoints: string[];
  documentType: 'contract' | 'report' | 'letter' | 'instruction' | 'other';
  urgency: 'low' | 'medium' | 'high' | 'critical';
  actionItems: string[];
  relatedDocuments?: string[];
  entities: {
    persons: string[];
    organizations: string[];
    dates: string[];
    amounts: string[];
    locations: string[];
  };
  confidence: number;
}


export class EnhancedDocumentService {
  constructor(
    private googleService: GoogleService,
    private aiService: AIService
  ) {}

  /**
   * Глибокий аналіз документа з використанням AI
   */
  async analyzeDocument(fileId: string): Promise<DocumentInsight> {
    try {
      // Отримуємо контент документа
      const content = await this.extractDocumentContent(fileId);
      
      // Створюємо спеціалізований промпт для аналізу документів
      const analysisPrompt = this.buildDocumentAnalysisPrompt(content);
      
      // Отримуємо AI аналіз
      const aiResponse = await this.aiService.generateResponse(analysisPrompt, {
        model: 'gpt-4',
        temperature: 0.2,
        maxTokens: 1000,
        useCache: true
      });

      // Парсимо та структуруємо результат
      const insights = this.parseDocumentInsights(aiResponse.content);
      
      logger.info('Документ проаналізовано', {
        component: 'EnhancedDocumentService',
        fileId,
        documentType: insights.documentType,
        urgency: insights.urgency,
        actionItems: insights.actionItems.length
      });

      return insights;
    } catch (error) {
      logger.error('Помилка аналізу документа', {
        component: 'EnhancedDocumentService',
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Побудова спеціалізованого промпта для аналізу документів
   */
  private buildDocumentAnalysisPrompt(content: string): string {
    return `
Ти — експерт-аналітик документів для українських державних та військових установ. 
Проаналізуй наданий документ та надай структуровану відповідь.

📄 ДОКУМЕНТ ДЛЯ АНАЛІЗУ:
${content.substring(0, 4000)}

🎯 ЗАВДАННЯ АНАЛІЗУ:
Проаналізуй документ та визнач:

1. ТИП ДОКУМЕНТА (contract/report/letter/instruction/other)
2. РІВЕНЬ ТЕРМІНОВОСТІ (low/medium/high/critical)
3. ОСНОВНІ ПУНКТИ (до 5 найважливіших)
4. НЕОБХІДНІ ДІЇ (конкретні завдання)
5. СУТНОСТІ (особи, організації, дати, суми, локації)

📋 ФОРМАТ ВІДПОВІДІ (JSON):
{
  "summary": "Короткий зміст документа (2-3 речення)",
  "keyPoints": ["Ключовий пункт 1", "Ключовий пункт 2"],
  "documentType": "тип_документа",
  "urgency": "рівень_терміновості",
  "actionItems": ["Дія 1", "Дія 2"],
  "entities": {
    "persons": ["Особа 1", "Особа 2"],
    "organizations": ["Організація 1"],
    "dates": ["2024-01-15", "2024-02-01"],
    "amounts": ["10000 грн", "5000 доларів"],
    "locations": ["Київ", "Харків"]
  },
  "confidence": 0.95
}

🔍 ОСОБЛИВОСТІ АНАЛІЗУ:
- Враховуй українську специфіку та термінологію
- Виділяй важливі дедлайни та терміни
- Розпізнавай військові та адміністративні документи
- Звертай увагу на підписи та печатки
- Визначай пріоритетність документів

💡 ВІДПОВІДЬ УКРАЇНСЬКОЮ МОВОЮ В JSON ФОРМАТІ:`;
  }

  /**
   * Витяг контенту з різних типів документів
   */
  private async extractDocumentContent(fileId: string): Promise<string> {
    // Тут інтеграція з Google Drive API для отримання контенту
    // Підтримка різних форматів: PDF, DOCX, TXT, Google Docs
    
    try {
      // Отримуємо файл через GoogleService
      const fileInfo = await this.googleService.getDriveFileMetadata(fileId);
      
      if (fileInfo.mimeType?.includes('document')) {
        const result = await this.googleService.extractTextForChat(fileId);
        return result.text;
      } else if (fileInfo.mimeType?.includes('pdf')) {
        return await this.extractPdfText(fileId);
      } else if (fileInfo.mimeType?.includes('spreadsheet')) {
        return await this.extractSpreadsheetText(fileId);
      }
      
      return 'Не вдалося витягти текст з документа';
    } catch (error) {
      logger.warn('Помилка витягу контенту документа', {
        component: 'EnhancedDocumentService',
        fileId,
        error
      });
      return 'Помилка читання документа';
    }
  }

  /**
   * Парсинг AI відповіді в структуровані insights
   */
  private parseDocumentInsights(aiResponse: string): DocumentInsight {
    try {
      // Витягуємо JSON з відповіді
      const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        return {
          summary: parsed.summary || 'Аналіз недоступний',
          keyPoints: parsed.keyPoints || [],
          documentType: parsed.documentType || 'other',
          urgency: parsed.urgency || 'low',
          actionItems: parsed.actionItems || [],
          entities: {
            persons: parsed.entities?.persons || [],
            organizations: parsed.entities?.organizations || [],
            dates: parsed.entities?.dates || [],
            amounts: parsed.entities?.amounts || [],
            locations: parsed.entities?.locations || []
          },
          confidence: parsed.confidence || 0.5
        };
      }
    } catch (error) {
      logger.warn('Помилка парсингу AI відповіді', {
        component: 'EnhancedDocumentService',
        error
      });
    }

    // Fallback структура
    return {
      summary: 'Автоматичний аналіз недоступний',
      keyPoints: [],
      documentType: 'other',
      urgency: 'low',
      actionItems: [],
      entities: { persons: [], organizations: [], dates: [], amounts: [], locations: [] },
      confidence: 0.1
    };
  }

  /**
   * Витяг тексту з PDF (заглушка для майбутньої реалізації)
   */
  private async extractPdfText(_fileId: string): Promise<string> {
    // TODO: Інтеграція з PDF парсером
    return 'PDF обробка буде реалізована';
  }

  /**
   * Витяг тексту зі spreadsheet
   */
  private async extractSpreadsheetText(_fileId: string): Promise<string> {
    // TODO: Інтеграція з Google Sheets API
    return 'Spreadsheet обробка буде реалізована';
  }

  /**
   * Пошук схожих документів на основі аналізу
   */
  async findRelatedDocuments(_insights: DocumentInsight): Promise<string[]> {
    // TODO: Реалізація пошуку схожих документів
    // Використання векторного пошуку або keywords matching
    return [];
  }

  /**
   * Створення автоматичного звіту по документу
   */
  async generateDocumentReport(fileId: string): Promise<string> {
    const insights = await this.analyzeDocument(fileId);
    
    return `
📄 **ЗВІТ ПО ДОКУМЕНТУ**

📋 **Загальна інформація:**
- Тип документа: ${this.translateDocumentType(insights.documentType)}
- Рівень терміновості: ${this.translateUrgency(insights.urgency)}
- Впевненість аналізу: ${Math.round(insights.confidence * 100)}%

📝 **Короткий зміст:**
${insights.summary}

🔑 **Ключові пункти:**
${insights.keyPoints.map((point, i) => `${i + 1}. ${point}`).join('\n')}

✅ **Необхідні дії:**
${insights.actionItems.map((action, i) => `${i + 1}. ${action}`).join('\n')}

👥 **Виявлені сутності:**
- **Особи:** ${insights.entities.persons.join(', ') || 'не виявлено'}
- **Організації:** ${insights.entities.organizations.join(', ') || 'не виявлено'}
- **Дати:** ${insights.entities.dates.join(', ') || 'не виявлено'}
- **Суми:** ${insights.entities.amounts.join(', ') || 'не виявлено'}
- **Локації:** ${insights.entities.locations.join(', ') || 'не виявлено'}
`;
  }

  private translateDocumentType(type: string): string {
    const translations = {
      contract: 'Договір',
      report: 'Звіт',
      letter: 'Лист',
      instruction: 'Інструкція',
      other: 'Інший'
    };
    return translations[type as keyof typeof translations] || type;
  }

  private translateUrgency(urgency: string): string {
    const translations = {
      low: 'Низький',
      medium: 'Середній',
      high: 'Високий',
      critical: 'Критичний'
    };
    return translations[urgency as keyof typeof translations] || urgency;
  }
}