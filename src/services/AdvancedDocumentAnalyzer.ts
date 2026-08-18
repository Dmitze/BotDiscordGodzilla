/**
 * 🧠 Розширений аналізатор документів з AI та машинним навчанням
 * Advanced Document Analyzer with AI & ML capabilities
 */

import type { AIService } from './AIService';
import type { GoogleService } from './GoogleService';

import logger from '@/utils/logger';

interface DocumentEntity {
  type: 'person' | 'organization' | 'date' | 'amount' | 'location' | 'document' | 'law' | 'project';
  value: string;
  confidence: number;
  context: string;
  position?: { start: number; end: number };
}

interface DocumentRelationship {
  source: string;
  target: string;
  type: 'references' | 'supersedes' | 'amends' | 'related_to' | 'part_of';
  confidence: number;
  description: string;
}

interface DocumentInsight {
  summary: string;
  keyTopics: string[];
  documentType: 'business_contract' | 'administrative_doc' | 'legal_contract' | 'financial_report' | 'technical_spec' | 'communication' | 'other';
  urgencyLevel: 'critical' | 'high' | 'medium' | 'low';
  actionItems: Array<{
    action: string;
    priority: 'high' | 'medium' | 'low';
    deadline?: Date;
    assignee?: string;
    status: 'pending' | 'in_progress' | 'completed';
  }>;
  entities: DocumentEntity[];
  relationships: DocumentRelationship[];
  compliance: {
    score: number; // 0-100
    issues: string[];
    recommendations: string[];
  };
  readability: {
    score: number; // 0-100
    level: 'elementary' | 'middle' | 'high' | 'college' | 'graduate';
    suggestions: string[];
  };
  sentiment: {
    overall: 'positive' | 'neutral' | 'negative';
    confidence: number;
    aspects: Record<string, number>; // topic -> sentiment score
  };
  riskAssessment: {
    level: 'low' | 'medium' | 'high' | 'critical';
    factors: string[];
    mitigation: string[];
  };
}

interface DocumentPattern {
  name: string;
  description: string;
  regex?: string;
  aiPrompt?: string;
  confidence: number;
  category: 'structure' | 'content' | 'format' | 'language';
}

export class AdvancedDocumentAnalyzer {
  private patterns = new Map<string, DocumentPattern>();
  // private entityCache = new Map<string, DocumentEntity[]>();
  private insightCache = new Map<string, { insight: DocumentInsight; timestamp: number }>();
  private readonly CACHE_TTL = 24 * 60 * 60 * 1000; // 24 години

  constructor(
    private aiService: AIService,
    private googleService: GoogleService
  ) {
    this.initializePatterns();
  }

  /**
   * 🎯 Глибокий аналіз документа з машинним навчанням
   */
  async analyzeDocument(fileId: string, options: {
    includeEntities?: boolean;
    includeRelationships?: boolean;
    includeCompliance?: boolean;
    includeSentiment?: boolean;
    includeRiskAssessment?: boolean;
    language?: 'uk' | 'en';
  } = {}): Promise<DocumentInsight> {
    const {
      includeEntities = true,
      includeRelationships = true,
      includeCompliance = true,
      includeSentiment = false,
      includeRiskAssessment = true,
      language = 'uk'
    } = options;

    try {
      // Перевірка кешу
      const cacheKey = `analysis_${fileId}_${JSON.stringify(options)}`;
      const cached = this.insightCache.get(cacheKey);
      if (cached && (Date.now() - cached.timestamp) < this.CACHE_TTL) {
        logger.debug('Повернення кешованого аналізу документа', { fileId });
        return cached.insight;
      }

      // Витяг контенту та метаданих
      const [content, metadata] = await Promise.all([
        this.extractDocumentContent(fileId),
        this.googleService.getDriveFileMetadata(fileId)
      ]);

      // Паралельний аналіз різних аспектів
      const analysisPromises: Promise<any>[] = [
        this.analyzeDocumentStructure(content, language),
        this.classifyDocument(content, metadata, language)
      ];

      if (includeEntities) {
        analysisPromises.push(this.extractEntities(content, language));
      }
      if (includeRelationships) {
        analysisPromises.push(this.findRelationships(content, language));
      }
      if (includeCompliance) {
        analysisPromises.push(this.assessCompliance(content, language));
      }
      if (includeSentiment) {
        analysisPromises.push(this.analyzeSentiment(content, language));
      }
      if (includeRiskAssessment) {
        analysisPromises.push(this.assessRisks(content, language));
      }

      const results = await Promise.all(analysisPromises);
      
      // Об'єднання результатів аналізу
      const insight: DocumentInsight = {
        summary: results[0].summary,
        keyTopics: results[0].keyTopics,
        documentType: results[1].type,
        urgencyLevel: results[1].urgency,
        actionItems: results[0].actionItems || [],
        entities: includeEntities ? (results[2] || []) : [],
        relationships: includeRelationships ? (results[3] || []) : [],
        compliance: includeCompliance ? (results[4] || { score: 0, issues: [], recommendations: [] }) : { score: 0, issues: [], recommendations: [] },
        readability: results[0].readability || { score: 50, level: 'middle', suggestions: [] },
        sentiment: includeSentiment ? (results[5] || { overall: 'neutral', confidence: 0, aspects: {} }) : { overall: 'neutral', confidence: 0, aspects: {} },
        riskAssessment: includeRiskAssessment ? (results[6] || { level: 'low', factors: [], mitigation: [] }) : { level: 'low', factors: [], mitigation: [] }
      };

      // Кешування результату
      this.insightCache.set(cacheKey, {
        insight,
        timestamp: Date.now()
      });

      logger.info('Документ проаналізовано', {
        component: 'AdvancedDocumentAnalyzer',
        fileId,
        documentType: insight.documentType,
        urgency: insight.urgencyLevel,
        entitiesCount: insight.entities.length,
        actionItemsCount: insight.actionItems.length
      });

      return insight;

    } catch (error) {
      logger.error('Помилка аналізу документа', {
        component: 'AdvancedDocumentAnalyzer',
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * 🔍 Структурний аналіз документа
   */
  private async analyzeDocumentStructure(content: string, language: 'uk' | 'en'): Promise<any> {
    const prompt = language === 'uk' ? `
Проаналізуй структуру та зміст документа. Визнач основні розділи, теми та дії.

📄 ДОКУМЕНТ:
${content.substring(0, 4000)}

🎯 ЗАВДАННЯ:
1. Створи короткий зміст (2-3 речення)
2. Виділи ключові теми
3. Знайди дії які потребують виконання
4. Оціни читабельність тексту

📋 ВІДПОВІДЬ У JSON:
{
  "summary": "Короткий зміст документа",
  "keyTopics": ["Тема 1", "Тема 2", "Тема 3"],
  "actionItems": [
    {
      "action": "Опис дії",
      "priority": "high|medium|low",
      "deadline": "2024-12-31T23:59:59Z",
      "assignee": "Відповідальна особа"
    }
  ],
  "readability": {
    "score": 75,
    "level": "middle",
    "suggestions": ["Поради для покращення"]
  }
}` : `
Analyze document structure and content. Identify main sections, topics, and actions.

📄 DOCUMENT:
${content.substring(0, 4000)}

🎯 TASK:
1. Create brief summary (2-3 sentences)
2. Extract key topics
3. Find actions requiring execution
4. Assess text readability

📋 JSON RESPONSE:
{
  "summary": "Brief document summary",
  "keyTopics": ["Topic 1", "Topic 2", "Topic 3"],
  "actionItems": [
    {
      "action": "Action description",
      "priority": "high|medium|low", 
      "deadline": "2024-12-31T23:59:59Z",
      "assignee": "Responsible person"
    }
  ],
  "readability": {
    "score": 75,
    "level": "middle",
    "suggestions": ["Improvement suggestions"]
  }
}`;

    const response = await this.aiService.generateResponse(prompt, {
      temperature: 0.2,
      maxTokens: 1000,
      useCache: true
    });

    return this.parseAIResponse(response.content);
  }

  /**
   * 📊 Класифікація документа та визначення терміновості
   */
  private async classifyDocument(content: string, metadata: any, language: 'uk' | 'en'): Promise<any> {
    const contextPrompt = this.buildClassificationPrompt(content, metadata, language);
    
    const response = await this.aiService.generateResponse(contextPrompt, {
      temperature: 0.1,
      maxTokens: 500,
      useCache: true
    });

    return this.parseAIResponse(response.content);
  }

  /**
   * 🏷️ Витяг сутностей з тексту
   */
  private async extractEntities(content: string, language: 'uk' | 'en'): Promise<DocumentEntity[]> {
    const entityPrompt = language === 'uk' ? `
Витягни всі важливі сутності з документа:

📄 ТЕКСТ:
${content.substring(0, 3000)}

🎯 ЗНАЙДИ:
- Особи (імена, посади)
- Організації (назви установ, компаній)  
- Дати (терміни, дедлайни)
- Суми (гроші, кількості)
- Локації (міста, адреси)
- Документи (назви, номери)
- Закони (статті, постанови)
- Проекти (назви програм, ініціатив)

📋 JSON ФОРМАТ:
{
  "entities": [
    {
      "type": "person|organization|date|amount|location|document|law|project",
      "value": "Значення",
      "confidence": 0.95,
      "context": "Контекст в якому зустрічається"
    }
  ]
}` : `
Extract all important entities from the document:

📄 TEXT:
${content.substring(0, 3000)}

🎯 FIND:
- Persons (names, positions)
- Organizations (institutions, companies)
- Dates (deadlines, terms)
- Amounts (money, quantities)
- Locations (cities, addresses)
- Documents (titles, numbers)
- Laws (articles, regulations)
- Projects (program names, initiatives)

📋 JSON FORMAT:
{
  "entities": [
    {
      "type": "person|organization|date|amount|location|document|law|project", 
      "value": "Value",
      "confidence": 0.95,
      "context": "Context where it appears"
    }
  ]
}`;

    const response = await this.aiService.generateResponse(entityPrompt, {
      temperature: 0.1,
      maxTokens: 1500,
      useCache: true
    });

    const parsed = this.parseAIResponse(response.content);
    return parsed.entities || [];
  }

  /**
   * 🔗 Пошук зв'язків між документами
   */
  private async findRelationships(_content: string, _language: 'uk' | 'en'): Promise<DocumentRelationship[]> {
    // Спрощена імплементація для прикладу
    // TODO: Повна реалізація пошуку зв'язків
    return [];
  }

  /**
   * ⚖️ Оцінка відповідності регулятивним вимогам
   */
  private async assessCompliance(content: string, language: 'uk' | 'en'): Promise<any> {
    const compliancePrompt = language === 'uk' ? `
Перевір документ на відповідність українському законодавству:

📄 ДОКУМЕНТ:
${content.substring(0, 3000)}

🔍 ПЕРЕВІР:
- Наявність обов'язкових реквізитів
- Відповідність формальним вимогам
- Дотримання процедур
- Потенційні правові ризики

📋 JSON ВІДПОВІДЬ:
{
  "score": 85,
  "issues": ["Проблема 1", "Проблема 2"],
  "recommendations": ["Рекомендація 1", "Рекомендація 2"]
}` : `
Check document compliance with regulations:

📄 DOCUMENT:
${content.substring(0, 3000)}

🔍 CHECK:
- Required elements presence
- Formal requirements compliance
- Procedure adherence
- Potential legal risks

📋 JSON RESPONSE:
{
  "score": 85,
  "issues": ["Issue 1", "Issue 2"],
  "recommendations": ["Recommendation 1", "Recommendation 2"]
}`;

    const response = await this.aiService.generateResponse(compliancePrompt, {
      temperature: 0.2,
      maxTokens: 800,
      useCache: true
    });

    return this.parseAIResponse(response.content);
  }

  /**
   * 😊 Аналіз тону та настрою документа
   */
  private async analyzeSentiment(_content: string, _language: 'uk' | 'en'): Promise<any> {
    // Спрощена імплементація
    return {
      overall: 'neutral' as const,
      confidence: 0.5,
      aspects: {}
    };
  }

  /**
   * ⚠️ Оцінка ризиків
   */
  private async assessRisks(content: string, language: 'uk' | 'en'): Promise<any> {
    const riskPrompt = language === 'uk' ? `
Оціни потенційні ризики в документі:

📄 ДОКУМЕНТ:
${content.substring(0, 2500)}

⚠️ РИЗИКИ:
- Фінансові ризики
- Правові ризики  
- Операційні ризики
- Репутаційні ризики

📋 JSON:
{
  "level": "low|medium|high|critical",
  "factors": ["Фактор ризику 1"],
  "mitigation": ["Заходи пом'якшення 1"]
}` : `
Assess potential risks in document:

📄 DOCUMENT: 
${content.substring(0, 2500)}

⚠️ RISKS:
- Financial risks
- Legal risks
- Operational risks
- Reputational risks

📋 JSON:
{
  "level": "low|medium|high|critical",
  "factors": ["Risk factor 1"],
  "mitigation": ["Mitigation measure 1"]
}`;

    const response = await this.aiService.generateResponse(riskPrompt, {
      temperature: 0.3,
      maxTokens: 600,
      useCache: true
    });

    return this.parseAIResponse(response.content);
  }

  /**
   * 📝 Побудова промпта для класифікації
   */
  private buildClassificationPrompt(content: string, metadata: any, language: 'uk' | 'en'): string {
    return language === 'uk' ? `
Класифікуй документ за типом та терміновістю:

📄 МЕТАДАНІ:
- Назва: ${metadata.name || 'Невідома'}
- Тип: ${metadata.mimeType || 'Невідомий'}
- Розмір: ${metadata.size || 'Невідомий'}

📝 ЗМІСТ:
${content.substring(0, 2000)}

🎯 ВИЗНАЧ:
- Тип документа
- Рівень терміновості

📋 JSON:
{
  "type": "business_contract|administrative_doc|legal_contract|financial_report|technical_spec|communication|other",
  "urgency": "critical|high|medium|low"
}` : `
Classify document by type and urgency:

📄 METADATA:
- Name: ${metadata.name || 'Unknown'}
- Type: ${metadata.mimeType || 'Unknown'}  
- Size: ${metadata.size || 'Unknown'}

📝 CONTENT:
${content.substring(0, 2000)}

🎯 DETERMINE:
- Document type
- Urgency level

📋 JSON:
{
  "type": "business_contract|administrative_doc|legal_contract|financial_report|technical_spec|communication|other",
  "urgency": "critical|high|medium|low"
}`;
  }

  /**
   * 🔧 Допоміжні методи
   */
  private async extractDocumentContent(fileId: string): Promise<string> {
    const result = await this.googleService.extractTextForChat(fileId);
    return result.text;
  }

  private parseAIResponse(response: string): any {
    try {
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        return JSON.parse(jsonMatch[0]);
      }
    } catch (error) {
      logger.warn('Помилка парсингу AI відповіді', { error });
    }
    return {};
  }

  private initializePatterns(): void {
    // Ініціалізація шаблонів для розпізнавання документів
    this.patterns.set('business_contract', {
      name: 'операційний наказ',
      description: 'Розпізнавання операційних наказів',
      confidence: 0.8,
      category: 'structure'
    });

    this.patterns.set('administrative_doc', {
      name: 'Адміністративний документ',
      description: 'Розпізнавання адмін документів',
      confidence: 0.7,
      category: 'structure'
    });
  }

  /**
   * 📊 Генерація звіту аналізу
   */
  async generateAnalysisReport(insight: DocumentInsight, language: 'uk' | 'en' = 'uk'): Promise<string> {
    if (language === 'uk') {
      return `
📊 **ЗВІТ АНАЛІЗУ ДОКУМЕНТА**

📝 **Загальна інформація:**
- Тип: ${this.translateDocumentType(insight.documentType)}
- Терміновість: ${this.translateUrgency(insight.urgencyLevel)}

📖 **Короткий зміст:**
${insight.summary}

🔑 **Ключові теми:**
${insight.keyTopics.map((topic, i) => `${i + 1}. ${topic}`).join('\n')}

✅ **Дії до виконання:**
${insight.actionItems.map((item, i) => 
  `${i + 1}. ${item.action} (Пріоритет: ${item.priority}${item.deadline ? `, Дедлайн: ${item.deadline.toLocaleDateString()}` : ''})`
).join('\n') || 'Немає конкретних дій'}

👥 **Виявлені сутності:**
${insight.entities.length > 0 ? 
  insight.entities.slice(0, 10).map(e => `- ${e.type}: ${e.value} (${Math.round(e.confidence * 100)}%)`).join('\n') 
  : 'Сутності не виявлено'}

⚖️ **Відповідність вимогам:** ${insight.compliance.score}%
${insight.compliance.issues.length > 0 ? 
  `\n🚨 **Проблеми:** ${insight.compliance.issues.join(', ')}` : ''}

⚠️ **Рівень ризику:** ${this.translateRiskLevel(insight.riskAssessment.level)}
${insight.riskAssessment.factors.length > 0 ? 
  `\n**Фактори ризику:** ${insight.riskAssessment.factors.join(', ')}` : ''}

📈 **Читабельність:** ${insight.readability.score}% (${this.translateReadabilityLevel(insight.readability.level)})
`;
    } else {
      return `
📊 **DOCUMENT ANALYSIS REPORT**

📝 **General Information:**
- Type: ${insight.documentType.replace(/_/g, ' ')}
- Urgency: ${insight.urgencyLevel}

📖 **Summary:**
${insight.summary}

🔑 **Key Topics:**
${insight.keyTopics.map((topic, i) => `${i + 1}. ${topic}`).join('\n')}

✅ **Action Items:**
${insight.actionItems.map((item, i) => 
  `${i + 1}. ${item.action} (Priority: ${item.priority}${item.deadline ? `, Deadline: ${item.deadline.toLocaleDateString()}` : ''})`
).join('\n') || 'No specific actions'}

👥 **Detected Entities:**
${insight.entities.length > 0 ? 
  insight.entities.slice(0, 10).map(e => `- ${e.type}: ${e.value} (${Math.round(e.confidence * 100)}%)`).join('\n') 
  : 'No entities detected'}

⚖️ **Compliance Score:** ${insight.compliance.score}%
${insight.compliance.issues.length > 0 ? 
  `\n🚨 **Issues:** ${insight.compliance.issues.join(', ')}` : ''}

⚠️ **Risk Level:** ${insight.riskAssessment.level}
${insight.riskAssessment.factors.length > 0 ? 
  `\n**Risk Factors:** ${insight.riskAssessment.factors.join(', ')}` : ''}

📈 **Readability:** ${insight.readability.score}% (${insight.readability.level})
`;
    }
  }

  // Допоміжні методи перекладу
  private translateDocumentType(type: string): string {
    const translations: Record<string, string> = {
      'business_contract': 'операційний наказ',
      'administrative_doc': 'Адміністративний документ', 
      'legal_contract': 'Юридичний договір',
      'financial_report': 'Фінансовий звіт',
      'technical_spec': 'Технічна специфікація',
      'communication': 'Листування',
      'other': 'Інший'
    };
    return translations[type] || type;
  }

  private translateUrgency(urgency: string): string {
    const translations: Record<string, string> = {
      'critical': 'Критична',
      'high': 'Висока',
      'medium': 'Середня',
      'low': 'Низька'
    };
    return translations[urgency] || urgency;
  }

  private translateRiskLevel(level: string): string {
    const translations: Record<string, string> = {
      'critical': 'Критичний',
      'high': 'Високий', 
      'medium': 'Середній',
      'low': 'Низький'
    };
    return translations[level] || level;
  }

  private translateReadabilityLevel(level: string): string {
    const translations: Record<string, string> = {
      'elementary': 'Початковий',
      'middle': 'Середній',
      'high': 'Високий',
      'college': 'Університетський',
      'graduate': 'Аспірантський'
    };
    return translations[level] || level;
  }
}