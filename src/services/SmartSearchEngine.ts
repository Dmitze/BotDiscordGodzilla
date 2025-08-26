/**
 * 🔍 Розумна система пошуку з AI та семантичним аналізом
 * Smart Search Engine with AI & Semantic Analysis
 */

import type { AIService } from './AIService';
import type { GoogleService } from './GoogleService';
import type { RagService } from './RagService';
import logger from '@/utils/logger';

interface SearchQuery {
  text: string;
  filters?: SearchFilters;
  options?: SearchOptions;
}

interface SearchFilters {
  documentType?: string[];
  dateRange?: { from: Date; to: Date };
  author?: string[];
  tags?: string[];
  mimeType?: string[];
  urgency?: string[];
  folder?: string;
}

interface SearchOptions {
  limit?: number;
  offset?: number;
  includeContent?: boolean;
  semanticSearch?: boolean;
  fuzzyMatch?: boolean;
  language?: 'uk' | 'en' | 'auto';
  sortBy?: 'relevance' | 'date' | 'name' | 'size';
  sortOrder?: 'asc' | 'desc';
}

interface SearchResult {
  fileId: string;
  name: string;
  mimeType: string;
  relevanceScore: number;
  highlights: SearchHighlight[];
  summary?: string;
  metadata: Record<string, any>;
  url?: string;
  path?: string;
  lastModified?: Date;
}

interface SearchHighlight {
  field: 'title' | 'content' | 'description';
  text: string;
  startOffset: number;
  endOffset: number;
  score: number;
}

interface SearchInsight {
  query: string;
  totalResults: number;
  searchTime: number;
  suggestions: string[];
  categories: Record<string, number>;
  semanticExpansions?: string[];
  filters: SearchFilters;
}

interface SearchAnalytics {
  popularQueries: Record<string, number>;
  averageResponseTime: number;
  totalSearches: number;
  successRate: number;
}

export class SmartSearchEngine {
  private searchCache = new Map<string, { results: SearchResult[]; insight: SearchInsight; timestamp: number }>();
  private analytics: SearchAnalytics = {
    popularQueries: {},
    averageResponseTime: 0,
    totalSearches: 0,
    successRate: 0
  };
  private readonly CACHE_TTL = 5 * 60 * 1000; // 5 хвилин

  constructor(
    private aiService: AIService,
    private googleService: GoogleService,
    private ragService: RagService
  ) {}

  /**
   * 🔎 Розумний пошук з AI обробкою
   */
  async search(query: SearchQuery): Promise<{ results: SearchResult[]; insight: SearchInsight }> {
    const startTime = Date.now();
    
    try {
      this.analytics.totalSearches++;
      this.trackQuery(query.text);

      // Перевірка кешу
      const cacheKey = this.generateCacheKey(query);
      const cached = this.searchCache.get(cacheKey);
      if (cached && (Date.now() - cached.timestamp) < this.CACHE_TTL) {
        logger.debug('Повернення кешованих результатів пошуку', { query: query.text });
        return { results: cached.results, insight: cached.insight };
      }

      // Обробка запиту за допомогою AI
      const processedQuery = await this.processQuery(query);
      
      // Паралельний пошук
      const [semanticResults, keywordResults, ragResults] = await Promise.all([
        this.performSemanticSearch(processedQuery),
        this.performKeywordSearch(processedQuery), 
        this.performRAGSearch(processedQuery)
      ]);

      // Об'єднання та ранжування результатів
      const combinedResults = await this.combineAndRankResults([
        ...semanticResults,
        ...keywordResults,
        ...ragResults
      ], processedQuery);

      // Створення insights
      const insight = await this.generateSearchInsight(query, combinedResults, Date.now() - startTime);

      // Кешування
      this.searchCache.set(cacheKey, {
        results: combinedResults,
        insight,
        timestamp: Date.now()
      });

      this.updateAnalytics(true, Date.now() - startTime);

      logger.info('Пошук завершено', {
        component: 'SmartSearchEngine',
        query: query.text,
        resultsCount: combinedResults.length,
        searchTime: Date.now() - startTime
      });

      return { results: combinedResults, insight };

    } catch (error) {
      this.updateAnalytics(false, Date.now() - startTime);
      
      logger.error('Помилка пошуку', {
        component: 'SmartSearchEngine',
        query: query.text,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * 🧠 Обробка запиту за допомогою AI
   */
  private async processQuery(query: SearchQuery): Promise<SearchQuery> {
    const language = query.options?.language || 'auto';
    
    const processingPrompt = language === 'uk' ? `
Покращи пошуковий запит для документів:

Оригінальний запит: "${query.text}"

🎯 ЗАВДАННЯ:
1. Виправ орфографічні помилки
2. Розшир синонімами
3. Додай релевантні терміни
4. Визнач тип документа
5. Визнач мову запиту

📋 JSON ВІДПОВІДЬ:
{
  "expandedQuery": "Розширений запит",
  "synonyms": ["синонім1", "синонім2"],
  "documentTypes": ["тип1", "тип2"],
  "language": "uk|en|auto",
  "intent": "search|filter|analyze",
  "keywords": ["ключове_слово1", "ключове_слово2"]
}` : `
Improve search query for documents:

Original query: "${query.text}"

🎯 TASK:
1. Fix spelling errors
2. Expand with synonyms
3. Add relevant terms
4. Determine document type
5. Detect query language

📋 JSON RESPONSE:
{
  "expandedQuery": "Expanded query",
  "synonyms": ["synonym1", "synonym2"],
  "documentTypes": ["type1", "type2"],
  "language": "uk|en|auto",
  "intent": "search|filter|analyze", 
  "keywords": ["keyword1", "keyword2"]
}`;

    const response = await this.aiService.generateResponse(processingPrompt, {
      temperature: 0.2,
      maxTokens: 500,
      useCache: true
    });

    const processed = this.parseAIResponse(response.content);
    
    return {
      ...query,
      text: processed.expandedQuery || query.text,
      options: {
        ...query.options,
        language: processed.language || query.options?.language
      }
    };
  }

  /**
   * 🔍 Семантичний пошук
   */
  private async performSemanticSearch(query: SearchQuery): Promise<SearchResult[]> {
    if (!query.options?.semanticSearch) return [];

    try {
      // Використання RAG для семантичного пошуку
      const ragResults = await this.ragService.searchRelevantDocuments(query.text, {
        limit: query.options?.limit || 20,
        scoreThreshold: 0.7
      });

      return ragResults.map(result => ({
        fileId: result.fileId,
        name: result.name,
        mimeType: result.mimeType || 'application/octet-stream',
        relevanceScore: result.score,
        highlights: [{
          field: 'content' as const,
          text: result.snippet,
          startOffset: 0,
          endOffset: result.snippet.length,
          score: result.score
        }],
        summary: result.snippet,
        metadata: result.metadata || {},
        url: result.url,
        lastModified: result.lastModified ? new Date(result.lastModified) : undefined
      }));

    } catch (error) {
      logger.warn('Помилка семантичного пошуку', {
        component: 'SmartSearchEngine',
        query: query.text,
        error
      });
      return [];
    }
  }

  /**
   * 🔤 Ключовий пошук
   */
  private async performKeywordSearch(query: SearchQuery): Promise<SearchResult[]> {
    try {
      // Побудова Google Drive запиту
      const driveQuery = this.buildDriveQuery(query);
      
      const driveResults = await this.googleService.listFiles({
        q: driveQuery,
        pageSize: query.options?.limit || 50,
        orderBy: this.mapSortOption(query.options?.sortBy, query.options?.sortOrder)
      });

      // Конвертація результатів Google Drive
      const searchResults: SearchResult[] = [];
      
      for (const file of driveResults.files || []) {
        const relevanceScore = this.calculateRelevanceScore(query.text, file);
        
        searchResults.push({
          fileId: file.id || '',
          name: file.name || 'Без назви',
          mimeType: file.mimeType || 'application/octet-stream',
          relevanceScore,
          highlights: await this.generateHighlights(query.text, file),
          summary: file.description || undefined,
          metadata: {
            size: file.size,
            createdTime: file.createdTime,
            modifiedTime: file.modifiedTime,
            owners: file.owners,
            permissions: file.permissions
          },
          url: file.webViewLink,
          lastModified: file.modifiedTime ? new Date(file.modifiedTime) : undefined
        });
      }

      return searchResults;

    } catch (error) {
      logger.warn('Помилка ключового пошуку', {
        component: 'SmartSearchEngine',
        query: query.text,
        error
      });
      return [];
    }
  }

  /**
   * 📚 RAG пошук
   */
  private async performRAGSearch(query: SearchQuery): Promise<SearchResult[]> {
    // Спрощена імплементація через RAG service
    return await this.performSemanticSearch(query);
  }

  /**
   * 🏆 Об'єднання та ранжування результатів
   */
  private async combineAndRankResults(results: SearchResult[], query: SearchQuery): Promise<SearchResult[]> {
    // Видалення дублікатів
    const uniqueResults = new Map<string, SearchResult>();
    
    for (const result of results) {
      const existing = uniqueResults.get(result.fileId);
      if (!existing || result.relevanceScore > existing.relevanceScore) {
        uniqueResults.set(result.fileId, result);
      }
    }

    // Сортування за релевантністю
    const sortedResults = Array.from(uniqueResults.values())
      .sort((a, b) => b.relevanceScore - a.relevanceScore);

    // Застосування лімітів
    const limit = query.options?.limit || 20;
    const offset = query.options?.offset || 0;
    
    return sortedResults.slice(offset, offset + limit);
  }

  /**
   * 💡 Генерація insights пошуку
   */
  private async generateSearchInsight(
    query: SearchQuery,
    results: SearchResult[],
    searchTime: number
  ): Promise<SearchInsight> {
    // Категоризація результатів
    const categories: Record<string, number> = {};
    for (const result of results) {
      const category = this.categorizeDocument(result.mimeType);
      categories[category] = (categories[category] || 0) + 1;
    }

    // Генерація пропозицій
    const suggestions = await this.generateSearchSuggestions(query.text, results);

    return {
      query: query.text,
      totalResults: results.length,
      searchTime,
      suggestions,
      categories,
      filters: query.filters || {}
    };
  }

  /**
   * 🔧 Допоміжні методи
   */
  private buildDriveQuery(query: SearchQuery): string {
    let driveQuery = `fullText contains '${query.text}'`;
    
    if (query.filters) {
      if (query.filters.mimeType && query.filters.mimeType.length > 0) {
        const mimeConditions = query.filters.mimeType
          .map(mime => `mimeType = '${mime}'`)
          .join(' or ');
        driveQuery += ` and (${mimeConditions})`;
      }
      
      if (query.filters.dateRange) {
        driveQuery += ` and modifiedTime >= '${query.filters.dateRange.from.toISOString()}'`;
        driveQuery += ` and modifiedTime <= '${query.filters.dateRange.to.toISOString()}'`;
      }
    }

    return driveQuery;
  }

  private calculateRelevanceScore(queryText: string, file: any): number {
    let score = 0;
    
    const queryWords = queryText.toLowerCase().split(/\s+/);
    const fileName = (file.name || '').toLowerCase();
    const description = (file.description || '').toLowerCase();

    // Точне співпадіння в назві
    for (const word of queryWords) {
      if (fileName.includes(word)) score += 0.5;
      if (description.includes(word)) score += 0.3;
    }

    // Бонус за актуальність файлу
    if (file.modifiedTime) {
      const daysSinceModified = (Date.now() - new Date(file.modifiedTime).getTime()) / (1000 * 60 * 60 * 24);
      if (daysSinceModified < 7) score += 0.2;
      else if (daysSinceModified < 30) score += 0.1;
    }

    return Math.min(score, 1.0);
  }

  private async generateHighlights(queryText: string, file: any): Promise<SearchHighlight[]> {
    const highlights: SearchHighlight[] = [];
    
    if (file.name && file.name.toLowerCase().includes(queryText.toLowerCase())) {
      highlights.push({
        field: 'title',
        text: file.name,
        startOffset: 0,
        endOffset: file.name.length,
        score: 1.0
      });
    }

    return highlights;
  }

  private categorizeDocument(mimeType: string): string {
    if (mimeType.includes('document') || mimeType.includes('text')) return 'Документи';
    if (mimeType.includes('spreadsheet')) return 'Таблиці';
    if (mimeType.includes('presentation')) return 'Презентації';
    if (mimeType.includes('image')) return 'Зображення';
    if (mimeType.includes('pdf')) return 'PDF';
    return 'Інші';
  }

  private async generateSearchSuggestions(queryText: string, results: SearchResult[]): Promise<string[]> {
    if (results.length === 0) {
      return [`Спробуйте "${queryText}" з іншими словами`, 'Перевірте правопис'];
    }

    // Спрощені пропозиції на основі результатів
    const suggestions = [];
    const categories = new Set(results.map(r => this.categorizeDocument(r.mimeType)));
    
    if (categories.size > 1) {
      suggestions.push(`Фільтрувати тільки ${Array.from(categories)[0]}`);
    }

    return suggestions.slice(0, 3);
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

  private mapSortOption(sortBy?: string, sortOrder?: string): string {
    const order = sortOrder === 'asc' ? '' : ' desc';
    
    switch (sortBy) {
      case 'date': return 'modifiedTime' + order;
      case 'name': return 'name' + order;
      case 'size': return 'quotaBytesUsed' + order;
      default: return 'relevance' + order;
    }
  }

  private generateCacheKey(query: SearchQuery): string {
    return `search_${JSON.stringify(query)}`.replace(/\s/g, '_');
  }

  private trackQuery(queryText: string): void {
    this.analytics.popularQueries[queryText] = (this.analytics.popularQueries[queryText] || 0) + 1;
  }

  private updateAnalytics(success: boolean, responseTime: number): void {
    this.analytics.averageResponseTime = 
      (this.analytics.averageResponseTime * (this.analytics.totalSearches - 1) + responseTime) / 
      this.analytics.totalSearches;
    
    if (success) {
      this.analytics.successRate = 
        (this.analytics.successRate * (this.analytics.totalSearches - 1) + 1) / 
        this.analytics.totalSearches;
    }
  }

  /**
   * 📊 Отримання аналітики пошуку
   */
  getSearchAnalytics(): SearchAnalytics {
    return { ...this.analytics };
  }

  /**
   * 🔍 Автокомплетація запитів
   */
  async getQuerySuggestions(partialQuery: string, limit: number = 5): Promise<string[]> {
    const popularQueries = Object.entries(this.analytics.popularQueries)
      .filter(([query]) => query.toLowerCase().includes(partialQuery.toLowerCase()))
      .sort(([, countA], [, countB]) => countB - countA)
      .slice(0, limit)
      .map(([query]) => query);

    return popularQueries;
  }
}