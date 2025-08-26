/**
 * 📚 Knowledge Base Service
 * Comprehensive knowledge management with semantic search, categorization, and AI-powered insights
 */

import { GoogleService } from './GoogleService';
import { AIService } from './AIService';
import { RagService } from './RagService';
import { ResponseCacheService } from './ResponseCacheService';
import logger from '@/utils/logger';

export interface KnowledgeEntry {
  id: string;
  title: string;
  content: string;
  category: string;
  tags: string[];
  source: {
    type: 'google_drive' | 'manual' | 'imported' | 'ai_generated';
    url?: string;
    fileId?: string;
    createdBy?: string;
  };
  metadata: {
    language: string;
    confidence: number;
    lastUpdated: Date;
    accessCount: number;
    rating?: number;
    verified: boolean;
  };
  embeddings?: number[];
  summary?: string;
  keywords?: string[];
}

export interface KnowledgeSearchOptions {
  query: string;
  categories?: string[];
  tags?: string[];
  language?: string;
  minConfidence?: number;
  limit?: number;
  useSemanticSearch?: boolean;
  includeSummary?: boolean;
}

export interface KnowledgeSearchResult {
  entry: KnowledgeEntry;
  relevanceScore: number;
  matchedKeywords: string[];
  snippet: string;
}

export interface KnowledgeStats {
  totalEntries: number;
  categoriesCount: number;
  mostPopularCategory: string;
  averageConfidence: number;
  recentlyAdded: number;
  verifiedEntries: number;
  languages: Record<string, number>;
}

export class KnowledgeBaseService {
  private knowledgeEntries: Map<string, KnowledgeEntry> = new Map();
  private categoryIndex: Map<string, Set<string>> = new Map();
  private tagIndex: Map<string, Set<string>> = new Map();
  private keywordIndex: Map<string, Set<string>> = new Map();

  constructor(
    private readonly googleService: GoogleService,
    private readonly aiService: AIService,
    private readonly ragService: RagService,
    private readonly cacheService: ResponseCacheService
  ) {}

  /**
   * 📖 Add knowledge entry
   */
  async addEntry(
    title: string,
    content: string,
    category: string,
    tags: string[] = [],
    source: KnowledgeEntry['source'],
    metadata?: Partial<KnowledgeEntry['metadata']>
  ): Promise<string> {
    const entryId = this.generateEntryId();
    
    // Generate AI-powered summary and keywords
    const summary = await this.generateSummary(content);
    const keywords = await this.extractKeywords(content, title);
    
    // Generate embeddings for semantic search
    const embeddings = await this.generateEmbeddings(content);

    const entry: KnowledgeEntry = {
      id: entryId,
      title,
      content,
      category,
      tags,
      source,
      metadata: {
        language: metadata?.language || 'uk',
        confidence: metadata?.confidence || 0.8,
        lastUpdated: new Date(),
        accessCount: 0,
        verified: metadata?.verified || false,
        ...metadata
      },
      embeddings,
      summary,
      keywords
    };

    this.knowledgeEntries.set(entryId, entry);
    this.updateIndices(entry);

    logger.info('Knowledge entry added', {
      component: 'KnowledgeBaseService',
      entryId,
      title,
      category,
      tagsCount: tags.length
    });

    return entryId;
  }

  /**
   * 🔍 Search knowledge base
   */
  async search(options: KnowledgeSearchOptions): Promise<KnowledgeSearchResult[]> {
    const cacheKey = ResponseCacheService.generateKey(
      'knowledge_search',
      options.query,
      JSON.stringify(options)
    );

    // Check cache first
    const cachedResults = this.cacheService.get<KnowledgeSearchResult[]>(cacheKey);
    if (cachedResults) {
      return cachedResults;
    }

    let results: KnowledgeSearchResult[] = [];

    if (options.useSemanticSearch && this.ragService) {
      results = await this.performSemanticSearch(options);
    } else {
      results = await this.performKeywordSearch(options);
    }

    // Apply filters
    results = this.applyFilters(results, options);

    // Sort by relevance and limit
    results.sort((a, b) => b.relevanceScore - a.relevanceScore);
    
    if (options.limit) {
      results = results.slice(0, options.limit);
    }

    // Update access counts
    for (const result of results) {
      result.entry.metadata.accessCount++;
    }

    // Cache results
    this.cacheService.set(cacheKey, results, 15, { // 15 minutes cache
      source: 'knowledge_search',
      tags: ['search', 'knowledge']
    });

    logger.debug('Knowledge search completed', {
      component: 'KnowledgeBaseService',
      query: options.query,
      resultsCount: results.length,
      useSemanticSearch: options.useSemanticSearch
    });

    return results;
  }

  /**
   * 📝 Get entry by ID
   */
  getEntry(entryId: string): KnowledgeEntry | null {
    const entry = this.knowledgeEntries.get(entryId);
    if (entry) {
      entry.metadata.accessCount++;
    }
    return entry || null;
  }

  /**
   * ✏️ Update entry
   */
  async updateEntry(
    entryId: string,
    updates: Partial<Pick<KnowledgeEntry, 'title' | 'content' | 'category' | 'tags'>>
  ): Promise<boolean> {
    const entry = this.knowledgeEntries.get(entryId);
    if (!entry) {
      return false;
    }

    // Remove from old indices
    this.removeFromIndices(entry);

    // Apply updates
    if (updates.title) entry.title = updates.title;
    if (updates.content) {
      entry.content = updates.content;
      // Regenerate AI-powered content
      entry.summary = await this.generateSummary(updates.content);
      entry.keywords = await this.extractKeywords(updates.content, entry.title);
      entry.embeddings = await this.generateEmbeddings(updates.content);
    }
    if (updates.category) entry.category = updates.category;
    if (updates.tags) entry.tags = updates.tags;

    entry.metadata.lastUpdated = new Date();

    // Update indices
    this.updateIndices(entry);

    logger.info('Knowledge entry updated', {
      component: 'KnowledgeBaseService',
      entryId,
      updates: Object.keys(updates)
    });

    return true;
  }

  /**
   * 🗑️ Delete entry
   */
  deleteEntry(entryId: string): boolean {
    const entry = this.knowledgeEntries.get(entryId);
    if (!entry) {
      return false;
    }

    this.removeFromIndices(entry);
    this.knowledgeEntries.delete(entryId);

    logger.info('Knowledge entry deleted', {
      component: 'KnowledgeBaseService',
      entryId
    });

    return true;
  }

  /**
   * 📊 Get knowledge base statistics
   */
  getStats(): KnowledgeStats {
    const categories = new Map<string, number>();
    const languages = new Map<string, number>();
    let totalConfidence = 0;
    let verifiedCount = 0;
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    let recentCount = 0;

    for (const entry of this.knowledgeEntries.values()) {
      // Categories
      categories.set(entry.category, (categories.get(entry.category) || 0) + 1);
      
      // Languages
      languages.set(entry.metadata.language, (languages.get(entry.metadata.language) || 0) + 1);
      
      // Confidence
      totalConfidence += entry.metadata.confidence;
      
      // Verified
      if (entry.metadata.verified) verifiedCount++;
      
      // Recent entries
      if (entry.metadata.lastUpdated > thirtyDaysAgo) recentCount++;
    }

    const mostPopularCategory = Array.from(categories.entries())
      .sort((a, b) => b[1] - a[1])[0]?.[0] || '';

    return {
      totalEntries: this.knowledgeEntries.size,
      categoriesCount: categories.size,
      mostPopularCategory,
      averageConfidence: this.knowledgeEntries.size > 0 
        ? Math.round((totalConfidence / this.knowledgeEntries.size) * 100) / 100 
        : 0,
      recentlyAdded: recentCount,
      verifiedEntries: verifiedCount,
      languages: Object.fromEntries(languages)
    };
  }

  /**
   * 📂 Get entries by category
   */
  getEntriesByCategory(category: string): KnowledgeEntry[] {
    const entryIds = this.categoryIndex.get(category) || new Set();
    return Array.from(entryIds)
      .map(id => this.knowledgeEntries.get(id))
      .filter((entry): entry is KnowledgeEntry => entry !== undefined);
  }

  /**
   * 🏷️ Get entries by tag
   */
  getEntriesByTag(tag: string): KnowledgeEntry[] {
    const entryIds = this.tagIndex.get(tag) || new Set();
    return Array.from(entryIds)
      .map(id => this.knowledgeEntries.get(id))
      .filter((entry): entry is KnowledgeEntry => entry !== undefined);
  }

  /**
   * 📈 Get trending topics
   */
  getTrendingTopics(limit: number = 10): Array<{ topic: string; count: number; growth: number }> {
    const topicCounts = new Map<string, number>();
    const recentTopicCounts = new Map<string, number>();
    const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);

    for (const entry of this.knowledgeEntries.values()) {
      // Count all keywords
      for (const keyword of entry.keywords || []) {
        topicCounts.set(keyword, (topicCounts.get(keyword) || 0) + entry.metadata.accessCount);
        
        // Count recent access
        if (entry.metadata.lastUpdated > sevenDaysAgo) {
          recentTopicCounts.set(keyword, (recentTopicCounts.get(keyword) || 0) + entry.metadata.accessCount);
        }
      }
    }

    const trending = Array.from(topicCounts.entries())
      .map(([topic, count]) => {
        const recentCount = recentTopicCounts.get(topic) || 0;
        const growth = count > 0 ? (recentCount / count) * 100 : 0;
        return { topic, count, growth };
      })
      .sort((a, b) => b.growth - a.growth)
      .slice(0, limit);

    return trending;
  }

  /**
   * 🤖 Generate AI-powered insights
   */
  async generateInsights(category?: string): Promise<string> {
    const entries = category 
      ? this.getEntriesByCategory(category)
      : Array.from(this.knowledgeEntries.values());

    if (entries.length === 0) {
      return 'Недостатньо даних для генерації інсайтів.';
    }

    const summaries = entries
      .slice(0, 20) // Limit for performance
      .map(entry => entry.summary || entry.title)
      .join('\n\n');

    const prompt = `
Проаналізуйте наступні записи бази знань та надайте інсайти:

${summaries}

Будь ласка, надайте:
1. Ключові теми та тренди
2. Потенційні прогалини в знаннях
3. Рекомендації для поліпшення
4. Зв'язки між різними темами

Відповідь українською мовою:`;

    try {
      const insights = await this.aiService.generateResponse(prompt, {
        maxTokens: 1000,
        temperature: 0.7
      });

      return typeof insights === 'string' ? insights : String(insights?.content || '');
    } catch (error) {
      logger.error('Failed to generate AI insights', {
        component: 'KnowledgeBaseService',
        error: error instanceof Error ? error.message : String(error)
      });
      return 'Помилка генерації інсайтів. Спробуйте пізніше.';
    }
  }

  /**
   * 🔍 Perform semantic search using RAG
   */
  private async performSemanticSearch(options: KnowledgeSearchOptions): Promise<KnowledgeSearchResult[]> {
    try {
      // Using answer method from RagService instead of non-existent search method
      const ragResponse = await this.ragService.answer(options.query, {
        k: options.limit || 10
      });

      const results: KnowledgeSearchResult[] = [];

      // Try to match RAG chunks with knowledge entries
      for (const chunk of ragResponse.chunks || []) {
        for (const entry of this.knowledgeEntries.values()) {
          const chunkText = String(chunk);
          if (entry.content.includes(chunkText.substring(0, 100))) {
            results.push({
              entry,
              relevanceScore: 0.8, // Default score since RAG doesn't provide scores
              matchedKeywords: entry.keywords || [],
              snippet: chunkText.substring(0, 200) + '...'
            });
            break;
          }
        }
      }

      return results;
    } catch (error) {
      logger.error('Semantic search failed, falling back to keyword search', {
        component: 'KnowledgeBaseService',
        error: error instanceof Error ? error.message : String(error)
      });
      return this.performKeywordSearch(options);
    }
  }

  /**
   * 🔤 Perform keyword-based search
   */
  private async performKeywordSearch(options: KnowledgeSearchOptions): Promise<KnowledgeSearchResult[]> {
    const queryWords = options.query.toLowerCase().split(/\s+/);
    const results: KnowledgeSearchResult[] = [];

    for (const entry of this.knowledgeEntries.values()) {
      let relevanceScore = 0;
      const matchedKeywords: string[] = [];

      // Search in title (higher weight)
      for (const word of queryWords) {
        if (entry.title.toLowerCase().includes(word)) {
          relevanceScore += 3;
          matchedKeywords.push(word);
        }
      }

      // Search in content
      for (const word of queryWords) {
        if (entry.content.toLowerCase().includes(word)) {
          relevanceScore += 1;
          if (!matchedKeywords.includes(word)) {
            matchedKeywords.push(word);
          }
        }
      }

      // Search in keywords
      for (const keyword of entry.keywords || []) {
        for (const word of queryWords) {
          if (keyword.toLowerCase().includes(word)) {
            relevanceScore += 2;
            if (!matchedKeywords.includes(word)) {
              matchedKeywords.push(word);
            }
          }
        }
      }

      if (relevanceScore > 0) {
        const snippet = this.generateSnippet(entry.content, queryWords);
        results.push({
          entry,
          relevanceScore,
          matchedKeywords,
          snippet
        });
      }
    }

    return results;
  }

  /**
   * 📄 Generate content snippet with highlighted matches
   */
  private generateSnippet(content: string, queryWords: string[]): string {
    const sentences = content.split(/[.!?]+/);
    let bestSentence = '';
    let maxMatches = 0;

    for (const sentence of sentences) {
      const matches = queryWords.filter(word => 
        sentence.toLowerCase().includes(word.toLowerCase())
      ).length;

      if (matches > maxMatches) {
        maxMatches = matches;
        bestSentence = sentence.trim();
      }
    }

    return bestSentence.length > 200 
      ? bestSentence.substring(0, 200) + '...'
      : bestSentence;
  }

  /**
   * 🎯 Apply search filters
   */
  private applyFilters(
    results: KnowledgeSearchResult[],
    options: KnowledgeSearchOptions
  ): KnowledgeSearchResult[] {
    return results.filter(result => {
      const entry = result.entry;

      // Category filter
      if (options.categories && !options.categories.includes(entry.category)) {
        return false;
      }

      // Tags filter
      if (options.tags && !options.tags.some(tag => entry.tags.includes(tag))) {
        return false;
      }

      // Language filter
      if (options.language && entry.metadata.language !== options.language) {
        return false;
      }

      // Confidence filter
      if (options.minConfidence && entry.metadata.confidence < options.minConfidence) {
        return false;
      }

      return true;
    });
  }

  /**
   * 🧠 Generate summary using AI
   */
  private async generateSummary(content: string): Promise<string> {
    if (content.length < 200) {
      return content;
    }

    try {
      const prompt = `Створіть стислий саммарі наступного тексту українською мовою (до 100 слів):\n\n${content}`;
      const summary = await this.aiService.generateResponse(prompt, {
        maxTokens: 150,
        temperature: 0.3
      });
      const responseText = typeof summary === 'string' ? summary : String((summary as any)?.content || '');
      return responseText.trim();
    } catch (error) {
      logger.error('Failed to generate summary', {
        component: 'KnowledgeBaseService',
        error: error instanceof Error ? error.message : String(error)
      });
      return content.substring(0, 200) + '...';
    }
  }

  /**
   * 🔤 Extract keywords using AI
   */
  private async extractKeywords(content: string, title: string): Promise<string[]> {
    try {
      const prompt = `Вилучіть до 10 ключових слів та фраз з наступного тексту. Відповідь надайте у форматі: слово1, слово2, слово3\n\nЗаголовок: ${title}\nТекст: ${content}`;
      
      const response = await this.aiService.generateResponse(prompt, {
        maxTokens: 100,
        temperature: 0.3
      });

      const responseText = typeof response === 'string' ? response : String(response?.content || '');
      const keywords = responseText
        .split(',')
        .map((k: string) => k.trim())
        .filter((k: string) => k.length > 2)
        .slice(0, 10);

      return keywords;
    } catch (error) {
      logger.error('Failed to extract keywords', {
        component: 'KnowledgeBaseService',
        error: error instanceof Error ? error.message : String(error)
      });
      // Fallback: simple word extraction
      return content
        .toLowerCase()
        .split(/\s+/)
        .filter(word => word.length > 4)
        .slice(0, 5);
    }
  }

  /**
   * 🔢 Generate embeddings for semantic search
   */
  private async generateEmbeddings(content: string): Promise<number[]> {
    try {
      // This would typically use the embeddings service
      // For now, return empty array as placeholder
      return [];
    } catch (error) {
      logger.error('Failed to generate embeddings', {
        component: 'KnowledgeBaseService',
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * 📇 Update search indices
   */
  private updateIndices(entry: KnowledgeEntry): void {
    // Category index
    if (!this.categoryIndex.has(entry.category)) {
      this.categoryIndex.set(entry.category, new Set());
    }
    this.categoryIndex.get(entry.category)!.add(entry.id);

    // Tag index
    for (const tag of entry.tags) {
      if (!this.tagIndex.has(tag)) {
        this.tagIndex.set(tag, new Set());
      }
      this.tagIndex.get(tag)!.add(entry.id);
    }

    // Keyword index
    for (const keyword of entry.keywords || []) {
      if (!this.keywordIndex.has(keyword)) {
        this.keywordIndex.set(keyword, new Set());
      }
      this.keywordIndex.get(keyword)!.add(entry.id);
    }
  }

  /**
   * 🗑️ Remove from search indices
   */
  private removeFromIndices(entry: KnowledgeEntry): void {
    // Category index
    this.categoryIndex.get(entry.category)?.delete(entry.id);

    // Tag index
    for (const tag of entry.tags) {
      this.tagIndex.get(tag)?.delete(entry.id);
    }

    // Keyword index
    for (const keyword of entry.keywords || []) {
      this.keywordIndex.get(keyword)?.delete(entry.id);
    }
  }

  /**
   * 🆔 Generate unique entry ID
   */
  private generateEntryId(): string {
    return `kb_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
}