import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';
import type { AIService } from './AIService';

export interface DocumentSummary {
  fileId: string;
  fileName: string;
  summary: string;
  keyPoints: string[];
  entities: string[];
  sentiment: 'positive' | 'negative' | 'neutral';
  wordCount: number;
  readingTime: number; // in minutes
  generatedAt: Date;
}

export interface SummaryOptions {
  maxLength?: number;
  includeKeyPoints?: boolean;
  includeEntities?: boolean;
  language?: string;
  style?: 'concise' | 'detailed' | 'bullet-points';
}

export class DocumentSummarizationService extends BaseService {
  private aiService: AIService | null = null;
  private summaries: Map<string, DocumentSummary> = new Map();
  private readonly MAX_SUMMARY_LENGTH = 500;
  private readonly MAX_CACHE_ENTRIES = 1000;

  constructor(config: BotConfig) {
    super('DocumentSummarizationService', config);
  }

  /**
   * Initialize service with required dependencies
   */
  initializeServices(aiService: AIService): void {
    this.aiService = aiService;
  }

  /**
   * Generate automatic summary for a document
   */
  async summarizeDocument(
    file: DriveFile,
    content: string,
    options: SummaryOptions = {}
  ): Promise<DocumentSummary> {
    try {
      // Check if we have a cached summary
      const cacheKey = `${file.id}-${file.modifiedTime}`;
      const cachedSummary = this.summaries.get(cacheKey);
      
      if (cachedSummary) {
        logger.debug('Returning cached document summary', {
          component: 'DocumentSummarizationService',
          fileId: file.id,
        });
        return cachedSummary;
      }

      // Validate content
      if (!content || content.trim().length === 0) {
        throw new Error('Document content is empty');
      }

      // Prepare AI prompt based on options
      const prompt = this.createSummaryPrompt(content, options);
      
      // Generate summary using AI service
      if (!this.aiService) {
        throw new Error('AI service not initialized');
      }

      const aiResponse = await this.aiService.generateResponse(prompt, {
        useCache: true,
        maxTokens: options.maxLength || this.MAX_SUMMARY_LENGTH,
      });

      // Parse AI response
      const summary = this.parseSummaryResponse(aiResponse.content, options);
      
      // Create document summary object
      const documentSummary: DocumentSummary = {
        fileId: file.id,
        fileName: file.name || 'Untitled',
        summary: summary.text,
        keyPoints: summary.keyPoints,
        entities: summary.entities,
        sentiment: summary.sentiment,
        wordCount: content.split(/\s+/).length,
        readingTime: Math.ceil(content.split(/\s+/).length / 200), // avg 200 words per minute
        generatedAt: new Date(),
      };

      // Cache the summary
      this.cacheSummary(cacheKey, documentSummary);
      
      logger.info('Document summary generated successfully', {
        component: 'DocumentSummarizationService',
        fileId: file.id,
        wordCount: documentSummary.wordCount,
        readingTime: documentSummary.readingTime,
      });

      return documentSummary;
    } catch (error) {
      logger.error('Error generating document summary', {
        component: 'DocumentSummarizationService',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error),
      });
      
      throw error;
    }
  }

  /**
   * Create AI prompt for document summarization
   */
  private createSummaryPrompt(content: string, options: SummaryOptions): string {
    const language = options.language || 'Ukrainian';
    const style = options.style || 'concise';
    
    let instruction = `Summarize the following document in ${language}. `;
    
    switch (style) {
      case 'detailed':
        instruction += 'Provide a comprehensive summary with all important details. ';
        break;
      case 'bullet-points':
        instruction += 'Present the summary as bullet points with key information. ';
        break;
      case 'concise':
      default:
        instruction += 'Provide a concise summary of the main points. ';
        break;
    }
    
    if (options.includeKeyPoints) {
      instruction += 'Include a list of key points. ';
    }
    
    if (options.includeEntities) {
      instruction += 'Identify important entities (names, organizations, dates, etc.). ';
    }
    
    instruction += `Summary should be no more than ${options.maxLength || this.MAX_SUMMARY_LENGTH} words. `;
    instruction += 'Also determine the overall sentiment (positive, negative, or neutral).\n\n';
    instruction += `Document content:\n${content.substring(0, 5000)}...`; // Limit content sent to AI
    
    return instruction;
  }

  /**
   * Parse AI response into structured summary
   */
  private parseSummaryResponse(
    response: string,
    options: SummaryOptions
  ): { text: string; keyPoints: string[]; entities: string[]; sentiment: 'positive' | 'negative' | 'neutral' } {
    // For now, we'll use a simple parsing approach
    // In a more advanced implementation, we could use structured output from the AI
    
    // Extract summary text (first part)
    const lines = response.split('\n').filter(line => line.trim() !== '');
    let summaryText = lines[0] || response.substring(0, 200);
    
    // Extract key points if requested
    let keyPoints: string[] = [];
    if (options.includeKeyPoints) {
      // Look for bullet points or numbered lists
      keyPoints = lines
        .filter(line => line.match(/^[\*\-\d]/) || line.length > 20)
        .map(line => line.replace(/^[\*\-\d\.\s]+/, '').trim())
        .slice(0, 5); // Limit to 5 key points
    }
    
    // Extract entities if requested
    let entities: string[] = [];
    if (options.includeEntities) {
      // Simple entity extraction (in a real implementation, this would be more sophisticated)
      const entityRegex = /\b[A-Z][a-z]+(?:\s[A-Z][a-z]+)*\b/g;
      entities = Array.from(new Set(response.match(entityRegex) || []))
        .filter(entity => entity.length > 2)
        .slice(0, 10); // Limit to 10 entities
    }
    
    // Determine sentiment (simple approach)
    let sentiment: 'positive' | 'negative' | 'neutral' = 'neutral';
    const positiveWords = ['good', 'great', 'excellent', 'positive', 'success', 'добре', 'чудово', 'успіх'];
    const negativeWords = ['bad', 'poor', 'negative', 'fail', 'problem', 'погано', 'проблема', 'невдача'];
    
    const lowerResponse = response.toLowerCase();
    const positiveCount = positiveWords.filter(word => lowerResponse.includes(word)).length;
    const negativeCount = negativeWords.filter(word => lowerResponse.includes(word)).length;
    
    if (positiveCount > negativeCount) {
      sentiment = 'positive';
    } else if (negativeCount > positiveCount) {
      sentiment = 'negative';
    }
    
    return {
      text: summaryText,
      keyPoints,
      entities,
      sentiment,
    };
  }

  /**
   * Cache summary with size management
   */
  private cacheSummary(key: string, summary: DocumentSummary): void {
    // Remove oldest entries if we're at capacity
    if (this.summaries.size >= this.MAX_CACHE_ENTRIES) {
      const firstKey = this.summaries.keys().next().value;
      if (firstKey) {
        this.summaries.delete(firstKey);
      }
    }
    
    this.summaries.set(key, summary);
  }

  /**
   * Get cached summary for a document
   */
  getDocumentSummary(fileId: string, modifiedTime?: string): DocumentSummary | null {
    const cacheKey = modifiedTime ? `${fileId}-${modifiedTime}` : fileId;
    return this.summaries.get(cacheKey) || null;
  }

  /**
   * Clear cached summaries for a specific document
   */
  clearDocumentSummary(fileId: string): boolean {
    let deleted = false;
    for (const key of this.summaries.keys()) {
      if (key.startsWith(fileId)) {
        this.summaries.delete(key);
        deleted = true;
      }
    }
    return deleted;
  }

  /**
   * Clear all cached summaries
   */
  clearAllSummaries(): void {
    this.summaries.clear();
  }

  /**
   * Get summary statistics
   */
  getSummaryStats(): {
    totalSummaries: number;
    cacheSize: number;
    averageWordCount: number;
  } {
    const summaries = Array.from(this.summaries.values());
    const totalWordCount = summaries.reduce((sum, summary) => sum + summary.wordCount, 0);
    const averageWordCount = summaries.length > 0 ? Math.round(totalWordCount / summaries.length) : 0;
    
    return {
      totalSummaries: summaries.length,
      cacheSize: this.summaries.size,
      averageWordCount,
    };
  }
}