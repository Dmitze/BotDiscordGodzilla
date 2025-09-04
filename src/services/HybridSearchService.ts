import type { SearchIndex, SearchQuery, SearchHit } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import logger from '@/utils/logger';

export interface HybridSearchResult extends SearchHit {
  vectorScore?: number;
  textScore?: number;
  combinedScore?: number; // Made combinedScore optional to fix exactOptionalPropertyTypes issue
}

export interface HybridSearchOptions {
  limit?: number;
  minScore?: number;
  filters?: Record<string, any>;
  vectorWeight?: number; // Weight for vector search score (0-1)
  textWeight?: number;   // Weight for text search score (0-1)
  useCache?: boolean;
}

export class HybridSearchService {
  constructor(
    private readonly searchIndex: SearchIndex,
    private readonly aiService: AIService,
    private readonly embeddingsService?: { embed: (text: string) => Promise<number[]> }
  ) {}

  /**
   * Perform hybrid search combining vector search and full-text search
   */
  async search(
    query: string,
    options: HybridSearchOptions = {}
  ): Promise<HybridSearchResult[]> {
    try {
      const limit = options.limit || 20;
      const vectorWeight = options.vectorWeight ?? 0.7;
      const textWeight = options.textWeight ?? 0.3;

      // Validate weights
      if (vectorWeight < 0 || vectorWeight > 1 || textWeight < 0 || textWeight > 1) {
        throw new Error('Weights must be between 0 and 1');
      }

      if (Math.abs((vectorWeight + textWeight) - 1) > 0.001) {
        throw new Error('Vector weight and text weight must sum to 1');
      }

      logger.info('Starting hybrid search', {
        component: 'HybridSearchService',
        query,
        limit,
        vectorWeight,
        textWeight
      });

      // Perform vector search if embeddings service is available
      let vectorResults: SearchHit[] = [];
      if (this.embeddingsService) {
        vectorResults = await this.performVectorSearch(query, limit * 2); // Get more results for reranking
        logger.debug('Vector search completed', {
          component: 'HybridSearchService',
          resultCount: vectorResults.length
        });
      }

      // Perform full-text search
      const textResults = await this.performTextSearch(query, limit * 2); // Get more results for reranking
      logger.debug('Text search completed', {
        component: 'HybridSearchService',
        resultCount: textResults.length
      });

      // Combine and rerank results
      const combinedResults = this.combineAndRerank(
        vectorResults,
        textResults,
        vectorWeight,
        textWeight,
        limit
      );

      logger.info('Hybrid search completed', {
        component: 'HybridSearchService',
        finalResultCount: combinedResults.length
      });

      return combinedResults;
    } catch (error) {
      logger.error('Error during hybrid search', {
        component: 'HybridSearchService',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined
      });
      throw error;
    }
  }

  /**
   * Perform vector search using embeddings
   */
  private async performVectorSearch(query: string, limit: number): Promise<SearchHit[]> {
    if (!this.embeddingsService) {
      return [];
    }

    try {
      // Generate query embedding
      const queryEmbedding = await this.embeddingsService.embed(query);
      
      // For now, we'll simulate vector search by using the existing search
      // In a real implementation, this would query a vector database
      const searchQuery: SearchQuery = {
        text: query,
        limit
      };

      const result = await this.searchIndex.search(searchQuery);
      return result.hits.map(hit => ({
        ...hit,
        score: hit.score ? hit.score * 0.8 : 0.5 // Simulate vector score
      }));
    } catch (error) {
      logger.warn('Vector search failed, falling back to text search only', {
        component: 'HybridSearchService',
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * Perform full-text search
   */
  private async performTextSearch(query: string, limit: number): Promise<SearchHit[]> {
    try {
      const searchQuery: SearchQuery = {
        text: query,
        limit
      };

      const result = await this.searchIndex.search(searchQuery);
      return result.hits;
    } catch (error) {
      logger.error('Text search failed', {
        component: 'HybridSearchService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Combine vector and text search results and rerank them
   */
  private combineAndRerank(
    vectorResults: SearchHit[],
    textResults: SearchHit[],
    vectorWeight: number,
    textWeight: number,
    limit: number
  ): HybridSearchResult[] {
    // Create a map to store combined results
    const combinedMap = new Map<string, HybridSearchResult>();

    // Process vector results
    for (const result of vectorResults) {
      const existing = combinedMap.get(result.fileId);
      if (existing) {
        existing.vectorScore = result.score ?? 0;
        existing.combinedScore = this.calculateCombinedScore(
          result.score,
          existing.textScore,
          vectorWeight,
          textWeight
        );
      } else {
        combinedMap.set(result.fileId, {
          ...result,
          vectorScore: result.score ?? 0,
          combinedScore: this.calculateCombinedScore(
            result.score,
            undefined,
            vectorWeight,
            textWeight
          )
        });
      }
    }

    // Process text results
    for (const result of textResults) {
      const existing = combinedMap.get(result.fileId);
      if (existing) {
        existing.textScore = result.score ?? 0;
        existing.combinedScore = this.calculateCombinedScore(
          existing.vectorScore,
          result.score,
          vectorWeight,
          textWeight
        );
      } else {
        combinedMap.set(result.fileId, {
          ...result,
          textScore: result.score ?? 0,
          combinedScore: this.calculateCombinedScore(
            undefined,
            result.score,
            vectorWeight,
            textWeight
          )
        });
      }
    }

    // Convert map to array and sort by combined score
    const combinedArray = Array.from(combinedMap.values());
    combinedArray.sort((a, b) => (b.combinedScore ?? 0) - (a.combinedScore ?? 0));

    // Return top results
    return combinedArray.slice(0, limit);
  }

  /**
   * Calculate combined score based on vector and text scores
   */
  private calculateCombinedScore(
    vectorScore: number | undefined,
    textScore: number | undefined,
    vectorWeight: number,
    textWeight: number
  ): number {
    // Normalize scores (they come as BM25 scores where lower is better)
    // For BM25, we invert and normalize to 0-1 range
    const normalizedVectorScore = vectorScore !== undefined ? 
      Math.max(0, Math.min(1, 1 - (vectorScore / 100))) : 0;
    
    const normalizedTextScore = textScore !== undefined ? 
      Math.max(0, Math.min(1, 1 - (textScore / 100))) : 0;

    // Calculate weighted combination
    return (normalizedVectorScore * vectorWeight) + (normalizedTextScore * textWeight);
  }
}