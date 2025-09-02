import type { AIService } from '@/services/AIService';
import type { RetrievedDoc } from './types';
import logger from '@/utils/logger';

export interface RerankerOptions {
  model?: string;
  limit?: number;
  temperature?: number;
}

export class Reranker {
  constructor(private readonly ai: AIService) {}

  /**
   * Rerank documents based on their relevance to the query
   * Uses AI to evaluate semantic similarity between query and documents
   */
  async rerank(query: string, docs: RetrievedDoc[], options: RerankerOptions = {}): Promise<RetrievedDoc[]> {
    try {
      const limit = options.limit ?? docs.length;
      
      // If we have no documents or only one, no need to rerank
      if (docs.length <= 1) {
        return docs.slice(0, limit);
      }

      logger.info('Starting document reranking', {
        component: 'Reranker',
        queryLength: query.length,
        documentCount: docs.length,
        limit
      });

      // Create pairs of query and document content for reranking
      const docPairs = docs.map(doc => ({
        doc,
        content: this.getDocumentContent(doc)
      }));

      // Generate prompts for reranking
      const prompts = docPairs.map(pair => 
        this.createRerankingPrompt(query, pair.content)
      );

      // Score each document using AI
      const scores = await Promise.all(
        prompts.map((prompt, index) => 
          this.scoreDocumentRelevance(prompt, options)
            .then(score => ({ index, score }))
            .catch(err => {
              logger.warn('Failed to score document relevance', {
                component: 'Reranker',
                error: err instanceof Error ? err.message : String(err),
                documentIndex: index
              });
              // Return a neutral score if scoring fails
              return { index, score: 0.5 };
            })
        )
      );

      // Apply scores to documents
      const scoredDocs = scores.map(({ index, score }) => ({
        ...docs[index],
        rerankScore: score,
        fusedScore: this.combineScores(docs[index].fusedScore, score)
      }));

      // Sort by rerank score (higher is better)
      scoredDocs.sort((a, b) => (b.rerankScore ?? 0) - (a.rerankScore ?? 0));

      logger.info('Document reranking completed', {
        component: 'Reranker',
        finalDocumentCount: scoredDocs.length
      });

      return scoredDocs.slice(0, limit);
    } catch (error) {
      logger.error('Error during document reranking', {
        component: 'Reranker',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined
      });
      // If reranking fails, return original documents
      return docs.slice(0, limit);
    }
  }

  /**
   * Extract content from a document for reranking
   */
  private getDocumentContent(doc: RetrievedDoc): string {
    // Prefer snippet if available, otherwise use name and available metadata
    if (doc.snippet) {
      return doc.snippet;
    }
    
    // Build content from available fields
    let content = doc.name || '';
    
    if (doc.mimeType) {
      content += ` [${doc.mimeType}]`;
    }
    
    return content.trim();
  }

  /**
   * Create a prompt for evaluating document relevance
   */
  private createRerankingPrompt(query: string, documentContent: string): string {
    return `Оціни релевантність наступного документа до запитання користувача.
Запит: "${query}"
Документ: "${documentContent}"

Оцініть релевантність від 0 до 1, де:
0 - зовсім не релевантно
1 - повністю релевантно

Поверни лише число від 0 до 1. Наприклад: 0.85`;
  }

  /**
   * Score document relevance using AI
   */
  private async scoreDocumentRelevance(prompt: string, options: RerankerOptions): Promise<number> {
    try {
      // Use AI service to generate a relevance score
      const response = await this.ai.generateResponse(prompt, {
        model: options.model,
        maxTokens: 10,
        temperature: options.temperature || 0.1, // Low temperature for consistent scoring
        useCache: false // Don't cache reranking scores
      });

      // Extract numeric score from response
      const scoreText = response.content.trim();
      const score = parseFloat(scoreText);
      
      // Validate score is between 0 and 1
      if (isNaN(score) || score < 0 || score > 1) {
        logger.warn('Invalid relevance score, using default', {
          component: 'Reranker',
          scoreText,
          parsedScore: score
        });
        return 0.5; // Neutral score
      }
      
      return score;
    } catch (error) {
      logger.error('Error scoring document relevance', {
        component: 'Reranker',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Combine original fused score with rerank score
   */
  private combineScores(originalScore: number | undefined, rerankScore: number): number {
    // If we don't have an original score, use the rerank score
    if (originalScore === undefined) {
      return rerankScore;
    }
    
    // Combine scores with weighted average (70% rerank, 30% original)
    return (rerankScore * 0.7) + (originalScore * 0.3);
  }
}