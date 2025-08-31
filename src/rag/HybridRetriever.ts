import type { SearchIndex, SearchQuery, SearchHit } from '@/search/SearchIndex';
import type { RetrieverOptions, RetrievedDoc } from './types';
import { cosineSimilarity } from '@/utils/cosine';
import { countTokens } from '@/utils/token';

/**
 * Enhanced Hybrid Retriever with Reranking
 * Implements advanced retrieval with hybrid search and intelligent reranking
 */
export class HybridRetriever {
  constructor(
    private readonly search: SearchIndex,
    private readonly embeddings?: { embed: (text: string) => Promise<number[]> }
  ) {}

  /**
   * Retrieve documents with hybrid search and reranking
   * @param queryText The query text
   * @param opts Retrieval options
   * @returns Retrieved documents with scores
   */
  async retrieve(queryText: string, opts: RetrieverOptions = {}): Promise<RetrievedDoc[]> {
    const mode = opts.mode ?? (process.env['EMBEDDINGS_ENABLE'] === 'true' ? 'hybrid' : 'fts');
    // Get more candidates for reranking (top 20 as specified in the plan)
    const initialK = Math.max(20, opts.k ?? Number(process.env['RETRIEVER_K'] ?? 6));
    const finalK = opts.k ?? Number(process.env['RETRIEVER_K'] ?? 6);

    const q: SearchQuery = {
      text: queryText,
      limit: initialK,
      ...(opts.filters ? { filters: opts.filters } as Pick<SearchQuery, 'filters'> : {}),
    };
    
    const { hits } = await this.search.search(q);

    // If no embeddings or mode is fts => rank/score based
    const useHybrid = (mode === 'hybrid') && this.embeddings;
    if (!useHybrid) {
      const docs: RetrievedDoc[] = hits.map((h: SearchHit, i) => ({
        ...h,
        fusedScore: this.normalizeFtsScore(h.score, i, hits.length),
      }));
      return docs.slice(0, finalK);
    }

    // Hybrid: compute query embedding and snippet embeddings, then fuse
    const alpha = this.clamp01(opts.alpha ?? Number(process.env['RETRIEVER_ALPHA'] ?? 0.5));
    const queryVec = await this.embeddings.embed(queryText);
    const docsWithEmbed: RetrievedDoc[] = [];
    
    for (let i = 0; i < hits.length; i++) {
      const h = hits[i]!; // noUncheckedIndexedAccess: ensured by bounds
      const snippet = h.snippet ?? '';
      const vec = await this.embeddings.embed(snippet);
      const sim = cosineSimilarity(queryVec, vec);
      const fts = this.normalizeFtsScore(h.score, i, hits.length);
      const fused = alpha * sim + (1 - alpha) * fts;
      docsWithEmbed.push({ ...h, embedScore: sim, fusedScore: fused });
    }

    // Sort by initial fused score
    docsWithEmbed.sort((a, b) => (b.fusedScore ?? 0) - (a.fusedScore ?? 0));
    
    // Apply reranking to top candidates
    const rerankedDocs = await this.rerank(queryText, docsWithEmbed.slice(0, initialK));
    
    return rerankedDocs.slice(0, finalK);
  }

  /**
   * Rerank documents using a more sophisticated approach
   * @param queryText The query text
   * @param docs Documents to rerank
   * @returns Reranked documents
   */
  private async rerank(queryText: string, docs: RetrievedDoc[]): Promise<RetrievedDoc[]> {
    if (!this.embeddings || docs.length === 0) {
      return docs;
    }

    // Get query embedding for reranking
    const queryVec = await this.embeddings.embed(queryText);
    
    // Enhanced reranking with multiple factors
    const reranked: RetrievedDoc[] = [];
    
    for (const doc of docs) {
      const snippet = doc.snippet ?? '';
      
      // Get document embedding
      const docVec = await this.embeddings.embed(snippet);
      
      // Calculate cosine similarity
      const cosineSim = cosineSimilarity(queryVec, docVec);
      
      // Calculate token overlap as additional signal
      const tokenOverlap = this.calculateTokenOverlap(queryText, snippet);
      
      // Length normalization (prefer medium-length documents)
      const tokenCount = countTokens(snippet);
      const lengthScore = this.normalizeLength(tokenCount);
      
      // Combine scores with weighted average
      // Weights can be adjusted based on experimentation
      const weights = {
        cosine: 0.5,
        fts: 0.3,
        overlap: 0.1,
        length: 0.1
      };
      
      const rerankedScore = 
        weights.cosine * cosineSim +
        weights.fts * (doc.fusedScore ?? 0) +
        weights.overlap * tokenOverlap +
        weights.length * lengthScore;
      
      reranked.push({
        ...doc,
        embedScore: cosineSim,
        fusedScore: rerankedScore,
        rerankMetadata: {
          tokenOverlap,
          lengthScore,
          weights
        }
      });
    }
    
    // Sort by reranked score
    return reranked.sort((a, b) => (b.fusedScore ?? 0) - (a.fusedScore ?? 0));
  }

  /**
   * Calculate token overlap between query and document
   * @param queryText Query text
   * @param docText Document text
   * @returns Overlap score between 0 and 1
   */
  private calculateTokenOverlap(queryText: string, docText: string): number {
    const queryTokens = new Set(queryText.toLowerCase().split(/\s+/));
    const docTokens = docText.toLowerCase().split(/\s+/);
    
    if (queryTokens.size === 0 || docTokens.length === 0) {
      return 0;
    }
    
    let overlapCount = 0;
    for (const token of docTokens) {
      if (queryTokens.has(token)) {
        overlapCount++;
      }
    }
    
    // Jaccard similarity
    return overlapCount / Math.max(queryTokens.size, docTokens.length);
  }

  /**
   * Normalize document length score (prefers medium-length documents)
   * @param tokenCount Number of tokens in document
   * @returns Normalized score between 0 and 1
   */
  private normalizeLength(tokenCount: number): number {
    // Ideal length is between 100-500 tokens
    const idealMin = 100;
    const idealMax = 500;
    
    if (tokenCount < idealMin) {
      // Too short - linear penalty
      return Math.max(0, tokenCount / idealMin);
    } else if (tokenCount > idealMax) {
      // Too long - inverse penalty
      return Math.max(0, idealMax / tokenCount);
    } else {
      // Ideal length - full score
      return 1;
    }
  }

  /**
   * Clamp value between 0 and 1
   * @param x Value to clamp
   * @returns Clamped value
   */
  private clamp01(x: number): number {
    return Math.max(0, Math.min(1, x));
  }

  /**
   * Normalize FTS score
   * @param score Raw FTS score
   * @param rankIndex Rank index
   * @param total Total number of results
   * @returns Normalized score between 0 and 1
   */
  private normalizeFtsScore(score: number | undefined, rankIndex: number, total: number): number {
    // Prefer FTS numeric score if present; else fall back to rank-based
    if (typeof score === 'number' && Number.isFinite(score)) {
      // Heuristic min-max over positive scores is unknown; use logistic-like squash
      const s = Math.max(0, score);
      return s / (1 + s); // maps [0, +inf) to (0,1)
    }
    const denom = Math.max(1, total - 1);
    return 1 - rankIndex / denom;
  }
}