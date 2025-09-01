import type { SearchFilters, SearchHit } from '@/search/SearchIndex';

export type RetrieverMode = 'fts' | 'hybrid';

export interface RetrieverOptions {
  mode?: RetrieverMode;
  k?: number; // top-k
  alpha?: number; // hybrid weighting [0..1] for FTS vs embeddings
  filters?: SearchFilters;
}

export interface RetrievedDoc extends SearchHit {
  // normalized score after fusion (0..1)
  fusedScore?: number;
  // embedding score if available (cosine similarity 0..1)
  embedScore?: number;
}

export interface ContextChunk {
  fileId: string;
  name: string;
  snippet: string;
  score?: number;
  url?: string;
}

export interface AugmentOptions {
  maxTokens?: number; // budget for context
  maxChunks?: number; // safety cap
  maskPII?: boolean;
}

export interface GenerateWithContextOptions {
  temperature?: number;
  maxTokens?: number;
  model?: string;
}
