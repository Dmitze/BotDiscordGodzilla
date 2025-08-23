import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import { RagPipeline, type RagAnswer } from '@/rag/RagPipeline';
import type { RetrieverOptions, AugmentOptions, GenerateWithContextOptions } from '@/rag/types';

export class RagService {
  private pipeline: RagPipeline;

  constructor(
    searchIndex: SearchIndex,
    ai: AIService,
    embeddings?: { embed: (text: string) => Promise<number[]> }
  ) {
    this.pipeline = new RagPipeline(searchIndex, ai, embeddings);
  }

  async answer(
    query: string,
    retriever: RetrieverOptions = {},
    augment: AugmentOptions = {},
    generate: GenerateWithContextOptions = {}
  ): Promise<RagAnswer> {
    return this.pipeline.answer(query, retriever, augment, generate);
  }
}
