import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import { RagPipeline } from '@/rag/RagPipeline';
import type { RetrieverOptions, AugmentOptions, GenerateWithContextOptions } from '@/rag/types';

export class RagService {
  private pipeline: RagPipeline;

  constructor(searchIndex: SearchIndex, ai: AIService) {
    this.pipeline = new RagPipeline(searchIndex, ai);
  }

  async answer(
    query: string,
    retriever: RetrieverOptions = {},
    augment: AugmentOptions = {},
    generate: GenerateWithContextOptions = {}
  ) {
    return this.pipeline.answer(query, retriever, augment, generate);
  }
}
