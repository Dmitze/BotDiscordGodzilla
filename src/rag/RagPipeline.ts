import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from '@/services/AIService';
import { Retriever } from './Retriever';
import { Augmenter } from './Augmenter';
import type { RetrieverOptions, AugmentOptions, ContextChunk, GenerateWithContextOptions } from './types';
import logger from '@/utils/logger';

export interface RagAnswer {
  answer: string;
  chunks: ContextChunk[];
  citations?: { index: number; fileId: string; name: string; url?: string }[];
  provider: string;
  model?: string;
  tokens?: number;
}

export class RagPipeline {
  private readonly retriever: Retriever;
  private readonly augmenter: Augmenter;

  constructor(
    search: SearchIndex,
    private readonly ai: AIService,
    embeddings?: { embed: (text: string) => Promise<number[]> }
  ) {
    this.retriever = new Retriever(search, embeddings);
    this.augmenter = new Augmenter();
  }

  async answer(
    query: string,
    retrieverOpts: RetrieverOptions = {},
    augmentOpts: AugmentOptions = {},
    genOpts: GenerateWithContextOptions = {}
  ): Promise<RagAnswer> {
    const t0 = Date.now();
    const docs = await this.retriever.retrieve(query, retrieverOpts);
    const chunks = this.augmenter.buildContext(docs, augmentOpts);
    logger.info('RAG retrieve+augment complete', {
      service: 'RagPipeline',
      operation: 'answer',
      stage: 'context_ready',
      status: 'ok',
      docs: docs.length,
      chunks: chunks.length,
      elapsedMs: Date.now() - t0,
    });

    const args: Parameters<AIService['generateWithContext']>[0] = {
      prompt: query,
      contextChunks: chunks,
      citations: true,
      locale: 'uk',
    };
    if (typeof genOpts.model === 'string') (args as any).model = genOpts.model;
    if (typeof genOpts.maxTokens === 'number') (args as any).maxTokens = genOpts.maxTokens;
    if (typeof genOpts.temperature === 'number') (args as any).temperature = genOpts.temperature;
    const resp = await this.ai.generateWithContext(args);
    logger.info('RAG generation complete', {
      service: 'RagPipeline',
      operation: 'answer',
      stage: 'generate_done',
      status: 'ok',
      provider: resp.provider,
      model: resp.model,
      tokens: resp.tokens,
      chunks: chunks.length,
    });
    const out: any = {
      answer: resp.text,
      chunks,
      citations: resp.citations,
      provider: resp.provider,
      tokens: resp.tokens,
    };
    if (typeof resp.model === 'string') out.model = resp.model;
    return out as import('./RagPipeline').RagAnswer;
  }
}
