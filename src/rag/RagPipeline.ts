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

    // Fallback: unit tests may provide AI mock without generateWithContext
    const aiAny = this.ai as unknown as {
      generateWithContext?: (a: typeof args) => Promise<{ text: string; citations: any[]; provider: string; model?: string; tokens?: number }>;
      generateResponse?: (prompt: string, opts?: any) => Promise<{ content: string; provider: string; model?: string; tokens?: number }>;
    };
    let resp: { text: string; citations: any[]; provider: string; model?: string; tokens?: number };
    if (typeof aiAny.generateWithContext === 'function') {
      resp = await aiAny.generateWithContext(args);
    } else if (typeof aiAny.generateResponse === 'function') {
      // Build prompt similarly to AIService.generateWithContext
      const system = 'Ти — помічник, який відповідає стисло, українською, з посиланнями на джерела.';
      const sources = chunks
        .map((c, i) => `(${i + 1}) ${c.name} [${c.fileId}]\n${c.snippet}`)
        .join('\n\n');
      const fullPrompt = `${system}\n\nПитання:\n${query}\n\nКонтекст (релевантні уривки):\n${sources}\n\nВідповідь: наведи коротку відповідь та в кінці перелік джерел у форматі [1], [2], ... з короткими назвами.`;
      const aiResp = await aiAny.generateResponse(fullPrompt, {
        useCache: true,
        model: (args as any).model,
        maxTokens: (args as any).maxTokens,
        temperature: (args as any).temperature,
      });
      const cites = chunks.map((c, i) => ({ index: i + 1, fileId: c.fileId, name: c.name, url: (c as any).url }));
      const base: { text: string; citations: any[]; provider: string; model?: string; tokens?: number } = {
        text: aiResp.content,
        citations: cites,
        provider: aiResp.provider,
      };
      if (typeof aiResp.model === 'string') (base as any).model = aiResp.model;
      if (typeof aiResp.tokens === 'number') (base as any).tokens = aiResp.tokens;
      resp = base;
    } else {
      throw new Error('AI service does not implement generateWithContext or generateResponse');
    }
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
