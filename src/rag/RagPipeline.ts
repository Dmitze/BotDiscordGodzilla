import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from '@/services/AIService';
import type { AIRequestOptions } from '@/types';
import { Retriever } from './Retriever';
import { Augmenter } from './Augmenter';
import type { RetrieverOptions, AugmentOptions, ContextChunk, GenerateWithContextOptions } from './types';

export interface RagAnswer {
  answer: string;
  chunks: ContextChunk[];
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
    const docs = await this.retriever.retrieve(query, retrieverOpts);
    const chunks = this.augmenter.buildContext(docs, augmentOpts);

    const sources = chunks
      .map((c, i) => `(${i + 1}) ${c.name} [${c.fileId}]\n${c.snippet}`)
      .join('\n\n');

    const system = 'Ти — помічник, який відповідає стисло, українською, з посиланнями на джерела.';
    const prompt = `${system}\n\nПитання:\n${query}\n\nКонтекст (релевантні уривки):\n${sources}\n\nВідповідь: наведи коротку відповідь та в кінці перелік джерел у форматі [1], [2], ... з короткими назвами.`;

    const req: AIRequestOptions = { useCache: true };
    if (typeof genOpts.model === 'string') req.model = genOpts.model;
    if (typeof genOpts.maxTokens === 'number') req.maxTokens = genOpts.maxTokens;
    if (typeof genOpts.temperature === 'number') req.temperature = genOpts.temperature;

    const resp = await this.ai.generateResponse(prompt, req);

    return {
      answer: resp.content,
      chunks,
      provider: resp.provider,
      model: resp.model,
      tokens: resp.tokens,
    };
  }
}
