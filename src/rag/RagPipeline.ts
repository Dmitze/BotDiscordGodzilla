import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from '@/services/AIService';
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

  constructor(private readonly search: SearchIndex, private readonly ai: AIService) {
    this.retriever = new Retriever(search);
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

    const resp = await this.ai.generateResponse(prompt, {
      model: genOpts.model,
      maxTokens: genOpts.maxTokens,
      temperature: genOpts.temperature,
      useCache: true,
    });

    return {
      answer: resp.content,
      chunks,
      provider: resp.provider,
      model: resp.model,
      tokens: resp.tokens,
    };
  }
}
