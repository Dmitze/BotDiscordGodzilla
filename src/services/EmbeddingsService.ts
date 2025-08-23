import type { BotConfig } from '@/types';
import OpenAI from 'openai';

export interface EmbeddingsProvider {
  embed(text: string): Promise<number[]>;
}

export class EmbeddingsService implements EmbeddingsProvider {
  private provider: 'openai' | 'mock';
  private model: string;
  private client?: OpenAI;

  constructor(private readonly config: BotConfig) {
    this.provider = (process.env['EMBEDDINGS_PROVIDER'] as any) || 'mock';
    this.model = process.env['EMBEDDINGS_MODEL'] || 'text-embedding-3-small';

    if (this.provider === 'openai') {
      const apiKey = process.env['OPENAI_API_KEY'] || this.config?.ai?.openai?.apiKey;
      if (!apiKey) {
        // Fallback to mock if no key
        this.provider = 'mock';
      } else {
        this.client = new OpenAI({ apiKey });
      }
    }
  }

  async embed(text: string): Promise<number[]> {
    if (!text) return [];

    if (this.provider === 'openai' && this.client) {
      const resp = await this.client.embeddings.create({ model: this.model, input: text });
      const vec = resp.data?.[0]?.embedding || [];
      // Ensure it's a copy of number[]
      return Array.from(vec);
    }

    // Mock deterministic embedding: hash -> PRNG -> fixed-size vector
    return this.mockEmbed(text, 384);
  }

  private mockEmbed(text: string, dim: number): number[] {
    let h = 2166136261 >>> 0; // FNV-1a 32-bit
    for (let i = 0; i < text.length; i++) {
      h ^= text.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    // xorshift to generate pseudo-random but deterministic
    let x = h || 123456789;
    const out = new Array<number>(dim);
    let norm = 0;
    for (let i = 0; i < dim; i++) {
      x ^= x << 13; x ^= x >>> 17; x ^= x << 5; // xorshift32
      const v = (x >>> 0) / 0xffffffff; // [0,1)
      const centered = (v - 0.5) * 2;    // (-1,1)
      out[i] = centered;
      norm += centered * centered;
    }
    // Normalize to unit length
    norm = Math.sqrt(norm) || 1;
    for (let i = 0; i < dim; i++) {
      const v = out[i];
      out[i] = (v ?? 0) / norm;
    }
    return out;
  }
}
