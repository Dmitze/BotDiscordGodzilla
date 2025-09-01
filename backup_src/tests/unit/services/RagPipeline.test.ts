import { RagPipeline } from '@/rag/RagPipeline';
import type { SearchIndex, SearchHit } from '@/search/SearchIndex';
import type { AIService } from '@/services/AIService';

// Lightweight AIResponse type to avoid deep imports
interface MockAIResponse {
  content: string;
  provider: string;
  model?: string;
  tokens?: number;
}

describe('RagPipeline', () => {
  const makeSearchIndex = (hits: SearchHit[]): SearchIndex => ({
    async upsert() { /* no-op */ },
    async search() { return { hits, total: hits.length }; },
    async getDiff() { return { latestHash: 'x' }; },
  });

  const makeAI = (fn?: (prompt: string) => MockAIResponse): AIService => ({
    // @ts-expect-error partial mock is enough for tests
    generateResponse: async (prompt: string, _opts?: any) => {
      const base: MockAIResponse = fn ? fn(prompt) : { content: 'Відповідь (мок)', provider: 'mock', model: 'mock-model', tokens: 42 };
      return base as any;
    },
  });

  it('builds context with PII masked and returns Ukrainian answer with chunks', async () => {
    const hits: SearchHit[] = [
      {
        fileId: 'f1',
        name: 'Документ 1',
        mimeType: 'text/plain',
        ownerEmail: 'owner@example.com',
        modifiedTime: Date.now(),
        snippet: 'Email користувача: john.doe@example.com, телефон: +380 67 123 45 67',
        contentHash: 'h1',
        textLen: 120,
        score: 0.9,
      },
      {
        fileId: 'f2',
        name: 'Документ 2',
        mimeType: 'application/pdf',
        ownerEmail: 'boss@example.com',
        modifiedTime: Date.now(),
        snippet: 'Контакт: manager@company.ua та номер 093-111-22-33',
        contentHash: 'h2',
        textLen: 200,
        score: 0.8,
      },
    ];

    // Ensure security masking is enabled
    process.env['RAG_MAX_CONTEXT_TOKENS'] = '1200';

    const search = makeSearchIndex(hits);
    const ai = makeAI((prompt) => {
      // Basic sanity that prompt is Ukrainian-ish and contains context header
      expect(prompt).toMatch(/Питання:/);
      expect(prompt).toMatch(/Контекст \(релевантні уривки\)/);
      return { content: 'Коротка відповідь з посиланнями [1], [2].', provider: 'mock', model: 'mock', tokens: 10 };
    });

    const pipeline = new RagPipeline(search, ai as any);
    const res = await pipeline.answer('Що відомо про контакти?', { k: 5 }, { maskPII: true, maxChunks: 5 }, { maxTokens: 256 });

    expect(res.answer).toMatch(/Коротка відповідь/);
    expect(res.chunks.length).toBeGreaterThan(0);

    // PII masked: email local part shortened and phone largely masked by utils
    const joined = res.chunks.map((c) => c.snippet).join('\n');
    expect(joined).not.toMatch(/john\.doe@example\.com/);
    expect(joined).not.toMatch(/093-111-22-33/);

    expect(res.provider).toBe('mock');
    expect(res.model).toBe('mock');
  });

  it('respects token and chunk budgets', async () => {
    const hugeText = 'a'.repeat(5000);
    const hits: SearchHit[] = Array.from({ length: 10 }).map((_, i) => ({
      fileId: `f${i}`,
      name: `Документ ${i}`,
      contentHash: `h${i}`,
      textLen: hugeText.length,
      snippet: hugeText,
      score: 1 - i * 0.01,
    }));

    const search = makeSearchIndex(hits as any);
    const ai = makeAI();

    const pipeline = new RagPipeline(search, ai as any);
    const res = await pipeline.answer('Запит', { k: 10 }, { maxTokens: 200, maxChunks: 3, maskPII: true }, { maxTokens: 64 });

    expect(res.chunks.length).toBeLessThanOrEqual(3);
  });
});
