import { RagService } from '@/services/RagService';
import type { SearchIndex, SearchHit } from '@/search/SearchIndex';
import type { AIService } from '@/services/AIService';

describe('RagService', () => {
  const makeSearchIndex = (hits: SearchHit[]): SearchIndex => ({
    async upsert() { /* no-op */ },
    async search() { return { hits, total: hits.length }; },
    async getDiff() { return { latestHash: 'x' }; },
  });

  const makeAI = (): AIService => ({
    // @ts-expect-error partial mock is enough for tests
    generateResponse: async (prompt: string) => {
      // ensure prompt contains Ukrainian system text markers
      expect(prompt).toMatch(/Питання:/);
      return {
        content: 'Відповідь з джерелами [1].',
        provider: 'mock',
        model: 'mock',
        tokens: 5,
      } as any;
    },
  });

  it('answers via pipeline and returns chunks with citations context', async () => {
    const hits: SearchHit[] = [
      {
        fileId: 'file-123',
        name: 'Прикладовий документ',
        snippet: 'Текст з email test@example.com і номером 050-222-33-44',
        contentHash: 'hash',
        textLen: 100,
        score: 0.95,
      },
    ];

    const svc = new RagService(makeSearchIndex(hits), makeAI() as any);
    const res = await svc.answer('Поясни контакти', { k: 3 }, { maskPII: true }, { maxTokens: 128 });

    expect(res.answer).toContain('Відповідь');
    expect(res.chunks.length).toBe(1);
    // PII masked in chunk snippet
    expect(res.chunks[0].snippet).not.toMatch(/test@example.com/);
  });
});
