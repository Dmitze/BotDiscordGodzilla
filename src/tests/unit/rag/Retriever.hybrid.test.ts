import { Retriever } from '@/rag/Retriever';
import type { SearchIndex, SearchQuery } from '@/search/SearchIndex';

class MemSearch implements SearchIndex {
  private docs: { id: string; name: string; text: string }[] = [];
  async upsert(doc: any): Promise<void> {
    const i = this.docs.findIndex(d => d.id === doc.fileId);
    const entry = { id: doc.fileId, name: doc.name, text: doc.text };
    if (i >= 0) this.docs[i] = entry; else this.docs.push(entry);
  }
  async search(q: SearchQuery): Promise<{ hits: any[]; total: number }> {
    const term = (q.text || '').toLowerCase();
    const hits = this.docs
      .filter(d => d.text.toLowerCase().includes(term))
      .map((d, i) => ({ fileId: d.id, name: d.name, snippet: d.text.slice(0, 60), score: i + 1 }));
    return { hits: hits.slice(0, q.limit ?? 10), total: hits.length };
  }
}

describe('Retriever: hybrid ranking', () => {
  const oldEnv = { ...process.env };
  afterAll(() => { process.env = oldEnv; });

  it('fuses FTS and embeddings (mock) scores', async () => {
    process.env['EMBEDDINGS_ENABLE'] = 'true';
    process.env['RETRIEVER_ALPHA'] = '0.7';

    const idx = new MemSearch();
    await idx.upsert({ fileId: '1', name: 'Doc1', text: 'cats and dogs' });
    await idx.upsert({ fileId: '2', name: 'Doc2', text: 'only cats here' });

    const embeddings = { embed: async (text: string) => {
      // simple deterministic small vectors for test
      const arr = Array.from(text).map(c => ((c.charCodeAt(0) % 5) - 2) / 2);
      const dim = Math.max(8, arr.length);
      const out = new Array<number>(dim).fill(0);
      for (let i = 0; i < arr.length; i++) out[i] = arr[i];
      return out;
    }};

    const retr = new Retriever(idx as any, embeddings);
    const res = await retr.retrieve('cats', { mode: 'hybrid', k: 2 });

    expect(res.length).toBe(2);
    // Has fusedScore present
    expect(typeof res[0].fusedScore).toBe('number');
    // Ordering should be stable and based on fused score
    expect((res[0].fusedScore ?? 0) >= (res[1].fusedScore ?? 0)).toBe(true);
  });
});
