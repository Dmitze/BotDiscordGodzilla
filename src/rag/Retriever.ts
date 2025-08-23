import type { SearchIndex, SearchQuery, SearchHit } from '@/search/SearchIndex';
import type { RetrieverOptions, RetrievedDoc } from './types';

export class Retriever {
  constructor(private readonly search: SearchIndex) {}

  async retrieve(queryText: string, opts: RetrieverOptions = {}): Promise<RetrievedDoc[]> {
    const mode = opts.mode ?? (process.env['EMBEDDINGS_ENABLE'] === 'true' ? 'hybrid' : 'fts');
    const k = Math.max(1, Math.min(100, opts.k ?? Number(process.env['RETRIEVER_K'] ?? 6)));

    const q: SearchQuery = { text: queryText, filters: opts.filters, limit: k };
    const { hits } = await this.search.search(q);

    // For MVP: FTS-only scoring passthrough. If embeddings enabled, fusion can be added later.
    const docs: RetrievedDoc[] = hits.map((h: SearchHit, i) => ({ ...h, fusedScore: h.score ?? (1 - i / hits.length) }));
    return docs.slice(0, k);
  }
}
