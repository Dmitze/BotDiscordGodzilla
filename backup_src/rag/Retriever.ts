import type { SearchIndex, SearchQuery, SearchHit } from '@/search/SearchIndex';
import type { RetrieverOptions, RetrievedDoc } from './types';
import { cosineSimilarity } from '@/utils/cosine';

export class Retriever {
  constructor(
    private readonly search: SearchIndex,
    private readonly embeddings?: { embed: (text: string) => Promise<number[]> }
  ) {}

  async retrieve(queryText: string, opts: RetrieverOptions = {}): Promise<RetrievedDoc[]> {
    const mode = opts.mode ?? (process.env['EMBEDDINGS_ENABLE'] === 'true' ? 'hybrid' : 'fts');
    const k = Math.max(1, Math.min(100, opts.k ?? Number(process.env['RETRIEVER_K'] ?? 6)));

    const q: SearchQuery = {
      text: queryText,
      limit: k,
      ...(opts.filters ? { filters: opts.filters } as Pick<SearchQuery, 'filters'> : {}),
    };
    const { hits } = await this.search.search(q);

    // If no embeddings or mode is fts => rank/score based
    const useHybrid = (mode === 'hybrid') && this.embeddings;
    if (!useHybrid) {
      const docs: RetrievedDoc[] = hits.map((h: SearchHit, i) => ({
        ...h,
        fusedScore: normalizeFtsScore(h.score, i, hits.length),
      }));
      return docs.slice(0, k);
    }

    // Hybrid: compute query embedding and snippet embeddings, then fuse
    const alpha = clamp01(opts.alpha ?? Number(process.env['RETRIEVER_ALPHA'] ?? 0.5));
    const queryVec = await this.embeddings.embed(queryText);
    const docsWithEmbed: RetrievedDoc[] = [];
    for (let i = 0; i < hits.length; i++) {
      const h = hits[i]!; // noUncheckedIndexedAccess: ensured by bounds
      const snippet = h.snippet ?? '';
      const vec = await this.embeddings.embed(snippet);
      const sim = cosineSimilarity(queryVec, vec);
      const fts = normalizeFtsScore(h.score, i, hits.length);
      const fused = alpha * sim + (1 - alpha) * fts;
      docsWithEmbed.push({ ...h, embedScore: sim, fusedScore: fused });
    }

    docsWithEmbed.sort((a, b) => (b.fusedScore ?? 0) - (a.fusedScore ?? 0));
    return docsWithEmbed.slice(0, k);
  }
}

function clamp01(x: number): number {
  return Math.max(0, Math.min(1, x));
}

function normalizeFtsScore(score: number | undefined, rankIndex: number, total: number): number {
  // Prefer FTS numeric score if present; else fall back to rank-based
  if (typeof score === 'number' && Number.isFinite(score)) {
    // Heuristic min-max over positive scores is unknown; use logistic-like squash
    const s = Math.max(0, score);
    return s / (1 + s); // maps [0, +inf) to (0,1)
  }
  const denom = Math.max(1, total - 1);
  return 1 - rankIndex / denom;
}
