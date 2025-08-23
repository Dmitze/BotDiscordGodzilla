import { maskText } from '@/utils/pii';
import type { RetrievedDoc, ContextChunk, AugmentOptions } from './types';

function estimateTokens(s: string): number {
  // crude: ~4 chars per token
  return Math.ceil((s?.length ?? 0) / 4);
}

export class Augmenter {
  buildContext(docs: RetrievedDoc[], opts: AugmentOptions = {}): ContextChunk[] {
    const maxTokens = Math.max(128, Math.min(8000, opts.maxTokens ?? Number(process.env['RAG_MAX_CONTEXT_TOKENS'] ?? 1200)));
    const maxChunks = Math.max(1, Math.min(50, opts.maxChunks ?? 8));
    const mask = opts.maskPII ?? true;

    const sorted = [...docs].sort((a, b) => (b.fusedScore ?? 0) - (a.fusedScore ?? 0));

    const out: ContextChunk[] = [];
    let budget = maxTokens;
    for (const d of sorted) {
      if (!d.snippet) continue;
      const snippet = mask ? maskText(d.snippet) : d.snippet;
      const cost = estimateTokens(snippet) + estimateTokens(d.name);
      if (cost > budget) continue;
      out.push({ fileId: d.fileId, name: d.name, snippet, score: d.fusedScore, url: undefined });
      budget -= cost;
      if (out.length >= maxChunks) break;
      if (budget <= 64) break;
    }
    return out;
  }
}
