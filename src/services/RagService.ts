import type { SearchIndex, SearchHit, SearchQuery } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import { RagPipeline, type RagAnswer } from '@/rag/RagPipeline';
import type { RetrieverOptions, AugmentOptions, GenerateWithContextOptions } from '@/rag/types';
import logger from '@/utils/logger';

type CacheValue = RagAnswer & { __expiresAt?: number };

export class RagService {
  private pipeline: RagPipeline;
  // Simple LRU/TTL cache
  private cache = new Map<string, CacheValue>();
  private readonly maxSize: number;
  private readonly defaultTTLms: number;
  private externalCache?: { get<T = unknown>(key: string): Promise<T | null>; set<T = unknown>(key: string, value: T, ttlSec?: number): Promise<unknown> };
  // Metrics
  private cacheHits = 0;
  private cacheMisses = 0;
  private readonly debug = process.env['RAG_DEBUG'] === '1' || process.env['RAG_DEBUG'] === 'true';

  constructor(
    private readonly searchIndex: SearchIndex,
    ai: AIService,
    embeddings?: { embed: (text: string) => Promise<number[]> },
    opts?: { maxSize?: number; ttlSec?: number; cache?: { get<T = unknown>(key: string): Promise<T | null>; set<T = unknown>(key: string, value: T, ttlSec?: number): Promise<unknown> } }
  ) {
    this.pipeline = new RagPipeline(searchIndex, ai, embeddings);
    this.maxSize = Math.max(16, opts?.maxSize ?? 256);
    this.defaultTTLms = Math.max(1, (opts?.ttlSec ?? 900)) * 1000; // 15m by default
    if (opts?.cache) this.externalCache = opts.cache;
  }

  private makeKey(
    model: string | undefined,
    query: string,
    filters: RetrieverOptions['filters'] | undefined,
    docIds: string[],
    maxModified: number | undefined
  ): string {
    const base = JSON.stringify({ model: model ?? 'default', query, filters: filters ?? null });
    const idPart = docIds.sort().join(',');
    const ver = maxModified ?? 0;
    return `${base}::${idPart}::m${ver}`;
  }

  private async setCache(key: string, value: RagAnswer, ttlMs?: number) {
    const ttlSec = Math.ceil((ttlMs ?? this.defaultTTLms) / 1000);
    if (this.externalCache) {
      try {
        await this.externalCache.set(key, value, ttlSec);
        return;
      } catch {
        // fall back to local if external fails
      }
    }
    const expiresAt = Date.now() + (ttlMs ?? this.defaultTTLms);
    if (this.cache.has(key)) this.cache.delete(key);
    this.cache.set(key, { ...value, __expiresAt: expiresAt });
    if (this.cache.size > this.maxSize) {
      const iter = this.cache.keys().next();
      const firstKey = iter.done ? undefined : iter.value;
      if (firstKey !== undefined) this.cache.delete(firstKey);
    }
  }

  private async getCache(key: string): Promise<RagAnswer | null> {
    if (this.externalCache) {
      try {
        const external = await this.externalCache.get<RagAnswer>(key);
        if (external) {
          this.cacheHits++;
          if (this.debug) logger.debug('RagService external cache hit', { key });
          return external;
        }
      } catch {
        // ignore and try local
      }
    }
    const v = this.cache.get(key);
    if (!v) return null;
    if (v.__expiresAt && Date.now() > v.__expiresAt) {
      this.cache.delete(key);
      return null;
    }
    this.cache.delete(key);
    const { __expiresAt: _exp, ...rest } = v;
    this.cache.set(key, v);
    const pure: RagAnswer = {
      provider: rest.provider,
      answer: rest.answer,
      chunks: rest.chunks,
    };
    if (typeof rest.model === 'string') pure.model = rest.model;
    if (typeof rest.tokens === 'number') pure.tokens = rest.tokens;
    if (Array.isArray(rest.citations)) pure.citations = rest.citations;
    this.cacheHits++;
    if (this.debug) logger.debug('RagService local cache hit', { key });
    return pure;
  }

  async answer(
    query: string,
    retriever: RetrieverOptions = {},
    augment: AugmentOptions = {},
    generate: GenerateWithContextOptions = {}
  ): Promise<RagAnswer> {
    // Always retrieve first to compute cache key with fileIds and modifiedTime
    const baseQuery: SearchQuery = {
      text: query,
      limit: Math.max(1, retriever.k ?? 5),
      offset: 0,
    };
    if (retriever.filters) baseQuery.filters = retriever.filters;
    const retrieved = await this.searchIndex.search(baseQuery);
    const hits: SearchHit[] = retrieved.hits;
    const docIds = hits.map(h => h.fileId);
    const maxModified = hits.reduce<number | undefined>((acc, h) =>
      h.modifiedTime ? Math.max(acc ?? h.modifiedTime, h.modifiedTime) : acc,
    undefined);

    const key = this.makeKey(generate.model, query, retriever.filters, docIds, maxModified);
    const cached = await this.getCache(key);
    if (cached) {
      logger.info('RagService cache hit', {
        service: 'RagService',
        operation: 'answer',
        status: 'ok',
        cache: 'hit',
        size: this.cache.size,
      });
      return cached;
    }

    logger.info('RagService cache miss', {
      service: 'RagService',
      operation: 'answer',
      status: 'ok',
      cache: 'miss',
      size: this.cache.size,
    });
    this.cacheMisses++;

    // Fallback to full pipeline (it will re-run retrieval/augment with embeddings/hybrid if configured)
    const res = await this.pipeline.answer(query, retriever, augment, generate);
    await this.setCache(key, res);
    return res;
  }

  getMetrics(): { cacheHits: number; cacheMisses: number; size: number; capacity: number } {
    return {
      cacheHits: this.cacheHits,
      cacheMisses: this.cacheMisses,
      size: this.cache.size,
      capacity: this.maxSize,
    };
  }
}
