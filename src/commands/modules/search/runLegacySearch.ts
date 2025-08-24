import type logger from '@/utils/logger';
import type { SearchParams } from '@/types';
import type { SearchResult } from './types';

export type PerformSearchFn = (params: SearchParams) => Promise<SearchResult>;

export interface CacheLike<V> {
  get(key: string): V | undefined;
  set(key: string, val: V): void;
  delete(key: string): void;
  size: number;
  keys(): IterableIterator<string>;
}

export interface StatsLike {
  cacheHits: number;
  cacheMisses: number;
}

export interface CreatePerformSearchWithCacheDeps {
  generateKey: (params: SearchParams) => string;
  cache: CacheLike<{ result: SearchResult; timestamp: number }>;
  stats: StatsLike;
  log: typeof logger;
  cacheTtlSec: number;
  searchFn: PerformSearchFn;
  maxCacheSize?: number; // default 100
}

export type PerformSearchWithCache = (params: SearchParams, userId: string) => Promise<SearchResult>;

export function createPerformSearchWithCache(deps: CreatePerformSearchWithCacheDeps): PerformSearchWithCache {
  const { generateKey, cache, stats, log, cacheTtlSec, searchFn, maxCacheSize = 100 } = deps;

  return async (searchParams: SearchParams, _userId: string): Promise<SearchResult> => {
    const cacheKey = generateKey(searchParams);

    // Cache lookup
    const cached = cache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < cacheTtlSec * 1000) {
      stats.cacheHits++;
      try {
        log.debug('search.cache.hit', { type: 'performance', component: 'SearchCommand', cacheKey });
      } catch {}
      return { ...cached.result, cacheHit: true };
    }

    stats.cacheMisses++;

    const searchResult = await searchFn(searchParams);

    cache.set(cacheKey, { result: searchResult, timestamp: Date.now() });

    if (cache.size > maxCacheSize) {
      const oldestKey = cache.keys().next().value as string | undefined;
      if (oldestKey) cache.delete(oldestKey);
    }

    return { ...searchResult, cacheHit: false };
  };
}
