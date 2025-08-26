/**
 * 💾 Response Cache Service
 * High-performance caching with TTL for AI responses and API calls
 */

import logger from '@/utils/logger';

export interface CacheEntry<T = any> {
  key: string;
  value: T;
  createdAt: Date;
  expiresAt: Date;
  hits: number;
  lastAccessed: Date;
  metadata?: {
    source?: string;
    size?: number;
    tags?: string[];
  };
}

export interface CacheStats {
  totalEntries: number;
  totalHits: number;
  totalMisses: number;
  hitRate: number;
  memoryUsage: {
    approximateSizeBytes: number;
    largestEntry: string | null;
    oldestEntry: Date | null;
  };
  expirationInfo: {
    expiredCount: number;
    nextExpiration: Date | null;
  };
}

export class ResponseCacheService {
  private static readonly DEFAULT_TTL_MINUTES = 30;
  private static readonly MAX_ENTRIES = 1000;
  private static readonly CLEANUP_INTERVAL_MINUTES = 5;

  private cache: Map<string, CacheEntry> = new Map();
  private cleanupInterval?: NodeJS.Timeout | undefined;
  private stats = {
    hits: 0,
    misses: 0
  };

  constructor(
    private readonly defaultTtlMinutes: number = ResponseCacheService.DEFAULT_TTL_MINUTES,
    private readonly maxEntries: number = ResponseCacheService.MAX_ENTRIES
  ) {
    this.startCleanupTask();
  }

  /**
   * 💾 Store value in cache
   */
  set<T>(
    key: string,
    value: T,
    ttlMinutes?: number,
    metadata?: CacheEntry['metadata']
  ): void {
    const now = new Date();
    const ttl = ttlMinutes ?? this.defaultTtlMinutes;
    const expiresAt = new Date(now.getTime() + ttl * 60 * 1000);

    // Check cache size limit
    if (this.cache.size >= this.maxEntries && !this.cache.has(key)) {
      this.evictOldestEntry();
    }

    const entry: CacheEntry<T> = {
      key,
      value,
      createdAt: now,
      expiresAt,
      hits: 0,
      lastAccessed: now
    };
    
    if (metadata) {
      entry.metadata = metadata;
    }

    this.cache.set(key, entry);

    logger.debug('Value cached', {
      component: 'ResponseCacheService',
      key,
      ttlMinutes: ttl,
      expiresAt: expiresAt.toISOString(),
      cacheSize: this.cache.size
    });
  }

  /**
   * 🔍 Get value from cache
   */
  get<T>(key: string): T | null {
    const entry = this.cache.get(key);
    
    if (!entry) {
      this.stats.misses++;
      logger.debug('Cache miss', {
        component: 'ResponseCacheService',
        key
      });
      return null;
    }

    // Check if expired
    if (new Date() > entry.expiresAt) {
      this.cache.delete(key);
      this.stats.misses++;
      logger.debug('Cache expired', {
        component: 'ResponseCacheService',
        key,
        expiredAt: entry.expiresAt.toISOString()
      });
      return null;
    }

    // Update access stats
    entry.hits++;
    entry.lastAccessed = new Date();
    this.stats.hits++;

    logger.debug('Cache hit', {
      component: 'ResponseCacheService',
      key,
      hits: entry.hits
    });

    return entry.value as T;
  }

  /**
   * ❓ Check if key exists and is valid
   */
  has(key: string): boolean {
    const entry = this.cache.get(key);
    if (!entry) {
      return false;
    }

    // Check if expired
    if (new Date() > entry.expiresAt) {
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * 🗑️ Delete specific key
   */
  delete(key: string): boolean {
    const deleted = this.cache.delete(key);
    
    if (deleted) {
      logger.debug('Cache entry deleted', {
        component: 'ResponseCacheService',
        key
      });
    }

    return deleted;
  }

  /**
   * 🧹 Clear all cache
   */
  clear(): void {
    const size = this.cache.size;
    this.cache.clear();
    this.stats.hits = 0;
    this.stats.misses = 0;

    logger.debug('Cache cleared', {
      component: 'ResponseCacheService',
      clearedEntries: size
    });
  }

  /**
   * 🔍 Search cache by pattern or tags
   */
  findByPattern(pattern: RegExp): CacheEntry[] {
    const results: CacheEntry[] = [];
    
    for (const entry of this.cache.values()) {
      if (pattern.test(entry.key)) {
        // Check if not expired
        if (new Date() <= entry.expiresAt) {
          results.push(entry);
        }
      }
    }

    return results;
  }

  /**
   * 🏷️ Find entries by tags
   */
  findByTags(tags: string[]): CacheEntry[] {
    const results: CacheEntry[] = [];
    
    for (const entry of this.cache.values()) {
      if (entry.metadata?.tags) {
        const hasTag = tags.some(tag => entry.metadata!.tags!.includes(tag));
        if (hasTag && new Date() <= entry.expiresAt) {
          results.push(entry);
        }
      }
    }

    return results;
  }

  /**
   * ⏰ Extend TTL for existing entry
   */
  extendTtl(key: string, additionalMinutes: number): boolean {
    const entry = this.cache.get(key);
    
    if (!entry || new Date() > entry.expiresAt) {
      return false;
    }

    entry.expiresAt = new Date(entry.expiresAt.getTime() + additionalMinutes * 60 * 1000);

    logger.debug('Cache TTL extended', {
      component: 'ResponseCacheService',
      key,
      newExpiresAt: entry.expiresAt.toISOString(),
      additionalMinutes
    });

    return true;
  }

  /**
   * 📊 Get cache statistics
   */
  getStats(): CacheStats {
    const now = new Date();
    let totalSize = 0;
    let largestEntry: string | null = null;
    let largestSize = 0;
    let oldestEntry: Date | null = null;
    let expiredCount = 0;
    let nextExpiration: Date | null = null;

    for (const [key, entry] of this.cache.entries()) {
      // Calculate approximate size
      const entrySize = JSON.stringify(entry.value).length;
      totalSize += entrySize;

      if (entrySize > largestSize) {
        largestSize = entrySize;
        largestEntry = key;
      }

      if (!oldestEntry || entry.createdAt < oldestEntry) {
        oldestEntry = entry.createdAt;
      }

      // Check expiration
      if (now > entry.expiresAt) {
        expiredCount++;
      } else {
        if (!nextExpiration || entry.expiresAt < nextExpiration) {
          nextExpiration = entry.expiresAt;
        }
      }
    }

    const totalRequests = this.stats.hits + this.stats.misses;
    const hitRate = totalRequests > 0 ? (this.stats.hits / totalRequests) * 100 : 0;

    return {
      totalEntries: this.cache.size,
      totalHits: this.stats.hits,
      totalMisses: this.stats.misses,
      hitRate: Math.round(hitRate * 100) / 100,
      memoryUsage: {
        approximateSizeBytes: totalSize,
        largestEntry,
        oldestEntry
      },
      expirationInfo: {
        expiredCount,
        nextExpiration
      }
    };
  }

  /**
   * 🔄 Start periodic cleanup task
   */
  private startCleanupTask(): void {
    this.cleanupInterval = setInterval(() => {
      this.cleanupExpiredEntries();
    }, ResponseCacheService.CLEANUP_INTERVAL_MINUTES * 60 * 1000);

    logger.debug('Cache cleanup task started', {
      component: 'ResponseCacheService',
      intervalMinutes: ResponseCacheService.CLEANUP_INTERVAL_MINUTES
    });
  }

  /**
   * 🧹 Clean up expired entries
   */
  private cleanupExpiredEntries(): void {
    const now = new Date();
    const expiredKeys: string[] = [];

    for (const [key, entry] of this.cache.entries()) {
      if (now > entry.expiresAt) {
        expiredKeys.push(key);
      }
    }

    for (const key of expiredKeys) {
      this.cache.delete(key);
    }

    if (expiredKeys.length > 0) {
      logger.debug('Expired cache entries cleaned', {
        component: 'ResponseCacheService',
        expiredCount: expiredKeys.length,
        remainingEntries: this.cache.size
      });
    }
  }

  /**
   * 🚮 Evict oldest entry when cache is full
   */
  private evictOldestEntry(): void {
    let oldestKey: string | null = null;
    let oldestTime: Date | null = null;

    for (const [key, entry] of this.cache.entries()) {
      if (!oldestTime || entry.lastAccessed < oldestTime) {
        oldestTime = entry.lastAccessed;
        oldestKey = key;
      }
    }

    if (oldestKey) {
      this.cache.delete(oldestKey);
      logger.debug('Oldest cache entry evicted', {
        component: 'ResponseCacheService',
        evictedKey: oldestKey,
        lastAccessed: oldestTime?.toISOString()
      });
    }
  }

  /**
   * 🔑 Generate cache key from components
   */
  static generateKey(...components: (string | number | boolean)[]): string {
    return components
      .map(c => String(c))
      .join(':')
      .replace(/[^a-zA-Z0-9:_-]/g, '_');
  }

  /**
   * 🛑 Shutdown service
   */
  shutdown(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = undefined;
    }

    this.clear();

    logger.debug('ResponseCacheService shutdown completed', {
      component: 'ResponseCacheService'
    });
  }
}