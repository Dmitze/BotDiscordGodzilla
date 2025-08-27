import { createClient, type RedisClientType } from 'redis';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

export interface CacheOptions {
  ttl?: number; // Time to live in seconds
  compress?: boolean;
  tags?: string[]; // For tag-based invalidation
}

export interface CacheStats {
  hits: number;
  misses: number;
  evictions: number;
  memoryUsage: number;
  keys: number;
}

export class EnhancedCacheService {
  private redisClient: RedisClientType | null = null;
  private localCache: Map<string, { value: any; expiry: number; tags: string[] }> = new Map();
  private stats: CacheStats = {
    hits: 0,
    misses: 0,
    evictions: 0,
    memoryUsage: 0,
    keys: 0,
  };
  private readonly LOCAL_CACHE_MAX_SIZE = 1000;
  private readonly DEFAULT_TTL = 300; // 5 minutes

  constructor(private config: BotConfig) {
    this.initializeRedis();
  }

  /**
   * Initialize Redis connection
   */
  private async initializeRedis(): Promise<void> {
    try {
      if (this.config.redis?.enabled) {
        this.redisClient = createClient({
          host: this.config.redis.host,
          port: this.config.redis.port,
          password: this.config.redis.password,
          database: this.config.redis.database,
        }) as RedisClientType;

        await this.redisClient.connect();
        logger.info('Redis cache initialized successfully', {
          component: 'EnhancedCacheService',
        });
      }
    } catch (error) {
      logger.error('Failed to initialize Redis cache', {
        component: 'EnhancedCacheService',
        error: error instanceof Error ? error.message : String(error),
      });
      this.redisClient = null;
    }
  }

  /**
   * Get value from cache
   */
  async get<T>(key: string): Promise<T | null> {
    // Try local cache first
    const localEntry = this.localCache.get(key);
    if (localEntry) {
      if (Date.now() < localEntry.expiry) {
        this.stats.hits++;
        return localEntry.value as T;
      } else {
        // Expired entry
        this.localCache.delete(key);
        this.stats.evictions++;
      }
    }

    // Try Redis if available
    if (this.redisClient) {
      try {
        const value = await this.redisClient.get(key);
        if (value !== null) {
          this.stats.hits++;
          return JSON.parse(value) as T;
        }
      } catch (error) {
        logger.warn('Redis cache get failed, falling back to local cache', {
          component: 'EnhancedCacheService',
          key,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    // Cache miss
    this.stats.misses++;
    return null;
  }

  /**
   * Set value in cache
   */
  async set<T>(key: string, value: T, options: CacheOptions = {}): Promise<void> {
    const ttl = options.ttl || this.DEFAULT_TTL;
    const expiry = Date.now() + ttl * 1000;
    const tags = options.tags || [];

    // Store in local cache
    this.manageLocalCacheSize();
    this.localCache.set(key, { value, expiry, tags });

    // Store in Redis if available
    if (this.redisClient) {
      try {
        const serializedValue = JSON.stringify(value);
        await this.redisClient.setEx(key, ttl, serializedValue);
        
        // Store tags for invalidation
        if (tags.length > 0) {
          const tagKey = `tags:${key}`;
          await this.redisClient.sAdd(tagKey, ...tags);
          // Set expiration for tag key
          await this.redisClient.expire(tagKey, ttl);
        }
      } catch (error) {
        logger.warn('Redis cache set failed', {
          component: 'EnhancedCacheService',
          key,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
  }

  /**
   * Delete value from cache
   */
  async delete(key: string): Promise<boolean> {
    let deleted = false;

    // Delete from local cache
    if (this.localCache.has(key)) {
      this.localCache.delete(key);
      deleted = true;
    }

    // Delete from Redis if available
    if (this.redisClient) {
      try {
        const result = await this.redisClient.del(key);
        if (result > 0) {
          deleted = true;
        }
        
        // Delete associated tags
        const tagKey = `tags:${key}`;
        await this.redisClient.del(tagKey);
      } catch (error) {
        logger.warn('Redis cache delete failed', {
          component: 'EnhancedCacheService',
          key,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    return deleted;
  }

  /**
   * Invalidate cache entries by tag
   */
  async invalidateByTag(tag: string): Promise<number> {
    let invalidatedCount = 0;

    // Invalidate local cache entries with this tag
    for (const [key, entry] of this.localCache.entries()) {
      if (entry.tags.includes(tag)) {
        this.localCache.delete(key);
        invalidatedCount++;
      }
    }

    // Invalidate Redis entries with this tag
    if (this.redisClient) {
      try {
        // Find all keys with this tag
        const tagPattern = `tags:*`;
        const tagKeys = await this.redisClient.keys(tagPattern);
        
        for (const tagKey of tagKeys) {
          if (await this.redisClient.sIsMember(tagKey, tag)) {
            // Extract the actual key from tag key (tags:actual_key)
            const actualKey = tagKey.substring(5);
            await this.redisClient.del(actualKey);
            await this.redisClient.del(tagKey);
            invalidatedCount++;
          }
        }
      } catch (error) {
        logger.warn('Redis cache tag invalidation failed', {
          component: 'EnhancedCacheService',
          tag,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    return invalidatedCount;
  }

  /**
   * Clear all cache entries
   */
  async clear(): Promise<void> {
    // Clear local cache
    this.localCache.clear();

    // Clear Redis cache if available
    if (this.redisClient) {
      try {
        await this.redisClient.flushDb();
      } catch (error) {
        logger.warn('Redis cache clear failed', {
          component: 'EnhancedCacheService',
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    // Reset stats
    this.stats = {
      hits: 0,
      misses: 0,
      evictions: 0,
      memoryUsage: 0,
      keys: 0,
    };
  }

  /**
   * Get cache statistics
   */
  getStats(): CacheStats {
    return { ...this.stats, keys: this.localCache.size };
  }

  /**
   * Manage local cache size to prevent memory issues
   */
  private manageLocalCacheSize(): void {
    if (this.localCache.size >= this.LOCAL_CACHE_MAX_SIZE) {
      // Remove oldest entries
      const keys = Array.from(this.localCache.keys());
      const keysToRemove = keys.slice(0, Math.floor(this.LOCAL_CACHE_MAX_SIZE * 0.1)); // Remove 10%
      
      for (const key of keysToRemove) {
        this.localCache.delete(key);
        this.stats.evictions++;
      }
    }
  }

  /**
   * Close Redis connection
   */
  async shutdown(): Promise<void> {
    if (this.redisClient) {
      try {
        await this.redisClient.quit();
        logger.info('Redis cache connection closed', {
          component: 'EnhancedCacheService',
        });
      } catch (error) {
        logger.warn('Error closing Redis cache connection', {
          component: 'EnhancedCacheService',
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
  }

  /**
   * Check if cache is healthy
   */
  async isHealthy(): Promise<boolean> {
    if (!this.config.redis?.enabled) {
      return true; // Local cache is always healthy
    }

    if (!this.redisClient) {
      return false;
    }

    try {
      await this.redisClient.ping();
      return true;
    } catch (error) {
      logger.error('Redis cache health check failed', {
        component: 'EnhancedCacheService',
        error: error instanceof Error ? error.message : String(error),
      });
      return false;
    }
  }
}