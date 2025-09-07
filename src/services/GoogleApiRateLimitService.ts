import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

export interface RateLimitInfo {
  limit: number;
  remaining: number;
  resetTime: number; // Unix timestamp
  window: number; // Window size in seconds
}

export interface RateLimitConfig {
  maxRetries: number;
  initialDelay: number; // in milliseconds
  maxDelay: number; // in milliseconds
  backoffMultiplier: number;
  jitter: boolean;
}

export interface ApiCallMetrics {
  endpoint: string;
  callCount: number;
  errorCount: number;
  averageResponseTime: number;
  lastCall: Date;
}

export interface RateLimitBucket {
  key: string;
  limit: number;
  remaining: number;
  resetTime: number;
  lastUpdated: Date;
}

export class GoogleApiRateLimitService extends BaseService {
  private rateLimits: Map<string, RateLimitInfo> = new Map();
  private metrics: Map<string, ApiCallMetrics> = new Map();
  private buckets: Map<string, RateLimitBucket> = new Map();
  private rateLimitConfig: RateLimitConfig;
  private readonly DEFAULT_RATE_LIMIT_WINDOW = 100; // 100 requests per 100 seconds
  
  constructor(config: BotConfig) {
    super('GoogleApiRateLimitService', config);
    
    // Extract rate limit config from the main config
    const rateLimitSettings = (config as any).rateLimit || {};
    
    this.rateLimitConfig = {
      maxRetries: rateLimitSettings.maxRetries || 3,
      initialDelay: rateLimitSettings.initialDelay || 1000,
      maxDelay: rateLimitSettings.maxDelay || 60000,
      backoffMultiplier: rateLimitSettings.backoffMultiplier || 2,
      jitter: rateLimitSettings.jitter ?? true
    };
  }

  /**
   * Check if we can make an API call without hitting rate limits
   */
  canMakeCall(endpoint: string): boolean {
    const rateLimit = this.rateLimits.get(endpoint);
    
    if (!rateLimit) {
      // No rate limit info yet, allow the call
      return true;
    }
    
    const now = Date.now();
    
    // If we're past the reset time, reset the counter
    if (now >= rateLimit.resetTime) {
      rateLimit.remaining = rateLimit.limit;
      rateLimit.resetTime = now + (rateLimit.window * 1000);
    }
    
    // Check if we have remaining calls
    return rateLimit.remaining > 0;
  }

  /**
   * Update rate limit information based on API response headers
   */
  updateRateLimit(endpoint: string, headers: Record<string, string>): void {
    try {
      // Parse rate limit headers from Google API responses
      const limit = parseInt(headers['x-ratelimit-limit'] || headers['ratelimit-limit'] || '100');
      const remaining = parseInt(headers['x-ratelimit-remaining'] || headers['ratelimit-remaining'] || '100');
      const reset = parseInt(headers['x-ratelimit-reset'] || headers['ratelimit-reset'] || '100');
      
      // If we couldn't parse the headers, use default values
      const rateLimit: RateLimitInfo = {
        limit: isNaN(limit) ? this.DEFAULT_RATE_LIMIT_WINDOW : limit,
        remaining: isNaN(remaining) ? this.DEFAULT_RATE_LIMIT_WINDOW : remaining,
        resetTime: isNaN(reset) ? Date.now() + 100000 : reset * 1000, // Convert to milliseconds
        window: 100 // Default window of 100 seconds
      };
      
      this.rateLimits.set(endpoint, rateLimit);
      
      logger.debug('Rate limit updated', {
        component: 'GoogleApiRateLimitService',
        endpoint,
        limit: rateLimit.limit,
        remaining: rateLimit.remaining,
        resetTime: new Date(rateLimit.resetTime).toISOString()
      });
    } catch (error) {
      logger.warn('Error updating rate limit info', {
        component: 'GoogleApiRateLimitService',
        endpoint,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Wait for rate limit to reset if necessary
   */
  async waitForRateLimit(endpoint: string): Promise<void> {
    const rateLimit = this.rateLimits.get(endpoint);
    
    if (!rateLimit) {
      // No rate limit info, no need to wait
      return;
    }
    
    const now = Date.now();
    
    // If we're past the reset time, reset the counter
    if (now >= rateLimit.resetTime) {
      rateLimit.remaining = rateLimit.limit;
      return;
    }
    
    // If we have remaining calls, no need to wait
    if (rateLimit.remaining > 0) {
      rateLimit.remaining--;
      return;
    }
    
    // Calculate wait time
    const waitTime = rateLimit.resetTime - now;
    
    if (waitTime > 0) {
      logger.info('Rate limit reached, waiting for reset', {
        component: 'GoogleApiRateLimitService',
        endpoint,
        waitTime: `${(waitTime / 1000).toFixed(1)}s`
      });
      
      // Wait for the reset time
      await new Promise(resolve => setTimeout(resolve, waitTime));
      
      // Reset the counter after waiting
      rateLimit.remaining = rateLimit.limit;
    }
  }

  /**
   * Execute an API call with rate limit handling and exponential backoff
   */
  async executeWithRateLimit<T>(
    endpoint: string,
    apiCall: () => Promise<T>
  ): Promise<T> {
    let lastError: Error | null = null;
    
    for (let attempt = 1; attempt <= this.rateLimitConfig.maxRetries; attempt++) {
      try {
        // Check rate limit before making the call
        if (!this.canMakeCall(endpoint)) {
          await this.waitForRateLimit(endpoint);
        }
        
        // Update metrics before the call
        this.updateMetricsBeforeCall(endpoint);
        
        // Make the API call
        const startTime = Date.now();
        const result = await apiCall();
        const endTime = Date.now();
        
        // Update metrics after successful call
        this.updateMetricsAfterCall(endpoint, endTime - startTime, true);
        
        // Update rate limit info (would come from actual response headers in real implementation)
        // For now, we'll just decrement the remaining count
        const rateLimit = this.rateLimits.get(endpoint);
        if (rateLimit && rateLimit.remaining > 0) {
          rateLimit.remaining--;
        }
        
        return result;
      } catch (error) {
        lastError = error as Error;
        
        // Update metrics after failed call
        this.updateMetricsAfterCall(endpoint, 0, false);
        
        // If this is the last attempt, throw the error
        if (attempt === this.rateLimitConfig.maxRetries) {
          logger.error('API call failed after max retries', {
            component: 'GoogleApiRateLimitService',
            endpoint,
            attempts: attempt,
            error: error instanceof Error ? error.message : String(error)
          });
          
          throw error;
        }
        
        // Calculate delay with exponential backoff
        const delay = this.calculateDelay(attempt);
        
        logger.warn('API call failed, retrying', {
          component: 'GoogleApiRateLimitService',
          endpoint,
          attempt,
          maxRetries: this.rateLimitConfig.maxRetries,
          delay: `${delay}ms`,
          error: error instanceof Error ? error.message : String(error)
        });
        
        // Wait before retrying
        await new Promise(resolve => setTimeout(resolve, delay));
      }
    }
    
    // This should never be reached, but just in case
    throw lastError || new Error('API call failed');
  }

  /**
   * Calculate delay with exponential backoff and optional jitter
   */
  private calculateDelay(attempt: number): number {
    let delay = this.rateLimitConfig.initialDelay * Math.pow(this.rateLimitConfig.backoffMultiplier, attempt - 1);
    
    // Cap the delay at the maximum
    delay = Math.min(delay, this.rateLimitConfig.maxDelay);
    
    // Add jitter if enabled
    if (this.rateLimitConfig.jitter) {
      delay = delay * (0.5 + Math.random() * 0.5); // 50-100% of calculated delay
    }
    
    return Math.round(delay);
  }

  /**
   * Update metrics before making an API call
   */
  private updateMetricsBeforeCall(endpoint: string): void {
    if (!this.metrics.has(endpoint)) {
      this.metrics.set(endpoint, {
        endpoint,
        callCount: 0,
        errorCount: 0,
        averageResponseTime: 0,
        lastCall: new Date()
      });
    }
    
    const metrics = this.metrics.get(endpoint)!;
    metrics.callCount++;
    metrics.lastCall = new Date();
  }

  /**
   * Update metrics after an API call
   */
  private updateMetricsAfterCall(endpoint: string, responseTime: number, success: boolean): void {
    const metrics = this.metrics.get(endpoint);
    
    if (!metrics) {
      return;
    }
    
    if (!success) {
      metrics.errorCount++;
      return;
    }
    
    // Update average response time
    if (metrics.callCount > 1) {
      metrics.averageResponseTime = 
        ((metrics.averageResponseTime * (metrics.callCount - 1)) + responseTime) / metrics.callCount;
    } else {
      metrics.averageResponseTime = responseTime;
    }
  }

  /**
   * Get current rate limit information for an endpoint
   */
  getRateLimitInfo(endpoint: string): RateLimitInfo | null {
    return this.rateLimits.get(endpoint) || null;
  }

  /**
   * Get API call metrics
   */
  getMetrics(endpoint?: string): ApiCallMetrics | Map<string, ApiCallMetrics> {
    if (endpoint) {
      return this.metrics.get(endpoint) || {
        endpoint,
        callCount: 0,
        errorCount: 0,
        averageResponseTime: 0,
        lastCall: new Date(0)
      };
    }
    
    return new Map(this.metrics);
  }

  /**
   * Reset rate limit information for an endpoint
   */
  resetRateLimit(endpoint: string): void {
    this.rateLimits.delete(endpoint);
    logger.info('Rate limit reset for endpoint', {
      component: 'GoogleApiRateLimitService',
      endpoint
    });
  }

  /**
   * Reset all rate limit information
   */
  resetAllRateLimits(): void {
    this.rateLimits.clear();
    logger.info('All rate limits reset', {
      component: 'GoogleApiRateLimitService'
    });
  }

  /**
   * Get service statistics
   */
  override getStats(): any {
    let totalCalls = 0;
    let totalErrors = 0;
    let totalResponseTime = 0;
    let responseTimeCount = 0;
    
    for (const metrics of this.metrics.values()) {
      totalCalls += metrics.callCount;
      totalErrors += metrics.errorCount;
      
      if (metrics.averageResponseTime > 0) {
        totalResponseTime += metrics.averageResponseTime;
        responseTimeCount++;
      }
    }
    
    const averageResponseTime = responseTimeCount > 0 
      ? totalResponseTime / responseTimeCount 
      : 0;
    
    return {
      trackedEndpoints: this.rateLimits.size,
      totalCalls,
      totalErrors,
      averageResponseTime
    };
  }

  /**
   * Add a custom rate limit bucket
   */
  addRateLimitBucket(bucket: RateLimitBucket): void {
    this.buckets.set(bucket.key, bucket);
    logger.info('Rate limit bucket added', {
      component: 'GoogleApiRateLimitService',
      bucketKey: bucket.key,
      limit: bucket.limit
    });
  }

  /**
   * Check if a bucket has available capacity
   */
  canUseBucket(bucketKey: string): boolean {
    const bucket = this.buckets.get(bucketKey);
    
    if (!bucket) {
      // No bucket found, allow the call
      return true;
    }
    
    const now = Date.now();
    
    // If we're past the reset time, reset the counter
    if (now >= bucket.resetTime) {
      bucket.remaining = bucket.limit;
      bucket.lastUpdated = new Date();
    }
    
    // Check if we have remaining capacity
    return bucket.remaining > 0;
  }

  /**
   * Use capacity from a bucket
   */
  useBucketCapacity(bucketKey: string, amount: number = 1): boolean {
    const bucket = this.buckets.get(bucketKey);
    
    if (!bucket) {
      // No bucket found, allow the call
      return true;
    }
    
    const now = Date.now();
    
    // If we're past the reset time, reset the counter
    if (now >= bucket.resetTime) {
      bucket.remaining = bucket.limit;
      bucket.lastUpdated = new Date();
    }
    
    // Check if we have enough remaining capacity
    if (bucket.remaining >= amount) {
      bucket.remaining -= amount;
      bucket.lastUpdated = new Date();
      return true;
    }
    
    return false;
  }

  /**
   * Get all rate limit buckets
   */
  getBuckets(): Map<string, RateLimitBucket> {
    return new Map(this.buckets);
  }

  /**
   * Clear expired buckets (not updated in the last hour)
   */
  clearExpiredBuckets(): void {
    const now = Date.now();
    const oneHour = 60 * 60 * 1000; // 1 hour in milliseconds
    
    for (const [key, bucket] of this.buckets.entries()) {
      if (now - bucket.lastUpdated.getTime() > oneHour) {
        this.buckets.delete(key);
      }
    }
    
    logger.info('Expired rate limit buckets cleared', {
      component: 'GoogleApiRateLimitService',
      remainingBuckets: this.buckets.size
    });
  }

  /**
   * Generate a rate limit report
   */
  generateReport(): {
    rateLimits: Map<string, RateLimitInfo>;
    metrics: Map<string, ApiCallMetrics>;
    buckets: Map<string, RateLimitBucket>;
    stats: {
      trackedEndpoints: number;
      totalCalls: number;
      totalErrors: number;
      errorRate: number;
      averageResponseTime: number;
    };
  } {
    const stats = this.getStats();
    const errorRate = stats.totalCalls > 0 
      ? (stats.totalErrors / stats.totalCalls) * 100 
      : 0;
    
    return {
      rateLimits: new Map(this.rateLimits),
      metrics: new Map(this.metrics),
      buckets: new Map(this.buckets),
      stats: {
        ...stats,
        errorRate
      }
    };
  }

  // === BaseService required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('GoogleApiRateLimitService initialized', {
      component: 'GoogleApiRateLimitService'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('GoogleApiRateLimitService shutdown', {
      component: 'GoogleApiRateLimitService'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      trackedEndpoints: this.rateLimits.size
    };
  }

  protected onGetStats(): any {
    return this.getStats();
  }
}
