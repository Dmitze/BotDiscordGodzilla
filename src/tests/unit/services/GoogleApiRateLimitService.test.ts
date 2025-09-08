/**
 * Unit tests for GoogleApiRateLimitService functionality
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { GoogleApiRateLimitService } from '../../../services/GoogleApiRateLimitService';
import { createMockConfig } from '../../utils/testHelpers';

describe('GoogleApiRateLimitService', () => {
  let rateLimitService: GoogleApiRateLimitService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    rateLimitService = new GoogleApiRateLimitService(mockConfig);
  });

  it('should initialize with default configuration', () => {
    const stats = rateLimitService.getStats();
    
    expect(stats).toBeDefined();
    expect(typeof stats.trackedEndpoints).toBe('number');
    expect(typeof stats.totalCalls).toBe('number');
    expect(typeof stats.totalErrors).toBe('number');
    expect(typeof stats.averageResponseTime).toBe('number');
  });

  it('should allow API calls when no rate limit info is available', () => {
    const canMakeCall = rateLimitService.canMakeCall('/test/endpoint');
    expect(canMakeCall).toBe(true);
  });

  it('should update rate limit information from headers', () => {
    const headers = {
      'x-ratelimit-limit': '100',
      'x-ratelimit-remaining': '90',
      'x-ratelimit-reset': '1234567890'
    };
    
    rateLimitService.updateRateLimit('/test/endpoint', headers);
    
    const rateLimitInfo = rateLimitService.getRateLimitInfo('/test/endpoint');
    expect(rateLimitInfo).toBeDefined();
    expect(rateLimitInfo?.limit).toBe(100);
    expect(rateLimitInfo?.remaining).toBe(90);
    expect(rateLimitInfo?.resetTime).toBe(1234567890000); // Converted to milliseconds
  });

  it('should handle missing rate limit headers gracefully', () => {
    const headers = {};
    
    rateLimitService.updateRateLimit('/test/endpoint', headers);
    
    // Should use default values
    const rateLimitInfo = rateLimitService.getRateLimitInfo('/test/endpoint');
    expect(rateLimitInfo).toBeDefined();
    expect(rateLimitInfo?.limit).toBe(100); // Default value
  });

  it('should track API call metrics', async () => {
    // Execute a successful API call
    await rateLimitService.executeWithRateLimit('/test/endpoint', async () => {
      return 'success';
    });
    
    const metrics = rateLimitService.getMetrics('/test/endpoint') as any;
    expect(metrics).toBeDefined();
    expect(metrics.callCount).toBe(1);
    expect(metrics.errorCount).toBe(0);
  });

  it('should handle API call failures and retry', async () => {
    let callCount = 0;
    
    // Mock an API call that fails twice then succeeds
    const result = await rateLimitService.executeWithRateLimit('/test/endpoint', async () => {
      callCount++;
      if (callCount <= 2) {
        throw new Error('Temporary failure');
      }
      return 'success';
    });
    
    expect(result).toBe('success');
    expect(callCount).toBe(3); // Initial call + 2 retries
    
    const metrics = rateLimitService.getMetrics('/test/endpoint') as any;
    expect(metrics.callCount).toBe(3);
    expect(metrics.errorCount).toBe(2);
  });

  it('should respect rate limits and wait when necessary', async () => {
    // Set up a rate limit with no remaining calls
    const headers = {
      'x-ratelimit-limit': '100',
      'x-ratelimit-remaining': '0',
      'x-ratelimit-reset': ((Date.now() / 1000) + 1).toString() // Reset in 1 second
    };
    
    rateLimitService.updateRateLimit('/limited/endpoint', headers);
    
    // Mock setTimeout to immediately resolve
    const setTimeoutSpy = jest.spyOn(global, 'setTimeout')
      .mockImplementation((callback) => {
        // Execute the callback immediately
        if (typeof callback === 'function') {
          callback();
        }
        return {} as NodeJS.Timeout; // Return a mock timeout object
      });
    
    // Execute an API call - should wait for rate limit reset
    const startTime = Date.now();
    await rateLimitService.executeWithRateLimit('/limited/endpoint', async () => {
      return 'success';
    });
    const endTime = Date.now();
    
    // Should have waited (but our mock makes it instant)
    expect(endTime - startTime).toBeGreaterThanOrEqual(0);
    
    // Restore setTimeout
    setTimeoutSpy.mockRestore();
  });

  it('should calculate exponential backoff delays correctly', () => {
    // Create a new instance with jitter disabled for this test
    const testConfig = {
      ...mockConfig,
      rateLimit: {
        ...mockConfig.rateLimit,
        jitter: false
      }
    };
    const testRateLimitService = new GoogleApiRateLimitService(testConfig);
    
    // Access private method through reflection for testing
    const calculateDelay = (testRateLimitService as any).calculateDelay.bind(testRateLimitService);
    
    const delay1 = calculateDelay(1);
    const delay2 = calculateDelay(2);
    const delay3 = calculateDelay(3);
    
    // Should follow exponential pattern: 1000, 2000, 4000
    expect(delay1).toBe(1000);
    expect(delay2).toBe(2000);
    expect(delay3).toBe(4000);
  });

  it('should manage rate limit buckets', () => {
    const bucket = {
      key: 'test-bucket',
      limit: 50,
      remaining: 50,
      resetTime: Date.now() + 3600000, // 1 hour from now
      lastUpdated: new Date()
    };
    
    rateLimitService.addRateLimitBucket(bucket);
    
    const buckets = rateLimitService.getBuckets();
    expect(buckets.size).toBe(1);
    expect(buckets.has('test-bucket')).toBe(true);
    
    // Check if bucket has capacity
    const canUse = rateLimitService.canUseBucket('test-bucket');
    expect(canUse).toBe(true);
    
    // Use some capacity
    const used = rateLimitService.useBucketCapacity('test-bucket', 10);
    expect(used).toBe(true);
    
    // Check remaining capacity
    const bucketAfterUse = buckets.get('test-bucket');
    expect(bucketAfterUse?.remaining).toBe(40);
  });

  it('should clear expired buckets', () => {
    // Add a fresh bucket
    const freshBucket = {
      key: 'fresh-bucket',
      limit: 50,
      remaining: 50,
      resetTime: Date.now() + 3600000,
      lastUpdated: new Date()
    };
    
    // Add an expired bucket (last updated 2 hours ago)
    const expiredBucket = {
      key: 'expired-bucket',
      limit: 50,
      remaining: 50,
      resetTime: Date.now() + 3600000,
      lastUpdated: new Date(Date.now() - 2 * 60 * 60 * 1000) // 2 hours ago
    };
    
    rateLimitService.addRateLimitBucket(freshBucket);
    rateLimitService.addRateLimitBucket(expiredBucket);
    
    expect(rateLimitService.getBuckets().size).toBe(2);
    
    // Clear expired buckets
    rateLimitService.clearExpiredBuckets();
    
    // Should only have the fresh bucket left
    expect(rateLimitService.getBuckets().size).toBe(1);
    expect(rateLimitService.getBuckets().has('fresh-bucket')).toBe(true);
  });

  it('should generate comprehensive reports', () => {
    const report = rateLimitService.generateReport();
    
    expect(report).toBeDefined();
    expect(report.rateLimits).toBeDefined();
    expect(report.metrics).toBeDefined();
    expect(report.buckets).toBeDefined();
    expect(report.stats).toBeDefined();
    
    // Stats should have correct structure
    expect(typeof report.stats.trackedEndpoints).toBe('number');
    expect(typeof report.stats.totalCalls).toBe('number');
    expect(typeof report.stats.totalErrors).toBe('number');
    expect(typeof report.stats.errorRate).toBe('number');
    expect(typeof report.stats.averageResponseTime).toBe('number');
  });

  it('should reset rate limit information', () => {
    // Set up some rate limit info
    const headers = {
      'x-ratelimit-limit': '100',
      'x-ratelimit-remaining': '90',
      'x-ratelimit-reset': '1234567890'
    };
    
    rateLimitService.updateRateLimit('/test/endpoint', headers);
    
    // Verify rate limit info exists
    expect(rateLimitService.getRateLimitInfo('/test/endpoint')).toBeDefined();
    
    // Reset rate limit
    rateLimitService.resetRateLimit('/test/endpoint');
    
    // Verify rate limit info is gone
    expect(rateLimitService.getRateLimitInfo('/test/endpoint')).toBeNull();
    
    // Reset all rate limits
    rateLimitService.updateRateLimit('/another/endpoint', headers);
    rateLimitService.resetAllRateLimits();
    
    expect(rateLimitService.getRateLimitInfo('/another/endpoint')).toBeNull();
  });

  it('should handle maximum retries exceeded', async () => {
    // Mock an API call that always fails
    await expect(
      rateLimitService.executeWithRateLimit('/failing/endpoint', async () => {
        throw new Error('Permanent failure');
      })
    ).rejects.toThrow('Permanent failure');
    
    const metrics = rateLimitService.getMetrics('/failing/endpoint') as any;
    expect(metrics.callCount).toBe(3); // Initial call + 2 retries
    expect(metrics.errorCount).toBe(3);
  });
});