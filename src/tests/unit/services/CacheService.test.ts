/**
 * Unit тесты для CacheService
 */

import { jest, describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import { CacheService } from '../../../services/CacheService';
import { createMockConfig } from '../../utils/testHelpers';

// Моки для Redis
jest.mock('redis', () => ({
  createClient: jest.fn(() => ({
    connect: jest.fn(),
    disconnect: jest.fn(),
    get: jest.fn(),
    set: jest.fn(),
    del: jest.fn(),
    exists: jest.fn(),
    keys: jest.fn(),
    flushDb: jest.fn(),
    ping: jest.fn(),
  })),
}));

describe('CacheService', () => {
  let cacheService: CacheService;
  let mockConfig: any;
  let mockRedisClient: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    cacheService = new CacheService(mockConfig);
    
    // Мокаем Redis клиент
    mockRedisClient = {
      connect: jest.fn(),
      disconnect: jest.fn(),
      get: jest.fn(),
      set: jest.fn(),
      del: jest.fn(),
      exists: jest.fn(),
      keys: jest.fn(),
      flushDb: jest.fn(),
      ping: jest.fn(),
    };
    
    (cacheService as any).client = mockRedisClient;
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('constructor', () => {
    it('should create CacheService instance', () => {
      expect(cacheService).toBeInstanceOf(CacheService);
    });

    it('should have correct service name', () => {
      expect(cacheService.getName()).toBe('CacheService');
    });
  });

  describe('initialization', () => {
    it('should initialize successfully when Redis is enabled', async () => {
      mockConfig.cache.enabled = true;
      mockRedisClient.connect.mockResolvedValue(undefined);
      mockRedisClient.ping.mockResolvedValue('PONG');

      await expect(cacheService.initialize()).resolves.not.toThrow();
      expect(mockRedisClient.connect).toHaveBeenCalled();
    });

    it('should skip initialization when Redis is disabled', async () => {
      mockConfig.cache.enabled = false;

      await expect(cacheService.initialize()).resolves.not.toThrow();
      expect(mockRedisClient.connect).not.toHaveBeenCalled();
    });

    it('should handle connection error', async () => {
      mockConfig.cache.enabled = true;
      mockRedisClient.connect.mockRejectedValue(new Error('Connection failed'));

      await expect(cacheService.initialize()).rejects.toThrow('Connection failed');
    });
  });

  describe('cache operations', () => {
    beforeEach(async () => {
      mockConfig.cache.enabled = true;
      await cacheService.initialize();
    });

    it('should set cache value', async () => {
      const key = 'test_key';
      const value = { data: 'test_value' };
      const ttl = 3600;

      mockRedisClient.set.mockResolvedValue('OK');

      await cacheService.set(key, value, ttl);

      expect(mockRedisClient.set).toHaveBeenCalledWith(key, JSON.stringify(value), {
        EX: ttl,
      });
    });

    it('should get cache value', async () => {
      const key = 'test_key';
      const cachedValue = JSON.stringify({ data: 'test_value' });

      mockRedisClient.get.mockResolvedValue(cachedValue);

      const result = await cacheService.get(key);

      expect(mockRedisClient.get).toHaveBeenCalledWith(key);
      expect(result).toEqual({ data: 'test_value' });
    });

    it('should return null for non-existent key', async () => {
      const key = 'non_existent_key';

      mockRedisClient.get.mockResolvedValue(null);

      const result = await cacheService.get(key);

      expect(result).toBeNull();
    });

    it('should delete cache value', async () => {
      const key = 'test_key';

      mockRedisClient.del.mockResolvedValue(1);

      await cacheService.delete(key);

      expect(mockRedisClient.del).toHaveBeenCalledWith(key);
    });

    it('should check if key exists', async () => {
      const key = 'test_key';

      mockRedisClient.exists.mockResolvedValue(1);

      const result = await cacheService.exists(key);

      expect(mockRedisClient.exists).toHaveBeenCalledWith(key);
      expect(result).toBe(true);
    });

    it('should return false for non-existent key', async () => {
      const key = 'non_existent_key';

      mockRedisClient.exists.mockResolvedValue(0);

      const result = await cacheService.exists(key);

      expect(result).toBe(false);
    });
  });

  describe('cache statistics', () => {
    beforeEach(async () => {
      mockConfig.cache.enabled = true;
      await cacheService.initialize();
    });

    it('should get cache statistics', async () => {
      const mockKeys = ['key1', 'key2', 'key3'];
      mockRedisClient.keys.mockResolvedValue(mockKeys);

      const stats = await cacheService.getStats();

      expect(mockRedisClient.keys).toHaveBeenCalledWith('*');
      expect(stats).toEqual({
        hits: 0,
        misses: 0,
        size: 3,
        hitRate: 0,
      });
    });

    it('should calculate hit rate correctly', async () => {
      // Симулируем hits и misses
      (cacheService as any).hits = 80;
      (cacheService as any).misses = 20;

      const mockKeys = ['key1'];
      mockRedisClient.keys.mockResolvedValue(mockKeys);

      const stats = await cacheService.getStats();

      expect(stats.hitRate).toBe(0.8); // 80 / (80 + 20) = 0.8
    });

    it('should handle zero requests', async () => {
      const mockKeys = ['key1'];
      mockRedisClient.keys.mockResolvedValue(mockKeys);

      const stats = await cacheService.getStats();

      expect(stats.hitRate).toBe(0);
    });
  });

  describe('cache management', () => {
    beforeEach(async () => {
      mockConfig.cache.enabled = true;
      await cacheService.initialize();
    });

    it('should clear all cache', async () => {
      mockRedisClient.flushDb.mockResolvedValue('OK');

      await cacheService.clear();

      expect(mockRedisClient.flushDb).toHaveBeenCalled();
    });

    it('should get cache size', async () => {
      const mockKeys = ['key1', 'key2', 'key3', 'key4'];
      mockRedisClient.keys.mockResolvedValue(mockKeys);

      const size = await cacheService.getSize();

      expect(mockRedisClient.keys).toHaveBeenCalledWith('*');
      expect(size).toBe(4);
    });
  });

  describe('error handling', () => {
    beforeEach(async () => {
      mockConfig.cache.enabled = true;
      await cacheService.initialize();
    });

    it('should handle Redis get error', async () => {
      const key = 'test_key';
      mockRedisClient.get.mockRejectedValue(new Error('Redis error'));

      await expect(cacheService.get(key)).rejects.toThrow('Redis error');
    });

    it('should handle Redis set error', async () => {
      const key = 'test_key';
      const value = { data: 'test' };
      mockRedisClient.set.mockRejectedValue(new Error('Redis error'));

      await expect(cacheService.set(key, value)).rejects.toThrow('Redis error');
    });

    it('should handle JSON parse error', async () => {
      const key = 'test_key';
      const invalidJson = 'invalid json';

      mockRedisClient.get.mockResolvedValue(invalidJson);

      const result = await cacheService.get(key);

      expect(result).toBeNull();
    });
  });

  describe('health check', () => {
    it('should return healthy status when Redis is connected', async () => {
      mockConfig.cache.enabled = true;
      mockRedisClient.ping.mockResolvedValue('PONG');
      await cacheService.initialize();

      const health = await cacheService.getHealthStatus();

      expect(health.healthy).toBe(true);
      expect(health.service).toBe('CacheService');
    });

    it('should return unhealthy status when Redis is not connected', async () => {
      mockConfig.cache.enabled = true;
      mockRedisClient.ping.mockRejectedValue(new Error('Connection failed'));

      const health = await cacheService.getHealthStatus();

      expect(health.healthy).toBe(false);
      expect(health.service).toBe('CacheService');
    });

    it('should return healthy status when Redis is disabled', async () => {
      mockConfig.cache.enabled = false;

      const health = await cacheService.getHealthStatus();

      expect(health.healthy).toBe(true);
      expect(health.service).toBe('CacheService');
    });
  });

  describe('shutdown', () => {
    it('should disconnect Redis client on shutdown', async () => {
      mockConfig.cache.enabled = true;
      await cacheService.initialize();

      await cacheService.shutdown();

      expect(mockRedisClient.disconnect).toHaveBeenCalled();
    });

    it('should handle shutdown when Redis is disabled', async () => {
      mockConfig.cache.enabled = false;

      await expect(cacheService.shutdown()).resolves.not.toThrow();
      expect(mockRedisClient.disconnect).not.toHaveBeenCalled();
    });
  });
}); 