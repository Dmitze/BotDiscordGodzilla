import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { OllamaService } from '../OllamaService';
import type { BotConfig } from '@/types';
import { CacheService } from '../CacheService';

// Mock the logger
jest.mock('@/utils/logger', () => ({
  __esModule: true,
  default: {
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn(),
    log: jest.fn(),
    apiRequest: jest.fn(),
    apiError: jest.fn(),
    security: jest.fn(),
    performance: jest.fn(),
    system: jest.fn(),
    logStructured: jest.fn(),
    startStructuredTimer: jest.fn().mockReturnValue({ end: jest.fn() }),
    getStats: jest.fn(),
    getLogBuffer: jest.fn(),
    cleanup: jest.fn(),
    isHealthy: jest.fn(),
  },
}));

describe('OllamaService', () => {
  let ollamaService: OllamaService;
  let mockConfig: BotConfig;
  let mockCacheService: CacheService;

  beforeEach(() => {
    mockConfig = {
      ai: {
        ollama: {
          host: 'http://localhost:11434',
          model: 'llama3',
          ctx: 2048,
          chatMaxLength: 500,
        },
      },
      discord: {
        token: 'test-token',
        enableSlash: true,
        clientId: 'test-client-id',
      },
      google: {
        spreadsheetId: 'test-spreadsheet-id',
        sheetName: 'test-sheet-name',
      },
      redis: {
        host: 'localhost',
        port: 6379,
      },
      metrics: {
        enabled: true,
        port: 9090,
      },
      features: {
        enableUserWorkspace: true,
      },
      drive: {
        pageSize: 10,
      },
    } as unknown as BotConfig;

    mockCacheService = {
      get: jest.fn(),
      set: jest.fn(),
      delete: jest.fn(),
    } as unknown as CacheService;

    ollamaService = new OllamaService(mockConfig, mockCacheService);
  });

  describe('constructor', () => {
    it('should create an instance with default config', () => {
      expect(ollamaService).toBeInstanceOf(OllamaService);
    });

    it('should use default values when config is missing', () => {
      const configWithoutOllama = {
        ai: {
          ollama: {}
        }
      } as BotConfig;
      const service = new OllamaService(configWithoutOllama);
      
      // We can't directly access private properties, but we can test through methods
      expect(service).toBeInstanceOf(OllamaService);
    });
  });

  describe('getStats', () => {
    it('should return initial stats', () => {
      // Set the start time to a known value to test uptime calculation
      (ollamaService as any).startTime = Date.now() - 1000;
      
      const stats = ollamaService.getStats();
      expect(stats.service).toBe('OllamaService');
      expect(stats.requests).toBe(0);
      expect(stats.errors).toBe(0);
      expect(stats.avgResponseTime).toBe(0);
      // Uptime should be approximately 1000ms (±100ms for test timing)
      expect(stats.uptime).toBeGreaterThanOrEqual(900);
      expect(stats.uptime).toBeLessThanOrEqual(1100);
    });
  });

  describe('resetChannelHistory', () => {
    it('should call cache delete with correct key', async () => {
      const channelId = 'test-channel';
      await ollamaService.resetChannelHistory(channelId);
      
      expect(mockCacheService.delete).toHaveBeenCalledWith(`ollama:channel:${channelId}`);
    });

    it('should handle cache service being null', async () => {
      const service = new OllamaService(mockConfig, undefined);
      await expect(service.resetChannelHistory('test-channel')).resolves.toBeUndefined();
    });
  });

  describe('healthCheck', () => {
    it('should return healthy status when fetch succeeds', async () => {
      (global as any).fetch = jest.fn().mockImplementation(async () => {
        return Promise.resolve({
          ok: true,
        });
      });

      const result = await ollamaService.healthCheck();
      expect(result.healthy).toBe(true);
      expect(result.message).toBe('Ollama is available');
    });

    it('should return unhealthy status when fetch fails', async () => {
      (global as any).fetch = jest.fn().mockImplementation(async () => {
        return Promise.reject(new Error('Network error'));
      });

      const result = await ollamaService.healthCheck();
      expect(result.healthy).toBe(false);
    });
  });
});

export {};
