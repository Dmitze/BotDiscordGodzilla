/**
 * Unit тести для AIService (оновлено під актуальний API)
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { AIService } from '../../../services/AIService';
import { createMockConfig } from '../../utils/testHelpers';

describe('AIService', () => {
  let aiService: AIService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    aiService = new AIService(mockConfig);
    // Підміняємо CacheService на легкий мок без зовнішніх ресурсів
    (aiService as any).cacheService = {
      initialize: jest.fn(async () => {}),
      get: jest.fn(async () => undefined),
      set: jest.fn(async () => {}),
      cleanup: jest.fn(async () => {}),
    };
  });

  afterEach(async () => {
    // Акуратно завершуємо сервіс, якщо він позначений як ініціалізований
    if ((aiService as any)._initialized) {
      // Уникаємо реальної зупинки залежностей
      jest.spyOn(aiService as any, 'onShutdown').mockResolvedValue(undefined);
      try {
        await (aiService as any).shutdown();
      } catch {
        // ignore
      }
    }
    jest.clearAllTimers();
  });

  describe('constructor', () => {
    it('should create AIService instance', () => {
      expect(aiService).toBeInstanceOf(AIService);
    });

    it('should have correct service name', () => {
      expect(aiService.getName()).toBe('AIService');
    });
  });

  describe('initialization', () => {
    it('should initialize successfully', async () => {
      // Спрощено: підміняємо приватні хелпери щоб не чіпати зовнішні сервіси
      jest.spyOn(aiService as any, 'createProviders').mockResolvedValue(undefined);
      jest.spyOn(aiService as any, 'validateConfiguration').mockImplementation(() => {});
      jest.spyOn(aiService as any, 'startMemoryCleanup').mockImplementation(() => {});
      jest.spyOn(aiService as any, 'startHealthCheck').mockImplementation(() => {});
      await expect(aiService.initialize()).resolves.not.toThrow();
    });

    it('should handle initialization error', async () => {
      // Емулюємо помилку під час createProviders
      jest.spyOn(aiService as any, 'createProviders').mockRejectedValue(
        new Error('provider create error')
      );
      // Відключаємо ретраї базового сервісу, щоб уникнути таймерів/відкритих хендлів
      (aiService as any).retryCount = 3;
      await expect(aiService.initialize()).rejects.toThrow('provider create error');
    });
  });

  describe('generateResponse', () => {
    beforeEach(async () => {
      await aiService.initialize();
    });

    it('should generate response successfully', async () => {
      const mockProvider = {
        generate: jest.fn(async () => ({
          content: 'AI generated response',
          provider: 'openai',
          model: 'gpt-test',
          tokens: 10,
          duration: 5,
        })),
        isHealthy: jest.fn(async () => true),
      };
      (aiService as any).providers = { openai: mockProvider };

      const result = await aiService.generateResponse('Hello, how are you?', {
        useCache: false,
        retryAttempts: 0,
      });

      expect(result.content).toBe('AI generated response');
      expect(mockProvider.generate).toHaveBeenCalled();
    });

    it('should handle empty response', async () => {
      const mockProvider = {
        generate: jest.fn(async () => ({
          content: '',
          provider: 'openai',
          model: 'gpt-test',
          tokens: 0,
          duration: 1,
        })),
        isHealthy: jest.fn(async () => true),
      };
      (aiService as any).providers = { openai: mockProvider };

      const result = await aiService.generateResponse('test', {
        useCache: false,
        retryAttempts: 0,
      });
      expect(result.content).toBe('');
    });

    it('should handle API error', async () => {
      const mockProvider = {
        generate: jest.fn(async () => {
          throw new Error('OpenAI API error');
        }),
        isHealthy: jest.fn(async () => true),
      };
      (aiService as any).providers = { openai: mockProvider };

      await expect(
        aiService.generateResponse('test', { useCache: false, retryAttempts: 0 })
      ).rejects.toThrow('OpenAI API error');
    });

    it('should handle rate limit error', async () => {
      const mockProvider = {
        generate: jest.fn(async () => {
          throw new Error('Rate limit exceeded');
        }),
        isHealthy: jest.fn(async () => true),
      };
      (aiService as any).providers = { openai: mockProvider };

      await expect(
        aiService.generateResponse('test', { useCache: false, retryAttempts: 0 })
      ).rejects.toThrow('Rate limit exceeded');
    });
  });

  describe('health check', () => {
    it('should return healthy status when initialized', async () => {
      // Імітуємо ініціалізацію без справжніх залежностей
      (aiService as any)._initialized = true;
      jest
        .spyOn(aiService as any, 'onHealthCheck')
        .mockResolvedValue({ healthy: true });
      const health = await aiService.getHealthStatus();
      
      expect(health.healthy).toBe(true);
      expect(health.service).toBe('AIService');
    });

    it('should return unhealthy status when not initialized', async () => {
      const health = await aiService.getHealthStatus();
      
      expect(health.healthy).toBe(false);
      expect(health.service).toBe('AIService');
    });
  });

  describe('configuration', () => {
    it('should use OpenAI provider by default', () => {
      expect(aiService.getProvider()).toBe('openai');
    });

    it('should handle different providers', () => {
      const ollamaConfig = {
        ...mockConfig,
        ai: {
          ...mockConfig.ai,
          provider: 'ollama',
        },
      };

      const ollamaService = new AIService(ollamaConfig);
      expect(ollamaService.getProvider()).toBe('ollama');
    });
  });
}); 