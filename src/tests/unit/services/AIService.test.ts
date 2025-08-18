/**
 * Unit тесты для AIService
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { AIService } from '../../../services/AIService';
import { createMockConfig } from '../../utils/testHelpers';

// Моки для OpenAI
jest.mock('openai', () => ({
  OpenAI: jest.fn().mockImplementation(() => ({
    chat: {
      completions: {
        create: jest.fn(),
      },
    },
  })),
}));

describe('AIService', () => {
  let aiService: AIService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    aiService = new AIService(mockConfig);
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
      await expect(aiService.initialize()).resolves.not.toThrow();
    });

    it('should handle initialization error', async () => {
      // Мокаем ошибку инициализации
      jest.spyOn(aiService as any, 'setupOpenAI').mockImplementation(() => {
        throw new Error('OpenAI setup error');
      });

      await expect(aiService.initialize()).rejects.toThrow('OpenAI setup error');
    });
  });

  describe('generateResponse', () => {
    beforeEach(async () => {
      await aiService.initialize();
    });

    it('should generate response successfully', async () => {
      const mockResponse: any = {
        choices: [
          {
            message: {
              content: 'AI generated response',
            },
          },
        ],
      };

      // Мокаем OpenAI API
      const mockOpenAI = {
        chat: {
          completions: {
            create: jest.fn().mockResolvedValue(mockResponse as any),
          },
        },
      };

      (aiService as any).openai = mockOpenAI;

      const result = await aiService.generateResponse('Hello, how are you?');

      expect(result).toBe('AI generated response');
      expect(mockOpenAI.chat.completions.create).toHaveBeenCalled();
    });

    it('should handle empty response', async () => {
      const mockResponse: any = {
        choices: [
          {
            message: {
              content: '',
            },
          },
        ],
      };

      const mockOpenAI = {
        chat: {
          completions: {
            create: jest.fn().mockResolvedValue(mockResponse),
          },
        },
      };

      (aiService as any).openai = mockOpenAI;

      const result = await aiService.generateResponse('test');

      expect(result).toBe('Відповідь не отримана');
    });

    it('should handle API error', async () => {
      const mockOpenAI = {
        chat: {
          completions: {
            create: jest.fn().mockRejectedValue(new Error('OpenAI API error') as any),
          },
        },
      };

      (aiService as any).openai = mockOpenAI;

      await expect(aiService.generateResponse('test')).rejects.toThrow('OpenAI API error');
    });

    it('should handle rate limit error', async () => {
      const mockOpenAI = {
        chat: {
          completions: {
            create: jest.fn().mockRejectedValue(new Error('Rate limit exceeded') as any),
          },
        },
      };

      (aiService as any).openai = mockOpenAI;

      await expect(aiService.generateResponse('test')).rejects.toThrow('Rate limit exceeded');
    });
  });

  describe('health check', () => {
    it('should return healthy status when initialized', async () => {
      await aiService.initialize();
      
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