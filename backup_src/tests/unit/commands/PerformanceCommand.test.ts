/**
 * Unit тесты для PerformanceCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { PerformanceCommand } from '../../../commands/PerformanceCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('PerformanceCommand', () => {
  let performanceCommand: PerformanceCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    performanceCommand = new PerformanceCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create PerformanceCommand instance', () => {
      expect(performanceCommand).toBeInstanceOf(PerformanceCommand);
    });

    it('should have correct name', () => {
      expect(performanceCommand.getName()).toBe('продуктивність');
    });

    it('should have correct description', () => {
      expect(performanceCommand.getDescription()).toBe('📊 Моніторинг продуктивності системи');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = performanceCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('продуктивність');
    });
  });

  describe('execute', () => {
    it('should handle status subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('статус');

      // Выполнение
      await performanceCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle cache subcommand', async () => {
      // Настройка моков
      const mockCacheService = {
        getCacheStats: jest.fn().mockReturnValue({
          hits: 80,
          misses: 20,
          sets: 10,
          deletes: 2,
          errors: 0,
        }),
      } as any;
      // В реализации команда обращается к interaction.client.bot.serviceContainer
      (mockInteraction.client as any).bot = {
        serviceContainer: {
          get: jest.fn().mockReturnValue(mockCacheService),
        },
      };
      mockInteraction.options.getSubcommand.mockReturnValue('кеш');

      // Выполнение
      await performanceCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockCacheService.getCacheStats).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });
  });
}); 