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
      // Настройка моков
      const mockMetricsService = {
        getStats: (jest.fn() as any).mockResolvedValue({
          uptime: 3600,
          requests: 100,
          errors: 5,
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockMetricsService);
      mockInteraction.options.getSubcommand.mockReturnValue('статус');

      // Выполнение
      await performanceCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockMetricsService.getStats).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle cache subcommand', async () => {
      // Настройка моков
      const mockCacheService = {
        getStats: (jest.fn() as any).mockResolvedValue({
          hits: 80,
          misses: 20,
          size: 1024,
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockCacheService);
      mockInteraction.options.getSubcommand.mockReturnValue('кеш');

      // Выполнение
      await performanceCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockCacheService.getStats).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });
  });
}); 