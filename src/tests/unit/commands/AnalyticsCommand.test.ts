/**
 * Unit тесты для AnalyticsCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { AnalyticsCommand } from '../../../commands/AnalyticsCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('AnalyticsCommand', () => {
  let analyticsCommand: AnalyticsCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    analyticsCommand = new AnalyticsCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create AnalyticsCommand instance', () => {
      expect(analyticsCommand).toBeInstanceOf(AnalyticsCommand);
    });

    it('should have correct name', () => {
      expect(analyticsCommand.getName()).toBe('аналітика');
    });

    it('should have correct description', () => {
      expect(analyticsCommand.getDescription()).toBe('Аналітика та звіти про використання бота');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = analyticsCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('аналітика');
    });
  });

  describe('execute', () => {
    it('should handle report subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        generateReport: jest.fn().mockResolvedValue({
          type: 'daily',
          data: { users: 100, commands: 500 },
          timestamp: new Date(),
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('report');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('report', true);
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('format');
      expect(mockAnalyticsService.generateReport).toHaveBeenCalledWith('usage', 10, 'text');
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle statistics subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getStatistics: jest.fn().mockResolvedValue({
          totalUsers: 1000,
          totalCommands: 5000,
          popularCommands: ['пошук', 'ai_асистент'],
          dailyActiveUsers: 150,
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('statistics');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockAnalyticsService.getStatistics).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle trends subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getTrends: jest.fn().mockResolvedValue({
          period: '7d',
          trends: [
            { date: '2024-01-01', users: 100, commands: 500 },
            { date: '2024-01-02', users: 120, commands: 600 },
          ],
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('trends');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockAnalyticsService.getTrends).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle insights subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getInsights: jest.fn().mockResolvedValue({
          insights: [
            'Популярність команди /пошук зросла на 25%',
            'Середній час відповіді AI зменшився на 15%',
          ],
          recommendations: [
            'Додати більше фільтрів для пошуку',
            'Оптимізувати AI відповіді',
          ],
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('insights');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockAnalyticsService.getInsights).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('invalid');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.editReply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Invalid subcommand'),
        })
      );
    });

    it('should handle service error', async () => {
      // Настройка моков с ошибкой
      mockInteraction.client.serviceContainer.get.mockImplementation(() => {
        throw new Error('Analytics service error');
      });
      mockInteraction.options.getSubcommand.mockReturnValue('report');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.editReply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Error in analytics command'),
        })
      );
    });
  });
}); 